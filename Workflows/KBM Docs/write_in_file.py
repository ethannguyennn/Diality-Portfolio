from __future__ import annotations

import argparse
import configparser
import io
import json
import re
import sys
from pathlib import Path

import msal
import pdfplumber
import requests

CLIENT_ID = "YOUR_CLIENT_ID_HERE"  # Azure AD App Registration (Application) ID
TENANT_ID = "YOUR_TENANT_ID_HERE"  # Azure AD Directory (Tenant) ID
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SHAREPOINT_SCOPE = "https://kbmazurekebormed.sharepoint.com/sites/KebormedQMS/.default"  # replace 'mycompany' with your SharePoint domain
TOKEN_CACHE_PATH = Path(__file__).resolve().parent / "token_cache.bin"

_VERSION_RE = re.compile(r"\s+v\.(\d+)\.(\d+)", re.IGNORECASE)


def _parse_version_key(filename: str) -> tuple[str, tuple[int, int]] | None:
    """
    Splits a filename like
      "DIA EF-732L.1 Trace Matrix Cloud Extensions Project Whitney v.17.0.url"
    into a (group_key, version) pair, where group_key is everything before
    " v.<major>.<minor>" (the DIA identifier + title) and version is
    (major, minor) as ints.

    Returns None if no "v.<major>.<minor>" pattern is found, meaning the
    file can't be compared against other versions.
    """
    match = None
    for m in _VERSION_RE.finditer(filename):
        match = m  # use the last match in case "v." appears earlier too

    if match is None:
        return None

    group_key = filename[:match.start()].strip().lower()
    version = (int(match.group(1)), int(match.group(2)))
    return group_key, version


def _filter_outdated(files: list[Path]) -> tuple[list[Path], list[Path]]:
    """
    Splits files into (current, outdated), where "outdated" files are
    superseded by another file with the same DIA identifier + title but a
    higher "v.<major>.<minor>" version number.
    """
    parsed = {p: _parse_version_key(p.name) for p in files}

    latest_version: dict[str, tuple[int, int]] = {}
    for info in parsed.values():
        if info is None:
            continue
        key, version = info
        if key not in latest_version or version > latest_version[key]:
            latest_version[key] = version

    current: list[Path] = []
    outdated: list[Path] = []
    for p in files:
        info = parsed[p]
        if info is not None and info[1] < latest_version[info[0]]:
            outdated.append(p)
        else:
            current.append(p)

    return current, outdated


_DIA_ID_RE = re.compile(r"\bDIA\s+([A-Z]{2}-[A-Z0-9]+(?:\.\d+)?)")
_PDF_VERSION_RE = re.compile(r"Version:\s*([\d]+\.[\d]+)", re.IGNORECASE)
_PDF_DATE_RE = re.compile(r"(?:Effective Date|Date of Version):\s*([\d]{1,2}/[\d]{1,2}/[\d]{2,4})", re.IGNORECASE)


def _load_propel_map() -> dict[str, str]:
    """
    Loads the DIA-identifier -> Propel # mapping from kbmmap.json, located
    next to this script. Returns an empty dict if the file is missing.
    """
    map_path = Path(__file__).resolve().parent / "kbmmap.json"
    if not map_path.exists():
        return {}
    with open(map_path, "r", encoding="utf-8") as f:
        return json.load(f)


def _read_url_file(url_path: Path) -> str | None:
    """
    Reads the URL= value from the [InternetShortcut] section of a Windows
    Internet Shortcut (.url) file.

    Returns the raw (still percent-encoded) URL string, or None if it
    couldn't be found/read. Interpolation is disabled because SharePoint
    URLs contain '%' characters that ConfigParser would otherwise try to
    interpret.
    """
    parser = configparser.ConfigParser(interpolation=None)
    try:
        parser.read(url_path, encoding="utf-8")
        return parser.get("InternetShortcut", "URL", fallback=None)
    except configparser.Error:
        return None


def _get_access_token() -> str:
    """
    Acquires an access token for SHAREPOINT_SCOPE using MSAL's device-code
    flow, reusing a cached token (and refresh token) from TOKEN_CACHE_PATH
    when possible so the user isn't prompted on every run.
    """
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        cache.deserialize(TOKEN_CACHE_PATH.read_text(encoding="utf-8"))

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent([SHAREPOINT_SCOPE], account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=[SHAREPOINT_SCOPE])
        if "user_code" not in flow:
            raise RuntimeError(f"Failed to start device flow: {flow.get('error_description', flow)}")
        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    if cache.has_state_changed:
        TOKEN_CACHE_PATH.write_text(cache.serialize(), encoding="utf-8")

    if "access_token" not in result:
        raise RuntimeError(f"Could not acquire access token: {result.get('error_description', result)}")

    return result["access_token"]


def _fetch_pdf_bytes(url: str, token: str) -> io.BytesIO | None:
    """
    Fetches `url` with the given bearer token, streaming the response body
    into an in-memory BytesIO. Returns None if the response isn't actually a
    PDF (e.g. an HTML login redirect was returned instead, indicating auth
    failed).
    """
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers, stream=True, timeout=60)
    resp.raise_for_status()

    buffer = io.BytesIO()
    for chunk in resp.iter_content(chunk_size=8192):
        buffer.write(chunk)

    if not buffer.getvalue().startswith(b"%PDF"):
        return None

    buffer.seek(0)
    return buffer


def _extract_pdf_pages(pdf_bytes: io.BytesIO) -> list[str]:
    """Returns a list of extracted text, one entry per page."""
    with pdfplumber.open(pdf_bytes) as pdf:
        return [page.extract_text() or "" for page in pdf.pages]


def extract_pdf_metadata(pages_text: list[str], propel_map: dict[str, str]) -> dict:
    """
    Returns a dict with keys: dia_id, version, date, propel, error.
    'error' is None on success, or a message string if any field couldn't
    be found in the PDF text.

    The cover page (first page) is checked first, falling back to the full
    document text if a field isn't found there.
    """
    first_page = pages_text[0] if pages_text else ""
    full_text = "\n".join(pages_text)

    dia_match = _DIA_ID_RE.search(first_page) or _DIA_ID_RE.search(full_text)
    dia_id = f"DIA {dia_match.group(1)}" if dia_match else None

    version_match = _PDF_VERSION_RE.search(first_page) or _PDF_VERSION_RE.search(full_text)
    version = version_match.group(1) if version_match else None

    date_match = _PDF_DATE_RE.search(first_page) or _PDF_DATE_RE.search(full_text)
    date = date_match.group(1) if date_match else None

    result = {
        "dia_id": dia_id,
        "version": version,
        "date": date,
        "propel": propel_map.get(dia_id) if dia_id else None,
        "error": None,
    }

    missing = [
        label for label, value in (
            ("DIA identifier", dia_id),
            ("version", version),
            ("date", date),
        ) if value is None
    ]
    if missing:
        result["error"] = "Could not find " + ", ".join(missing) + " in PDF text."

    return result


def _empty_result(error: str) -> dict:
    return {"dia_id": None, "version": None, "date": None, "propel": None, "error": error}


def process_url_file(url_path: Path, token: str, propel_map: dict[str, str]) -> dict:
    """
    Runs the parse -> fetch -> validate -> extract -> process pipeline for a
    single .url file. Any failure along the way produces an all-'-' result
    row carrying an explanatory error message.
    """
    url = _read_url_file(url_path)
    if url is None:
        return _empty_result(f"Could not read URL from {url_path.name}.")

    try:
        pdf_bytes = _fetch_pdf_bytes(url, token)
    except requests.RequestException as exc:
        return _empty_result(f"PDF fetch failed: {exc}")

    if pdf_bytes is None:
        return _empty_result("Fetch did not return a PDF (auth failure or login redirect).")

    try:
        pages_text = _extract_pdf_pages(pdf_bytes)
    except Exception as exc:
        return _empty_result(f"Failed to read PDF: {exc}")

    return extract_pdf_metadata(pages_text, propel_map)


def run(folder: Path) -> list[dict]:
    url_files = sorted(folder.glob("*.url"))

    if not url_files:
        print(f"[!] No .url files found in: {folder.resolve()}")
        sys.exit(1)

    url_files, outdated_files = _filter_outdated(url_files)

    if outdated_files:
        print(f"[i] Skipping {len(outdated_files)} outdated file(s):")
        for p in outdated_files:
            print(f"    - {p.name}")
        print()

    if not url_files:
        print(f"[!] No current (non-outdated) .url files found in: {folder.resolve()}")
        sys.exit(1)

    print(f"Found {len(url_files)} .url file(s) in '{folder}'\n")

    token = _get_access_token()
    propel_map = _load_propel_map()

    results = []
    for url_path in url_files:
        print(f"Processing: {url_path.name}")
        results.append(process_url_file(url_path, token, propel_map))

    id_w = max([len(r["dia_id"]) for r in results if r["dia_id"]] + [len("DIA Identifier")]) + 2

    print()
    print(f"{'DIA Identifier':<{id_w}}  {'Version':<12}  {'Date':<16}  {'Propel #':<12}  Notes")
    print("-" * (id_w + 50))

    error_count = 0
    for meta in results:
        dia_id = meta["dia_id"] or "—"
        version = meta["version"] or "—"
        date_val = meta["date"] or "—"
        propel = meta["propel"] or "—"
        notes = meta["error"] or ""
        if meta["error"]:
            error_count += 1

        print(f"{dia_id:<{id_w}}  {version:<12}  {date_val:<16}  {propel:<12}  {notes}")

    print(
        f"\nProcessed {len(url_files)} file(s) "
        f"({error_count} with errors), skipped {len(outdated_files)} outdated file(s)."
    )

    return results


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Extract DIA Identifier + Version + Date + Propel # from KeborMed SharePoint PDFs via .url shortcuts."
    )
    parser.add_argument(
        "--folder",
        default="Records",
        help="Path to the folder containing .url shortcuts (default: 'Records')",
    )
    args = parser.parse_args()

    folder = Path(args.folder)
    if not folder.exists():
        print(f"[!] Folder not found: {folder.resolve()}")
        sys.exit(1)
    if not folder.is_dir():
        print(f"[!] Not a directory: {folder.resolve()}")
        sys.exit(1)

    run(folder)
