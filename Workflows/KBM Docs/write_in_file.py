from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

# Matches the version suffix in a filename, e.g. " v.17.0"
VERSION_PATTERN = re.compile(r"\s+v\.(\d+)\.(\d+)", re.IGNORECASE)

# Matches a DIA identifier in a filename, e.g. "DIA EF-732L.1"
DIA_ID_PATTERN = re.compile(r"\bDIA\s+([A-Z]{2}-[A-Z0-9]+(?:\.\d+)?)")

def get_dia_id(filename: str) -> str | None:
    """Returns the DIA identifier found in the filename, or None."""
    match = DIA_ID_PATTERN.search(filename)
    return f"DIA {match.group(1)}" if match else None


def get_version(filename: str) -> str | None:
    """Returns the version string (e.g. '17.0') found in the filename, or None."""
    match = None
    for m in VERSION_PATTERN.finditer(filename):
        match = m  # keep the last match in case the pattern appears more than once

    if match is None:
        return None
    return f"{match.group(1)}.{match.group(2)}"


def get_version_tuple(filename: str) -> tuple[str, tuple[int, int]] | None:
    """
    Returns (group_key, (major, minor)) for version-comparison purposes,
    where group_key is the filename text before the version suffix.
    Used to identify which files are superseded by a newer version.
    """
    match = None
    for m in VERSION_PATTERN.finditer(filename):
        match = m

    if match is None:
        return None

    group_key = filename[:match.start()].strip().lower()
    version   = (int(match.group(1)), int(match.group(2)))
    return group_key, version

def split_current_and_outdated(files: list[Path]) -> tuple[list[Path], list[Path]]:
    """
    Separates files into (current, outdated).
    A file is outdated if another file in the list shares the same
    DIA identifier + title but has a higher version number.
    """
    parsed = {f: get_version_tuple(f.name) for f in files}

    # Find the highest version seen for each group key
    latest: dict[str, tuple[int, int]] = {}
    for info in parsed.values():
        if info is None:
            continue
        key, version = info
        if key not in latest or version > latest[key]:
            latest[key] = version

    current:  list[Path] = []
    outdated: list[Path] = []
    for f in files:
        info = parsed[f]
        if info is not None and info[1] < latest[info[0]]:
            outdated.append(f)
        else:
            current.append(f)

    return current, outdated

def load_propel_map() -> dict[str, str]:
    """Loads DIA identifier -> Propel # from kbmmap.json (next to this script)."""
    path = Path(__file__).resolve().parent / "kbmmap.json"
    if not path.exists():
        return {}
    with open(path, encoding="utf-8") as f:
        return json.load(f)

def run(folder: Path) -> None:
    # Collect and filter .url files
    all_files = sorted(folder.glob("*.url"))
    if not all_files:
        print(f"[!] No .url files found in: {folder.resolve()}")
        sys.exit(1)

    current_files, outdated_files = split_current_and_outdated(all_files)

    if outdated_files:
        print(f"[i] Skipping {len(outdated_files)} outdated file(s):")
        for f in outdated_files:
            print(f"    - {f.name}")
        print()

    if not current_files:
        print(f"[!] No current .url files found in: {folder.resolve()}")
        sys.exit(1)

    print(f"Found {len(current_files)} .url file(s) in '{folder}'\n")

    propel_map = load_propel_map()

    # Column width for the DIA Identifier column
    dia_ids  = [get_dia_id(f.name) for f in current_files]
    col_width = max((len(d) for d in dia_ids if d), default=0)
    col_width = max(col_width, len("DIA Identifier")) + 2

    # Header
    print(f"{'DIA Identifier':<{col_width}}  {'Version':<12}  {'Date':<16}  {'Propel #':<12}  Notes")
    print("-" * (col_width + 50))

    # Rows
    for f, dia_id in zip(current_files, dia_ids):
        version = get_version(f.name)
        propel  = propel_map.get(dia_id) if dia_id else None
        notes   = "Could not extract DIA identifier from filename." if dia_id is None else ""

        print(
            f"{dia_id or '—':<{col_width}}  "
            f"{version or '—':<12}  "
            f"{'—':<16}  "
            f"{propel or '—':<12}  "
            f"{notes}"
        )

    print(f"\nProcessed {len(current_files)} file(s), skipped {len(outdated_files)} outdated file(s).")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Extract DIA Identifier, Version, and Propel # from KeborMed .url shortcuts."
    )
    parser.add_argument(
        "--folder",
        default="Records",
        help="Folder containing .url shortcuts (default: 'Records')",
    )
    args = parser.parse_args()

    folder = Path(args.folder)
    if not folder.exists() or not folder.is_dir():
        print(f"[!] Not a valid folder: {folder.resolve()}")
        sys.exit(1)

    run(folder)
