"""SharePoint Excel file operations via Microsoft Graph API."""

import json
import logging
import os
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Optional

import requests
from dotenv import load_dotenv
from openpyxl import load_workbook

from config import (
    AUTH_ENDPOINT,
    GRAPH_API_BASE,
    LOG_FORMAT,
    LOG_LEVEL,
    SHAREPOINT_DRIVE_ID,
    SHAREPOINT_SITE_ID,
    TOKEN_CACHE_FILE,
    TOKEN_EXPIRY_BUFFER,
)

load_dotenv()

logger = logging.getLogger(__name__)
logging.basicConfig(level=LOG_LEVEL, format=LOG_FORMAT)


class SharePointAuthError(Exception):
    pass


class SharePointOperationError(Exception):
    pass


class TokenCache:
    """OAuth token cache with expiration."""

    def __init__(self, cache_file: str = TOKEN_CACHE_FILE):
        self.cache_file = cache_file

    def save(self, token: str, expires_in: int) -> None:
        expiry = datetime.now() + timedelta(seconds=expires_in)
        try:
            with open(self.cache_file, "w") as f:
                json.dump({"token": token, "expiry": expiry.isoformat()}, f)
        except Exception as e:
            logger.error(f"Failed to save token cache: {e}")

    def load(self) -> Optional[str]:
        try:
            if not Path(self.cache_file).exists():
                return None
            with open(self.cache_file, "r") as f:
                cache_data = json.load(f)
            expiry = datetime.fromisoformat(cache_data["expiry"])
            if datetime.now() < expiry - timedelta(seconds=TOKEN_EXPIRY_BUFFER):
                return cache_data["token"]
            return None
        except Exception:
            return None

    def clear(self) -> None:
        try:
            p = Path(self.cache_file)
            if p.exists():
                p.unlink()
        except Exception:
            pass


class SharePointHelper:
    """Read/write Excel files on SharePoint via Microsoft Graph API."""

    def __init__(self, tenant_id=None, client_id=None, client_secret=None,
                 site_id=SHAREPOINT_SITE_ID, drive_id=SHAREPOINT_DRIVE_ID):
        self.tenant_id = tenant_id or os.getenv("TENANT_ID")
        self.client_id = client_id or os.getenv("CLIENT_ID")
        self.client_secret = client_secret or os.getenv("CLIENT_SECRET")
        self.site_id = site_id
        self.drive_id = drive_id

        if not all([self.tenant_id, self.client_id, self.client_secret]):
            raise SharePointAuthError(
                "Missing credentials. Set TENANT_ID, CLIENT_ID, CLIENT_SECRET in .env."
            )

        self.token_cache = TokenCache()
        self.access_token: Optional[str] = None

    def _get_access_token(self) -> str:
        cached = self.token_cache.load()
        if cached:
            self.access_token = cached
            return cached

        url = AUTH_ENDPOINT.format(tenant_id=self.tenant_id)
        payload = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }

        try:
            resp = requests.post(url, data=payload,
                                 headers={"Content-Type": "application/x-www-form-urlencoded"},
                                 timeout=10)
            resp.raise_for_status()
            data = resp.json()
            token = data.get("access_token")
            if not token:
                raise SharePointAuthError("No access token in response")
            self.token_cache.save(token, data.get("expires_in", 3600))
            self.access_token = token
            return token
        except requests.RequestException as e:
            raise SharePointAuthError(f"Authentication failed: {e}")

    def _get_headers(self) -> dict:
        if not self.access_token:
            self.access_token = self._get_access_token()
        return {"Authorization": f"Bearer {self.access_token}", "Content-Type": "application/json"}

    def _file_url(self, file_path: str) -> str:
        return f"{GRAPH_API_BASE}/sites/{self.site_id}/drives/{self.drive_id}/root:/{file_path}:/content"

    # ── Internal helpers ──

    def _drive_base(self) -> str:
        return f"{GRAPH_API_BASE}/sites/{self.site_id}/drives/{self.drive_id}"

    def get_item_id(self, file_path: str) -> str:
        """Resolve a file path to its driveItem id."""
        url = f"{self._drive_base()}/root:/{file_path}"
        try:
            resp = requests.get(url, headers=self._get_headers(), timeout=15)
            resp.raise_for_status()
            return resp.json()["id"]
        except requests.RequestException as e:
            raise SharePointOperationError(f"Could not resolve item id for {file_path}: {e}")

    # ── File I/O ──

    def read_excel(self, file_path: str) -> bytes:
        """Download file bytes from SharePoint."""
        try:
            resp = requests.get(self._file_url(file_path), headers=self._get_headers(), timeout=30)
            resp.raise_for_status()
            return resp.content
        except requests.RequestException as e:
            raise SharePointOperationError(f"Read failed for {file_path}: {e}")

    def write_excel(self, file_path: str, file_content: bytes) -> None:
        """Upload file bytes to SharePoint via a simple PUT.

        Used for archive copies written to fresh (unlocked) paths. The live
        sprint file is updated incrementally through the Workbook API instead
        (see the workbook session methods below), because a co-authored file
        cannot be overwritten wholesale.
        """
        url = f"{self._drive_base()}/root:/{file_path}:/content"
        headers = self._get_headers()
        headers["Content-Type"] = "application/octet-stream"
        try:
            resp = requests.put(url, data=file_content, headers=headers, timeout=60)
            resp.raise_for_status()
        except requests.RequestException as e:
            raise SharePointOperationError(f"Write failed for {file_path}: {e}")

    # ── Workbook API (incremental, co-authoring safe) ──

    def open_workbook_session(self, item_id: str, persist: bool = True) -> str:
        """Open a workbook session and return its id."""
        url = f"{self._drive_base()}/items/{item_id}/workbook/createSession"
        try:
            resp = requests.post(url, json={"persistChanges": persist},
                                 headers=self._get_headers(), timeout=30)
            resp.raise_for_status()
            return resp.json()["id"]
        except requests.RequestException as e:
            raise SharePointOperationError(f"Failed to open workbook session: {e}")

    def close_workbook_session(self, item_id: str, session_id: str) -> None:
        url = f"{self._drive_base()}/items/{item_id}/workbook/closeSession"
        try:
            requests.post(url, headers=self._wb_headers(session_id), timeout=30)
        except requests.RequestException as e:
            logger.warning(f"Failed to close workbook session: {e}")

    def _wb_headers(self, session_id: str) -> dict:
        headers = self._get_headers()
        headers["workbook-session-id"] = session_id
        return headers

    def workbook_request(self, method: str, item_id: str, session_id: str,
                         rel_url: str, json_body: dict = None) -> dict:
        """Issue a Workbook API call relative to .../workbook/. Returns parsed JSON ({} if empty)."""
        url = f"{self._drive_base()}/items/{item_id}/workbook/{rel_url.lstrip('/')}"
        try:
            resp = requests.request(method, url, json=json_body,
                                    headers=self._wb_headers(session_id), timeout=60)
            resp.raise_for_status()
            return resp.json() if resp.content else {}
        except requests.HTTPError as e:
            body = ""
            try:
                body = e.response.json().get("error", {}).get("message", "")
            except Exception:
                pass
            raise SharePointOperationError(
                f"Workbook {method} {rel_url} failed: {e}" +
                (f" | {body}" if body else ""))
        except requests.RequestException as e:
            raise SharePointOperationError(f"Workbook {method} {rel_url} failed: {e}")

    def archive_file(self, file_path: str, archive_folder: str = None) -> str:
        """Create a timestamped copy of a file on SharePoint."""
        try:
            file_bytes = self.read_excel(file_path)
            stem = Path(file_path).stem
            ext = Path(file_path).suffix
            date_suffix = datetime.now().strftime("%m%d%Y")

            if archive_folder:
                dest = f"{archive_folder}/{stem}_{date_suffix}{ext}"
            else:
                parent = str(Path(file_path).parent)
                dest = f"{stem}_{date_suffix}{ext}" if parent == "." else f"{parent}/{stem}_{date_suffix}{ext}"

            self.write_excel(dest, file_bytes)
            return dest
        except Exception as e:
            raise SharePointOperationError(f"Archive failed: {e}")

    # ── Convenience wrappers ──

    def load_workbook(self, file_path: str, keep_vba: bool = None):
        """Load an openpyxl Workbook from SharePoint."""
        if keep_vba is None:
            keep_vba = file_path.lower().endswith('.xlsm')
        wb_bytes = self.read_excel(file_path)
        return load_workbook(BytesIO(wb_bytes), keep_vba=keep_vba)

    def save_workbook(self, file_path: str, workbook) -> None:
        """Save an openpyxl Workbook back to SharePoint."""
        buf = BytesIO()
        workbook.save(buf)
        self.write_excel(file_path, buf.getvalue())

    def read_excel_to_dataframe(self, file_path: str):
        """Read SharePoint Excel file into a pandas DataFrame."""
        try:
            import pandas as pd
        except ImportError:
            raise SharePointOperationError("pandas not installed")
        content = self.read_excel(file_path)
        return pd.read_excel(BytesIO(content))

    def write_dataframe_to_excel(self, df, file_path: str) -> None:
        """Write a pandas DataFrame as Excel to SharePoint."""
        try:
            import pandas as pd
        except ImportError:
            raise SharePointOperationError("pandas not installed")
        buf = BytesIO()
        df.to_excel(buf, index=False, engine="openpyxl")
        self.write_excel(file_path, buf.getvalue())

    def clear_token_cache(self) -> None:
        self.token_cache.clear()
        self.access_token = None
"""
TO DO:
- upload file
- download
- rename
- make dir_and_upload

"""
