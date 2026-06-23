"""
SharePoint Helper Module - Complete Excel Workbook Management via Microsoft Graph API

OVERVIEW
========
This module provides a comprehensive API for reading, writing, and manipulating Excel files
on SharePoint. It exposes openpyxl functionality through helper methods, enabling team members
to run complex Excel automation (like update_sprint_management.py) directly against SharePoint
without local files.

KEY FEATURES
============
✓ OAuth2 authentication with automatic token caching
✓ Load/save Excel workbooks (XLSX and XLSM with VBA preservation)
✓ Full cell styling support (fonts, colors, borders, hyperlinks, comments, data validation)
✓ Row operations with complete state preservation (snapshot/restore)
✓ Excel table XML management for complex spreadsheets
✓ SharePoint file operations (archiving, listing, existence checks)
✓ Comprehensive logging and error handling
✓ Complete openpyxl API exposure for advanced use cases

QUICK START
===========
    from sharepoint_helper import SharePointHelper

    # Initialize
    helper = SharePointHelper()

    # Load workbook from SharePoint
    wb = helper.load_workbook("General/17_SW Sprint/sprint testing.xlsx")

    # Access and modify
    ws = helper.get_sheet(wb, "Sprint Template")
    helper.set_cell(ws, 4, 1, "Value")
    helper.set_freeze_panes(ws, "A4")

    # Save back
    helper.save_workbook("General/17_SW Sprint/sprint testing.xlsx", wb)

PHASES
======

PHASE 1: WORKBOOK I/O
    Methods: load_workbook, save_workbook
    Purpose: Load Excel files from SharePoint into memory, save workbooks back
    Usage:
        wb = helper.load_workbook("path/to/file.xlsx", keep_vba=True)
        helper.save_workbook("path/to/file.xlsx", wb)

PHASE 2: SHEET OPERATIONS
    Methods: get_sheet, get_sheet_names, sheet_exists, get_max_row, get_max_column,
             set_freeze_panes, set_row_height, set_column_width
    Purpose: Access sheets, configure layout, get metadata
    Usage:
        ws = helper.get_sheet(wb, "Sheet1")
        helper.set_freeze_panes(ws, "A4")
        helper.set_column_width(ws, "A", 9.29)
        max_row = helper.get_max_row(ws)

PHASE 3: CELL OPERATIONS
    Methods: get_cell, set_cell, set_cell_styled, get_cell_style, set_cell_style
    Purpose: Read/write cell values with complete formatting (font, fill, border,
             alignment, protection, hyperlink, comment)
    Usage:
        helper.set_cell(ws, 4, 1, "Value")
        helper.set_cell_styled(ws, 4, 8, "[LDT-123] Title",
                              font=Font(color="1155CC", underline="single"),
                              hyperlink="https://jira.com/browse/LDT-123")
        style = helper.get_cell_style(ws, 1, 1)
        helper.set_cell_style(ws, 2, 1, style)

PHASE 4: ROW OPERATIONS
    Methods: delete_rows, insert_rows, snapshot_row, restore_row, clear_row
    Purpose: Manipulate rows while preserving all formatting and state
    Usage:
        # Capture row state before modification
        snapshot = helper.snapshot_row(ws, 10, num_cols=12)

        # Do operations...

        # Restore row exactly as it was
        helper.restore_row(ws, 20, snapshot)

        # Insert/delete rows
        helper.insert_rows(ws, 10, count=3)
        helper.delete_rows(ws, 5, count=2)

        # Clear row
        helper.clear_row(ws, 10, start_col=1, end_col=9)

PHASE 5: DATA VALIDATION
    Methods: get_data_validations, reset_data_validations, add_data_validation
    Purpose: Manage dropdown lists and validation rules
    Usage:
        from openpyxl.worksheet.datavalidation import DataValidation

        helper.reset_data_validations(ws)

        dv = DataValidation(type="list", formula1="Selection!$F$2:$F$101",
                           allow_blank=True)
        dv.add("A4:A1000")
        helper.add_data_validation(ws, dv)

PHASE 6: TABLE OPERATIONS
    Methods: get_tables, capture_table_xml, patch_table_xml
    Purpose: Manage Excel table formatting and structure
    Usage:
        # Before modification, capture original table XML
        path, xml = helper.capture_table_xml(wb_bytes)

        # ... modify spreadsheet ...

        # After save, patch table to maintain Excel compatibility
        col_names = ["Sprint", "Team", "Assignee", ...]
        wb_bytes = helper.patch_table_xml(wb_bytes, final_row=100,
                                          col_names=col_names, orig_table_xml=xml)

PHASE 7: SHAREPOINT FILE OPERATIONS
    Methods: file_exists, archive_file, list_files
    Purpose: Manage files on SharePoint (archiving, listing, checks)
    Usage:
        # Check if file exists
        if helper.file_exists("General/17_SW Sprint/sprint testing.xlsx"):
            wb = helper.load_workbook("General/17_SW Sprint/sprint testing.xlsx")

        # Create timestamped backup
        archive = helper.archive_file("General/17_SW Sprint/sprint testing.xlsx",
                                     archive_folder="General/17_SW Sprint/archive")
        # Creates: .../archive/sprint testing_06232026.xlsx

        # List files in folder
        files = helper.list_files("General/17_SW Sprint/archive")

INTEGRATION WITH update_sprint_management.py
=============================================
The helper enables update_sprint_management.py to run on SharePoint with ZERO logic changes.

Before (local file):
    wb = load_workbook(input_path, keep_vba=...)
    ws = wb[SHEET_NAME]
    ws.freeze_panes = "A4"
    ws.cell(row=4, column=1).value = 123
    ws.delete_rows(5)
    wb.save(input_path)

After (SharePoint):
    helper = SharePointHelper()
    wb = helper.load_workbook("General/17_SW Sprint/sprint testing.xlsx", keep_vba=True)
    ws = helper.get_sheet(wb, "Sprint Template")
    helper.set_freeze_panes(ws, "A4")
    helper.set_cell(ws, 4, 1, 123)
    helper.delete_rows(ws, 5)
    helper.save_workbook("General/17_SW Sprint/sprint testing.xlsx", wb)

The logic stays identical - only I/O changes from local file to SharePoint helper calls.

ERROR HANDLING
==============
- Operations log warnings and continue where possible
- Errors are raised only when operation cannot proceed
- All exceptions are wrapped in SharePointAuthError or SharePointOperationError
- Check logs for detailed error information

LOGGING
=======
All operations are logged at appropriate levels:
- DEBUG: Low-level operations (cell access, row snapshots)
- INFO: Major operations (workbook load/save, archive)
- WARNING: Non-fatal issues (failed to set styling, missing properties)
- ERROR: Fatal issues (authentication failure, file not found)

Monitor logs to understand what's happening:
    import logging
    logging.basicConfig(level=logging.DEBUG)

PERFORMANCE NOTES
=================
- Token caching means only first operation authenticates
- Subsequent operations reuse cached token (no API overhead)
- Table XML patching requires temporary ZIP decompression (minimal overhead)
- Row snapshots capture all 12 columns by default (configurable)

REQUIREMENTS
============
- python-dotenv: For loading .env credentials
- requests: For Microsoft Graph API calls
- openpyxl: For Excel workbook manipulation
- Standard library: copy, io, json, logging, os, re, uuid, zipfile, datetime

All in requirements.txt
"""

import copy
import io
import json
import logging
import os
import re
import uuid
import zipfile
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Optional, List, Dict, Tuple, Any

import requests
from dotenv import load_dotenv
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Protection, PatternFill, Border
from openpyxl.worksheet.datavalidation import DataValidation, DataValidationList
from openpyxl.worksheet.table import Table, TableStyleInfo, TableColumn
from openpyxl.utils import range_boundaries

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

# Load environment variables
load_dotenv()

# Configure logging
logger = logging.getLogger(__name__)
logging.basicConfig(level=LOG_LEVEL, format=LOG_FORMAT)

# Constants
SNAPSHOT_NUM_COLS = 12  # Default number of columns to snapshot per row
SHEET_NAME_DEFAULT = "Sprint Template"

# SharePoint File Paths (can be overridden from .env)
SHAREPOINT_SYNC_PATH = os.getenv("SHAREPOINT_SYNC_PATH", "General/17_SW Sprint/Leahi Sprint Management.xlsx")
ARCHIVE_PATH = os.getenv("ARCHIVE_PATH", "General/17_SW Sprint/archive")
SPRINT_TEMPLATE_FILE = os.getenv("SPRINT_TEMPLATE_FILE", "General/17_SW Sprint/sprint testing.xlsx")


class SharePointAuthError(Exception):
    """Raised when authentication fails."""
    pass


class SharePointOperationError(Exception):
    """Raised when SharePoint operation fails."""
    pass


class TokenCache:
    """Manages OAuth token caching with expiration."""

    def __init__(self, cache_file: str = TOKEN_CACHE_FILE):
        self.cache_file = cache_file

    def save(self, token: str, expires_in: int) -> None:
        """Save token with expiration timestamp."""
        expiry = datetime.now() + timedelta(seconds=expires_in)
        cache_data = {
            "token": token,
            "expiry": expiry.isoformat()
        }
        try:
            with open(self.cache_file, "w") as f:
                json.dump(cache_data, f)
            logger.debug(f"Token cached successfully, expires at {expiry.isoformat()}")
        except Exception as e:
            logger.error(f"Failed to save token cache: {e}")

    def load(self) -> Optional[str]:
        """Load token if valid (not expired)."""
        try:
            if not Path(self.cache_file).exists():
                return None

            with open(self.cache_file, "r") as f:
                cache_data = json.load(f)

            expiry = datetime.fromisoformat(cache_data["expiry"])
            buffer_time = expiry - timedelta(seconds=TOKEN_EXPIRY_BUFFER)

            if datetime.now() < buffer_time:
                logger.debug("Valid token found in cache")
                return cache_data["token"]

            logger.debug("Token in cache has expired or is about to expire")
            return None
        except Exception as e:
            logger.warning(f"Failed to load token cache: {e}")
            return None

    def clear(self) -> None:
        """Clear cached token."""
        try:
            if Path(self.cache_file).exists():
                Path(self.cache_file).unlink()
                logger.debug("Token cache cleared")
        except Exception as e:
            logger.error(f"Failed to clear token cache: {e}")


class SharePointHelper:
    """Helper class for SharePoint operations via Microsoft Graph API."""

    def __init__(
        self,
        tenant_id: Optional[str] = None,
        client_id: Optional[str] = None,
        client_secret: Optional[str] = None,
        site_id: str = SHAREPOINT_SITE_ID,
        drive_id: str = SHAREPOINT_DRIVE_ID,
    ):
        
        self.tenant_id = tenant_id or os.getenv("TENANT_ID")
        self.client_id = client_id or os.getenv("CLIENT_ID")
        self.client_secret = client_secret or os.getenv("CLIENT_SECRET")
        self.site_id = site_id
        self.drive_id = drive_id

        if not all([self.tenant_id, self.client_id, self.client_secret]):
            raise SharePointAuthError(
                "Missing credentials. Ensure TENANT_ID, CLIENT_ID, and CLIENT_SECRET "
                "are set in .env file or passed as arguments."
            )

        self.token_cache = TokenCache()
        self.access_token: Optional[str] = None
        self.token_expiry: Optional[datetime] = None

    def _get_access_token(self) -> str:
        """
        Get access token from cache or request new one.

        Returns:
            Valid access token

        Raises:
            SharePointAuthError: If token retrieval fails
        """
        # Try to get cached token first
        cached_token = self.token_cache.load()
        if cached_token:
            self.access_token = cached_token
            return cached_token

        # Request new token
        url = AUTH_ENDPOINT.format(tenant_id=self.tenant_id)
        payload = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }
        headers = {"Content-Type": "application/x-www-form-urlencoded"}

        try:
            response = requests.post(url, data=payload, headers=headers, timeout=10)
            response.raise_for_status()

            data = response.json()
            token = data.get("access_token")
            expires_in = data.get("expires_in", 3600)

            if not token:
                raise SharePointAuthError("No access token in response")

            # Cache the token
            self.token_cache.save(token, expires_in)
            self.access_token = token
            logger.info("Successfully obtained new access token")
            return token

        except requests.RequestException as e:
            logger.error(f"Failed to get access token: {e}")
            raise SharePointAuthError(f"Authentication failed: {e}")

    def _get_headers(self) -> dict:
        """Get headers with authorization token."""
        if not self.access_token:
            self.access_token = self._get_access_token()

        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
        }

    def _build_file_url(self, file_path: str) -> str:
        """
        Build SharePoint Graph API URL for file operations.

        Args:
            file_path: Path within SharePoint (e.g., 'General/test.xlsx')

        Returns:
            Full Graph API URL for the file
        """
        return (
            f"{GRAPH_API_BASE}/sites/{self.site_id}/drives/{self.drive_id}"
            f"/root:/{file_path}:/content"
        )

    def write_excel(self, file_path: str, file_content: bytes) -> None:
        """
        Write Excel file to SharePoint.

        Args:
            file_path: Path in SharePoint (e.g., 'General/report.xlsx')
            file_content: Binary content of the Excel file

        Raises:
            SharePointOperationError: If write operation fails
        """
        url = self._build_file_url(file_path)
        headers = self._get_headers()
        headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        try:
            logger.info(f"Writing Excel file to SharePoint: {file_path}")
            response = requests.put(url, data=file_content, headers=headers, timeout=30)
            response.raise_for_status()
            logger.info(f"Successfully wrote Excel file: {file_path}")

        except requests.RequestException as e:
            logger.error(f"Failed to write Excel file to SharePoint: {e}")
            raise SharePointOperationError(f"Write operation failed: {e}")

    def read_excel(self, file_path: str) -> bytes:
        """
        Read Excel file from SharePoint.

        Args:
            file_path: Path in SharePoint (e.g., 'General/report.xlsx')

        Returns:
            Binary content of the Excel file

        Raises:
            SharePointOperationError: If read operation fails
        """
        url = self._build_file_url(file_path)
        headers = self._get_headers()

        try:
            logger.info(f"Reading Excel file from SharePoint: {file_path}")
            response = requests.get(url, headers=headers, timeout=30)
            response.raise_for_status()
            logger.info(f"Successfully read Excel file: {file_path}")
            return response.content

        except requests.RequestException as e:
            logger.error(f"Failed to read Excel file from SharePoint: {e}")
            raise SharePointOperationError(f"Read operation failed: {e}")

    def read_excel_to_dataframe(self, file_path: str):
        """
        Read Excel file from SharePoint into a pandas DataFrame.

        Args:
            file_path: Path in SharePoint

        Returns:
            pandas DataFrame

        Raises:
            SharePointOperationError: If operation fails
        """
        try:
            import pandas as pd
        except ImportError:
            logger.error("pandas is required for this operation. Install it with: pip install pandas openpyxl")
            raise SharePointOperationError("pandas not installed")

        try:
            content = self.read_excel(file_path)
            df = pd.read_excel(BytesIO(content))
            logger.info(f"Excel file read into DataFrame: {file_path}")
            return df
        except Exception as e:
            logger.error(f"Failed to read Excel file into DataFrame: {e}")
            raise

    def write_dataframe_to_excel(self, df, file_path: str) -> None:
        """
        Write pandas DataFrame as Excel file to SharePoint.

        Args:
            df: pandas DataFrame
            file_path: Destination path in SharePoint

        Raises:
            SharePointOperationError: If operation fails
        """
        try:
            import pandas as pd
        except ImportError:
            logger.error("pandas is required for this operation. Install it with: pip install pandas openpyxl")
            raise SharePointOperationError("pandas not installed")

        try:
            buffer = BytesIO()
            df.to_excel(buffer, index=False, engine="openpyxl")
            buffer.seek(0)
            self.write_excel(file_path, buffer.getvalue())
            logger.info(f"DataFrame written to SharePoint as Excel: {file_path}")
        except Exception as e:
            logger.error(f"Failed to write DataFrame to Excel: {e}")
            raise

    def clear_token_cache(self) -> None:
        """Clear cached authentication token."""
        self.token_cache.clear()
        self.access_token = None
        logger.info("Token cache cleared")

    # ==================== PHASE 1: WORKBOOK I/O ====================

    def load_workbook(self, file_path: str, keep_vba: bool = False):
        """
        Load Excel file from SharePoint into openpyxl.Workbook object.

        Args:
            file_path: Path in SharePoint (e.g., 'General/sprint.xlsx')
            keep_vba: If True, preserve VBA macros (XLSM files)

        Returns:
            openpyxl.Workbook object

        Raises:
            SharePointOperationError: If read/load fails
        """
        try:
            logger.info(f"Loading workbook from SharePoint: {file_path}")
            bytes_content = self.read_excel(file_path)
            wb = load_workbook(BytesIO(bytes_content), keep_vba=keep_vba)
            logger.info(f"Workbook loaded successfully. Sheets: {wb.sheetnames}")
            return wb
        except Exception as e:
            logger.error(f"Failed to load workbook: {e}")
            raise SharePointOperationError(f"Failed to load workbook: {e}")

    def save_workbook(self, file_path: str, workbook) -> None:
        """
        Save openpyxl.Workbook back to SharePoint.

        Args:
            file_path: Destination path in SharePoint
            workbook: openpyxl.Workbook object

        Raises:
            SharePointOperationError: If save/write fails
        """
        try:
            logger.info(f"Saving workbook to SharePoint: {file_path}")
            buffer = BytesIO()
            workbook.save(buffer)
            buffer.seek(0)
            self.write_excel(file_path, buffer.getvalue())
            logger.info("Workbook saved successfully")
        except Exception as e:
            logger.error(f"Failed to save workbook: {e}")
            raise SharePointOperationError(f"Failed to save workbook: {e}")

    # ==================== PHASE 2: SHEET OPERATIONS ====================

    def get_sheet(self, workbook, sheet_name: str):
        """
        Get worksheet by name with validation.

        Args:
            workbook: openpyxl.Workbook object
            sheet_name: Name of sheet to retrieve

        Returns:
            openpyxl.Worksheet object

        Raises:
            ValueError: If sheet doesn't exist
        """
        if sheet_name not in workbook.sheetnames:
            logger.error(f"Sheet '{sheet_name}' not found. Available: {workbook.sheetnames}")
            raise ValueError(
                f"Sheet '{sheet_name}' not found in workbook. "
                f"Available sheets: {workbook.sheetnames}"
            )
        return workbook[sheet_name]

    def get_sheet_names(self, workbook) -> List[str]:
        """Get list of all sheet names in workbook."""
        return workbook.sheetnames

    def sheet_exists(self, workbook, sheet_name: str) -> bool:
        """Check if sheet exists in workbook."""
        return sheet_name in workbook.sheetnames

    def get_max_row(self, worksheet) -> int:
        """Get the maximum row number with data in worksheet."""
        return worksheet.max_row

    def get_max_column(self, worksheet) -> int:
        """Get the maximum column number with data in worksheet."""
        return worksheet.max_column

    def set_freeze_panes(self, worksheet, cell_ref: str) -> None:
        """
        Set freeze panes at specified cell.

        Args:
            worksheet: openpyxl.Worksheet object
            cell_ref: Cell reference (e.g., 'A4' to freeze first 3 rows)
        """
        try:
            worksheet.freeze_panes = cell_ref
            logger.debug(f"Freeze panes set to {cell_ref}")
        except Exception as e:
            logger.warning(f"Failed to set freeze panes: {e}")

    def set_row_height(self, worksheet, row: int, height: float) -> None:
        """Set height of specified row."""
        try:
            worksheet.row_dimensions[row].height = height
        except Exception as e:
            logger.warning(f"Failed to set row {row} height: {e}")

    def set_column_width(self, worksheet, col_letter: str, width: float) -> None:
        """Set width of specified column."""
        try:
            worksheet.column_dimensions[col_letter].width = width
        except Exception as e:
            logger.warning(f"Failed to set column {col_letter} width: {e}")

    # ==================== PHASE 3: CELL OPERATIONS ====================

    def get_cell(self, worksheet, row: int, col: int) -> Any:
        """
        Get cell value.

        Args:
            worksheet: openpyxl.Worksheet object
            row: Row number (1-indexed)
            col: Column number (1-indexed)

        Returns:
            Cell value (any type)
        """
        return worksheet.cell(row=row, column=col).value

    def set_cell(self, worksheet, row: int, col: int, value: Any) -> None:
        """
        Set cell value.

        Args:
            worksheet: openpyxl.Worksheet object
            row: Row number (1-indexed)
            col: Column number (1-indexed)
            value: Value to set
        """
        worksheet.cell(row=row, column=col).value = value

    def set_cell_styled(self, worksheet, row: int, col: int, value: Any,
                       font: Font = None, fill: PatternFill = None,
                       border: Border = None, alignment: Alignment = None,
                       protection: Protection = None, hyperlink: str = None) -> None:
        """
        Set cell value with styling (font, fill, border, alignment, protection, hyperlink).

        Args:
            worksheet: openpyxl.Worksheet object
            row: Row number
            col: Column number
            value: Cell value
            font: openpyxl.Font object
            fill: openpyxl.PatternFill object
            border: openpyxl.Border object
            alignment: openpyxl.Alignment object
            protection: openpyxl.Protection object
            hyperlink: URL string for hyperlink
        """
        cell = worksheet.cell(row=row, column=col)
        cell.value = value
        if font:
            cell.font = font
        if fill:
            cell.fill = fill
        if border:
            cell.border = border
        if alignment:
            cell.alignment = alignment
        if protection:
            cell.protection = protection
        if hyperlink:
            cell.hyperlink = hyperlink

    def get_cell_style(self, worksheet, row: int, col: int) -> dict:
        """
        Get cell styling (font, fill, border, alignment, protection, hyperlink, comment).

        Returns:
            Dictionary with all style properties
        """
        cell = worksheet.cell(row=row, column=col)
        return {
            "value": cell.value,
            "font": copy.copy(cell.font),
            "fill": copy.copy(cell.fill),
            "border": copy.copy(cell.border),
            "alignment": copy.copy(cell.alignment),
            "number_format": cell.number_format,
            "protection": copy.copy(cell.protection),
            "hyperlink": copy.copy(cell.hyperlink),
            "comment": copy.copy(cell.comment),
        }

    def set_cell_style(self, worksheet, row: int, col: int, style: dict) -> None:
        """
        Apply styling to a cell.

        Args:
            worksheet: openpyxl.Worksheet object
            row: Row number
            col: Column number
            style: Dictionary with style properties (from get_cell_style)
        """
        cell = worksheet.cell(row=row, column=col)
        if "value" in style:
            cell.value = style["value"]
        if "font" in style:
            cell.font = style["font"]
        if "fill" in style:
            cell.fill = style["fill"]
        if "border" in style:
            cell.border = style["border"]
        if "alignment" in style:
            cell.alignment = style["alignment"]
        if "number_format" in style:
            cell.number_format = style["number_format"]
        if "protection" in style:
            cell.protection = style["protection"]
        if "hyperlink" in style:
            cell.hyperlink = style["hyperlink"]
        if "comment" in style:
            cell.comment = style["comment"]

    # ==================== PHASE 4: ROW OPERATIONS ====================

    def delete_rows(self, worksheet, start_row: int, count: int = 1) -> None:
        """
        Delete rows starting at start_row.

        Args:
            worksheet: openpyxl.Worksheet object
            start_row: Starting row number to delete
            count: Number of rows to delete
        """
        try:
            worksheet.delete_rows(start_row, count)
            logger.debug(f"Deleted {count} rows starting at row {start_row}")
        except Exception as e:
            logger.warning(f"Failed to delete rows: {e}")

    def insert_rows(self, worksheet, row: int, count: int = 1) -> None:
        """
        Insert blank rows at specified position.

        Args:
            worksheet: openpyxl.Worksheet object
            row: Row number to insert at (1-indexed)
            count: Number of rows to insert
        """
        try:
            worksheet.insert_rows(row, count)
            logger.debug(f"Inserted {count} rows at row {row}")
        except Exception as e:
            logger.warning(f"Failed to insert rows: {e}")

    def snapshot_row(self, worksheet, row: int, num_cols: int = SNAPSHOT_NUM_COLS) -> List[dict]:
        """
        Capture full-fidelity per-cell state (value, style, hyperlink, comment) for a row.

        Mirrors update_sprint_management.py lines 680-696.

        Args:
            worksheet: openpyxl.Worksheet object
            row: Row number to snapshot
            num_cols: Number of columns to capture

        Returns:
            List of dictionaries with complete cell state for each column
        """
        cells = []
        for col in range(1, num_cols + 1):
            cell = worksheet.cell(row=row, column=col)
            cells.append({
                "value": cell.value,
                "font": copy.copy(cell.font),
                "fill": copy.copy(cell.fill),
                "border": copy.copy(cell.border),
                "alignment": copy.copy(cell.alignment),
                "number_format": cell.number_format,
                "protection": copy.copy(cell.protection),
                "hyperlink": copy.copy(cell.hyperlink),
                "comment": copy.copy(cell.comment),
            })
        logger.debug(f"Row {row} snapshotted ({num_cols} columns)")
        return cells

    def restore_row(self, worksheet, row: int, snapshot: List[dict]) -> None:
        """
        Write a snapshot captured by snapshot_row back into row.

        Mirrors update_sprint_management.py lines 699-711.

        Args:
            worksheet: openpyxl.Worksheet object
            row: Target row number
            snapshot: List of cell dictionaries from snapshot_row()
        """
        try:
            for col, state in enumerate(snapshot, start=1):
                cell = worksheet.cell(row=row, column=col)
                cell.value = state["value"]
                cell.font = state["font"]
                cell.fill = state["fill"]
                cell.border = state["border"]
                cell.alignment = state["alignment"]
                cell.number_format = state["number_format"]
                cell.protection = state["protection"]
                cell.hyperlink = state["hyperlink"]
                cell.comment = state["comment"]
            logger.debug(f"Row {row} restored from snapshot")
        except Exception as e:
            logger.warning(f"Failed to restore row {row}: {e}")

    def clear_row(self, worksheet, row: int, start_col: int = 1, end_col: int = None) -> None:
        """Clear all values and formatting from a row range."""
        if end_col is None:
            end_col = worksheet.max_column
        try:
            for col in range(start_col, end_col + 1):
                cell = worksheet.cell(row=row, column=col)
                cell.value = None
                cell.font = Font()
                cell.fill = PatternFill()
                cell.border = Border()
                cell.alignment = Alignment()
                cell.protection = Protection()
                cell.hyperlink = None
                cell.comment = None
            logger.debug(f"Row {row} cleared")
        except Exception as e:
            logger.warning(f"Failed to clear row {row}: {e}")

    # ==================== PHASE 5: DATA VALIDATION ====================

    def get_data_validations(self, worksheet):
        """
        Get the DataValidationList from worksheet.

        Returns:
            openpyxl.worksheet.datavalidation.DataValidationList
        """
        return worksheet.data_validations

    def reset_data_validations(self, worksheet) -> None:
        """Clear all data validation rules from worksheet."""
        try:
            worksheet.data_validations = DataValidationList()
            logger.debug("Data validations reset")
        except Exception as e:
            logger.warning(f"Failed to reset data validations: {e}")

    def add_data_validation(self, worksheet, validation: DataValidation) -> None:
        """
        Add a data validation rule to worksheet.

        Args:
            worksheet: openpyxl.Worksheet object
            validation: openpyxl.worksheet.datavalidation.DataValidation object
        """
        try:
            worksheet.add_data_validation(validation)
            logger.debug(f"Data validation added")
        except Exception as e:
            logger.warning(f"Failed to add data validation: {e}")

    # ==================== PHASE 6: TABLE OPERATIONS ====================

    def get_tables(self, worksheet) -> dict:
        """
        Get all tables in worksheet.

        Returns:
            Dictionary of table_name -> Table objects
        """
        return worksheet.tables

    def capture_table_xml(self, wb_bytes: bytes) -> Tuple[Optional[str], Optional[str]]:
        """
        Extract table XML from XLSX zipfile.

        Mirrors update_sprint_management.py lines 611-618.

        Args:
            wb_bytes: Binary content of XLSX file

        Returns:
            Tuple of (table_path, table_xml_string) or (None, None) if no table
        """
        try:
            with zipfile.ZipFile(BytesIO(wb_bytes), "r") as z:
                table_paths = [n for n in z.namelist()
                               if n.startswith("xl/tables/") and n.endswith(".xml")]
                if not table_paths:
                    logger.debug("No table found in XLSX")
                    return None, None
                table_path = table_paths[0]
                xml_content = z.read(table_path).decode("utf-8", errors="replace")
                logger.debug(f"Captured table XML from {table_path}")
                return table_path, xml_content
        except Exception as e:
            logger.warning(f"Failed to capture table XML: {e}")
            return None, None

    def _generate_table_xml(self, ref: str, col_names: List[str], uid: str = None) -> bytes:
        """
        Generate table XML for Excel table.

        Mirrors update_sprint_management.py lines 588-608.

        Args:
            ref: Table reference (e.g., 'A3:L100')
            col_names: List of column header names
            uid: Optional UUID for table (generates new one if not provided)

        Returns:
            XML as bytes
        """
        xr_uid = uid or ("{" + str(uuid.uuid4()).upper() + "}")

        def _xml_attr(value: str) -> str:
            """Escape special XML characters."""
            return (value.replace("&", "&amp;").replace("<", "&lt;")
                         .replace(">", "&gt;").replace('"', "&quot;"))

        cols_xml = "".join(
            f'<tableColumn id="{i + 1}" name="{_xml_attr(n)}"/>'
            for i, n in enumerate(col_names)
        )

        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
            ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"'
            ' xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"'
            ' mc:Ignorable="xr xr3"'
            f' id="1" name="Table1" displayName="Table1"'
            f' ref="{ref}" totalsRowShown="0" xr:uid="{xr_uid}">'
            f'<autoFilter ref="{ref}"/>'
            f'<tableColumns count="{len(col_names)}">{cols_xml}</tableColumns>'
            f'<tableStyleInfo name="TableStyleMedium7" showFirstColumn="0" showLastColumn="0"'
            ' showRowStripes="1" showColumnStripes="0"/>'
            '</table>'
        ).encode("utf-8")

    def patch_table_xml(self, wb_bytes: bytes, final_row: int,
                       col_names: List[str], orig_table_xml: str = None) -> bytes:
        """
        Regenerate table XML and inject back into XLSX zipfile.

        Mirrors update_sprint_management.py lines 621-646.

        Args:
            wb_bytes: Binary content of XLSX file
            final_row: Last row with data (used to set table ref)
            col_names: List of column names for table
            orig_table_xml: Original table XML (to extract UID)

        Returns:
            Modified XLSX as bytes
        """
        try:
            ref = f"A3:L{final_row}"

            with zipfile.ZipFile(BytesIO(wb_bytes), "r") as zin:
                table_paths = [n for n in zin.namelist()
                               if n.startswith("xl/tables/") and n.endswith(".xml")]
                if not table_paths:
                    logger.debug("No table in XLSX to patch")
                    return wb_bytes

                table_path = table_paths[0]
                raw_xml = zin.read(table_path).decode("utf-8", errors="replace")
                entries = [(info, zin.read(info.filename)) for info in zin.infolist()]

            # Extract UID from original or current XML
            source = orig_table_xml or raw_xml
            uid_match = re.search(r'xr:uid="([^"]*)"', source)
            uid = uid_match.group(1) if uid_match else None

            # Generate new table XML
            clean_xml = self._generate_table_xml(ref, col_names, uid=uid)

            # Write back to zip
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
                for info, data in entries:
                    zout.writestr(info.filename, clean_xml if info.filename == table_path else data)

            result = buf.getvalue()
            logger.debug(f"Table XML patched for {final_row} rows")
            return result

        except Exception as e:
            logger.warning(f"Failed to patch table XML: {e}")
            return wb_bytes

    # ==================== PHASE 7: SHAREPOINT FILE OPERATIONS ====================

    def file_exists(self, file_path: str) -> bool:
        """
        Check if file exists on SharePoint by attempting to read it.

        Args:
            file_path: Path in SharePoint

        Returns:
            True if file exists, False otherwise
        """
        try:
            self.read_excel(file_path)
            return True
        except SharePointOperationError:
            return False
        except Exception:
            return False

    def archive_file(self, file_path: str, archive_folder: str = None) -> str:
        """
        Create timestamped copy of file on SharePoint.

        Creates a backup with format: filename_MMDDYYYY.ext

        Args:
            file_path: Path of file to archive
            archive_folder: SharePoint folder for archive (e.g., 'General/17_SW Sprint/archive')
                           If None, archives in same folder as source

        Returns:
            Path to created archive file

        Raises:
            SharePointOperationError: If archive fails
        """
        try:
            # Read source file
            logger.info(f"Archiving file: {file_path}")
            file_bytes = self.read_excel(file_path)

            # Generate archive path
            stem = Path(file_path).stem
            ext = Path(file_path).suffix
            date_suffix = datetime.now().strftime("%m%d%Y")

            if archive_folder:
                archive_path = f"{archive_folder}/{stem}_{date_suffix}{ext}"
            else:
                # Same folder as source
                parent = str(Path(file_path).parent)
                if parent == ".":
                    archive_path = f"{stem}_{date_suffix}{ext}"
                else:
                    archive_path = f"{parent}/{stem}_{date_suffix}{ext}"

            # Write archive
            self.write_excel(archive_path, file_bytes)
            logger.info(f"File archived to: {archive_path}")
            return archive_path

        except Exception as e:
            logger.error(f"Failed to archive file: {e}")
            raise SharePointOperationError(f"Failed to archive file: {e}")

    def list_files(self, folder_path: str) -> List[str]:
        """
        List files in SharePoint folder.

        Uses Microsoft Graph /children endpoint.

        Args:
            folder_path: Path to folder in SharePoint (e.g., 'General/archive')

        Returns:
            List of file paths in the folder

        Raises:
            SharePointOperationError: If listing fails
        """
        try:
            headers = self._get_headers()

            # Build URL to get folder's children
            folder_url = (
                f"{GRAPH_API_BASE}/sites/{self.site_id}/drives/{self.drive_id}"
                f"/root:/{folder_path}"
            )

            response = requests.get(f"{folder_url}/children", headers=headers, timeout=30)
            response.raise_for_status()

            data = response.json()
            files = []
            for item in data.get("value", []):
                if "file" in item:  # It's a file, not a folder
                    # Build full path
                    file_name = item.get("name", "")
                    files.append(f"{folder_path}/{file_name}")

            logger.debug(f"Listed {len(files)} files in {folder_path}")
            return files

        except Exception as e:
            logger.warning(f"Failed to list files in {folder_path}: {e}")
            return []
