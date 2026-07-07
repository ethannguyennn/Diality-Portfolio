"""Sprint Retro header stamping (separate sheet, separate session)."""

import logging
import re

from grid_utils import _gv
from sharepoint_helper import SharePointOperationError

logger = logging.getLogger(__name__)

# ── Sheet layout ──
RETRO_SHEET_NAME = "Sprint Retro"
RETRO_NUM_COLS = 3  # A "What went well?" .. C "What can we improve?"
_WS_RETRO = f"worksheets('{RETRO_SHEET_NAME}')"
_RETRO_LAST_COL = chr(ord('A') + RETRO_NUM_COLS - 1)  # 'C'

# Sprint Retro header block styling, confirmed against the existing
# manually-created Sprint 31/32 blocks already on the sheet (rowHeight is
# in points; font/fill matched exactly).
RETRO_TITLE_FILL = "#92D050"
RETRO_SUBHEADER_FILL = "#BFBFBF"
RETRO_TITLE_ROW_HEIGHT_PT = 23.25
RETRO_SUBHEADER_ROW_HEIGHT_PT = 15.0
RETRO_TITLE_RE = re.compile(r'Sprint (\d+) Retro')  # matched with .fullmatch()


def ensure_sprint_retro_header(helper, item_id, session_id, current_sprint_num):
    """Stamp the two-row retro header block for the current sprint, once.

    'Sprint Retro' is a manually-maintained scrapbook: each sprint gets a
    title row + column-header row that this function writes, and the team
    fills in retro notes underneath by hand -- this function never touches
    those notes. Cron runs nightly for the same sprint until it advances,
    so without an existence check we'd stamp a duplicate header every
    night. Detection scans column A for "Sprint N Retro" and skips once
    current_sprint_num is already covered.
    """
    def wb(method, rel, body=None):
        # Dispatch a Workbook API request through the SharePoint helper.
        return helper.workbook_request(method, item_id, session_id, rel, body)

    try:
        logger.info(f"Checking Sprint Retro sheet for an existing sprint {current_sprint_num} header")
        try:
            used = wb("GET", f"{_WS_RETRO}/usedRange?$select=values,rowCount,address")
        except SharePointOperationError as e:
            # Graph throws ItemNotFound on a genuinely empty sheet rather
            # than returning an empty grid; treat that specific case as empty.
            if "ItemNotFound" in str(e):
                logger.info(f"Sprint Retro usedRange lookup indicates an empty sheet ({e}); treating as empty")
                used = {}
            else:
                raise

        grid = used.get("values", [])
        total_rows = used.get("rowCount", len(grid))
        logger.debug(f"Sprint Retro grid read: address={used.get('address')}, rowCount={total_rows}")

        # Single pass: find the last content row (any of A-C) and the
        # highest sprint number already stamped in a title row.
        max_found = 0
        last_content_row = 0
        for r in range(1, total_rows + 1):
            if any(_gv(grid, r, c) is not None for c in range(1, RETRO_NUM_COLS + 1)):
                last_content_row = r
            a_val = _gv(grid, r, 1)
            if a_val is not None:
                m = RETRO_TITLE_RE.fullmatch(str(a_val).strip())
                if m:
                    max_found = max(max_found, int(m.group(1)))

        if current_sprint_num <= max_found:
            logger.info(
                f"Sprint Retro header for sprint {current_sprint_num} already present "
                f"(max found: {max_found}); skipping"
            )
            return

        title_row = 1 if last_content_row == 0 else last_content_row + 2
        subheader_row = title_row + 1
        logger.info(
            f"Writing Sprint Retro header for sprint {current_sprint_num} "
            f"at rows {title_row}-{subheader_row} (last_content_row={last_content_row})"
        )

        # ── Values ──
        wb("PATCH", f"{_WS_RETRO}/range(address='A{title_row}:{_RETRO_LAST_COL}{title_row}')",
           {"values": [[f"Sprint {current_sprint_num} Retro", "", ""]]})
        wb("PATCH", f"{_WS_RETRO}/range(address='A{subheader_row}:{_RETRO_LAST_COL}{subheader_row}')",
           {"values": [["What went well?", "What went wrong?", "What can we improve?"]]})

        # ── Font ──
        wb("PATCH", f"{_WS_RETRO}/range(address='A{title_row}:{_RETRO_LAST_COL}{title_row}')/format/font",
           {"name": "Calibri", "size": 18, "bold": True})
        wb("PATCH", f"{_WS_RETRO}/range(address='A{subheader_row}:{_RETRO_LAST_COL}{subheader_row}')/format/font",
           {"name": "Aptos Narrow", "size": 11, "bold": True})

        # ── Fill ──
        wb("PATCH", f"{_WS_RETRO}/range(address='A{title_row}:{_RETRO_LAST_COL}{title_row}')/format/fill",
           {"color": RETRO_TITLE_FILL})
        wb("PATCH", f"{_WS_RETRO}/range(address='A{subheader_row}:{_RETRO_LAST_COL}{subheader_row}')/format/fill",
           {"color": RETRO_SUBHEADER_FILL})

        # ── Row height (a row-level property; setting it on any cell in
        # the row sets the whole row) ──
        wb("PATCH", f"{_WS_RETRO}/range(address='A{title_row}:{_RETRO_LAST_COL}{title_row}')/format",
           {"rowHeight": RETRO_TITLE_ROW_HEIGHT_PT})
        wb("PATCH", f"{_WS_RETRO}/range(address='A{subheader_row}:{_RETRO_LAST_COL}{subheader_row}')/format",
           {"rowHeight": RETRO_SUBHEADER_ROW_HEIGHT_PT})

        logger.info(f"Sprint Retro header for sprint {current_sprint_num} written successfully")
    except Exception:
        # Non-fatal: the core Jira sync already succeeded, so log and move
        # on rather than failing the whole cron run over the manually-
        # maintained retro sheet.
        logger.exception("Could not ensure Sprint Retro header")
