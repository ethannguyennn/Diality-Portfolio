import argparse
import io
import os
import re
import zipfile
import requests
from copy import copy
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime

load_dotenv()

BASE_URL = "https://diality.atlassian.net"
PROJECT = "LDT"
SHEET_NAME = "Sprint Template"
DATA_START_ROW = 4  # Row 1=note banner, Row 2=sprint info, Row 3=headers, Row 4+=data

SWVV_TEAM = ["Raghu Kallala", "Thomas Lippold", "Tejaskumar Patel", "Tiffany Mejia",
             "Zoltan Miskolci", "Sarina Cheung", "Tisha Patel", "Ethan Nguyen"]
FW_TEAM   = ["Arpita Srivastava", "Jashwant Gantyada", "Michael Garthwaite", "Praneeth Bunne",
             "Sameer Poyil", "Varshini Nagabooshanam", "Vijay Pamula", "Suresh Dharnala", "Santhos Kumar Reddy", 
             "Vinayakam Mani"]
SW_TEAM   = ["Nicholas Ramirez", "Stephen Quong", "Dara Navaei", "Sean Nash", "Behrouz NematiPour"]
SYS_TEAM  = ["Eliza Petersen", "Caitlynn Chang", "Christina Heine"]
TEAMS = [FW_TEAM, SWVV_TEAM, SW_TEAM, SYS_TEAM]

# Full Jira display name → nickname used in the Selection sheet dropdowns.
# Non-obvious entries are listed first; straightforward first-name entries follow.
NICKNAME_MAP = {
    "Tejaskumar Patel":       "Tejas",
    "Nicholas Ramirez":       "Nico",
    "Thomas Lippold":         "Tom",
    "Jashwant Gantyada":      "Jashwant",       
    "Behrouz NematiPour":     "Behrouz",
    "Sean Nash":              "Sean",
    "Vinayakam Mani":         "Vinay",
    "Raghu Kallala":          "Raghu",
    "Tiffany Mejia":          "Tiffany",
    "Zoltan Miskolci":        "Zoltan",
    "Sarina Cheung":          "Sarina",
    "Tisha Patel":            "Tisha",
    "Ethan Nguyen":           "Ethan",
    "Arpita Srivastava":      "Arpita",
    "Michael Garthwaite":     "Michael",
    "Praneeth Bunne":         "Praneeth",
    "Sameer Poyil":           "Sameer",
    "Varshini Nagabooshanam": "Varshini",
    "Vijay Pamula":           "Vijay",
    "Suresh Dharnala":        "Suresh",
    "Dara Navaei":            "Dara",
    "Eliza Petersen":         "Eliza",
    "Stephen Quong":          "Stephen",
    "Christina Heine":        "Christina",
    "Caitlynn Chang":         "Caitlynn",
    "Santhos Kumar Reddy":    "Santhos"
}

# Lookup from assignee name (full Jira OR nickname) → team sort index.
_ASSIGNEE_TO_TEAM_IDX: dict = {}
for _ti, _team in enumerate(TEAMS):
    for _full in _team:
        _ASSIGNEE_TO_TEAM_IDX[_full] = _ti
        _nick = NICKNAME_MAP.get(_full)
        if _nick:
            _ASSIGNEE_TO_TEAM_IDX[_nick] = _ti

SHAREPOINT_SYNC_PATH = os.getenv("SHAREPOINT_SYNC_PATH", "").strip()

# Font colors for the Type column (C); types not listed use default black
TYPE_FONT_COLORS = {
    "Bug":     "FF0000",
    "Support": "00B050",
    "Dialin ticket": "4472C4",
    "Little-V": "7030A0",
}

# Light green used for every other data row (rows 5, 7, 9 … relative to DATA_START_ROW)
ALT_ROW_COLOR = "E2EFDA"

# Column headers written to row 3 — must match cell values exactly (leading spaces intentional).
# Used by both _ensure_sheet_setup and _generate_table_xml so they stay in sync.
SHEET_HEADERS = ["Sprint", "Assignee", " Type", "Scope Changes", "Description", "Jira #", " Status"]

TABLE_DISPLAY_NAME = "Table1"
TABLE_STYLE        = "TableStyleMedium2"

MIN_SPRINT = 28  # ignore every sprint before this number

SP_SYNC_COLS = (8, 10)  # H and J only — synced read-only from SharePoint; I is script-controlled
JIRA_NOTE_RE  = re.compile(r'xlsx<([^>]*)>xlsx')

# Thin vertical separator for white rows — left edge of every column after the first.
# Green rows use an empty border so the separator only appears on white rows.
_WHITE_COL_BORDER = Border(left=Side(style="thin", color="E0E0E0"))
_EMPTY_BORDER     = Border()

VALIDATION_ROW_END    = 1000  # dropdowns cover data rows up to this row
_SELECTION_SPRINT_MAX = 100   # Selection!F extended up to this sprint number

# Sprint Template column → Selection sheet source range for dropdown validation.
# Sprint 28 lives at Selection!F29 (row = sprint + 1; row 1 is header).
_COL_VALIDATIONS = {
    "A": f"Selection!$F$29:$F${_SELECTION_SPRINT_MAX + 1}",  # Sprint (28+)
    "B": "Selection!$A$2:$A$26",  # Assignee
    "C": "Selection!$D$2:$D$6",   # Type (Process Type)
    "D": "Selection!$H$2:$H$4",   # Scope Changes
    "E": "Selection!$K$2:$K$3",   # Description (Y/N)
    "G": "Selection!$E$2:$E$7",   # Status (Ticket Status)
}


# ── Jira fetch ────────────────────────────────────────────────────────────────

def _fetch_paginated(url, jql, auth):
    all_issues = []
    start_at = 0
    while True:
        params = {
            "fields": "summary,status,assignee,issuetype,customfield_10020,description,comment",
            "expand": "changelog",
            "jql": jql,
            "maxResults": 100,
            "startAt": start_at,
        }
        response = requests.get(url, params=params, auth=auth)
        response.raise_for_status()
        data = response.json()
        issues = data["issues"]
        all_issues.extend(issues)
        total = data.get("total", 0)
        start_at += len(issues)
        if start_at >= total or not issues:
            break
    return all_issues


def _sprint_names(s):
    return {n.strip() for n in s.split(",") if n.strip()} if s else set()


def _has_sprint_history(issue):
    """Return True if the changelog contains a Sprint transition into sprint >= MIN_SPRINT."""
    for history in issue["changelog"]["histories"]:
        for item in history["items"]:
            if item["field"] != "Sprint":
                continue
            for name in _sprint_names(item.get("toString") or ""):
                if _sprint_num(name) >= MIN_SPRINT:
                    return True
    return False


def _assignee_nickname(display_name: str) -> str:
    """
    Convert a full Jira display name to the Selection-sheet nickname.
    Checks NICKNAME_MAP first; falls back to the first word of the display name
    for people not yet in the map (so 'Jane Doe' becomes 'Jane').
    """
    if not display_name:
        return "Unassigned"
    nick = NICKNAME_MAP.get(display_name)
    if nick:
        return nick
    return display_name.split()[0]


def _adf_to_text(node) -> str:
    """Recursively extract plain text from an Atlassian Document Format node or plain string."""
    if isinstance(node, str):
        return node
    if not isinstance(node, dict):
        return ""
    if node.get("type") == "text":
        return node.get("text", "")
    return "".join(_adf_to_text(c) for c in node.get("content", []))


def _latest_jira_note(issue, auth) -> str:
    """
    Return the trimmed text from the latest xlsx<...>xlsx comment on this issue, or None.
    If the main issue fetch returned fewer comments than the total, fetches all comments
    via the comments endpoint before searching.
    """
    comment_field = issue["fields"].get("comment") or {}
    comments = list(comment_field.get("comments", []))
    total = comment_field.get("total", len(comments))

    if total > len(comments):
        url = f"{BASE_URL}/rest/api/2/issue/{issue['key']}/comment"
        resp = requests.get(url, params={"maxResults": 1000}, auth=auth)
        resp.raise_for_status()
        comments = resp.json().get("comments", [])

    best_text = None
    best_created = ""
    for comment in comments:
        body_text = _adf_to_text(comment.get("body") or "")
        m = JIRA_NOTE_RE.search(body_text)
        if m and comment.get("created", "") >= best_created:
            best_created = comment["created"]
            best_text = m.group(1).strip()

    return best_text


def get_jira_issues():
    auth = HTTPBasicAuth(os.getenv("JIRA_EMAIL"), os.getenv("JIRA_API_TOKEN"))

    board_issues = _fetch_paginated(
        f"{BASE_URL}/rest/agile/1.0/board/84/issue",
        "sprint in openSprints() OR sprint in closedSprints() OR sprint in futureSprints()",
        auth,
    )

    seen = {i["key"] for i in board_issues}

    # Fetch every backlog issue and keep only those with sprint changelog entries.
    all_backlog = _fetch_paginated(
        f"{BASE_URL}/rest/agile/1.0/board/84/backlog",
        f"project = {PROJECT}",
        auth,
    )
    backlog_with_history = [
        i for i in all_backlog
        if i["key"] not in seen and _has_sprint_history(i)
    ]

    all_issues = board_issues + backlog_with_history
    for issue in all_issues:
        issue["_jira_note"] = _latest_jira_note(issue, auth)
    return all_issues


# ── Sprint / issue field helpers ──────────────────────────────────────────────

def _sprint_num(name):
    m = re.search(r'(\d+)', name)
    return int(m.group(1)) if m else -1


def get_type(issue):
    issuetype = issue["fields"]["issuetype"]["name"].lower()
    summary   = issue["fields"].get("summary", "").lower()

    if "dialin" in issuetype or "dialin" in summary:
        return "Dialin ticket"
    elif issuetype == "bug":
        return "Bug"
    elif "little" in issuetype:
        return "Little-V"
    elif "support" in summary:
        return "Support"
    else:
        return "Big-V"


def parse_date(date_str):
    if not date_str:
        return None
    date_str = date_str.replace("Z", "+00:00")
    date_str = re.sub(r'([+-])(\d{2})(\d{2})$', r'\1\2:\3', date_str)
    return datetime.fromisoformat(date_str)


def _classify_sprint_timing(sprint, histories):
    sprint_name     = sprint["name"]
    sprint_start_dt = parse_date(sprint.get("startDate"))

    if not sprint_start_dt:
        return "Scoped"

    most_recent_add = None
    for history in histories:
        for item in history["items"]:
            if item["field"] != "Sprint":
                continue
            to_sprints   = _sprint_names(item.get("toString")  or "")
            from_sprints = _sprint_names(item.get("fromString") or "")
            if sprint_name in to_sprints and sprint_name not in from_sprints:
                added_dt = parse_date(history["created"])
                if added_dt and (most_recent_add is None or added_dt > most_recent_add):
                    most_recent_add = added_dt

    if most_recent_add is None:
        return "Scoped"

    sprint_start_eod = sprint_start_dt.replace(hour=23, minute=59, second=59, microsecond=0)
    return "Added" if most_recent_add > sprint_start_eod else "Scoped"


def get_sprint_history(issue):
    fields      = issue["fields"]
    histories   = issue["changelog"]["histories"]
    sprint_info = fields.get("customfield_10020") or []

    current_sprint_names = {s["name"] for s in sprint_info}
    current_sprint_objs  = {s["name"]: s for s in sprint_info}

    # Collect every sprint name ever referenced in the changelog
    all_sprint_names = set(current_sprint_names)
    for history in histories:
        for item in history["items"]:
            if item["field"] != "Sprint":
                continue
            for name in _sprint_names(item.get("fromString") or ""):
                all_sprint_names.add(name)
            for name in _sprint_names(item.get("toString") or ""):
                all_sprint_names.add(name)

    sprint_entries = {}  # sprint_num → scope

    # Pass 1: changelog-only sprints → ticket was manually removed from each
    for name in all_sprint_names:
        if name not in current_sprint_names:
            num = _sprint_num(name)
            if num >= MIN_SPRINT:
                sprint_entries[num] = "Move out of Sprint"

    # Pass 2: currently-held sprints → timing classification overwrites any conflict
    for name, sprint_obj in current_sprint_objs.items():
        num = _sprint_num(name)
        if num >= MIN_SPRINT:
            sprint_entries[num] = _classify_sprint_timing(sprint_obj, histories)

    return sorted(sprint_entries.items(), key=lambda x: x[0])  # [(sprint_num, scope), ...]


def get_status(issue):
    status       = issue["fields"]["status"]
    status_name  = status["name"].lower()
    category_key = status["statusCategory"]["key"]

    if "blocked" in status_name:
        return "Blocked"
    elif "on hold" in status_name or "on-hold" in status_name:
        return "On-hold"
    elif "in progress" in status_name:
        return "In-progress"
    elif category_key == "done" or "done" in status_name or "complete" in status_name:
        return "Done"
    else:
        return "Not started"


# ── Notification ──────────────────────────────────────────────────────────────

# ── Sort key ──────────────────────────────────────────────────────────────────

def _issue_sort_key(row_record):
    sprint_num = row_record["sprint_num"]
    fields     = row_record["issue"]["fields"]
    assignee   = fields["assignee"]["displayName"] if fields.get("assignee") else ""
    team_idx   = next((i for i, team in enumerate(TEAMS) if assignee in team), len(TEAMS))
    return (sprint_num, team_idx, assignee)


def _extract_key(label):
    """Extract Jira key from a cell value like '[LDT-123] Summary title'."""
    m = re.match(r'\[([A-Z]+-\d+)\]', str(label) if label is not None else '')
    return m.group(1) if m else None


# ── Excel helpers ─────────────────────────────────────────────────────────────

def _ensure_sheet_setup(ws):
    # Row 1: note banner — Calibri 14
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        if cell.value is not None:
            cell.font = Font(name="Calibri", size=14, bold=bool(cell.font.bold))

    # Row 2: sprint info — Calibri 11
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=2, column=col)
        if cell.value is not None:
            cell.font = Font(name="Calibri", size=11, bold=bool(cell.font.bold))

    header_fill = PatternFill(fill_type="solid", fgColor="FFC000")  
    header_font = Font(name="Calibri", size=11, bold=True)

    for col, header in enumerate(SHEET_HEADERS, start=1):
        cell = ws.cell(row=3, column=col)
        cell.value     = header
        cell.fill      = header_fill
        cell.font      = header_font
        cell.alignment = Alignment(horizontal="left", vertical="center")

    ws.row_dimensions[3].height = 18.75
    ws.freeze_panes = "A4"

    # Match actual file column widths; E (Description) left at default (narrow/unused)
    widths = {"A": 9.29, "B": 10.57, "C": 11.57, "D": 18.43, "F": 103.14, "G": 16.29}
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


def _ensure_selection_sprints(wb):
    """Extend Selection!F sprint list to _SELECTION_SPRINT_MAX if it falls short."""
    if "Selection" not in wb.sheetnames:
        return
    ws_sel = wb["Selection"]
    for sprint_num in range(1, _SELECTION_SPRINT_MAX + 1):
        row = sprint_num + 1  # row 1 is header; sprint N lives at row N+1
        if ws_sel.cell(row=row, column=6).value is None:
            ws_sel.cell(row=row, column=6).value = sprint_num


def _ensure_data_validations(ws):
    """Apply dropdown validation to cols A-E and G for all data rows."""
    ws.data_validations.dataValidation = []
    for col_letter, formula in _COL_VALIDATIONS.items():
        dv = DataValidation(
            type="list",
            formula1=formula,
            allow_blank=True,
            showErrorMessage=True,
            errorTitle="Invalid entry",
            error="Please select from the dropdown list.",
            errorStyle="warning",
        )
        dv.add(f"{col_letter}{DATA_START_ROW}:{col_letter}{VALIDATION_ROW_END}")
        ws.add_data_validation(dv)


def _row_bg(row_idx):
    """Return the alternating background color for a given absolute row index."""
    return ALT_ROW_COLOR if (row_idx - DATA_START_ROW) % 2 == 1 else "FFFFFF"


def _write_issue_row(ws, row_idx, row_record):
    """
    Write one (issue, sprint) entry into columns A–G of the given row.
    A=Sprint(int), B=Assignee, C=Type, D=Scope, E=empty, F=Jira#(hyperlink), G=Status
    row_record = {"issue": issue_obj, "sprint_num": int, "scope": str}
    """
    issue      = row_record["issue"]
    sprint_num = row_record["sprint_num"]
    scope      = row_record["scope"]

    fields = issue["fields"]
    key    = issue["key"]

    full_name    = fields["assignee"]["displayName"] if fields.get("assignee") else ""
    assignee     = _assignee_nickname(full_name)
    issue_type   = get_type(issue)
    ticket_url   = f"{BASE_URL}/browse/{key}"
    ticket_label = f"[{key}] {fields.get('summary', '')}"
    desc_text       = _adf_to_text(fields.get("description") or "")
    has_description = "Y" if len(desc_text) > 20 else "N"
    status          = get_status(issue)

    # Column order: A        B         C           D           E              F         G
    values = [sprint_num, assignee, issue_type, scope, has_description, ticket_label, status]

    bg         = _row_bg(row_idx)
    base_align = Alignment(horizontal="left", vertical="center")

    for col, value in enumerate(values, start=1):
        cell           = ws.cell(row=row_idx, column=col, value=value)
        cell.alignment = base_align

        if col == 6:  # F — Jira # (hyperlinked ticket title, blue underline)
            cell.fill      = PatternFill(fill_type="solid", fgColor=bg)
            cell.font      = Font(name="Calibri", size=11, color="1155CC", underline="single")
            cell.hyperlink = ticket_url

        elif col == 7:  # G — Status (bold Calibri 11, alternating background — no status color)
            cell.fill      = PatternFill(fill_type="solid", fgColor=bg)
            cell.font      = Font(name="Calibri", size=11, bold=True)
            cell.hyperlink = None

        elif col == 3:  # C — Type (colored font for Bug/Support; default otherwise)
            cell.fill      = PatternFill(fill_type="solid", fgColor=bg)
            type_color     = TYPE_FONT_COLORS.get(issue_type)
            cell.font      = Font(name="Calibri", size=11, color=type_color) if type_color \
                             else Font(name="Calibri", size=11)
            cell.hyperlink = None

        else:  # A, B, D, E — plain text with alternating background
            cell.fill      = PatternFill(fill_type="solid", fgColor=bg)
            cell.font      = Font(name="Calibri", size=11)
            cell.hyperlink = None

    # Column I (9) — latest xlsx<...>xlsx Jira comment; skipped when None (leaves existing value)
    jira_note = row_record.get("jira_note")
    if jira_note is not None:
        cell_i           = ws.cell(row=row_idx, column=9, value=jira_note)
        cell_i.alignment = base_align
        cell_i.fill      = PatternFill(fill_type="solid", fgColor=bg)
        cell_i.font      = Font(name="Calibri", size=11)
        cell_i.hyperlink = None


def _apply_alternating_fills(ws):
    """
    Re-apply position-based alternating fills to columns A–J (1–10) after a sort.
    K, L, M are unused and left blank. Only fill is changed; H–J values/fonts untouched.
    """
    for row_idx in range(DATA_START_ROW, ws.max_row + 1):
        if not any(ws.cell(row=row_idx, column=c).value is not None for c in range(1, 11)):
            continue  # skip rows with no content in A–J

        bg = _row_bg(row_idx)

        for col in range(1, 11):  # A through J (col 10); K+ are unused and left blank
            cell        = ws.cell(row=row_idx, column=col)
            cell.fill   = PatternFill(fill_type="solid", fgColor=bg)
            cell.border = _WHITE_COL_BORDER if (bg == "FFFFFF" and col > 1) else _EMPTY_BORDER


def _sort_data_rows(ws, api_map):
    max_row = ws.max_row
    if max_row < DATA_START_ROW:
        return

    max_col   = ws.max_column
    rows_data = []

    for row_idx in range(DATA_START_ROW, max_row + 1):
        cells = []
        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cells.append({
                "value":     cell.value,
                "hyperlink": cell.hyperlink.target if cell.hyperlink else None,
                "font":      copy(cell.font),
                "fill":      copy(cell.fill),
                "alignment": copy(cell.alignment),
            })

        # col F (0-based index 5) = "[LDT-XXXX] Summary"; col A (index 0) = sprint number
        key = _extract_key(cells[5]["value"])
        try:
            sprint_num = int(cells[0]["value"]) if cells[0]["value"] is not None else None
        except (ValueError, TypeError):
            sprint_num = None

        composite = (key, sprint_num) if key and sprint_num is not None else None

        if composite and composite in api_map:
            sort_key = _issue_sort_key(api_map[composite])
        else:
            sn       = sprint_num if sprint_num is not None else 9999
            assignee = cells[1]["value"] or ""
            # Cell holds the nickname; _ASSIGNEE_TO_TEAM_IDX handles both full names and
            # nicknames so manually entered or old rows still sort into the right team bucket.
            team_idx = _ASSIGNEE_TO_TEAM_IDX.get(assignee, len(TEAMS))
            sort_key = (sn, team_idx, assignee)

        rows_data.append((sort_key, cells))

    rows_data.sort(key=lambda x: x[0])

    for row_idx, (_, cells) in enumerate(rows_data, start=DATA_START_ROW):
        for col_idx, cd in enumerate(cells, start=1):
            cell           = ws.cell(row=row_idx, column=col_idx)
            cell.value     = cd["value"]
            cell.font      = cd["font"]
            cell.fill      = cd["fill"]
            cell.alignment = cd["alignment"]
            cell.hyperlink = cd["hyperlink"]

    # Fills are position-dependent; re-apply after rows have moved
    _apply_alternating_fills(ws)

def _ensure_table_in_model(ws, new_ref: str) -> None:
    """Update the existing table ref, or create a minimal Table if none exists."""
    if ws.tables:
        for tbl in ws.tables.values():
            tbl.ref = new_ref
    else:
        tbl = Table(displayName=TABLE_DISPLAY_NAME, ref=new_ref)
        tbl.tableStyleInfo = TableStyleInfo(name=TABLE_STYLE, showRowStripes=True)
        ws.add_table(tbl)


def _generate_table_xml(ref: str, tbl_id: int, name: str, display_name: str, style: str,
                        col_names: list) -> bytes:
    """Build a complete, valid OOXML table definition for the given ref and column list."""
    cols_xml = "".join(
        f'<tableColumn id="{i + 1}" name="{col}"/>'
        for i, col in enumerate(col_names)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        f' id="{tbl_id}" name="{name}" displayName="{display_name}"'
        f' ref="{ref}" totalsRowShown="0">'
        f'<autoFilter ref="{ref}"/>'
        f'<tableColumns count="{len(col_names)}">{cols_xml}</tableColumns>'
        f'<tableStyleInfo name="{style}" showFirstColumn="0" showLastColumn="0"'
        ' showRowStripes="1" showColumnStripes="0"/>'
        '</table>'
    ).encode("utf-8")


def _patch_xlsx_table(xlsx_path: str, final_row: int, col_names: list) -> None:
    """
    After openpyxl saves, replace its table XML with a clean generated version.
    Reads the table id/name/style from what openpyxl wrote so we don't change
    identifiers, then overwrites only that entry in the zip.
    col_names must match the actual row-3 header values for every table column.
    """
    ref = f"A3:J{final_row}"

    with zipfile.ZipFile(xlsx_path, "r") as zin:
        table_paths = [n for n in zin.namelist()
                       if n.startswith("xl/tables/") and n.endswith(".xml")]
        if not table_paths:
            return  # openpyxl wrote no table; nothing to patch

        table_path = table_paths[0]
        raw_xml    = zin.read(table_path).decode("utf-8", errors="replace")
        entries    = [(info, zin.read(info.filename)) for info in zin.infolist()]

    m_id      = re.search(r'\bid="(\d+)"',                   raw_xml)
    m_name    = re.search(r'<table\b[^>]+\bname="([^"]+)"',  raw_xml)
    m_display = re.search(r'\bdisplayName="([^"]+)"',         raw_xml)
    m_style   = re.search(r'<tableStyleInfo[^>]+\bname="([^"]+)"', raw_xml)

    clean_xml = _generate_table_xml(
        ref,
        tbl_id      = int(m_id.group(1))    if m_id      else 1,
        name        = m_name.group(1)        if m_name    else TABLE_DISPLAY_NAME,
        display_name= m_display.group(1)     if m_display else TABLE_DISPLAY_NAME,
        style       = m_style.group(1)       if m_style   else TABLE_STYLE,
        col_names   = col_names,
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in entries:
            zout.writestr(info.filename, clean_xml if info.filename == table_path else data)

    with open(xlsx_path, "wb") as f:
        f.write(buf.getvalue())


# ── Main update logic ─────────────────────────────────────────────────────────

def update_excel(input_path, issues):
    if not os.path.exists(input_path):
        raise FileNotFoundError(
            f"Required file not found: '{input_path}'. "
            "Ensure 'Leahi Sprint Management.xlsx' exists in the working directory."
        )

    wb = load_workbook(input_path)

    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(
            f"Required sheet '{SHEET_NAME}' not found in '{input_path}'. "
            f"Available sheets: {wb.sheetnames}"
        )

    ws = wb[SHEET_NAME]

    _ensure_sheet_setup(ws)
    _ensure_selection_sprints(wb)
    _ensure_data_validations(ws)

    # Build api_map: (jira_key, sprint_num) → row_record
    # Each issue expands into one row_record per sprint it has ever been in.
    api_map = {}
    for issue in issues:
        for sprint_num, scope in get_sprint_history(issue):
            composite        = (issue["key"], sprint_num)
            api_map[composite] = {"issue": issue, "sprint_num": sprint_num, "scope": scope, "jira_note": issue.get("_jira_note")}

    # Build existing map: (jira_key, sprint_num) → [row_idx, ...]
    # Using col F for the Jira key and col A for the sprint number.
    # A list per composite handles accidental duplicates from prior edits.
    existing = {}
    for row_idx in range(DATA_START_ROW, ws.max_row + 1):
        label      = ws.cell(row=row_idx, column=6).value  # col F
        sprint_val = ws.cell(row=row_idx, column=1).value  # col A
        key        = _extract_key(label)
        try:
            sprint_num = int(sprint_val) if sprint_val is not None else None
        except (ValueError, TypeError):
            sprint_num = None
        if key and sprint_num is not None:
            composite = (key, sprint_num)
            existing.setdefault(composite, []).append(row_idx)

    # Update the first (topmost) occurrence of each composite that is still in the API
    updated = set()
    for composite, row_indices in existing.items():
        if composite in api_map:
            _write_issue_row(ws, row_indices[0], api_map[composite])
            updated.add(composite)

    # Delete stale composites (gone from API) and duplicate rows (beyond the first).
    # Always delete in reverse row order so indices of earlier rows are not shifted.
    removed_count = sum(1 for c in existing if c not in api_map)
    to_delete = []
    for composite, row_indices in existing.items():
        if composite not in api_map:
            to_delete.extend(row_indices)        # all rows for a gone (key, sprint)
        else:
            to_delete.extend(row_indices[1:])    # duplicate rows beyond the first
    to_delete.sort(reverse=True)
    for row_idx in to_delete:
        ws.delete_rows(row_idx)

    # Append rows for (key, sprint) pairs not yet in the sheet
    new_composites = [c for c in api_map if c not in existing]
    next_row = ws.max_row + 1
    for composite in new_composites:
        _write_issue_row(ws, next_row, api_map[composite])
        next_row += 1

    _sort_data_rows(ws, api_map)

    # Remove trailing rows that carry fills/borders but no ticket data.
    # openpyxl counts formatted-but-empty cells toward max_row, so these rows
    # accumulate across runs unless explicitly deleted.
    for row_idx in range(ws.max_row, DATA_START_ROW - 1, -1):
        if any(ws.cell(row=row_idx, column=c).value is not None for c in range(1, 8)):
            break
        ws.delete_rows(row_idx)

    # Ensure the table covers every row written in this run. _ensure_table_in_model
    # creates the Table object if the file was previously corrupted and Excel had
    # removed it, so openpyxl always writes the worksheet/rels table references.
    # _patch_xlsx_table then replaces openpyxl's generated table XML with a clean
    # version, preventing the Excel repair dialog.
    if ws.max_row >= DATA_START_ROW:
        _ensure_table_in_model(ws, f"A3:J{ws.max_row}")

    # Read the actual row-3 header values for all 10 table columns (A–J).
    # H–J headers are user-maintained so we take them from the sheet rather than
    # hardcoding, ensuring the table XML always matches what is visible in the file.
    table_col_names = [
        ws.cell(row=3, column=c).value or f"Column{c}"
        for c in range(1, 11)
    ]

    final_row = ws.max_row
    wb.save(input_path)
    _patch_xlsx_table(input_path, final_row, table_col_names)

    print(
        f"Saved: {input_path} "
        f"({len(api_map)} sprint-rows across {len(issues)} tickets | "
        f"{len(updated)} updated, {len(new_composites)} added, {removed_count} removed)"
    )


def _is_file_locked(path: str) -> bool:
    try:
        with open(path, "r+b"):
            return False
    except (PermissionError, OSError):
        return True


def _sync_hplus(source_path: str, dest_path: str) -> None:
    """
    Read H–J note values from source (SharePoint) and write them into dest (X: drive),
    matched by (jira_key, sprint_num). SharePoint is the source of truth — even empty
    cells in source overwrite dest so deleted notes don't linger on the X: drive copy.
    """
    wb_src = load_workbook(source_path, data_only=True)
    ws_src = wb_src[SHEET_NAME]
    hplus_map = {}
    for row_idx in range(DATA_START_ROW, ws_src.max_row + 1):
        key = _extract_key(ws_src.cell(row=row_idx, column=6).value)
        try:
            sprint_num = int(ws_src.cell(row=row_idx, column=1).value)
        except (TypeError, ValueError):
            sprint_num = None
        if key and sprint_num is not None:
            hplus_map[(key, sprint_num)] = tuple(
                ws_src.cell(row=row_idx, column=c).value for c in SP_SYNC_COLS
            )

    wb_dst = load_workbook(dest_path)
    ws_dst = wb_dst[SHEET_NAME]
    for row_idx in range(DATA_START_ROW, ws_dst.max_row + 1):
        key = _extract_key(ws_dst.cell(row=row_idx, column=6).value)
        try:
            sprint_num = int(ws_dst.cell(row=row_idx, column=1).value)
        except (TypeError, ValueError):
            sprint_num = None
        if key and sprint_num is not None:
            vals = hplus_map.get((key, sprint_num))
            if vals is not None:
                for col_idx, val in zip(SP_SYNC_COLS, vals):
                    ws_dst.cell(row=row_idx, column=col_idx).value = val
    wb_dst.save(dest_path)
    print(f"H/J notes synced from SharePoint → X: drive ({len(hplus_map)} rows read)")


def main():
    parser = argparse.ArgumentParser(
        description="Update Leahi Sprint Management Excel from Jira."
    )
    parser.add_argument(
        "input",
        nargs="?",
        default=os.getenv("EXCEL_PATH"),
        help="Path to the Excel file",
    )
    args = parser.parse_args()

    if not SHAREPOINT_SYNC_PATH:
        raise RuntimeError(
            "SHAREPOINT_SYNC_PATH is not set in .env. "
            "Set it to the Linux path of your OneDrive-synced SharePoint file."
        )

    if _is_file_locked(SHAREPOINT_SYNC_PATH):
        raise RuntimeError(
            "SharePoint file is currently open or locked — close it in Excel and re-run."
        )

    print("Fetching issues from Jira...")
    issues = get_jira_issues()
    print(f"Found {len(issues)} issues")
    print("Updating Sprint Template on X: drive...")
    update_excel(args.input, issues)
    print("Syncing H-J notes from SharePoint to X: drive...")
    _sync_hplus(SHAREPOINT_SYNC_PATH, args.input)


if __name__ == "__main__":
    main()
