import argparse
import io
import os
import re
import smtplib
import traceback
import zipfile
import requests
from copy import copy
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import datetime

load_dotenv()

BASE_URL = "https://diality.atlassian.net"
PROJECT = "LDT"
SHEET_NAME = "Sprint Template"
DATA_START_ROW = 4  # Row 1=note banner, Row 2=sprint info, Row 3=headers, Row 4+=data

SWVV_TEAM = ["Raghu Kallala", "Thomas Lippold", "Tejaskumar Patel", "Tiffany Mejia",
             "Zoltan Miskolci", "Sarina Cheung", "Tisha Patel", "Ethan Nguyen"]
FW_TEAM   = ["Arpita Srivastava", "Jashwant Gantyada", "Michael Garthwaite", "Praneeth Bunne",
             "Sameer Poyil", "Varshini Nagabooshanam", "Vijay Pamula", "Suresh Dharnala"]
SW_TEAM   = ["Nicholas Ramirez", "Stephen Quong", "Dara Navaei",
             "Eliza Petersen", "Christina Heine", "Caitlynn Chang"]
TEAMS = [SWVV_TEAM, FW_TEAM, SW_TEAM]

# Full Jira display name → nickname used in the Selection sheet dropdowns.
# Non-obvious entries are listed first; straightforward first-name entries follow.
NICKNAME_MAP = {
    "Tejaskumar Patel":       "Tejas",
    "Nicholas Ramirez":       "Nico",
    "Thomas Lippold":         "Tom",
    "Jashwant Gantyada":      "Jaswant",       # Jira spells with 'h'; Selection does not
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
}

# Lookup from assignee name (full Jira OR nickname) → team sort index.
# Populated at import time so the sort fallback path works when only the cell
# nickname is available (the api_map lookup path uses the full Jira name directly).
_ASSIGNEE_TO_TEAM_IDX: dict = {}
for _ti, _team in enumerate(TEAMS):
    for _full in _team:
        _ASSIGNEE_TO_TEAM_IDX[_full] = _ti
        _nick = NICKNAME_MAP.get(_full)
        if _nick:
            _ASSIGNEE_TO_TEAM_IDX[_nick] = _ti

NOTIFICATION_EMAILS = [os.getenv("JIRA_EMAIL")]

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



# ── Jira fetch ────────────────────────────────────────────────────────────────

def _fetch_paginated(url, jql, auth):
    all_issues = []
    start_at = 0
    while True:
        params = {
            "fields": "summary,status,assignee,issuetype,customfield_10020,description",
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
    """Return True if the changelog contains at least one Sprint field transition."""
    for history in issue["changelog"]["histories"]:
        for item in history["items"]:
            if item["field"] == "Sprint":
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


def get_jira_issues():
    auth = HTTPBasicAuth(os.getenv("JIRA_EMAIL"), os.getenv("JIRA_API_TOKEN"))

    board_issues = _fetch_paginated(
        f"{BASE_URL}/rest/agile/1.0/board/84/issue",
        "sprint in openSprints() OR sprint in closedSprints() OR sprint in futureSprints()",
        auth,
    )

    seen = {i["key"] for i in board_issues}

    # Fetch every backlog issue and keep only those with sprint changelog entries.
    # This captures tickets that were fully removed from all sprints at any point in history,
    # not just those removed from the currently active sprint.
    all_backlog = _fetch_paginated(
        f"{BASE_URL}/rest/agile/1.0/board/84/backlog",
        f"project = {PROJECT}",
        auth,
    )
    backlog_with_history = [
        i for i in all_backlog
        if i["key"] not in seen and _has_sprint_history(i)
    ]

    return board_issues + backlog_with_history


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
    """
    Return a sorted list of (sprint_num, scope) for every sprint this issue has
    ever been in — one entry per sprint number, last-state-wins for repeated cycles.

    Scope decision rule:
      - Sprint is currently in customfield_10020  → ticket was never manually removed
                                                    → scope = _classify_sprint_timing
      - Sprint seen ONLY in the changelog         → ticket was manually removed from it
                                                    → scope = "Move out of Sprint"
    """
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
            if num >= 0:
                sprint_entries[num] = "Move out of Sprint"

    # Pass 2: currently-held sprints → timing classification overwrites any conflict
    for name, sprint_obj in current_sprint_objs.items():
        num = _sprint_num(name)
        if num >= 0:
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

def _send_notification(subject, body):
    if not NOTIFICATION_EMAILS:
        return
    smtp_host     = os.getenv("SMTP_HOST")
    smtp_port     = int(os.getenv("SMTP_PORT", 587))
    smtp_user     = os.getenv("SMTP_USER")
    smtp_password = os.getenv("SMTP_PASSWORD")
    smtp_from     = os.getenv("SMTP_FROM")
    if not all([smtp_host, smtp_user, smtp_password, smtp_from]):
        print("Warning: SMTP env vars not fully configured, skipping notification.")
        return
    msg = MIMEMultipart()
    msg["From"]    = smtp_from
    msg["To"]      = ", ".join(NOTIFICATION_EMAILS)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))
    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_from, NOTIFICATION_EMAILS, msg.as_string())


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
    """
    Set fonts for rows 1-3, freeze panes at A4, and column widths.
    Row 1 → Calibri 14 (note banner). Row 2 → Calibri 11 (sprint info).
    Row 3 headers A–G → Calibri 11 bold, golden yellow fill.
    """
    # Row 1: note banner — Calibri 14, preserve existing bold per cell
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        if cell.value is not None:
            cell.font = Font(name="Calibri", size=14, bold=bool(cell.font.bold))

    # Row 2: sprint info — Calibri 11, preserve existing bold per cell
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=2, column=col)
        if cell.value is not None:
            cell.font = Font(name="Calibri", size=11, bold=bool(cell.font.bold))

    header_fill = PatternFill(fill_type="solid", fgColor="FFC000")  # golden yellow
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
    has_description = "Yes" if fields.get("description") else "No"
    status       = get_status(issue)

    # Column order: A    B         C           D      E     F             G
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
            ws.cell(row=row_idx, column=col).fill = PatternFill(fill_type="solid", fgColor=bg)


def _sort_data_rows(ws, api_map):
    """
    Sort all data rows (row 4+) by (sprint_num, team_idx, assignee).
    api_map is keyed by (jira_key, sprint_num) — the composite row identity.
    Reads every column so that manually maintained H+ data moves with its row.
    Re-applies alternating fills after sorting.
    """
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


# ── Table XML preservation ────────────────────────────────────────────────────
# openpyxl regenerates table XML on save using only the attributes it knows,
# silently dropping others. A snapshot-and-restore strategy has a silent failure
# mode: if the source file is already corrupt (e.g. Excel previously repaired it
# by removing the table), the snapshot is empty and the restore is a no-op.
#
# Instead we let openpyxl write the table so the worksheet XML and relationship
# files stay correct, then replace only xl/tables/table1.xml with a freshly
# generated version that is always structurally valid.

def _ensure_table_in_model(ws, new_ref: str) -> None:
    """Update the existing table ref, or create a minimal Table if none exists."""
    if ws.tables:
        for tbl in ws.tables.values():
            tbl.ref = new_ref
    else:
        tbl = Table(displayName=TABLE_DISPLAY_NAME, ref=new_ref)
        tbl.tableStyleInfo = TableStyleInfo(name=TABLE_STYLE, showRowStripes=True)
        ws.add_table(tbl)


def _generate_table_xml(ref: str, tbl_id: int, name: str, display_name: str, style: str) -> bytes:
    """Build a complete, valid OOXML table definition for the given ref."""
    cols_xml = "".join(
        f'<tableColumn id="{i + 1}" name="{col}"/>'
        for i, col in enumerate(SHEET_HEADERS)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        f' id="{tbl_id}" name="{name}" displayName="{display_name}"'
        f' ref="{ref}" totalsRowShown="0">'
        f'<autoFilter ref="{ref}"/>'
        f'<tableColumns count="{len(SHEET_HEADERS)}">{cols_xml}</tableColumns>'
        f'<tableStyleInfo name="{style}" showFirstColumn="0" showLastColumn="0"'
        ' showRowStripes="1" showColumnStripes="0"/>'
        '</table>'
    ).encode("utf-8")


def _patch_xlsx_table(xlsx_path: str, final_row: int) -> None:
    """
    After openpyxl saves, replace its table XML with a clean generated version.
    Reads the table id/name/style from what openpyxl wrote so we don't change
    identifiers, then overwrites only that entry in the zip.
    """
    ref = f"A3:G{final_row}"

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

    # Build api_map: (jira_key, sprint_num) → row_record
    # Each issue expands into one row_record per sprint it has ever been in.
    api_map = {}
    for issue in issues:
        for sprint_num, scope in get_sprint_history(issue):
            composite        = (issue["key"], sprint_num)
            api_map[composite] = {"issue": issue, "sprint_num": sprint_num, "scope": scope}

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

    # Ensure the table covers every row written in this run. _ensure_table_in_model
    # creates the Table object if the file was previously corrupted and Excel had
    # removed it, so openpyxl always writes the worksheet/rels table references.
    # _patch_xlsx_table then replaces openpyxl's generated table XML with a clean
    # version, preventing the Excel repair dialog.
    if ws.max_row >= DATA_START_ROW:
        _ensure_table_in_model(ws, f"A3:G{ws.max_row}")

    final_row = ws.max_row
    wb.save(input_path)
    _patch_xlsx_table(input_path, final_row)

    print(
        f"Saved: {input_path} "
        f"({len(api_map)} sprint-rows across {len(issues)} tickets | "
        f"{len(updated)} updated, {len(new_composites)} added, {removed_count} removed)"
    )


def main():
    parser = argparse.ArgumentParser(
        description="Update Leahi Sprint Management Excel from Jira."
    )
    parser.add_argument(
        "input",
        nargs="?",
        default="Leahi Sprint Management.xlsx",
        help="Path to the Excel file (default: 'Leahi Sprint Management.xlsx')",
    )
    args = parser.parse_args()

    try:
        print("Fetching issues from Jira...")
        issues = get_jira_issues()
        print(f"Found {len(issues)} issues")
        print("Updating Excel...")
        update_excel(args.input, issues)
        print("Done!")
        _send_notification(
            "Sprint Management: Success",
            f"Script ran successfully.\n{len(issues)} tickets processed.",
        )
    except Exception:
        _send_notification(
            "Sprint Management: Failed",
            f"Script failed with the following error:\n\n{traceback.format_exc()}",
        )
        raise


if __name__ == "__main__":
    main()
