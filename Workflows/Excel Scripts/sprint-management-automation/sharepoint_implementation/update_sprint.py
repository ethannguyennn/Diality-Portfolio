"""Update Leahi Sprint Management on SharePoint from Jira API.

Uses the Graph Workbook API for incremental edits so the file can be
updated even while someone has it open in Excel Online.
"""

import os
import re
from datetime import datetime
from typing import Optional

import requests
from dotenv import load_dotenv
from requests.auth import HTTPBasicAuth

from config import ARCHIVE_PATH, SPRINT_TEMPLATE_FILE
from sharepoint_helper import SharePointHelper

load_dotenv()

# ── Jira ──
BASE_URL = "https://diality.atlassian.net"
PROJECT = "LDT"

# ── Sheet layout ──
SHEET_NAME = "Sprint Template"
DATA_START_ROW = 4
NUM_COLS = 12

# ── Teams ──
SWVV_TEAM = [
    "Raghu Kallala", "Thomas Lippold", "Tejaskumar Patel", "Tiffany Mejia",
    "Zoltan Miskolci", "Sarina Cheung", "Tisha Patel", "Ethan Nguyen",
]
FW_TEAM = [
    "Arpita Srivastava", "Jashwant Gantyada", "Michael Garthwaite", "Praneeth Bunne",
    "Sameer Poyil", "Varshini Nagabooshanam", "Vijay Pamula", "Suresh Dharnala",
    "Santhos Kumar Reddy", "Vinayakam Mani", "Sean Nash",
]
SW_TEAM = ["Nicholas Ramirez", "Stephen Quong", "Dara Navaei", "Behrouz NematiPour"]
SYS_TEAM = [
    "Eliza Petersen", "Caitlynn Chang", "Christina Heine", "Abhijit Barman",
    "Chris Yu", "Vitas Buenaventura", "Emiline Hernandez",
]
TEAMS = [FW_TEAM, SWVV_TEAM, SW_TEAM, SYS_TEAM]
TEAM_NAMES = ["FW", "SWVV", "SW", "SYS"]
_SYS_TEAM_IDX = len(TEAMS) - 1
_UNMATCHED_TEAM_IDX = len(TEAMS)

NICKNAME_MAP = {
    "Tejaskumar Patel": "Tejas",
    "Nicholas Ramirez": "Nico",
    "Thomas Lippold": "Tom",
    "Jashwant Gantyada": "Jashwant",
    "Behrouz NematiPour": "Behrouz",
    "Sean Nash": "Sean",
    "Vinayakam Mani": "Vinay",
    "Raghu Kallala": "Raghu",
    "Tiffany Mejia": "Tiffany",
    "Zoltan Miskolci": "Zoltan",
    "Sarina Cheung": "Sarina",
    "Tisha Patel": "Tisha",
    "Ethan Nguyen": "Ethan",
    "Arpita Srivastava": "Arpita",
    "Michael Garthwaite": "Michael",
    "Praneeth Bunne": "Praneeth",
    "Sameer Poyil": "Sameer",
    "Varshini Nagabooshanam": "Varshini",
    "Vijay Pamula": "Vijay",
    "Suresh Dharnala": "Suresh",
    "Dara Navaei": "Dara",
    "Eliza Petersen": "Eliza",
    "Stephen Quong": "Stephen",
    "Christina Heine": "Christina",
    "Caitlynn Chang": "Caitlynn",
    "Santhos Kumar Reddy": "Santhos",
    "Abhijit Barman": "Abhijit",
    "Chris Yu": "Chris",
    "Vitas Buenaventura": "Vitas",
    "Emiline Hernandez": "Emiline",
}

# Build a fast lookup from full name or nickname to team index.
_ASSIGNEE_TO_TEAM_IDX: dict = {}
for _ti, _team in enumerate(TEAMS):
    for _full in _team:
        _ASSIGNEE_TO_TEAM_IDX[_full] = _ti
        _nick = NICKNAME_MAP.get(_full)
        if _nick:
            _ASSIGNEE_TO_TEAM_IDX[_nick] = _ti
del _ti, _team, _full, _nick

# ── Type / status constants ──
TYPE_FONT_COLORS = {
    "Bug": "FF0000", "Dialin ticket": "4472C4",
    "Support": "00B050", "Little-V": "7030A0", "Big-V": "000000",
}
BUG_DIALIN_TYPES = {"Bug", "Dialin ticket"}
BUG_STATUS_MAP = {
    "More Info Needed": "More Info Needed",
    "On Hold": "On Hold",
    "Blocked": "Blocked",
    "Ready for SW/FW Implementation": "Ready for SW/FW Implementation",
    "Investigation In Progress (SW/FW)": "Investigation in Progress",
    "Fix In Progress": "Fix in Progress",
    "Ready for Fix to be Tested": "Ready for Fix to be Tested",
    "Testing Fix in Progress": "Testing Fix in Progress",
    "Test Passed - Ready for Updated Release": "Ready for Updated Release",
    "Test Failed - SW/FW Fix Needed": "SW/FW Fix Needed",
    "Test Passed - Needs Final Fix Details": "Needs Final Fix Details",
    "Done": "Done",
}
MIN_SPRINT = 28
JIRA_NOTE_RE = re.compile(r'xlsx<([^>]*)>xlsx')
MOVED_OUT = "Move out of Sprint"

# ── Workbook API constants ──
_WS = f"worksheets('{SHEET_NAME}')"
_LAST_COL = chr(ord('A') + NUM_COLS - 1)  # 'L'


# ═══════════════════════════════════════════════════════════════════
#  Jira helpers
# ═══════════════════════════════════════════════════════════════════

def _fetch_paginated(url, jql, auth):
    # Fetch all issues from a paginated Jira REST API endpoint.
    all_issues, start_at = [], 0
    while True:
        params = {
            "fields": "summary,status,assignee,issuetype,customfield_10020,description,comment",
            "expand": "changelog", "jql": jql, "maxResults": 100, "startAt": start_at,
        }
        resp = requests.get(url, params=params, auth=auth)
        resp.raise_for_status()
        data = resp.json()
        all_issues.extend(data["issues"])
        start_at += len(data["issues"])
        if start_at >= data.get("total", 0) or not data["issues"]:
            break
    return all_issues


def _sprint_names(s):
    # Parse a comma-separated sprint name string into a set of names.
    return {n.strip() for n in s.split(",") if n.strip()} if s else set()


def _sprint_num(name):
    # Extract the integer sprint number from a sprint name string.
    m = re.search(r'(\d+)', name)
    return int(m.group(1)) if m else -1


def _has_sprint_history(issue):
    # Return True if the issue has appeared in any sprint >= MIN_SPRINT.
    for history in issue["changelog"]["histories"]:
        for item in history["items"]:
            if item["field"] != "Sprint":
                continue
            for name in _sprint_names(item.get("toString") or ""):
                if _sprint_num(name) >= MIN_SPRINT:
                    return True
    return False


def _assignee_nickname(display_name: str) -> str:
    # Return the configured nickname for a display name, or the first name.
    if not display_name:
        return "Unassigned"
    return NICKNAME_MAP.get(display_name, display_name.split()[0])


def _team_for_nickname(nickname):
    # Look up the (team_name, team_index) pair for the given nickname.
    idx = _ASSIGNEE_TO_TEAM_IDX.get(nickname)
    if idx is None:
        return None, _UNMATCHED_TEAM_IDX
    return TEAM_NAMES[idx], idx


def _adf_to_text(node) -> str:
    # Recursively extract plain text from an Atlassian Document Format node.
    if isinstance(node, str):
        return node
    if not isinstance(node, dict):
        return ""
    if node.get("type") == "text":
        return node.get("text", "")
    return "".join(_adf_to_text(c) for c in node.get("content", []))


def _latest_jira_note(issue, auth) -> Optional[str]:
    # Return the most recent xlsx-tagged note text from the issue's comments.
    comment_field = issue["fields"].get("comment") or {}
    comments = list(comment_field.get("comments", []))
    if comment_field.get("total", len(comments)) > len(comments):
        url = f"{BASE_URL}/rest/api/2/issue/{issue['key']}/comment"
        comments = requests.get(url, params={"maxResults": 1000}, auth=auth).json().get("comments", [])

    best_text, best_created = None, ""
    for comment in comments:
        body_text = _adf_to_text(comment.get("body") or "")
        m = JIRA_NOTE_RE.search(body_text)
        if m and comment.get("created", "") >= best_created:
            best_created = comment["created"]
            best_text = m.group(1).strip()
    return best_text


def get_jira_issues():
    """Fetch all relevant issues from the Jira board and backlog.

    Collects issues from open and future sprints on board 84, plus
    any backlog items that have sprint history back to MIN_SPRINT.
    Attaches the most recent xlsx-tagged Jira note to each issue.
    """
    auth = HTTPBasicAuth(os.getenv("JIRA_EMAIL"), os.getenv("JIRA_API_TOKEN"))

    board_issues = _fetch_paginated(
        f"{BASE_URL}/rest/agile/1.0/board/84/issue",
        "sprint in openSprints() OR sprint in futureSprints()", auth,
    )
    seen = {i["key"] for i in board_issues}

    all_backlog = _fetch_paginated(
        f"{BASE_URL}/rest/agile/1.0/board/84/backlog",
        f"project = {PROJECT}", auth,
    )
    backlog_with_history = [
        i for i in all_backlog if i["key"] not in seen and _has_sprint_history(i)
    ]

    all_issues = board_issues + backlog_with_history
    for issue in all_issues:
        issue["_jira_note"] = _latest_jira_note(issue, auth)
    return all_issues


def parse_date(date_str):
    """Parse an ISO 8601 date string into a timezone-aware datetime.

    Returns None if date_str is empty or None.
    """
    if not date_str:
        return None
    date_str = date_str.replace("Z", "+00:00")
    date_str = re.sub(r'([+-])(\d{2})(\d{2})$', r'\1\2:\3', date_str)
    return datetime.fromisoformat(date_str)


def get_type(issue):
    """Return the display issue type string for the given Jira issue.

    Classifies the issue as one of: Bug, Dialin ticket, Little-V,
    Support, or Big-V based on issuetype name and summary text.
    """
    issuetype = issue["fields"]["issuetype"]["name"].lower()
    summary = issue["fields"].get("summary", "").lower()
    if "dialin" in issuetype or "dialin" in summary:
        return "Dialin ticket"
    if issuetype == "bug":
        return "Bug"
    if "little" in issuetype:
        return "Little-V"
    if "support" in summary:
        return "Support"
    return "Big-V"


def get_status(issue, issue_type) -> str:
    """Return the display status string for the given Jira issue.

    Applies BUG_STATUS_MAP for Bug and Dialin ticket issue types.
    For all other types, maps Jira status categories to simplified
    display strings such as 'Done', 'In-progress', or 'Not started'.
    """
    status = issue["fields"]["status"]
    status_name = status["name"]
    if issue_type in BUG_DIALIN_TYPES:
        return BUG_STATUS_MAP.get(status_name, status_name)

    status_name = status_name.lower()
    category_key = status["statusCategory"]["key"]
    if "blocked" in status_name:
        return "Blocked"
    if "on hold" in status_name or "on-hold" in status_name:
        return "On Hold"
    if status_name == "n/a":
        return "N/A"
    if category_key == "done" or "done" in status_name or "complete" in status_name:
        return "Done"
    if category_key == "new":
        return "Not started"
    return "In-progress"


# ═══════════════════════════════════════════════════════════════════
#  Sprint history / timing
# ═══════════════════════════════════════════════════════════════════

def _classify_sprint_timing(sprint, histories):
    # Return 'Added' if issue was added after sprint start, else 'Scoped'.
    sprint_name = sprint["name"]
    sprint_start_dt = parse_date(sprint.get("startDate"))
    if not sprint_start_dt:
        return "Scoped"

    most_recent_add = None
    for history in histories:
        for item in history["items"]:
            if item["field"] != "Sprint":
                continue
            to_sprints = _sprint_names(item.get("toString") or "")
            from_sprints = _sprint_names(item.get("fromString") or "")
            if sprint_name in to_sprints and sprint_name not in from_sprints:
                added_dt = parse_date(history["created"])
                if added_dt and (most_recent_add is None or added_dt > most_recent_add):
                    most_recent_add = added_dt

    if most_recent_add is None:
        return "Scoped"
    sprint_start_eod = sprint_start_dt.replace(hour=23, minute=59, second=59, microsecond=0)
    return "Added" if most_recent_add > sprint_start_eod else "Scoped"


def get_sprint_history(issue, max_sprint_num):
    """Return a sorted list of (sprint_num, scope) pairs for the issue.

    Scope values are 'Scoped', 'Added', or MOVED_OUT.  Only sprint
    numbers in the range [MIN_SPRINT, max_sprint_num] are included.
    """
    histories = issue["changelog"]["histories"]
    sprint_info = issue["fields"].get("customfield_10020") or []
    current_sprint_names = {s["name"] for s in sprint_info}
    current_sprint_objs = {s["name"]: s for s in sprint_info}

    all_sprint_names = set(current_sprint_names)
    for history in histories:
        for item in history["items"]:
            if item["field"] != "Sprint":
                continue
            for name in _sprint_names(item.get("fromString") or ""):
                all_sprint_names.add(name)
            for name in _sprint_names(item.get("toString") or ""):
                all_sprint_names.add(name)

    sprint_entries = {}
    for name in all_sprint_names:
        if name not in current_sprint_names:
            num = _sprint_num(name)
            if MIN_SPRINT <= num <= max_sprint_num:
                sprint_entries[num] = MOVED_OUT

    for name, sprint_obj in current_sprint_objs.items():
        num = _sprint_num(name)
        if num >= MIN_SPRINT:
            sprint_entries[num] = _classify_sprint_timing(sprint_obj, histories)

    return sorted(sprint_entries.items(), key=lambda x: x[0])


def _original_assignee_for_sprint(issue, sprint_obj) -> str:
    # Return the assignee nickname at the time the sprint first went active.
    fields = issue["fields"]
    histories = issue["changelog"]["histories"]
    current_full = fields["assignee"]["displayName"] if fields.get("assignee") else ""

    if not sprint_obj:
        return _assignee_nickname(current_full)

    sprint_name = sprint_obj.get("name", "")
    sprint_added_dt = None
    for history in histories:
        for item in history["items"]:
            if item["field"] != "Sprint":
                continue
            to_sprints = _sprint_names(item.get("toString") or "")
            from_sprints = _sprint_names(item.get("fromString") or "")
            if sprint_name in to_sprints and sprint_name not in from_sprints:
                dt = parse_date(history["created"])
                if sprint_added_dt is None or dt < sprint_added_dt:
                    sprint_added_dt = dt

    if sprint_added_dt is None:
        return _assignee_nickname(current_full)

    sprint_start_dt = parse_date(sprint_obj.get("startDate"))
    if sprint_start_dt:
        sprint_start_eod = sprint_start_dt.replace(hour=23, minute=59, second=59, microsecond=0)
        cutoff_dt = max(sprint_added_dt, sprint_start_eod)
    else:
        cutoff_dt = sprint_added_dt

    assignee_changes = []
    for history in histories:
        for item in history["items"]:
            if item["field"] == "assignee":
                assignee_changes.append((parse_date(history["created"]), item))
    assignee_changes.sort(key=lambda x: x[0])

    if not assignee_changes:
        return _assignee_nickname(current_full)

    current_state = assignee_changes[0][1].get("fromString") or ""
    for change_dt, item in assignee_changes:
        if change_dt <= cutoff_dt:
            current_state = item.get("toString") or ""
        else:
            break
    return _assignee_nickname(current_state) if current_state else "Unassigned"


def _sprint_started(sprint_num, sprint_map):
    # Return True if the given sprint number has already started or closed.
    s = sprint_map.get(sprint_num)
    if not s:
        return True
    state = s.get("state", "")
    if state in ("active", "closed"):
        return True
    if state == "future":
        return False
    start_dt = parse_date(s.get("startDate"))
    return datetime.now(start_dt.tzinfo) >= start_dt if start_dt else False


def _extract_key(label):
    # Extract the Jira issue key from a '[KEY] summary' label string.
    m = re.match(r'\[([A-Z]+-\d+)\]', str(label) if label is not None else '')
    return m.group(1) if m else None


# ═══════════════════════════════════════════════════════════════════
#  Workbook API update (incremental, co-authoring safe)
# ═══════════════════════════════════════════════════════════════════

def _gv(grid, row, col):
    """Read a 1-based cell from a usedRange values grid (0-based internally)."""
    r, c = row - 1, col - 1
    if 0 <= r < len(grid) and 0 <= c < len(grid[r]):
        v = grid[r][c]
        return v if v != "" else None
    return None


def _norm_sprint(v):
    # Normalize a sprint cell value to an integer, or None on failure.
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def _issue_row_cells(record):
    """Build the A..I values list for a Jira record."""
    issue = record["issue"]
    fields = issue["fields"]
    key = issue["key"]
    current_assignee = _assignee_nickname(
        fields["assignee"]["displayName"] if fields.get("assignee") else "")
    original_assignee = record.get("original_assignee") or current_assignee
    issue_type = get_type(issue)
    label = f"[{key}] {fields.get('summary', '')}"
    has_desc = "Y" if len(_adf_to_text(fields.get("description") or "")) > 20 else "N"
    status = get_status(issue, issue_type)
    return [record["sprint_num"], record.get("team"), original_assignee,
            current_assignee, issue_type, record["scope"], has_desc, label, status]


def _hyperlink_formula(label):
    """Build =HYPERLINK(url, label) from a '[KEY] summary' label."""
    key = _extract_key(label)
    if not key:
        return str(label) if label is not None else ""
    lbl = str(label).replace('"', '""').replace('\n', ' ').replace('\r', '')
    return f'=HYPERLINK("{BASE_URL}/browse/{key}","{lbl}")'


def update_via_workbook(helper, item_id, session_id, issues):
    """Incrementally update the Sprint Template sheet via the Workbook API.

    Reads the live sheet grid, reconciles it against current Jira state
    for the active and next sprints, then writes all changes (values,
    HYPERLINK formulas, and font/alignment formatting) in a single
    atomic pass.  Returns the current sprint number.
    """

    def wb(method, rel, body=None):
        # Dispatch a Workbook API request through the SharePoint helper.
        return helper.workbook_request(method, item_id, session_id, rel, body)

    # ── Read the live grid ──
    used = wb("GET", f"{_WS}/usedRange?$select=values,rowCount,address")
    grid = used.get("values", [])
    total_rows = used.get("rowCount", len(grid))
    print(f"[DIAG] usedRange address: {used.get('address')}, rowCount: {total_rows}, grid rows: {len(grid)}")

    # Find last row with actual data in A..I
    last_data_row = DATA_START_ROW - 1
    for r in range(DATA_START_ROW, total_rows + 1):
        if any(_gv(grid, r, c) is not None for c in range(1, 10)):
            last_data_row = r

    # ── Sprint targets (from Jira) ──
    active_sprint_nums = []
    for issue in issues:
        for s in (issue["fields"].get("customfield_10020") or []):
            num = _sprint_num(s.get("name", ""))
            if num >= MIN_SPRINT and s.get("state") == "active":
                active_sprint_nums.append(num)
    current_sprint_num = max(active_sprint_nums) if active_sprint_nums else MIN_SPRINT
    if not active_sprint_nums:
        b2 = _norm_sprint(_gv(grid, 2, 2))
        if b2 and b2 >= MIN_SPRINT:
            current_sprint_num = b2
    next_sprint_num = current_sprint_num + 1
    target_sprints = (current_sprint_num, next_sprint_num)
    target_set = set(target_sprints)

    # ── Sprint object lookup ──
    sprint_map = {}
    for issue in issues:
        for s in (issue["fields"].get("customfield_10020") or []):
            num = _sprint_num(s.get("name", ""))
            if num >= MIN_SPRINT and num not in sprint_map:
                sprint_map[num] = s

    # ── Build api_map (what Jira says should be in the sheet) ──
    api_map, sys_excluded = {}, set()
    for issue in issues:
        for sprint_num, scope in get_sprint_history(issue, next_sprint_num):
            if sprint_num not in target_set or scope == MOVED_OUT:
                continue
            composite = (issue["key"], sprint_num)
            sprint_obj = next(
                (s for s in (issue["fields"].get("customfield_10020") or [])
                 if _sprint_num(s.get("name", "")) == sprint_num), None)
            original_assignee = _original_assignee_for_sprint(issue, sprint_obj)
            current_full = (issue["fields"]["assignee"]["displayName"]
                            if issue["fields"].get("assignee") else "")
            current_assignee = _assignee_nickname(current_full)
            team_name, team_idx = _team_for_nickname(current_assignee)
            if team_idx == _SYS_TEAM_IDX:
                sys_excluded.add(composite)
                continue
            api_map[composite] = {
                "issue": issue, "sprint_num": sprint_num, "scope": scope,
                "jira_note": issue.get("_jira_note"),
                "original_assignee": original_assignee, "team": team_name,
            }

    # ── Index existing target-sprint rows from the live grid ──
    existing = {}  # (key, sprint) -> [{"row": int, "cells": list}, ...]
    print(f"\n[DIAG] === EXISTING GRID ROWS (col H raw values) ===")
    for r in range(DATA_START_ROW, last_data_row + 1):
        sn = _norm_sprint(_gv(grid, r, 1))
        raw_h = _gv(grid, r, 8)
        key = _extract_key(raw_h)
        if sn in target_set:
            print(f"  Excel row {r} | sprint={sn} | key={key} | raw H={repr(raw_h)}")
        if key and sn in target_set:
            cells = [_gv(grid, r, c) for c in range(1, NUM_COLS + 1)]
            existing.setdefault((key, sn), []).append({"row": r, "cells": cells})

    # ── Build prev_k_map: ticket key -> K value from the most recent older sprint ──
    prev_k_map = {}
    for r in range(DATA_START_ROW, last_data_row + 1):
        sn = _norm_sprint(_gv(grid, r, 1))
        if sn is not None and sn not in target_set:
            key = _extract_key(_gv(grid, r, 8))
            if key:
                k_val = _gv(grid, r, 11)
                if k_val is not None:
                    prev = prev_k_map.get(key)
                    if prev is None or sn > prev[0]:
                        prev_k_map[key] = (sn, k_val)

    # Find the contiguous region of target-sprint rows
    target_rows = [d["row"] for lst in existing.values() for d in lst]
    if target_rows:
        region_start = min(target_rows)
        region_cur_n = last_data_row - region_start + 1
    else:
        region_start = last_data_row + 1
        region_cur_n = 0

    # ── Build the final row set ──
    final = {sn: [] for sn in target_sprints}
    handled = set()
    updated = removed = deleted = 0

    for composite, lst in existing.items():
        sn = composite[1]
        first = lst[0]
        cur_scope = first["cells"][5]  # col F (0-indexed)

        if composite in api_map:
            handled.add(composite)
            if cur_scope == MOVED_OUT:
                final[sn].append(list(first["cells"]))
            else:
                a_to_i = _issue_row_cells(api_map[composite])
                cells = [None] * NUM_COLS
                cells[:9] = a_to_i
                cells[9] = first["cells"][9]    # preserve J (manual note)
                cells[11] = first["cells"][11]  # preserve L (manual note)
                existing_k = first["cells"][10]
                if not existing_k:
                    prev_entry = prev_k_map.get(composite[0])
                    existing_k = prev_entry[1] if prev_entry else None
                jira_note = api_map[composite].get("jira_note")
                if jira_note is not None:
                    cells[10] = f"{existing_k} Jira: {jira_note}" if existing_k else f"Jira: {jira_note}"
                else:
                    cells[10] = existing_k
                final[sn].append(cells)
                updated += 1
        else:
            if composite in sys_excluded and cur_scope != MOVED_OUT:
                deleted += 1
                continue
            if not _sprint_started(sn, sprint_map):
                deleted += 1
                continue
            cells = list(first["cells"])
            if cur_scope != MOVED_OUT:
                cells[5] = MOVED_OUT
                removed += 1
            final[sn].append(cells)

    new_composites = [c for c in api_map if c not in handled]
    for composite in new_composites:
        sn = composite[1]
        a_to_i = _issue_row_cells(api_map[composite])
        cells = [None] * NUM_COLS
        cells[:9] = a_to_i
        prev_entry = prev_k_map.get(composite[0])
        prev_k = prev_entry[1] if prev_entry else None
        jira_note = api_map[composite].get("jira_note")
        if jira_note is not None:
            cells[10] = f"{prev_k} Jira: {jira_note}" if prev_k else f"Jira: {jira_note}"
        else:
            cells[10] = prev_k
        final[sn].append(cells)

    # Sort each sprint block by team then assignee
    def _sort_key(cells):
        # Sort rows by team index, then alphabetically by original assignee name.
        orig = cells[2]  # col C = original assignee
        _, ti = _team_for_nickname(orig)
        return (ti, orig or "")

    final_rows = []
    for sn in target_sprints:
        final_rows.extend(sorted(final[sn], key=_sort_key))
    fin_n = len(final_rows)

    print(f"\n[DIAG] === FINAL ROWS TO WRITE ({fin_n} rows) ===")
    print(f"  {'Row':>4} | {'Sprint':>6} | {'Team':>4} | {'Assignee':<12} | {'Key':<12} | Col H (label)")
    print(f"  {'----':>4} | {'------':>6} | {'----':>4} | {'--------':<12} | {'---':<12} | -----")
    for i, row in enumerate(final_rows):
        label = row[7]
        key = _extract_key(label)
        print(f"  {region_start + i:>4} | {row[0]:>6} | {str(row[1]):>4} | {str(row[3]):<12} | {key or '???':<12} | {repr(label)}")

    # ── Reconcile row count: insert or delete rows as needed ──
    delta = fin_n - region_cur_n
    if delta > 0:
        ins_at = region_start + region_cur_n
        for _ in range(delta):
            wb("POST", f"{_WS}/range(address='{ins_at}:{ins_at}')/insert",
               {"shift": "Down"})
    elif delta < 0:
        del_start = region_start + fin_n
        del_end = region_start + region_cur_n - 1
        for r in range(del_end, del_start - 1, -1):
            wb("POST", f"{_WS}/range(address='{r}:{r}')/delete",
               {"shift": "Up"})

    if fin_n == 0:
        print("No active rows for target sprints; nothing to write.")
    else:
        end_row = region_start + fin_n - 1

        # ── Single atomic write: values + HYPERLINK formulas in col H ──
        formulas_grid = []
        for row in final_rows:
            frow = list(row)
            frow[7] = _hyperlink_formula(row[7])
            formulas_grid.append(frow)
        wb("PATCH", f"{_WS}/range(address='A{region_start}:{_LAST_COL}{end_row}')",
           {"formulas": formulas_grid})

        # ── Formatting: type color on col E (batched by contiguous same-color runs) ──
        types = [row[4] for row in final_rows]
        i = 0
        while i < fin_n:
            color = TYPE_FONT_COLORS.get(types[i], "000000")
            j = i
            while j + 1 < fin_n and TYPE_FONT_COLORS.get(types[j + 1], "000000") == color:
                j += 1
            r1, r2 = region_start + i, region_start + j
            wb("PATCH", f"{_WS}/range(address='E{r1}:E{r2}')/format/font",
               {"color": f"#{color}"})
            i = j + 1

        # Col H: blue underlined link style
        wb("PATCH", f"{_WS}/range(address='H{region_start}:H{end_row}')/format/font",
           {"color": "#1155CC", "underline": "Single"})

        # Alignment on all written cells
        wb("PATCH", f"{_WS}/range(address='A{region_start}:{_LAST_COL}{end_row}')/format",
           {"horizontalAlignment": "Left", "verticalAlignment": "Center"})

        print(
            f"Done via Workbook API "
            f"(sprints {current_sprint_num}/{next_sprint_num} | "
            f"{len(api_map)} active, {len(issues)} tickets | "
            f"{updated} updated, {len(new_composites)} added, "
            f"{removed} marked-removed, {deleted} deleted | "
            f"rows {region_start}-{end_row})"
        )

    # ── "Last Updated" timestamp in H2 ──
    now = datetime.now().strftime("%m/%d/%Y %I:%M %p")
    wb("PATCH", f"{_WS}/range(address='H2')",
       {"values": [[f"Last Updated: {now}"]]})

    return current_sprint_num


# ═══════════════════════════════════════════════════════════════════
#  Sprint filter (separate session so the Table range is up-to-date)
# ═══════════════════════════════════════════════════════════════════

def apply_sprint_filter(helper, item_id, session_id, current_sprint_num):
    """Filter the Table's first column to show only the current sprint.

    Must run in a separate session AFTER data writes are committed,
    otherwise the Table range won't include newly inserted rows.
    """
    def wb(method, rel, body=None):
        # Dispatch a Workbook API request through the SharePoint helper.
        return helper.workbook_request(method, item_id, session_id, rel, body)

    try:
        tables_resp = wb("GET", f"{_WS}/tables")
        table_list = tables_resp.get("value", [])
        if table_list:
            table_name = table_list[0]["name"]
            wb("POST",
               f"tables('{table_name}')/columns/itemAt(index=0)/filter/apply",
               {"criteria": {
                   "criterion1": f"={current_sprint_num}",
                   "filterOn": "Custom"
               }})
            print(f"Table filter applied: showing only Sprint {current_sprint_num}")
        else:
            print("Warning: no Excel Table found on sheet; "
                  "cannot set filter programmatically via Graph API")
    except Exception as e:
        print(f"Warning: could not apply filter: {e}")


# ═══════════════════════════════════════════════════════════════════
#  Entry point
# ═══════════════════════════════════════════════════════════════════

def main():
    """Archive the target file, sync Jira data, and apply the sprint filter."""
    helper = SharePointHelper()

    print("Fetching issues from Jira...")
    issues = get_jira_issues()
    print(f"Found {len(issues)} issues")

    print(f"Archiving: {SPRINT_TEMPLATE_FILE}")
    archive_path = helper.archive_file(SPRINT_TEMPLATE_FILE, archive_folder=ARCHIVE_PATH)
    print(f"Archived: {archive_path}")

    item_id = helper.get_item_id(SPRINT_TEMPLATE_FILE)

    print("Updating Sprint Template via Workbook API...")
    session_id = helper.open_workbook_session(item_id)
    try:
        current_sprint_num = update_via_workbook(helper, item_id, session_id, issues)
    finally:
        helper.close_workbook_session(item_id, session_id)

    print("Applying sprint filter...")
    session_id = helper.open_workbook_session(item_id)
    try:
        apply_sprint_filter(helper, item_id, session_id, current_sprint_num)
    finally:
        helper.close_workbook_session(item_id, session_id)


if __name__ == "__main__":
    main()
