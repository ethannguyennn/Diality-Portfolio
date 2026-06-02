import argparse
import os
import re
import smtplib
import traceback
import requests
from copy import copy
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime

load_dotenv()

BASE_URL = "https://diality.atlassian.net"
PROJECT = "LDT"

SWVV_TEAM = ["Raghu Kallala", "Thomas Lippold", "Tejaskumar Patel", "Tiffany Mejia", "Zoltan Miskolci", "Sarina Cheung", "Tisha Patel", "Ethan Nguyen"]
FW_TEAM = ["Arpita Srivastava", "Jashwant Gantyada", "Michael Garthwaite", "Praneeth Bunne", "Sameer Poyil", "Varshini Nagabooshanam", "Vijay Pamula", "Suresh Dharnala"]
SW_TEAM = ["Nicholas Ramirez", "Stephen Quong", "Dara Navaei", "Eliza Petersen", "Christina Heine", "Caitlynn Chang"]
TEAMS = [SWVV_TEAM, FW_TEAM, SW_TEAM]

NOTIFICATION_EMAILS = [os.getenv("JIRA_EMAIL")]

STATUS_COLORS = {
    "In Progress": "FFF2CC",
    "Done":        "D9EAD3",
    "Not Started": "F4CCCC",
    "Blocked":     "EA9999",
}


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


def _get_all_board_sprints(auth):
    """Return every sprint on board 84 across all states, handling pagination."""
    sprints = []
    start_at = 0
    while True:
        r = requests.get(
            f"{BASE_URL}/rest/agile/1.0/board/84/sprint",
            params={"state": "active,closed,future", "maxResults": 50, "startAt": start_at},
            auth=auth,
        )
        r.raise_for_status()
        data = r.json()
        page = data.get("values", [])
        sprints.extend(page)
        start_at += len(page)
        if data.get("isLast", True) or not page:
            break
    return sprints


def _get_active_sprint(auth):
    r = requests.get(
        f"{BASE_URL}/rest/agile/1.0/board/84/sprint",
        params={"state": "active"},
        auth=auth,
    )
    r.raise_for_status()
    values = r.json().get("values", [])
    return values[0] if values else None


def _sprint_names(s):
    return {n.strip() for n in s.split(",") if n.strip()} if s else set()


def _was_removed_from_sprint(issue, sprint_name):
    was_added = was_removed = False
    for history in issue["changelog"]["histories"]:
        for item in history["items"]:
            if item["field"] != "Sprint":
                continue
            to_sprints = _sprint_names(item.get("toString") or "")
            from_sprints = _sprint_names(item.get("fromString") or "")
            if sprint_name in to_sprints and sprint_name not in from_sprints:
                was_added = True
            if sprint_name in from_sprints and sprint_name not in to_sprints:
                was_removed = True
    return was_added and was_removed


def get_jira_issues():
    auth = HTTPBasicAuth(os.getenv("JIRA_EMAIL"), os.getenv("JIRA_API_TOKEN"))

    board_issues = _fetch_paginated(
        f"{BASE_URL}/rest/agile/1.0/board/84/issue",
        "sprint in openSprints() OR sprint in closedSprints() OR sprint in futureSprints()",
        auth,
    )

    # The search API 'was' operator is unavailable on this Jira instance.
    # Fetch recently-updated backlog items and identify removed-from-sprint
    # tickets via changelog analysis instead.
    seen = {i["key"] for i in board_issues}
    removed_issues = []
    active_sprint = _get_active_sprint(auth)
    if active_sprint:
        sprint_start = parse_date(active_sprint.get("startDate"))
        sprint_name = active_sprint["name"]
        if sprint_start:
            start_str = sprint_start.strftime("%Y-%m-%d")
            backlog_issues = _fetch_paginated(
                f"{BASE_URL}/rest/agile/1.0/board/84/backlog",
                f"project = {PROJECT} AND updated >= '{start_str}'",
                auth,
            )
            removed_issues = [
                i for i in backlog_issues
                if i["key"] not in seen and _was_removed_from_sprint(i, sprint_name)
            ]

    return board_issues + removed_issues


def _sprint_num(name):
    m = re.search(r'(\d+)', name)
    return int(m.group(1)) if m else -1


def get_latest_sprint(sprint_info):
    return max(sprint_info, key=lambda s: _sprint_num(s.get("name", "")))


def get_type(issue):
    issuetype = issue["fields"]["issuetype"]["name"].lower()
    summary = issue["fields"].get("summary", "").lower()

    if "dialin" in issuetype or "dialin" in summary:
        return "Dialin Ticket"
    elif issuetype == "bug":
        return "Bug"
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
    return "Added to Scope" if most_recent_add > sprint_start_eod else "Scoped"


def get_scope(issue):
    sprint_info = issue["fields"].get("customfield_10020")
    histories = issue["changelog"]["histories"]

    if not sprint_info:
        return "Removed from Scope"

    latest_sprint = get_latest_sprint(sprint_info)

    if latest_sprint.get("state") == "active":
        return _classify_sprint_timing(latest_sprint, histories)

    latest_num = _sprint_num(latest_sprint["name"])
    for history in histories:
        for item in history["items"]:
            if item["field"] != "Sprint":
                continue
            from_sprints = _sprint_names(item.get("fromString") or "")
            to_sprints = _sprint_names(item.get("toString") or "")
            for removed in from_sprints - to_sprints:
                if _sprint_num(removed) > latest_num:
                    return "Removed from Scope"

    return _classify_sprint_timing(latest_sprint, histories)


def get_status(issue):
    status = issue["fields"]["status"]
    status_name = status["name"].lower()
    category_key = status["statusCategory"]["key"]

    if "blocked" in status_name:
        return "Blocked"
    elif "in progress" in status_name:
        return "In Progress"
    elif category_key == "done" or "complete" in status_name or "done" in status_name:
        return "Done"
    else:
        return "Not Started"


def _send_notification(subject, body):
    if not NOTIFICATION_EMAILS:
        return
    smtp_host = os.getenv("SMTP_HOST")
    smtp_port = int(os.getenv("SMTP_PORT", 587))
    smtp_user = os.getenv("SMTP_USER")
    smtp_password = os.getenv("SMTP_PASSWORD")
    smtp_from = os.getenv("SMTP_FROM")
    if not all([smtp_host, smtp_user, smtp_password, smtp_from]):
        print("Warning: SMTP env vars not fully configured, skipping notification.")
        return
    msg = MIMEMultipart()
    msg["From"] = smtp_from
    msg["To"] = ", ".join(NOTIFICATION_EMAILS)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))
    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_from, NOTIFICATION_EMAILS, msg.as_string())


def _issue_sort_key(issue):
    fields = issue["fields"]
    sprint_info = fields.get("customfield_10020")
    sprint_num = _sprint_num(get_latest_sprint(sprint_info)["name"]) if sprint_info else 9999
    assignee = fields["assignee"]["displayName"] if fields.get("assignee") else ""
    team_idx = next((i for i, team in enumerate(TEAMS) if assignee in team), len(TEAMS))
    return (sprint_num, team_idx, assignee)


def _extract_key(label):
    """Extract Jira key from a cell label like '[LDT-123] Summary title'."""
    m = re.match(r'\[([A-Z]+-\d+)\]', label or '')
    return m.group(1) if m else None


def _ensure_sheet_setup(ws):
    headers = ["Sprint", "Assignee", "Type", "Scope Changes", "Jira #", "Description", "Status"]
    header_fill = PatternFill(fill_type="solid", fgColor="1F3864")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)

    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="left", vertical="center")

    ws.row_dimensions[1].height = 20
    ws.freeze_panes = "A2"

    widths = {"A": 8, "B": 22, "C": 15, "D": 18, "E": 75, "F": 14, "G": 15}
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


def _write_issue_row(ws, row_idx, issue):
    fields = issue["fields"]
    key = issue["key"]

    sprint_info = fields.get("customfield_10020")
    sprint = str(_sprint_num(get_latest_sprint(sprint_info)["name"])) if sprint_info else ""
    assignee = fields["assignee"]["displayName"] if fields.get("assignee") else "Unassigned"
    issue_type = get_type(issue)
    scope = get_scope(issue)
    ticket_url = f"{BASE_URL}/browse/{key}"
    ticket_label = f"[{key}] {fields.get('summary', '')}"
    has_description = "Yes" if fields.get("description") else "No"
    status = get_status(issue)

    values = [sprint, assignee, issue_type, scope, ticket_label, has_description, status]
    base_align = Alignment(horizontal="left", vertical="center")

    for col, value in enumerate(values, start=1):
        cell = ws.cell(row=row_idx, column=col, value=value)
        cell.alignment = base_align
        cell.fill = PatternFill()
        cell.hyperlink = None
        cell.font = Font(name="Arial", size=10)

        if col == 5:
            cell.hyperlink = ticket_url
            cell.font = Font(name="Arial", size=10, color="1155CC", underline="single")
        elif col == 6:
            cell.fill = PatternFill(fill_type="solid", fgColor="D9EAD3" if has_description == "Yes" else "F4CCCC")
        elif col == 7:
            cell.fill = PatternFill(fill_type="solid", fgColor=STATUS_COLORS.get(status, "FFFFFF"))


def _sort_data_rows(ws, api_map):
    max_row = ws.max_row
    if max_row < 3:
        return

    rows_data = []
    for row_idx in range(2, max_row + 1):
        cells = []
        for col_idx in range(1, 8):
            cell = ws.cell(row=row_idx, column=col_idx)
            cells.append({
                'value': cell.value,
                'hyperlink': cell.hyperlink.target if cell.hyperlink else None,
                'font': copy(cell.font),
                'fill': copy(cell.fill),
                'alignment': copy(cell.alignment),
            })

        key = _extract_key(cells[4]['value'])
        if key and key in api_map:
            sort_key = _issue_sort_key(api_map[key])
        else:
            try:
                sprint_num = int(cells[0]['value'] or 9999)
            except (ValueError, TypeError):
                sprint_num = 9999
            assignee = cells[1]['value'] or ''
            team_idx = next((i for i, t in enumerate(TEAMS) if assignee in t), len(TEAMS))
            sort_key = (sprint_num, team_idx, assignee)

        rows_data.append((sort_key, cells))

    rows_data.sort(key=lambda x: x[0])

    for row_idx, (_, cells) in enumerate(rows_data, start=2):
        for col_idx, cd in enumerate(cells, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = cd['value']
            cell.font = cd['font']
            cell.fill = cd['fill']
            cell.alignment = cd['alignment']
            cell.hyperlink = cd['hyperlink']


def update_excel(input_path, issues):
    try:
        wb = load_workbook(input_path)
    except FileNotFoundError:
        wb = Workbook()
        wb.active.title = "Sprint Management"

    if "Sprint Management" in wb.sheetnames:
        ws = wb["Sprint Management"]
    else:
        ws = wb.create_sheet("Sprint Management")

    _ensure_sheet_setup(ws)

    # Map every existing data row by its Jira key
    existing = {}
    for row_idx in range(2, ws.max_row + 1):
        label = ws.cell(row=row_idx, column=5).value
        key = _extract_key(label)
        if key:
            existing[key] = row_idx

    api_map = {issue["key"]: issue for issue in issues}

    # Update rows present in both the sheet and the API
    updated = set()
    for key, row_idx in existing.items():
        if key in api_map:
            _write_issue_row(ws, row_idx, api_map[key])
            updated.add(key)

    # Delete rows that no longer appear in the API (reverse order preserves indices)
    to_delete = sorted(
        [existing[k] for k in existing if k not in api_map],
        reverse=True,
    )
    for row_idx in to_delete:
        ws.delete_rows(row_idx)

    # Append rows for tickets not yet in the sheet
    new_keys = [k for k in api_map if k not in existing]
    next_row = ws.max_row + 1
    for key in new_keys:
        _write_issue_row(ws, next_row, api_map[key])
        next_row += 1

    _sort_data_rows(ws, api_map)

    wb.save(input_path)
    print(
        f"Saved: {input_path} "
        f"({len(api_map)} tickets | "
        f"{len(updated)} updated, {len(new_keys)} added, {len(to_delete)} removed)"
    )


def debug_sprints(auth):
    print("\n=== DEBUG: Sprint Diagnosis ===\n")

    # Step 1 — what the updated board JQL actually returns
    print("Step 1: Fetching via board JQL (open + closed + future)...")
    board_issues = _fetch_paginated(
        f"{BASE_URL}/rest/agile/1.0/board/84/issue",
        "sprint in openSprints() OR sprint in closedSprints() OR sprint in futureSprints()",
        auth,
    )
    sprint_counts: dict = {}
    for issue in board_issues:
        sprint_info = issue["fields"].get("customfield_10020")
        if sprint_info:
            num = _sprint_num(get_latest_sprint(sprint_info).get("name", ""))
            sprint_counts[num] = sprint_counts.get(num, 0) + 1
    print(f"  Total issues fetched: {len(board_issues)}")
    print(f"  Breakdown by sprint: {dict(sorted(sprint_counts.items()))}")
    for target in [31, 32]:
        count = sprint_counts.get(target, 0)
        status = "OK" if count > 0 else "MISSING"
        print(f"  Sprint {target}: {count} issues [{status}]")

    # Step 2 — list every sprint on board 84 with its state
    print("\nStep 2: All sprints on board 84...")
    all_sprints = _get_all_board_sprints(auth)
    if not all_sprints:
        print("  WARNING: No sprints returned — check board ID or auth.")
    for s in sorted(all_sprints, key=lambda x: _sprint_num(x.get("name", ""))):
        num = _sprint_num(s.get("name", ""))
        marker = " <-- TARGET" if num in [31, 32] else ""
        print(f"  Sprint {num:3d} | state={s.get('state', '?'):8s} | id={s['id']:6d} | {s.get('name', '')}{marker}")

    # Step 3 — direct per-sprint fetch for 31 and 32
    print("\nStep 3: Direct issue fetch for sprints 31 and 32...")
    board_keys = {i["key"] for i in board_issues}
    for target in [31, 32]:
        sprint_obj = next(
            (s for s in all_sprints if _sprint_num(s.get("name", "")) == target), None
        )
        if not sprint_obj:
            print(f"  Sprint {target}: NOT found on board 84 at all — board sub-filter may exclude it.")
            continue
        sid = sprint_obj["id"]
        state = sprint_obj["state"]
        name = sprint_obj["name"]
        direct = _fetch_paginated(
            f"{BASE_URL}/rest/agile/1.0/sprint/{sid}/issue",
            f"project = {PROJECT}",
            auth,
        )
        direct_keys = {i["key"] for i in direct}
        missing = direct_keys - board_keys
        print(f"  Sprint {target} (id={sid}, state={state}): '{name}'")
        print(f"    Direct sprint endpoint: {len(direct)} issues")
        print(f"    Already in board fetch: {len(direct_keys & board_keys)}")
        if missing:
            print(f"    MISSING from board fetch: {sorted(missing)}")
            print(f"    --> These tickets exist but the board JQL/filter is excluding them.")
        else:
            print(f"    All issues accounted for in board fetch.")

    print("\n=== END DEBUG ===\n")


def main():
    parser = argparse.ArgumentParser(description="Update sprint management Excel from Jira.")
    parser.add_argument(
        "input",
        nargs="?",
        default="sprint_management.xlsx",
        help="Path to the Excel file to update (default: sprint_management.xlsx)",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Run sprint diagnostics to identify missing tickets instead of updating Excel",
    )
    args = parser.parse_args()

    if args.debug:
        auth = HTTPBasicAuth(os.getenv("JIRA_EMAIL"), os.getenv("JIRA_API_TOKEN"))
        debug_sprints(auth)
        return

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
