import os
import re
import requests
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime

load_dotenv()

BASE_URL = "https://diality.atlassian.net"
PROJECT = "LDT"

SWVV_TEAM = ["Raghu Kallala", "Thomas Lippold", "Tejaskumar Patel", "Tiffany Mejia", "Zoltan Miskolci", "Sarina Cheung", "Tisha Patel", "Ethan Nguyen"]
FW_TEAM = ["Arpita Srivastava", "Jashwant Gantyada", "Michael Garthwaite", "Praneeth Bunne", "Sameer Poyil", "Varshini Nagabooshanam", "Vijay Pamula", "Suresh Dharnala"]
SW_TEAM = ["Nicholas Ramirez", "Stephen Quong", "Dara Navaei", "Eliza Petersen", "Christina Heine", "Caitlynn Chang"]
TEAMS = [SWVV_TEAM, FW_TEAM, SW_TEAM]


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
        "sprint in openSprints() OR sprint in closedSprints()",
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


def _issue_sort_key(issue):
    fields = issue["fields"]
    sprint_info = fields.get("customfield_10020")
    sprint_num = _sprint_num(get_latest_sprint(sprint_info)["name"]) if sprint_info else 9999
    assignee = fields["assignee"]["displayName"] if fields.get("assignee") else ""
    team_idx = next((i for i, team in enumerate(TEAMS) if assignee in team), len(TEAMS))
    return (sprint_num, team_idx, assignee)


def create_excel(issues):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sprint Management"

    headers = ["Sprint", "Assignee", "Type", "Scope Changes", "Jira #", "Description", "Status"]
    header_fill = PatternFill(fill_type="solid", fgColor="1F3864")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)

    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="left", vertical="center")

    ws.row_dimensions[1].height = 20

    status_colors = {
        "In Progress": "FFF2CC",
        "Done":        "D9EAD3",
        "Not Started": "F4CCCC",
        "Blocked":     "EA9999",
    }

    for row_idx, issue in enumerate(issues, start=2):
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

        row_data = [sprint, assignee, issue_type, scope, ticket_label, has_description, status]

        for col, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col, value=value)
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(horizontal="left", vertical="center")

            if col == 5:
                cell.hyperlink = ticket_url
                cell.font = Font(name="Arial", size=10, color="1155CC", underline="single")

            if col == 6:
                cell.fill = PatternFill(fill_type="solid", fgColor="D9EAD3" if has_description == "Yes" else "F4CCCC")

            if col == 7:
                cell.fill = PatternFill(fill_type="solid", fgColor=status_colors.get(status, "FFFFFF"))

    widths = {"A": 8, "B": 22, "C": 15, "D": 18, "E": 75, "F": 14, "G": 15}
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width

    ws.freeze_panes = "A2"

    output = "sprint_management.xlsx"
    wb.save(output)
    print(f"Saved: {output} ({row_idx - 1} tickets)")


def main():
    print("Fetching issues from Jira...")
    issues = get_jira_issues()
    print(f"Found {len(issues)} issues")

    issues.sort(key=_issue_sort_key)

    print("Building Excel...")
    create_excel(issues)
    print("Done!")


if __name__ == "__main__":
    main()
