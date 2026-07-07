"""Jira fetching and issue-field normalization.

Turns raw Jira API responses into the facts the sheet-update modules need:
issue type/status classification, and per-sprint scope/timing/assignee
history derived from each issue's changelog.
"""

import logging
import os
import re
from datetime import datetime
from typing import Optional

import requests
from dotenv import load_dotenv
from requests.auth import HTTPBasicAuth

from roster import _assignee_nickname

load_dotenv()

logger = logging.getLogger(__name__)

BASE_URL = "https://diality.atlassian.net"
PROJECT = "LDT"

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


# ═══════════════════════════════════════════════════════════════════
#  Jira helpers
# ═══════════════════════════════════════════════════════════════════

def _fetch_paginated(url, jql, auth):
    # Fetch all issues from a paginated Jira REST API endpoint.
    logger.debug(f"Fetching paginated Jira results from {url} | jql={jql}")
    all_issues, start_at = [], 0
    while True:
        params = {
            "fields": "summary,status,assignee,issuetype,customfield_10020,description,comment",
            "expand": "changelog", "jql": jql, "maxResults": 100, "startAt": start_at,
        }
        try:
            resp = requests.get(url, params=params, auth=auth, timeout=30)
            resp.raise_for_status()
        except requests.RequestException as e:
            logger.error(f"Jira request failed: {url} | jql={jql} | startAt={start_at} | {e}")
            raise
        data = resp.json()
        all_issues.extend(data["issues"])
        start_at += len(data["issues"])
        logger.debug(f"Fetched {start_at}/{data.get('total', '?')} issues from {url}")
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
        try:
            resp = requests.get(url, params={"maxResults": 1000}, auth=auth, timeout=30)
            resp.raise_for_status()
            comments = resp.json().get("comments", [])
        except requests.RequestException as e:
            logger.error(f"Failed to fetch full comment list for {issue['key']}: {e}")
            raise

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
    jira_email = os.getenv("JIRA_EMAIL")
    jira_token = os.getenv("JIRA_API_TOKEN")
    if not jira_email or not jira_token:
        logger.error("Missing JIRA_EMAIL or JIRA_API_TOKEN in environment")
        raise RuntimeError("Missing JIRA_EMAIL or JIRA_API_TOKEN in .env")
    auth = HTTPBasicAuth(jira_email, jira_token)

    logger.info("Fetching board issues (open/future sprints) from Jira")
    board_issues = _fetch_paginated(
        f"{BASE_URL}/rest/agile/1.0/board/84/issue",
        "sprint in openSprints() OR sprint in futureSprints()", auth,
    )
    logger.info(f"Fetched {len(board_issues)} board issues")
    seen = {i["key"] for i in board_issues}

    logger.info("Fetching backlog issues from Jira")
    all_backlog = _fetch_paginated(
        f"{BASE_URL}/rest/agile/1.0/board/84/backlog",
        f"project = {PROJECT}", auth,
    )
    backlog_with_history = [
        i for i in all_backlog if i["key"] not in seen and _has_sprint_history(i)
    ]
    logger.info(
        f"Fetched {len(all_backlog)} backlog issues, "
        f"{len(backlog_with_history)} with sprint history >= {MIN_SPRINT}"
    )

    all_issues = board_issues + backlog_with_history
    logger.info(f"Fetching Jira notes (xlsx<>xlsx comments) for {len(all_issues)} issues")
    for issue in all_issues:
        try:
            issue["_jira_note"] = _latest_jira_note(issue, auth)
        except requests.RequestException as e:
            logger.error(f"Failed to fetch Jira note for {issue.get('key', '?')}: {e}")
            raise
    logger.info(f"Completed Jira fetch: {len(all_issues)} total issues")
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
