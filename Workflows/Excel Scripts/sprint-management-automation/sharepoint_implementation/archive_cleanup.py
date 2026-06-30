"""Organize the SharePoint archive folder into per-sprint subfolders, and
prune old sprints down to just their first and last archived file.

Sprint date windows come from the Jira board so files are classified by
the real sprint they fall in, not just a flat date split.

Rules:
  - Each archive file's trailing _MMDDYYYY date determines which "Sprint N"
    subfolder it belongs in. Files with no parseable date, or whose date
    doesn't fall inside any known sprint window, are left at the top level.
  - The current sprint and the two sprints before it are never pruned.
  - For older sprint folders, every file is deleted except those matching
    the earliest and latest archived date present in that folder (all
    files sharing the min/max date are kept, not just one).

Usage:
    py archive_cleanup.py
"""

import logging
import os
import re
from datetime import date, datetime
from logging.handlers import TimedRotatingFileHandler

import requests
from dotenv import load_dotenv
from requests.auth import HTTPBasicAuth

from config import ARCHIVE_PATH, LOG_FORMAT, LOG_LEVEL, LOG_RETENTION_DAYS, LOGS_DIR
from sharepoint_helper import SharePointHelper

load_dotenv()

# ── Logging: console + daily rotating file (separate from main.py's log) ──
os.makedirs(LOGS_DIR, exist_ok=True)
_file_handler = TimedRotatingFileHandler(
    os.path.join(LOGS_DIR, "archive_cleanup.log"),
    when="midnight", backupCount=LOG_RETENTION_DAYS, encoding="utf-8",
)
_file_handler.suffix = "%Y-%m-%d"
logging.basicConfig(
    level=LOG_LEVEL, format=LOG_FORMAT,
    handlers=[logging.StreamHandler(), _file_handler], force=True,
)
logger = logging.getLogger(__name__)

BASE_URL = "https://diality.atlassian.net"
BOARD_ID = 84
PROTECTED_SPRINT_COUNT = 3  # current sprint + this many prior sprints are never pruned

FILE_DATE_RE = re.compile(r'_(\d{8})')
SPRINT_FOLDER_RE = re.compile(r'^Sprint (\d+)$')


def _sprint_num(name: str):
    m = re.search(r'(\d+)', name or "")
    return int(m.group(1)) if m else None


def _parse_jira_date(date_str):
    if not date_str:
        return None
    date_str = date_str.replace("Z", "+00:00")
    date_str = re.sub(r'([+-])(\d{2})(\d{2})$', r'\1\2:\3', date_str)
    try:
        return datetime.fromisoformat(date_str).date()
    except ValueError:
        return None


def _fetch_jira_sprints(board_id: int) -> dict:
    """Return {sprint_num: {"start": date, "end": date, "state": str}} for board_id."""
    jira_email = os.getenv("JIRA_EMAIL")
    jira_token = os.getenv("JIRA_API_TOKEN")
    if not jira_email or not jira_token:
        logger.error("Missing JIRA_EMAIL or JIRA_API_TOKEN in environment")
        raise RuntimeError("Missing JIRA_EMAIL or JIRA_API_TOKEN in .env")
    auth = HTTPBasicAuth(jira_email, jira_token)

    sprints, start_at = {}, 0
    url = f"{BASE_URL}/rest/agile/1.0/board/{board_id}/sprint"
    while True:
        try:
            resp = requests.get(url, params={"startAt": start_at, "maxResults": 50}, auth=auth, timeout=30)
            resp.raise_for_status()
        except requests.RequestException as e:
            logger.error(f"Failed to fetch sprints from board {board_id}: {e}")
            raise
        data = resp.json()
        values = data.get("values", [])
        for sprint in values:
            num = _sprint_num(sprint.get("name", ""))
            if num is None:
                continue
            start = _parse_jira_date(sprint.get("startDate"))
            end = _parse_jira_date(sprint.get("completeDate") or sprint.get("endDate"))
            if start and end:
                sprints[num] = {"start": start, "end": end, "state": sprint.get("state")}
        start_at += len(values)
        if data.get("isLast", True) or not values:
            break
    logger.info(f"Fetched {len(sprints)} sprints with valid date windows from Jira board {board_id}")
    return sprints


def _current_sprint_num(sprints: dict) -> int:
    active = [num for num, s in sprints.items() if s["state"] == "active"]
    if active:
        return max(active)
    if sprints:
        logger.warning("No active sprint found; falling back to the highest known sprint number")
        return max(sprints)
    raise RuntimeError("No sprints found on the board; cannot determine current sprint")


def _extract_file_date(filename: str):
    """Return the date encoded as _MMDDYYYY in filename, or None if absent/invalid."""
    for m in FILE_DATE_RE.finditer(filename):
        try:
            return datetime.strptime(m.group(1), "%m%d%Y").date()
        except ValueError:
            continue
    return None


def _classify_sprint(file_date: date, sprints: dict):
    """Return the sprint number whose [start, end] window contains file_date, or None."""
    for num, s in sprints.items():
        if s["start"] <= file_date <= s["end"]:
            return num
    return None


def organize_and_prune(helper: SharePointHelper) -> None:
    sprints = _fetch_jira_sprints(BOARD_ID)
    current = _current_sprint_num(sprints)
    protected = {current - i for i in range(PROTECTED_SPRINT_COUNT)}
    logger.info(f"Current sprint: {current}. Protected from pruning: {sorted(protected)}")

    top_level = helper.list_children(ARCHIVE_PATH)
    files = [f for f in top_level if "file" in f]
    existing_sprint_folders = {
        int(m.group(1))
        for f in top_level if "folder" in f
        for m in [SPRINT_FOLDER_RE.match(f["name"])] if m
    }
    logger.info(f"Found {len(files)} files at the top level of {ARCHIVE_PATH}")

    # ── Step 1: sort top-level files into their sprint folders ──
    moved = skipped_no_date = skipped_no_sprint = 0
    classified_sprint_nums = set()
    for f in files:
        name = f["name"]
        file_date = _extract_file_date(name)
        if file_date is None:
            logger.debug(f"Skipping (no date in name): {name}")
            skipped_no_date += 1
            continue
        sprint_num = _classify_sprint(file_date, sprints)
        if sprint_num is None:
            logger.debug(f"Skipping (date {file_date} doesn't match any sprint window): {name}")
            skipped_no_sprint += 1
            continue
        classified_sprint_nums.add(sprint_num)
        dest_path = f"{ARCHIVE_PATH}/Sprint {sprint_num}"
        dest_id = helper.ensure_folder(dest_path)
        if helper.get_item_id_if_exists(f"{dest_path}/{name}"):
            # File already exists in the subfolder; the root copy is the duplicate.
            helper.delete_item(f["id"])
            logger.info(f"Deleted root duplicate '{name}' — already exists in {dest_path}/")
        else:
            helper.move_item(f["id"], dest_id)
            logger.info(f"Moved '{name}' -> {dest_path}/")
        moved += 1
    logger.info(
        f"Organize step done: {moved} moved, "
        f"{skipped_no_date} skipped (no date in filename), "
        f"{skipped_no_sprint} skipped (date outside any sprint window)"
    )

    # ── Step 2: prune old sprint folders down to first/last archived day ──
    candidate_sprint_nums = existing_sprint_folders | classified_sprint_nums
    deleted_total = 0
    for sprint_num in sorted(candidate_sprint_nums):
        if sprint_num in protected or sprint_num > current:
            logger.debug(f"Sprint {sprint_num} is protected; skipping prune")
            continue

        folder_path = f"{ARCHIVE_PATH}/Sprint {sprint_num}"
        if not helper.get_item_id_if_exists(folder_path):
            logger.debug(f"{folder_path} doesn't exist; nothing to prune")
            continue
        children = helper.list_children(folder_path)

        dated = [(c, _extract_file_date(c["name"])) for c in children if "file" in c]
        dated_with_date = [(c, d) for c, d in dated if d is not None]
        if len(dated_with_date) < 2:
            logger.debug(f"Sprint {sprint_num} folder has <2 dated files; nothing to prune")
            continue

        min_date = min(d for _, d in dated_with_date)
        max_date = max(d for _, d in dated_with_date)
        to_delete = [c for c, d in dated_with_date if min_date < d < max_date]

        if not to_delete:
            logger.debug(f"Sprint {sprint_num} folder already pruned (only first/last day present)")
            continue

        for c in to_delete:
            helper.delete_item(c["id"])
            logger.info(f"Deleted '{c['name']}' from Sprint {sprint_num}")
        deleted_total += len(to_delete)

    logger.info(f"Prune step done: {deleted_total} files deleted")


def main():
    logger.info("=== Archive cleanup starting ===")
    try:
        helper = SharePointHelper()
        organize_and_prune(helper)
        logger.info("=== Archive cleanup completed successfully ===")
    except Exception:
        logger.exception("=== Archive cleanup FAILED ===")
        raise


if __name__ == "__main__":
    main()
