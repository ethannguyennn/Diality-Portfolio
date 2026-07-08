"""Update Leahi Sprint Management on SharePoint from Jira API.

Uses the Graph Workbook API for incremental edits so the file can be
updated even while someone has it open in Excel Online.
"""

import logging
import os
from logging.handlers import TimedRotatingFileHandler

from dotenv import load_dotenv

from config import (
    ARCHIVE_PATH, LOG_FORMAT, LOG_LEVEL, LOG_RETENTION_DAYS, LOGS_DIR,
    SPRINT_TEMPLATE_FILE,
)
from jira_client import get_jira_issues
from sharepoint_helper import SharePointHelper
from sprint_retro import ensure_sprint_retro_header
from sprint_summary import ensure_sprint_summary_row
from sprint_template import apply_sprint_filter, update_via_workbook

load_dotenv()

# ── Logging: console + daily rotating file, so cron runs leave a trail ──
os.makedirs(LOGS_DIR, exist_ok=True)
_file_handler = TimedRotatingFileHandler(
    os.path.join(LOGS_DIR, "leahi_sprint.log"),
    when="midnight", backupCount=LOG_RETENTION_DAYS, encoding="utf-8",
)
_file_handler.suffix = "%Y-%m-%d"
logging.basicConfig(
    level=LOG_LEVEL, format=LOG_FORMAT,
    handlers=[logging.StreamHandler(), _file_handler], force=True,
)
logger = logging.getLogger(__name__)


def main():
    """Archive the target file, sync Jira data, and apply the sprint filter."""
    logger.info("=== Leahi Sprint Management sync starting ===")
    try:
        helper = SharePointHelper()

        logger.info("Fetching issues from Jira...")
        issues = get_jira_issues()
        logger.info(f"Found {len(issues)} issues")

        logger.info(f"Archiving: {SPRINT_TEMPLATE_FILE}")
        archive_path = helper.archive_file(SPRINT_TEMPLATE_FILE, archive_folder=ARCHIVE_PATH)
        logger.info(f"Archived to: {archive_path}")

        item_id = helper.get_item_id(SPRINT_TEMPLATE_FILE)

        logger.info("Updating Sprint Template via Workbook API...")
        session_id = helper.open_workbook_session(item_id)
        try:
            current_sprint_num = update_via_workbook(helper, item_id, session_id, issues)
        except Exception:
            logger.exception("Failed updating Sprint Template via Workbook API")
            raise
        finally:
            helper.close_workbook_session(item_id, session_id)

        logger.info("Applying sprint filter...")
        session_id = helper.open_workbook_session(item_id)
        try:
            apply_sprint_filter(helper, item_id, session_id, current_sprint_num)
        finally:
            helper.close_workbook_session(item_id, session_id)

        logger.info("Ensuring Sprint Retro header...")
        session_id = helper.open_workbook_session(item_id)
        try:
            ensure_sprint_retro_header(helper, item_id, session_id, current_sprint_num)
        finally:
            helper.close_workbook_session(item_id, session_id)

        logger.info("Ensuring Sprint Summary row...")
        session_id = helper.open_workbook_session(item_id)
        try:
            ensure_sprint_summary_row(helper, item_id, session_id, current_sprint_num)
        finally:
            helper.close_workbook_session(item_id, session_id)

        logger.info("=== Leahi Sprint Management sync completed successfully ===")
    except Exception:
        logger.exception("=== Leahi Sprint Management sync FAILED ===")
        raise


if __name__ == "__main__":
    main()
