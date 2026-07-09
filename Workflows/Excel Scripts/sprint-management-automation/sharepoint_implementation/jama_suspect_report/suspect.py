"""Sprint Template sheet update (incremental, co-authoring safe) and the
sprint filter applied on top of it.
"""

import logging
import re
from datetime import datetime

from grid_utils import _gv, _norm_sprint
from jira_client import (
    BASE_URL, MIN_SPRINT, MOVED_OUT, _adf_to_text, _original_assignee_for_sprint,
    _sprint_num, _sprint_started, get_sprint_history, get_status, get_type,
)
from roster import _SYS_TEAM_IDX, _assignee_nickname, _team_for_nickname
from sharepoint_helper import SharePointOperationError

logger = logging.getLogger(__name__)

# ── Sheet layout ──
SHEET_NAME = "Sprint Template"
DATA_START_ROW = 4
NUM_COLS = 12

# ── Type font colors ──
TYPE_FONT_COLORS = {
    "Bug": "FF0000", "Dialin ticket": "4472C4",
    "Support": "00B050", "Little-V": "7030A0", "Big-V": "000000",
}

# ── Workbook API constants ──
_WS = f"worksheets('{SHEET_NAME}')"
_LAST_COL = chr(ord('A') + NUM_COLS - 1)  # 'L'


def _extract_key(label):
    # Extract the Jira issue key from a '[KEY] summary' label string.
    m = re.match(r'\[([A-Z]+-\d+)\]', str(label) if label is not None else '')
    return m.group(1) if m else None


def _strip_jira_suffix(k_val):
    """Remove the ' Jira: ...' tail this script appends, leaving only the human-authored portion."""
    if not k_val:
        return None
    stripped = re.sub(r'\s*Jira:.*$', '', str(k_val), flags=re.DOTALL).strip()
    return stripped or None


def _build_k(base_k, jira_note):
    """Combine an optional human base note with an optional Jira note into a K cell value."""
    if jira_note is not None:
        return f"{base_k} Jira: {jira_note}".strip() if base_k else f"Jira: {jira_note}"
    return base_k or ""


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
    logger.info("Reading live sheet grid")
    used = wb("GET", f"{_WS}/usedRange?$select=values,rowCount,address")
    grid = used.get("values", [])
    total_rows = used.get("rowCount", len(grid))
    logger.info(f"Grid read: address={used.get('address')}, rowCount={total_rows}")

    # Find last row with actual data in A..I
    last_data_row = DATA_START_ROW - 1
    for r in range(DATA_START_ROW, total_rows + 1):
        if any(_gv(grid, r, c) is not None for c in range(1, 10)):
            last_data_row = r
    logger.debug(f"Last data row: {last_data_row}")

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
    logger.info(f"Target sprints: current={current_sprint_num}, next={next_sprint_num}")

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
    logger.info(f"Built api_map: {len(api_map)} active tickets, {len(sys_excluded)} SYS-excluded")

    # ── Index existing target-sprint rows from the live grid ──
    existing = {}  # (key, sprint) -> [{"row": int, "cells": list}, ...]
    for r in range(DATA_START_ROW, last_data_row + 1):
        sn = _norm_sprint(_gv(grid, r, 1))
        key = _extract_key(_gv(grid, r, 8))
        if key and sn in target_set:
            cells = [_gv(grid, r, c) for c in range(1, NUM_COLS + 1)]
            existing.setdefault((key, sn), []).append({"row": r, "cells": cells})
    logger.info(f"Indexed {len(existing)} existing target-sprint rows from the sheet")

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
                cells[9] = first["cells"][9]   # preserve J (employee-maintained)
                cells[11] = first["cells"][11]  # preserve L (manual note)
                # K is hand-maintained too — strip the old Jira suffix we wrote
                # last run so we don't compound "Jira: x Jira: x" each run.
                base_k = _strip_jira_suffix(first["cells"][10])
                cells[10] = _build_k(base_k, api_map[composite].get("jira_note"))
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
        # J (index 9) left as None — new row starts blank, employees fill it in
        cells[10] = _build_k(None, api_map[composite].get("jira_note"))
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
    logger.info(
        f"Final row set built: {fin_n} rows "
        f"({updated} updated, {len(new_composites)} new, {removed} marked-removed, {deleted} deleted)"
    )

    # ── Reconcile row count: insert or delete rows as needed ──
    delta = fin_n - region_cur_n
    if delta > 0:
        logger.info(f"Inserting {delta} row(s) at row {region_start + region_cur_n}")
        ins_at = region_start + region_cur_n
        try:
            for _ in range(delta):
                wb("POST", f"{_WS}/range(address='{ins_at}:{ins_at}')/insert",
                   {"shift": "Down"})
        except SharePointOperationError:
            logger.error(f"Failed inserting rows at {ins_at}")
            raise
    elif delta < 0:
        logger.info(f"Deleting {-delta} row(s) from {region_start + fin_n} to {region_start + region_cur_n - 1}")
        del_start = region_start + fin_n
        del_end = region_start + region_cur_n - 1
        try:
            for r in range(del_end, del_start - 1, -1):
                wb("POST", f"{_WS}/range(address='{r}:{r}')/delete",
                   {"shift": "Up"})
        except SharePointOperationError:
            logger.error(f"Failed deleting rows {del_start}-{del_end}")
            raise

    if fin_n == 0:
        logger.info("No active rows for target sprints; nothing to write.")
    else:
        end_row = region_start + fin_n - 1

        # ── Clear stale formatting first ──
        # Row inserts (above) copy formatting from the adjacent row, so a
        # manually-struck-through or otherwise formatted row can propagate
        # its formatting onto brand-new rows forever, since nothing below
        # resets properties (e.g. strikethrough) that aren't explicitly set.
        # Wiping formats before reapplying our own guarantees a clean slate
        # every run regardless of what the row was inserted next to.
        logger.debug(f"Clearing stale formatting on A{region_start}:{_LAST_COL}{end_row}")
        try:
            wb("POST", f"{_WS}/range(address='A{region_start}:{_LAST_COL}{end_row}')/clear",
               {"applyTo": "Formats"})
        except SharePointOperationError:
            logger.error(f"Failed clearing formatting on A{region_start}:{_LAST_COL}{end_row}")
            raise

        # ── Single atomic write: values + HYPERLINK formulas in col H ──
        logger.info(f"Writing {fin_n} rows to A{region_start}:{_LAST_COL}{end_row}")
        formulas_grid = []
        for row in final_rows:
            frow = list(row)
            frow[1] = row[1] if row[1] is not None else ""   # B: team — write "" not None
            frow[7] = _hyperlink_formula(row[7])
            frow[9] = row[9] if row[9] is not None else ""   # J: clear stale on row shift
            frow[11] = row[11] if row[11] is not None else "" # L: clear stale on row shift
            formulas_grid.append(frow)
        try:
            wb("PATCH", f"{_WS}/range(address='A{region_start}:{_LAST_COL}{end_row}')",
               {"formulas": formulas_grid})
        except SharePointOperationError:
            logger.error(f"Failed writing rows to A{region_start}:{_LAST_COL}{end_row}")
            raise
        logger.info("Row data written successfully")

        # ── Force recalculation ──
        # The Workbook API writes formula text but doesn't evaluate it
        # server-side, so newly-written cells (e.g. col H's HYPERLINK)
        # keep a stale cached value (0 for a fresh formula cell) until
        # something recalculates them -- which also means the *next* run's
        # usedRange read would see that stale 0 instead of the real label,
        # breaking _extract_key's row matching. Force it now so both the
        # display and the next run's grid read reflect the real formula
        # results immediately.
        try:
            wb("POST", "application/calculate", {"calculationType": "Full"})
        except SharePointOperationError:
            logger.warning("Failed to force recalculation after writing rows")

        # ── Formatting: type color on col E (batched by contiguous same-color runs) ──
        logger.debug("Applying type color formatting to column E")
        types = [row[4] for row in final_rows]
        i = 0
        try:
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
        except SharePointOperationError:
            logger.error("Failed applying formatting to written rows")
            raise
        logger.debug("Formatting applied successfully")

        logger.info(
            f"Done via Workbook API "
            f"(sprints {current_sprint_num}/{next_sprint_num} | "
            f"{len(api_map)} active, {len(issues)} tickets | "
            f"{updated} updated, {len(new_composites)} added, "
            f"{removed} marked-removed, {deleted} deleted | "
            f"rows {region_start}-{end_row})"
        )

    # ── "Last Updated" timestamp in H2 ──
    now = datetime.now().strftime("%m/%d/%Y %I:%M %p")
    try:
        wb("PATCH", f"{_WS}/range(address='H2')",
           {"values": [[f"Last Updated: {now}"]]})
    except SharePointOperationError:
        logger.error("Failed writing 'Last Updated' timestamp to H2")
        raise
    logger.info(f"Last Updated timestamp written: {now}")

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

    logger.info(f"Applying sprint filter for sprint {current_sprint_num}")
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
            logger.info(f"Table filter applied: showing only Sprint {current_sprint_num}")
        else:
            logger.warning("No Excel Table found on sheet; cannot set filter programmatically via Graph API")
    except Exception:
        # Non-fatal: the data write already succeeded, so log and continue
        # rather than failing the whole run over a cosmetic filter step.
        logger.exception("Could not apply sprint filter")
