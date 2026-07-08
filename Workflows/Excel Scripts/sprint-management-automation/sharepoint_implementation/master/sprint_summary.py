"""Sprint Summary sheet (aggregate stats, separate sheet, separate session)."""

import logging
import time

from grid_utils import _col_letter, _gv, _norm_sprint
from sharepoint_helper import SharePointOperationError
from sprint_template import SHEET_NAME

logger = logging.getLogger(__name__)

# ── Sheet layout ──
SUMMARY_SHEET_NAME = "Sprint Summary"
_WS_SUMMARY = f"worksheets('{SUMMARY_SHEET_NAME}')"

# Bug/Dialin statuses where the fix work is effectively finished
# (just waiting on test sign-off or the next release cutover) --
# these count toward the Sprint Summary "Done" bucket alongside a
# literal "Done" status, even though the Sprint Template keeps
# showing the more granular status for visibility.
SUMMARY_DONE_ALIASES = [
    "Ready for Fix to be Tested", "Testing Fix in Progress",
    "Ready for Updated Release", "Needs Final Fix Details",
]

# Sprint Summary percentage columns (status, type, person) that get a
# green fill when their computed text reads "... (100.0%)". Column
# bounds are 1-based (start, end) per block: I..L (status), N..R
# (type), T..AP (person).
_SUMMARY_PCT_COL_BLOCKS = ((9, 12), (14, 18), (20, 42))
SUMMARY_COMPLETE_FILL = "#92D050"


def _f_scope_count(col_letter, r):
    # COUNTIF of Template col F (scope) over the sprint's row range, keyed on the column's own header.
    return (f"=COUNTIF(INDEX('{SHEET_NAME}'!$F:$F,$B{r}):INDEX('{SHEET_NAME}'!$F:$F,$C{r}),"
            f"{col_letter}$1)")


def _f_status_pct(col_letter, r, is_done=False):
    # "<count> (<pct of Committed>)" for a status column
    # (Done/In-progress/Not Started/Blocked). The Done bucket also
    # folds in SUMMARY_DONE_ALIASES so bug tickets sitting in a
    # post-fix/pre-release status read as "done" in the sprint
    # rollup, and excludes tickets whose scope (col F) is
    # "Move out of Sprint" -- those never count toward TotalTasks
    # ($G = Scoped+Added) either, so a ticket marked Done after
    # being moved out shouldn't inflate the Done count.
    status_range = (f"INDEX('{SHEET_NAME}'!$I:$I,$B{r}):"
                    f"INDEX('{SHEET_NAME}'!$I:$I,$C{r})")
    if is_done:
        scope_range = (f"INDEX('{SHEET_NAME}'!$F:$F,$B{r}):"
                       f"INDEX('{SHEET_NAME}'!$F:$F,$C{r})")
        scope_ok = f'{scope_range},"<>Move out of Sprint"'
        done_terms = [f"COUNTIFS({status_range},{col_letter}$1,{scope_ok})"]
        done_terms += [f'COUNTIFS({status_range},"{s}",{scope_ok})'
                       for s in SUMMARY_DONE_ALIASES]
        count_expr = f"SUM({','.join(done_terms)})"
    else:
        count_expr = f"COUNTIF({status_range},{col_letter}$1)"
    return (f"=LET(\nDoneCount,{count_expr},\nTotalTasks,$G{r},\n"
            f"IF(OR(DoneCount=0,TotalTasks=0),\"\",DoneCount&\" (\"&"
            f"TEXT(DoneCount/TotalTasks,\"0.0%\")&\")\")\n)")


def _f_type_pct(col_letter, r):
    # "<count> (<pct of Scoped+Added>)" for a ticket-type column (Little-V/Big-V/Dialin/Bug/Support).
    return (f"=LET(\nDoneCount,COUNTIF(INDEX('{SHEET_NAME}'!$E:$E,$B{r}):"
            f"INDEX('{SHEET_NAME}'!$E:$E,$C{r}),{col_letter}$1),\n"
            f"IF(DoneCount=0,\"\",DoneCount&\" (\"&"
            f"TEXT(DoneCount/($E{r}+$F{r}),\"0.0%\")&\")\")\n)")


def _f_person_pct(col_letter, r):
    # "<done> / <assigned> (<pct>)" for one team member's column, keyed on the column's own header.
    return (
        f"=LET(\nAssigned,\nSUM(\nCOUNTIFS(\n"
        f"INDEX('{SHEET_NAME}'!$D:$D,$B{r}):INDEX('{SHEET_NAME}'!$D:$D,$C{r}),{col_letter}$1,\n"
        f"INDEX('{SHEET_NAME}'!$F:$F,$B{r}):INDEX('{SHEET_NAME}'!$F:$F,$C{r}),{{\"Scoped\",\"Added\"}}\n"
        f")\n),\nDone,\nCOUNTIFS(\n"
        f"INDEX('{SHEET_NAME}'!$D:$D,$B{r}):INDEX('{SHEET_NAME}'!$D:$D,$C{r}),{col_letter}$1,\n"
        f"INDEX('{SHEET_NAME}'!$I:$I,$B{r}):INDEX('{SHEET_NAME}'!$I:$I,$C{r}),\"Done\"\n"
        f"),\nIF(\nAssigned=0,\n\"\",\nDone&\" / \"&Assigned&\" (\"&"
        f"TEXT(Done/Assigned,\"0.0%\")&\")\"\n)\n)"
    )


# Columns T(20)..AP(42): one per team member, in whatever order the sheet's
# header row already has them -- this mirrors the existing hand-built rows
# (29-31) exactly, just generated instead of copy-pasted.
_SUMMARY_PERSON_COLS = range(20, 43)


def _build_summary_row_formulas(sprint_num, r):
    """Build the A..AQ formula row for one Sprint Summary row.

    Mirrors the formulas already hand-built for the existing sprint rows
    (29-31) so the sheet stays self-consistent -- every cell recalculates
    live from the Sprint Template sheet via its Start/End row (B/C).
    """
    row = [""] * 43  # A..AQ; columns without a header (M, S, AQ) stay blank
    row[0] = sprint_num
    row[1] = f"=MATCH(A{r},'{SHEET_NAME}'!$A:$A,0)"
    row[2] = f"=LOOKUP(2,1/('{SHEET_NAME}'!$A:$A=A{r}),ROW('{SHEET_NAME}'!$A:$A))"
    row[3] = f"=C{r}-B{r}+1"
    row[4] = _f_scope_count("E", r)
    row[5] = _f_scope_count("F", r)
    row[6] = f"=E{r}+F{r}"
    row[7] = _f_scope_count("H", r)
    for col_idx in range(9, 13):  # I..L: Done, In-progress, Not Started, Blocked
        row[col_idx - 1] = _f_status_pct(
            _col_letter(col_idx), r, is_done=(col_idx == 9))
    for col_idx in range(14, 19):  # N..R
        row[col_idx - 1] = _f_type_pct(_col_letter(col_idx), r)
    for col_idx in _SUMMARY_PERSON_COLS:  # T..AP
        row[col_idx - 1] = _f_person_pct(_col_letter(col_idx), r)
    return row


# _apply_completion_highlight makes many sequential Workbook API calls (up
# to ~2 per percentage cell), long enough that a transient Gateway
# Timeout / Service Unavailable is a real occurrence -- retry those
# instead of aborting the whole row over one flaky request. Same
# thresholds as suspect.py's retry wrapper.
_SUMMARY_TRANSIENT_ERROR_MARKERS = (
    '502', '503', '504', 'Gateway Timeout', 'Service Unavailable', 'Bad Gateway',
)
_SUMMARY_MAX_RETRIES = 5


def ensure_sprint_summary_row(helper, item_id, session_id, current_sprint_num):
    """Write (or refresh) the current sprint's row on the Sprint Summary sheet.

    'Sprint Summary' is one row per sprint, formula-driven off the Sprint
    Template sheet's Start/End rows -- this only ever touches the row for
    current_sprint_num, appending a new one if it isn't there yet. Column A
    is scanned to find or place that row; every other sprint's row, the
    header, and the manually-authored Challenges column are left untouched.

    _apply_completion_highlight below makes up to ~2 Workbook API calls
    per percentage cell in the row (dozens of calls total), so wb()
    retries transient Gateway Timeout / Service Unavailable errors
    instead of letting one flaky call abort the whole row -- same
    rationale and thresholds as suspect.py's retry wrapper.
    """
    def wb(method, rel, body=None):
        for attempt in range(_SUMMARY_MAX_RETRIES + 1):
            try:
                return helper.workbook_request(method, item_id, session_id, rel, body)
            except SharePointOperationError as exc:
                is_transient = any(
                    marker in str(exc) for marker in _SUMMARY_TRANSIENT_ERROR_MARKERS)
                if not is_transient or attempt == _SUMMARY_MAX_RETRIES:
                    raise
                wait = 2 ** attempt
                logger.warning(
                    f"Transient error on Sprint Summary Workbook call, retrying "
                    f"in {wait}s (attempt {attempt + 1}/{_SUMMARY_MAX_RETRIES})..."
                )
                time.sleep(wait)

    try:
        logger.info(f"Ensuring Sprint Summary row for sprint {current_sprint_num}")
        used = wb("GET", f"{_WS_SUMMARY}/usedRange?$select=values,rowCount,address")
        grid = used.get("values", [])
        total_rows = used.get("rowCount", len(grid))

        target_row, last_row = None, 1  # row 1 is the header
        for r in range(2, total_rows + 1):
            sn = _norm_sprint(_gv(grid, r, 1))
            if sn is not None:
                last_row = r
                if sn == current_sprint_num:
                    target_row = r

        is_new = target_row is None
        if is_new:
            target_row = last_row + 1

        row_formulas = _build_summary_row_formulas(current_sprint_num, target_row)
        wb("PATCH", f"{_WS_SUMMARY}/range(address='A{target_row}:AQ{target_row}')",
           {"formulas": [row_formulas]})

        # Force recalculation so the percentage formulas just written
        # evaluate before _apply_completion_highlight reads their
        # computed text back below.
        try:
            wb("POST", "application/calculate", {"calculationType": "Full"})
        except SharePointOperationError:
            logger.warning("Failed to force recalculation after writing Sprint Summary row")

        _apply_completion_highlight(wb, target_row)

        logger.info(
            f"Sprint Summary row for sprint {current_sprint_num} "
            f"{'created' if is_new else 'refreshed'} at row {target_row}"
        )
    except Exception:
        # Non-fatal: the core Jira sync already succeeded, so log and move
        # on rather than failing the whole cron run over the summary sheet.
        logger.exception("Could not ensure Sprint Summary row")


def _apply_completion_highlight(wb, target_row):
    """Fill any percentage cell in target_row reading "...(100.0%)" with
    SUMMARY_COMPLETE_FILL, and clear the fill on every other cell in
    those columns.

    Excel's native conditional formatting isn't reachable through the
    Graph Workbook API -- workbookRange has no conditionalFormats
    relationship in either v1.0 or beta (confirmed against Microsoft's
    own resource reference; ConditionalFormat/FormatConditions only
    exist in the Excel JS API and the VBA object model, never in
    Graph). Reading back the computed text and setting a plain
    per-cell fill instead is equivalent here, since this row's values
    only ever change when this script runs.
    """
    for col_start, col_end in _SUMMARY_PCT_COL_BLOCKS:
        rng = f"{_col_letter(col_start)}{target_row}:{_col_letter(col_end)}{target_row}"
        text_row = wb("GET", f"{_WS_SUMMARY}/range(address='{rng}')?$select=text")["text"][0]
        for offset, cell_text in enumerate(text_row):
            addr = f"{_col_letter(col_start + offset)}{target_row}"
            if "(100.0%)" in str(cell_text):
                wb("PATCH", f"{_WS_SUMMARY}/range(address='{addr}')/format/fill",
                   {"color": SUMMARY_COMPLETE_FILL})
            else:
                wb("POST", f"{_WS_SUMMARY}/range(address='{addr}')/clear",
                   {"applyTo": "Formats"})
