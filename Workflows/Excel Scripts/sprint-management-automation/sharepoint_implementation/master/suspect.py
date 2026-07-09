"""
Build a Suspect Report sheet inside the live Sprint Management workbook
from a JAMA suspect report (JSON).

This script writes data only: cell values, the outline-level numbers
in column A, and the marker cells the VBA macro reads.  All
presentation (fills, fonts, alignment, merges, row grouping, column
widths) is applied inside the workbook by suspect_macro.vb, which
keeps the Graph API call count per build to a handful.
"""
import re
import json
import sys
import time
from datetime import datetime, timezone

from dotenv import load_dotenv

from config import SPRINT_TEMPLATE_FILE
from sharepoint_helper import SharePointHelper, SharePointOperationError

load_dotenv(r"c:\Scripts\Leahi Sprint Management Scripts\leahi-sprint-management\.env")

MODULE_ORDER = ('PRS', 'SRS', 'LVTC')
SHEET_NAME = 'Suspect Report'
HEADER_ROW = 3
DATA_START_ROW = HEADER_ROW + 1

LEVEL_MODULE = 1
LEVEL_ITEM = 2
LEVEL_SUSPECT = 3
MARKER_LABEL_CELL = 'N1'
BUILD_MARKER_CELL = 'O1'
# N2/O2 belong to the suspect.vb macro (it stamps O2 once it has grouped a
# build) -- never written by this script; listed here to document the
# shared marker-cell contract.
GROUPED_LABEL_CELL = 'N2'
GROUPED_MARKER_CELL = 'O2'
LAST_TREE_ROW_LABEL_CELL = 'N3'
LAST_TREE_ROW_CELL = 'O3'

# Leading spaces baked into column-B label text to simulate indentation
# per tree level, since the API has no real cell-indent property.
INDENT_SPACES = {
    'category': '',
    'module': '   ',
    'item': '      ',
    'suspect': '         ',
}

# Throttling (429) and gateway hiccups are transient service-side
# responses -- back off and retry instead of aborting the run over one
# flaky request.  Now that styling lives in the VBA macro a build is
# only a handful of calls, but large value writes can still be
# throttled when the service is busy.
_TRANSIENT_ERROR_MARKERS = ('429', 'Too Many Requests', '502', '503', '504',
                            'Gateway Timeout', 'Service Unavailable', 'Bad Gateway')
_MAX_RETRIES = 5


class SuspectReportError(Exception):
    """Raised when the input report is malformed or inconsistent."""


def natural_key(text):
    """Return a natural-sort key so 'ID-9' sorts before 'ID-10'."""
    parts = re.split(r'(\d+)', text)
    return [int(tok) if tok.isdigit() else tok for tok in parts]


def load_report(json_path):
    """Load and return the report dictionary from a JSON file."""
    with open(json_path, encoding='utf-8') as handle:
        return json.load(handle)


def _is_category_mapping(candidate):
    """Return True if candidate maps categories to item groups.

    A category group maps item IDs to item objects, so every value is a
    dict.  A bare item entry (old flat format) has scalar values such as
    the numeric 'ID', which distinguishes the two layouts.
    """
    if not isinstance(candidate, dict):
        return False
    for value in candidate.values():
        if not isinstance(value, dict):
            return False
    return True


def build_tree(report):
    """Return a nested tree of category -> module -> items.

    The returned structure is a list of tuples::

        (category, [(module, [(item_id, item_num, [(up_id, up_num)])])])

    Categories are sorted alphabetically; modules follow MODULE_ORDER;
    items are natural-sorted; suspect links keep their source order.
    """
    modules = report.get('Modules')
    if not isinstance(modules, dict):
        raise SuspectReportError("Report is missing a 'Modules' object.")

    # category -> module -> {item_id: (item_num, [(up_id, up_num), ...])}
    catalog = {}
    for module in modules:
        categories = modules[module]
        for category, items in categories.items():
            if not _is_category_mapping(items):
                raise SuspectReportError(
                    f"Module '{module}' is not organized by category. "
                    "This script expects Modules -> Module -> Category "
                    "-> Item. Use the by-category report format.")
            for item_id, info in items.items():
                links = list(info.get('Upstream Suspects', {}).items())
                catalog.setdefault(category, {}).setdefault(
                    module, {})[item_id] = (info.get('ID'), links)

    module_rank = {name: pos for pos, name in enumerate(MODULE_ORDER)}
    tree = []
    for category in sorted(catalog):
        module_map = catalog[category]
        module_names = sorted(
            module_map, key=lambda m: module_rank.get(m, len(module_rank)))
        module_rows = []
        for module in module_names:
            item_map = module_map[module]
            item_rows = []
            for item_id in sorted(item_map, key=natural_key):
                item_num, links = item_map[item_id]
                item_rows.append((item_id, item_num, links))
            module_rows.append((module, item_rows))
        tree.append((category, module_rows))
    return tree


def _count_links(module_rows):
    """Return total suspect links across every item in a module list."""
    return sum(
        len(links)
        for _module, item_rows in module_rows
        for _item_id, _item_num, links in item_rows)


# ═══════════════════════════════════════════════════════════════════
#  Row plan: pure data -> (matrix, row_level), no I/O
# ═══════════════════════════════════════════════════════════════════

def _plan_tree_rows(tree, matrix, row_level):
    """Fill matrix/row_level for the tree portion. Returns last row used."""
    row = DATA_START_ROW
    for category, module_rows in tree:
        matrix.setdefault(row, {})[2] = INDENT_SPACES['category'] + category
        matrix.setdefault(row, {})[4] = _count_links(module_rows)
        row += 1

        for module, item_rows in module_rows:
            module_links = sum(len(links) for _i, _n, links in item_rows)
            matrix.setdefault(row, {})[2] = INDENT_SPACES['module'] + module
            matrix.setdefault(row, {})[4] = module_links
            row_level[row] = LEVEL_MODULE
            row += 1

            for item_id, item_num, links in item_rows:
                matrix.setdefault(row, {})[2] = INDENT_SPACES['item'] + item_id
                matrix.setdefault(row, {})[3] = item_num
                matrix.setdefault(row, {})[4] = len(links)
                row_level[row] = LEVEL_ITEM
                row += 1

                for up_id, up_num in links:
                    matrix.setdefault(row, {})[2] = INDENT_SPACES['suspect'] + up_id
                    matrix.setdefault(row, {})[3] = up_num
                    row_level[row] = LEVEL_SUSPECT
                    row += 1

    return row - 1


def _plan_summary_rows(report, start_row, matrix):
    """Fill matrix for the module summary block (values only)."""
    title_row = start_row
    matrix.setdefault(title_row, {})[2] = 'Module Summary'

    header_row = title_row + 2
    headers = ('Module', 'Total Items', 'Items With Suspects',
               'Suspect Relations', 'Impacted Items')
    for offset, label in enumerate(headers):
        matrix.setdefault(header_row, {})[2 + offset] = label

    summary = report.get('Summary', {})
    module_names = [m for m in MODULE_ORDER if m in summary]
    module_names += [m for m in summary if m not in MODULE_ORDER]

    row = header_row + 1
    column_totals = [0, 0, 0, 0]
    for module in module_names:
        stats = summary[module]
        matrix.setdefault(row, {})[2] = module
        values = (
            stats.get('Total Items'),
            stats.get('Items With Suspects'),
            stats.get('Suspect Relations'),
            stats.get('Impacted Items Count'),
        )
        for offset, value in enumerate(values):
            matrix.setdefault(row, {})[3 + offset] = value
            column_totals[offset] += value or 0
        row += 1

    total_row = row
    matrix.setdefault(total_row, {})[2] = 'Total'
    for offset, total in enumerate(column_totals):
        matrix.setdefault(total_row, {})[3 + offset] = total


def build_row_plan(report, tree):
    """Return (last_row, matrix, row_level, last_tree_row).

    matrix is {row: {col: value}} for columns B-F (1-based col index).
    row_level maps row -> the outline level to stamp in column A (only
    set for tree rows below the always-visible category level).
    last_tree_row is the last row belonging to the tree itself, before
    the summary block.  All styling, merging, and grouping of these
    rows is applied later inside the workbook by suspect_macro.vb,
    driven by the column-A levels and the marker cells.
    """
    matrix = {}
    row_level = {}

    relation_counts = report.get('Suspect Relation Counts', {})
    counts_text = ' / '.join(
        f'{name} {relation_counts.get(name, 0)}'
        for name in MODULE_ORDER if name in relation_counts)
    matrix.setdefault(1, {})[2] = 'LEAHI - Upstream Suspect Links by Category'
    matrix.setdefault(2, {})[2] = (
        f"Project ID: {report.get('Project ID', '')}     "
        f"Generated: {report.get('Generated At', '')}     "
        f"Suspect links: {counts_text}")

    headers = ('Category / Module / Item / Suspect', 'Numeric ID', 'Suspect Links')
    for offset, label in enumerate(headers):
        matrix.setdefault(HEADER_ROW, {})[2 + offset] = label

    last_tree_row = _plan_tree_rows(tree, matrix, row_level)
    _plan_summary_rows(report, last_tree_row + 2, matrix)

    last_row = max(matrix)
    return last_row, matrix, row_level, last_tree_row


# ═══════════════════════════════════════════════════════════════════
#  Graph API write: turns the plan into workbook-session calls
# ═══════════════════════════════════════════════════════════════════

def _ensure_sheet(wb):
    sheets = wb('GET', 'worksheets')['value']
    if not any(s['name'] == SHEET_NAME for s in sheets):
        wb('POST', 'worksheets/add', {'name': SHEET_NAME})


def _clear_sheet(wb, clear_through_row):
    wb('POST', f"worksheets('{SHEET_NAME}')/range(address='A1:F{clear_through_row}')/clear",
       {'applyTo': 'All'})


def _write_values(wb, last_row, matrix):
    values = []
    for row in range(1, last_row + 1):
        cells = matrix.get(row, {})
        values.append([cells.get(col) for col in range(1, 7)])
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='A1:F{last_row}')",
       {'values': values})


def _write_levels(wb, last_row, row_level):
    values = [[row_level.get(row)] for row in range(DATA_START_ROW, last_row + 1)]
    if any(v[0] is not None for v in values):
        wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='A{DATA_START_ROW}:A{last_row}')",
           {'values': values})


def _write_build_marker(wb, last_tree_row):
    marker = datetime.now(timezone.utc).isoformat()
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{MARKER_LABEL_CELL}:{BUILD_MARKER_CELL}')",
       {'values': [['Build marker (do not edit):', marker]]})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{LAST_TREE_ROW_LABEL_CELL}:{LAST_TREE_ROW_CELL}')",
       {'values': [['Last tree row (do not edit):', last_tree_row]]})
    return marker


def write_suspect_report(helper, item_id, session_id, report, tree):
    """Clear and rewrite the Suspect Report sheet from scratch."""

    def wb(method, rel, body=None):
        for attempt in range(_MAX_RETRIES + 1):
            try:
                return helper.workbook_request(method, item_id, session_id, rel, body)
            except SharePointOperationError as exc:
                is_transient = any(marker in str(exc) for marker in _TRANSIENT_ERROR_MARKERS)
                if not is_transient or attempt == _MAX_RETRIES:
                    raise
                wait = 2 ** attempt
                print(f"  Transient error, retrying in {wait}s "
                      f"(attempt {attempt + 1}/{_MAX_RETRIES})...")
                time.sleep(wait)

    _ensure_sheet(wb)

    last_row, matrix, row_level, last_tree_row = build_row_plan(report, tree)

    existing = wb('GET', f"worksheets('{SHEET_NAME}')/usedRange?$select=rowCount")
    clear_through = max(last_row, existing.get('rowCount', 0), 200)
    _clear_sheet(wb, clear_through)

    _write_values(wb, last_row, matrix)
    _write_levels(wb, last_row, row_level)

    marker = _write_build_marker(wb, last_tree_row)

    return last_row, marker


def main(argv):
    """Command-line entry point."""
    if len(argv) < 2:
        print('Usage: python suspect.py <input.json>')
        return 1

    input_path = argv[1]

    try:
        report = load_report(input_path)
        tree = build_tree(report)
    except SuspectReportError as exc:
        print(f'Error: {exc}')
        return 1

    helper = SharePointHelper()
    item_id = helper.get_item_id(SPRINT_TEMPLATE_FILE)
    session_id = helper.open_workbook_session(item_id)
    try:
        last_row, marker = write_suspect_report(helper, item_id, session_id, report, tree)
    finally:
        helper.close_workbook_session(item_id, session_id)

    print(f"Wrote {last_row} rows to '{SHEET_NAME}' in {SPRINT_TEMPLATE_FILE} (marker={marker})")
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
