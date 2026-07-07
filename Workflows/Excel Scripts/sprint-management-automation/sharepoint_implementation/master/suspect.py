"""Build a Suspect Report sheet inside the live Sprint Management workbook
from a JAMA suspect report (JSON).

Writes flat rows -- no native Excel row grouping, since the Graph
Workbook API doesn't expose that action -- into the "Suspect Report"
sheet of SPRINT_TEMPLATE_FILE, via the same safe, co-authoring-friendly
workbook-session API pattern update_sprint.py uses. Column A carries
each row's intended outline level (1/2/3, blank for category rows); a
companion VBA macro (suspect.vba) reads that column to build real
collapsible grouping client-side the moment anyone opens the sheet,
since row grouping is an Excel-only feature the REST API can't create.

The whole sheet's content (columns A-F) is cleared and rewritten from
scratch every run. A fresh timestamp is stamped into the build-marker
cell (O1) each run so the companion macro knows new data has landed and
it needs to regroup -- see suspect.vba for the other half of this.

The expected JSON layout is::

    {
      "Summary": {"PRS": {"Total Items": <int>, ...}, ...},
      "Suspect Relation Counts": {"PRS": <int>, ...},
      "Modules": {
        "PRS": {
          "<Category>": {
            "<Item ID>": {
              "ID": <int>,
              "Upstream Suspects": {"<Upstream ID>": <int>, ...}
            }
          }
        }
      },
      "Generated At": "<str>",
      "Project ID": <int>
    }

Usage::

    python suspect.py <input.json>
"""
import re
import json
import sys
import time
from datetime import datetime, timezone

from dotenv import load_dotenv

sys.path.insert(0, r"c:\Scripts\Leahi Sprint Management Scripts\OOP")

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

# Marker cells the companion VBA macro (suspect.vba) reads to decide
# whether it needs to (re)build grouping. Python always stamps a fresh
# BUILD_MARKER value here on every run; the macro only regroups when
# that differs from what it last grouped (which it stores in the
# grouped-marker cell -- Python never touches that one). LAST_TREE_ROW
# tells the macro exactly where the tree ends and the always-visible
# Module Summary block begins, since column A alone can't reliably
# convey that boundary (it's blank for both category rows and the
# entire summary block).
MARKER_LABEL_CELL = 'N1'
BUILD_MARKER_CELL = 'O1'
GROUPED_LABEL_CELL = 'N2'
GROUPED_MARKER_CELL = 'O2'
LAST_TREE_ROW_LABEL_CELL = 'N3'
LAST_TREE_ROW_CELL = 'O3'

# Style specs per row "kind": fill color, font, and column extent
# (1-based start/end column) the fill+font apply across, plus one or
# more (start_col, end_col, horizontal_align, indent) alignment specs.
# NOTE: the Graph API's workbookRangeFormat has no working "indent"
# property (confirmed empirically -- a PATCH with "indentLevel" is
# silently accepted but never actually applied, and isn't even a real
# selectable property). Hierarchy indentation is done with leading
# spaces baked directly into the label text instead -- see INDENT_SPACES.
STYLES = {
    'header': {
        'fill': 'FFC000', 'font': {'bold': True, 'size': 10, 'color': 'FFFFFF'},
        'extent': (2, 4),
        'aligns': [(2, 2, 'Left'), (3, 4, 'Center')],
    },
    'category': {
        'fill': '107C41', 'font': {'bold': True, 'size': 11, 'color': 'FFFFFF'},
        'extent': (2, 4),
        'aligns': [(2, 2, 'Left'), (3, 4, 'Center')],
    },
    'module': {
        'fill': '70AD47', 'font': {'bold': True, 'size': 10, 'color': 'FFFFFF'},
        'extent': (2, 4),
        'aligns': [(2, 2, 'Left'), (3, 4, 'Center')],
    },
    'item': {
        'fill': 'A9D08E', 'font': {'bold': True, 'size': 10, 'color': '000000'},
        'extent': (2, 4),
        'aligns': [(2, 2, 'Left'), (3, 4, 'Center')],
    },
    'suspect': {
        'fill': 'E2EFDA', 'font': {'bold': False, 'size': 10, 'color': '000000'},
        'extent': (2, 4),
        'aligns': [(2, 2, 'Left'), (3, 4, 'Center')],
    },
    'summary_header': {
        'fill': '70AD47', 'font': {'bold': True, 'size': 11, 'color': 'FFFFFF'},
        'extent': (2, 6),
        'aligns': [(2, 6, 'Center')],
    },
    'summary_title': {
        'fill': '107C41', 'font': {'bold': True, 'size': 16, 'color': 'FFFFFF'},
        'extent': (2, 6),
        'aligns': [(2, 6, 'Left')],
    },
    'summary_data': {
        'fill': None, 'font': {'bold': False, 'size': 11, 'color': '000000'},
        'extent': (2, 6),
        'aligns': [(2, 2, 'Left'), (3, 6, 'Center')],
    },
    'summary_total': {
        'fill': 'A9D08E', 'font': {'bold': True, 'size': 11, 'color': '000000'},
        'extent': (2, 6),
        'aligns': [(2, 2, 'Left'), (3, 6, 'Center')],
    },
}

# Leading spaces baked into column-B label text to simulate indentation
# per tree level, since the API has no real cell-indent property.
INDENT_SPACES = {
    'category': '',
    'module': '   ',
    'item': '      ',
    'suspect': '         ',
}

# A tree this size needs ~2500 sequential Graph API calls to style, which
# runs long enough that a transient Gateway Timeout / Service Unavailable
# is a real occurrence, not a hypothetical -- retry those instead of
# aborting the whole run over one flaky request.
_TRANSIENT_ERROR_MARKERS = ('502', '503', '504', 'Gateway Timeout', 'Service Unavailable', 'Bad Gateway')
_MAX_RETRIES = 5
_COL_LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'


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
#  Row plan: pure data -> (matrix, row_kind, row_level), no I/O
# ═══════════════════════════════════════════════════════════════════

def _plan_tree_rows(tree, matrix, row_kind, row_level):
    """Fill matrix/row_kind/row_level for the tree portion. Returns last row used."""
    row = DATA_START_ROW
    for category, module_rows in tree:
        matrix.setdefault(row, {})[2] = INDENT_SPACES['category'] + category
        matrix.setdefault(row, {})[4] = _count_links(module_rows)
        row_kind[row] = 'category'
        row += 1

        for module, item_rows in module_rows:
            module_links = sum(len(links) for _i, _n, links in item_rows)
            matrix.setdefault(row, {})[2] = INDENT_SPACES['module'] + module
            matrix.setdefault(row, {})[4] = module_links
            row_kind[row] = 'module'
            row_level[row] = LEVEL_MODULE
            row += 1

            for item_id, item_num, links in item_rows:
                matrix.setdefault(row, {})[2] = INDENT_SPACES['item'] + item_id
                matrix.setdefault(row, {})[3] = item_num
                matrix.setdefault(row, {})[4] = len(links)
                row_kind[row] = 'item'
                row_level[row] = LEVEL_ITEM
                row += 1

                for up_id, up_num in links:
                    matrix.setdefault(row, {})[2] = INDENT_SPACES['suspect'] + up_id
                    matrix.setdefault(row, {})[3] = up_num
                    row_kind[row] = 'suspect'
                    row_level[row] = LEVEL_SUSPECT
                    row += 1

    return row - 1


def _plan_summary_rows(report, start_row, matrix, row_kind):
    """Fill matrix/row_kind for the module summary block. Returns (title_row, total_row)."""
    title_row = start_row
    matrix.setdefault(title_row, {})[2] = 'Module Summary'
    row_kind[title_row] = 'summary_title'

    header_row = title_row + 2
    headers = ('Module', 'Total Items', 'Items With Suspects',
               'Suspect Relations', 'Impacted Items')
    for offset, label in enumerate(headers):
        matrix.setdefault(header_row, {})[2 + offset] = label
    row_kind[header_row] = 'summary_header'

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
        row_kind[row] = 'summary_data'
        row += 1

    total_row = row
    matrix.setdefault(total_row, {})[2] = 'Total'
    for offset, total in enumerate(column_totals):
        matrix.setdefault(total_row, {})[3 + offset] = total
    row_kind[total_row] = 'summary_total'

    return title_row, total_row


def build_row_plan(report, tree):
    """Return (last_row, matrix, row_kind, row_level, summary_title_row, last_tree_row).

    matrix is {row: {col: value}} for columns B-F (1-based col index).
    row_kind maps row -> style key (see STYLES). row_level maps row ->
    the outline level to stamp in column A (only set for tree rows below
    the always-visible category level). summary_title_row is the row
    the "Module Summary" banner needs merged across B:F. last_tree_row is
    the last row belonging to the tree itself, before the summary block.
    """
    matrix = {}
    row_kind = {}
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
    row_kind[HEADER_ROW] = 'header'

    last_tree_row = _plan_tree_rows(tree, matrix, row_kind, row_level)
    summary_title_row, _summary_total_row = _plan_summary_rows(
        report, last_tree_row + 2, matrix, row_kind)

    last_row = max(matrix)
    return last_row, matrix, row_kind, row_level, summary_title_row, last_tree_row


# ═══════════════════════════════════════════════════════════════════
#  Graph API write: turns the plan into workbook-session calls
# ═══════════════════════════════════════════════════════════════════

def _col_letter(n):
    return _COL_LETTERS[n - 1]


def _iter_runs(row_kind, last_row):
    """Yield (start_row, end_row, kind) for maximal contiguous same-kind runs."""
    r = HEADER_ROW
    while r <= last_row:
        kind = row_kind.get(r)
        if kind is None:
            r += 1
            continue
        start = r
        while r + 1 <= last_row and row_kind.get(r + 1) == kind:
            r += 1
        yield start, r, kind
        r += 1


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


def _apply_run_style(wb, start_row, end_row, kind):
    style = STYLES[kind]
    col_start, col_end = style['extent']
    rng = f"{_col_letter(col_start)}{start_row}:{_col_letter(col_end)}{end_row}"

    if style['fill'] is not None:
        wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{rng}')/format/fill",
           {'color': f"#{style['fill']}"})
    if style['font'] is not None:
        wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{rng}')/format/font",
           style['font'])
    for a_start, a_end, halign in style['aligns']:
        a_rng = f"{_col_letter(a_start)}{start_row}:{_col_letter(a_end)}{end_row}"
        wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{a_rng}')/format",
           {'horizontalAlignment': halign, 'verticalAlignment': 'Center'})


def _write_banner(wb, report):
    title_rng = "B1:D1"
    wb('POST', f"worksheets('{SHEET_NAME}')/range(address='{title_rng}')/merge", {})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{title_rng}')/format/fill",
       {'color': '#107C41'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{title_rng}')/format/font",
       {'bold': True, 'size': 16, 'color': 'FFFFFF'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{title_rng}')/format",
       {'horizontalAlignment': 'Left', 'verticalAlignment': 'Center', 'rowHeight': 30})

    subtitle_rng = "B2:D2"
    wb('POST', f"worksheets('{SHEET_NAME}')/range(address='{subtitle_rng}')/merge", {})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{subtitle_rng}')/format/font",
       {'italic': True, 'size': 9, 'color': '666666'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{subtitle_rng}')/format",
       {'horizontalAlignment': 'Left', 'verticalAlignment': 'Center', 'rowHeight': 18})


def _write_summary_title_merge(wb, title_row):
    rng = f"B{title_row}:F{title_row}"
    wb('POST', f"worksheets('{SHEET_NAME}')/range(address='{rng}')/merge", {})


def _set_column_widths(wb):
    widths = {'A': 4, 'B': 52, 'C': 18, 'D': 16, 'E': 20, 'F': 20}
    for col, width in widths.items():
        try:
            wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{col}1:{col}1')/format",
               {'columnWidth': width})
        except SharePointOperationError:
            pass  # cosmetic only


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

    last_row, matrix, row_kind, row_level, summary_title_row, last_tree_row = \
        build_row_plan(report, tree)

    existing = wb('GET', f"worksheets('{SHEET_NAME}')/usedRange?$select=rowCount")
    clear_through = max(last_row, existing.get('rowCount', 0), 200)
    _clear_sheet(wb, clear_through)

    _write_values(wb, last_row, matrix)
    _write_levels(wb, last_row, row_level)

    for start_row, end_row, kind in _iter_runs(row_kind, last_row):
        _apply_run_style(wb, start_row, end_row, kind)

    _write_banner(wb, report)
    _write_summary_title_merge(wb, summary_title_row)
    _set_column_widths(wb)
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
