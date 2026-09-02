"""
Build a Suspect Report sheet inside the live Sprint Management workbook
from a JAMA suspect report (JSON).

The sheet holds two tables. The Need Change Summary (rows 1-61) is fully
self-contained here: this script writes both its values *and* its
formatting via the Graph API, no VBA at all.

The suspect-links tree (Module Summary + tree, from row 63 down) splits
the work differently: this script writes the cell values *and* all of
that table's visual formatting (fills, fonts, hyperlinks, banner,
column widths -- see the "Tree + Module Summary formatting" section
below), but the tree's collapsible row grouping (outline levels) is
applied separately, inside the workbook, by suspect_macro.vb's
ApplyTreeGrouping. That one piece stays in VBA because the Graph API
has no direct row-outline-level setter, only incremental group()/
ungroup() actions -- everything else moved here specifically so the
sheet is fully colored/linked the moment this script finishes, with no
dependency on someone opening the workbook afterward for a macro to
run; only the grouping still needs that. The Module Summary block
itself is currently disabled (see WRITE_MODULE_SUMMARY) -- its
row-plan/formatting code is intentionally left in place, just unused.

If the report also has an 'Unlinked Downstream' section (items with no
downstream link at all, not just a suspect one), those items are
appended to the end of their matching module's item list (e.g. a
"PRS Without SRS" entry lands at the end of that category's PRS
module, not in a section of its own), with an 'x' suffix on their
column-A level ('2x' instead of '2') so the per-row formatting pass
below gives them an extra highlight on top of the normal item-row
formatting.

See the layout constants just below the imports for the exact row
ranges each table owns.
"""
import re
import json
import sys
import time
from datetime import datetime, timezone

from dotenv import load_dotenv

import roster
from config import SPRINT_TEMPLATE_FILE
from sharepoint_helper import SharePointHelper, SharePointOperationError

load_dotenv(r"c:\Scripts\Leahi Sprint Management Scripts\leahi-sprint-management\.env")

MODULE_ORDER = ('PRS', 'VeTC', 'SRS', 'LVTC')
SHEET_NAME = 'Suspect Report'

# Temporarily disabled 2026-08-31: adding VeTC's own row to the Module
# Summary block pushed it to end at row 70 -- exactly one row before
# the tree starts at TREE_TITLE_ROW (71), leaving zero margin. Rather
# than reshuffle the fixed row layout for a block whose necessity is
# being reconsidered, this just skips writing/formatting it entirely
# (build_row_plan and write_suspect_report both check this flag). The
# row-plan and formatting logic (_plan_summary_rows / _format_summary_
# block) is left fully intact, untouched, so this can be flipped back
# on later with no other code changes.
WRITE_MODULE_SUMMARY = False

# The sheet has two independent tables, stacked top to bottom, each
# owning a fixed row range so neither ever touches the other's rows:
#   - Need Change Summary: rows 1-NEED_CHANGE_LAST_ALLOWED_ROW (61) -- one
#     row per feature/category, columns B-F. See build_need_change_matrix.
#   - (row 62 is a deliberate blank gap between the two tables)
#   - Module Summary + tree: rows 63 on -- the original suspect-links
#     tree. Module Summary is fixed at SUMMARY_TITLE_ROW (63) but
#     currently disabled -- see WRITE_MODULE_SUMMARY above; the tree
#     is fixed at TREE_TITLE_ROW (71), not computed from the summary's
#     size, so it stays put regardless of how many modules the summary
#     lists; build_row_plan raises if the summary ever grows into it.
RESERVED_TOP_ROWS = 62
FIRST_MANAGED_ROW = RESERVED_TOP_ROWS + 1   # 63 -- start of the summary/tree table
SUMMARY_TITLE_ROW = FIRST_MANAGED_ROW       # 63
TREE_TITLE_ROW = 71
SUBTITLE_ROW = TREE_TITLE_ROW + 1           # 72
HEADER_ROW = TREE_TITLE_ROW + 2             # 73
DATA_START_ROW = HEADER_ROW + 1             # 74

NEED_CHANGE_TITLE_ROW = 1
NEED_CHANGE_SUBTITLE_ROW = 2
NEED_CHANGE_HEADER_ROW = 3
NEED_CHANGE_DATA_START_ROW = 4
NEED_CHANGE_LAST_ALLOWED_ROW = RESERVED_TOP_ROWS - 1   # 61 -- row 62 stays a blank gap

LEVEL_MODULE = 1
LEVEL_ITEM = 2
LEVEL_SUSPECT = 3
MARKER_LABEL_CELL = f'N{TREE_TITLE_ROW}'
BUILD_MARKER_CELL = f'O{TREE_TITLE_ROW}'
GROUPED_LABEL_CELL = f'N{SUBTITLE_ROW}'
GROUPED_MARKER_CELL = f'O{SUBTITLE_ROW}'
LAST_TREE_ROW_LABEL_CELL = f'N{HEADER_ROW}'
LAST_TREE_ROW_CELL = f'O{HEADER_ROW}'

# Leading spaces baked into column-B label text to simulate indentation
# per tree level, since the API has no real cell-indent property.
INDENT_SPACES = {
    'category': '',
    'module': '     ',
    'item': '          ',
    'suspect': '               ',
}

_TRANSIENT_ERROR_MARKERS = ('429', 'Too Many Requests', '502', '503', '504',
                            'Gateway Timeout', 'Service Unavailable', 'Bad Gateway',
                            'Read timed out', 'timed out', 'ConnectionError',
                            'Connection aborted', 'Max retries exceeded')
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
    """Return True if candidate maps categories to item groups."""
    if not isinstance(candidate, dict):
        return False
    for value in candidate.values():
        if not isinstance(value, dict):
            return False
    return True


def build_tree(report):
    """Return a nested tree of category -> module -> items.

    Two sources feed the same per-category, per-module item list:
    'Modules' (normal suspect links, as before) and the optional
    'Unlinked Downstream' section (items with *no* downstream link at
    all). A group there named e.g. "PRS Without SRS" is a list of PRS
    items with no SRS downstream link -- those items are displayed
    under the DOWNSTREAM module's section instead of their own: a "PRS
    Without SRS" item appears under SRS, not PRS; "SRS Without LVTC"
    appears under LVTC; "PRS Without VeTC" (added to the feed as of the
    2026-08-31 revision) appears under VeTC -- entirely independently of
    whichever of SRS/LVTC that same PRS item does or doesn't have a link
    to, since this is driven purely by which named groups exist in the
    report, keyed by item id per group, not by any single "this PRS
    item's one true missing module" notion. This is deliberate: when
    deciding what a downstream module needs to add, you look at that
    module's own tree section and see exactly which upstream items don't
    have a counterpart yet, right where you'd build the new item. They're
    appended to the end of that section's item list (after the normal,
    natural_key-sorted items), each individually flagged as missing so
    the row-plan/writer can mark it for _format_tree_rows to highlight.
    Older reports with no
    'Unlinked Downstream' key are unaffected.
    """
    modules = report.get('Modules')
    if not isinstance(modules, dict):
        raise SuspectReportError("Report is missing a 'Modules' object.")

    # category -> module -> {item_id: (item_num, item_name, last_editor, links)}
    # 'links' is [(up_id, up_num, up_name, up_last_editor), ...].
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
                links = [
                    (up_id, up_info.get('ID'), up_info.get('Name'),
                     up_info.get('Last Editor'))
                    for up_id, up_info
                    in info.get('Upstream Suspects', {}).items()]
                catalog.setdefault(category, {}).setdefault(
                    module, {})[item_id] = (
                        info.get('ID'), info.get('Name'),
                        info.get('Last Editor'), links)

    # category -> module -> {item_id: (item_num, item_name)} -- items with
    # no downstream link at all. The group label ("PRS Without SRS") names
    # the module that's MISSING the link ("SRS", after " Without ") --
    # that's where these items get displayed, not under their own actual
    # module ("PRS"). See the docstring above for why. This section of
    # the report doesn't carry a 'Last Editor' field, so these items
    # never get one (see the None fill-in below).
    missing = report.get('Unlinked Downstream')
    missing_catalog = {}
    if isinstance(missing, dict):
        for group_label, categories in missing.items():
            if not _is_category_mapping(categories):
                continue
            parts = group_label.split(' Without ')
            if len(parts) != 2:
                continue
            display_module = parts[1].strip()
            for category, items in categories.items():
                for item_id, info in items.items():
                    missing_catalog.setdefault(category, {}).setdefault(
                        display_module, {})[item_id] = (
                            info.get('ID'), info.get('Name'))

    module_rank = {name: pos for pos, name in enumerate(MODULE_ORDER)}
    all_categories = sorted(set(catalog) | set(missing_catalog))
    tree = []
    for category in all_categories:
        module_map = catalog.get(category, {})
        missing_map = missing_catalog.get(category, {})
        module_names = sorted(
            set(module_map) | set(missing_map),
            key=lambda m: module_rank.get(m, len(module_rank)))
        module_rows = []
        for module in module_names:
            item_map = module_map.get(module, {})
            item_rows = [
                (item_id, *item_map[item_id], False)
                for item_id in sorted(item_map, key=natural_key)]

            miss_map = missing_map.get(module, {})
            item_rows += [
                (item_id, miss_map[item_id][0], miss_map[item_id][1], None, [], True)
                for item_id in sorted(miss_map, key=natural_key)]

            module_rows.append((module, item_rows))
        tree.append((category, module_rows))
    return tree


# ═══════════════════════════════════════════════════════════════════
#  Need Change Summary: a second, independent table on the same sheet.
#  Unlike the tree above, this one's values *and* formatting are both
#  written here in Python -- no VBA macro involved for it at all.
# ═══════════════════════════════════════════════════════════════════

def _collect_categories(report):
    """Return every category name in the report, sorted.

    Pulled from both 'Modules' (has suspect links) and 'Unlinked
    Downstream' (missing links) -- the same two sources build_tree()
    reads -- so this is the full feature list, not just the ones that
    currently have a suspect link.
    """
    categories = set()
    for section_key in ('Modules', 'Unlinked Downstream'):
        section = report.get(section_key)
        if not isinstance(section, dict):
            continue
        for group in section.values():
            if isinstance(group, dict):
                categories.update(group.keys())
    return sorted(categories)


def build_need_change_matrix(report):
    """Return the {row: {col: value}} matrix for the Need Change Summary,
    and its last used row.

    Raises SuspectReportError if the feature list is too long for the
    fixed rows reserved for it (NEED_CHANGE_DATA_START_ROW through
    NEED_CHANGE_LAST_ALLOWED_ROW).
    """
    categories = _collect_categories(report)
    last_row = NEED_CHANGE_DATA_START_ROW + len(categories) - 1
    if last_row > NEED_CHANGE_LAST_ALLOWED_ROW:
        raise SuspectReportError(
            f"Need Change Summary has {len(categories)} features, which would "
            f"run through row {last_row} -- past row "
            f"{NEED_CHANGE_LAST_ALLOWED_ROW}, the last one reserved for it. "
            "Raise NEED_CHANGE_LAST_ALLOWED_ROW (and shift the tables below "
            "it down) before writing."
        )

    matrix = {}
    matrix.setdefault(NEED_CHANGE_TITLE_ROW, {})[2] = 'LEAHI - Need Change Summary'
    matrix.setdefault(NEED_CHANGE_SUBTITLE_ROW, {})[2] = (
        f"Project ID: {report.get('Project ID', '')}     "
        f"Generated: {report.get('Generated At', '')}")

    headers = ('Category / Module / Item / Suspect', 'PRS Needs Change',
               'VeTC Needs Change', 'SRS Needs Change', 'LVTC Needs Change')
    for offset, label in enumerate(headers):
        matrix.setdefault(NEED_CHANGE_HEADER_ROW, {})[2 + offset] = label

    for offset, category in enumerate(categories):
        matrix.setdefault(NEED_CHANGE_DATA_START_ROW + offset, {})[2] = category

    return matrix, last_row


NEED_CHANGE_YES_TEXT = 'Yes'
NEED_CHANGE_YES_FILL = '#197c0a'   # a feature that needs a change, unclaimed
NEED_CHANGE_NO_FILL = '#d9d9d9'    # a feature that doesn't need a change

# ═══════════════════════════════════════════════════════════════════
#  Tree + Module Summary formatting -- ported 1:1 from what used to be
#  suspect_macro.vb's FormatBanner / per-row loop / FormatSummary /
#  AddJamaHyperlink / ApplyColumnWidths. suspect_macro.vb still exists
#  and still runs -- it now only sets the tree's row-outline levels
#  (ApplyTreeGrouping), since the Graph API has no direct way to set
#  those. Every color, font, and alignment value below matches the old
#  macro exactly; see suspect-report-formatting-bug project memory for
#  the debugging history that led to this split (a
#  ws.Outline.ShowLevels(1) call was hiding freshly-colored rows,
#  unrelated to any of the values themselves).
# ═══════════════════════════════════════════════════════════════════
TREE_BANNER_FILL = '#385899'         # tree title banner + category rows
TREE_HEADER_FILL = '#FFC000'         # tree column header row
TREE_SUBTITLE_COLOR = '#666666'
TREE_MODULE_FILL = '#197C0A'
TREE_ITEM_FILL = '#FCE4D6'           # normal item (has a real downstream link)
TREE_MISSING_FILL = '#E2EFDA'        # item appended from Unlinked Downstream ('2x')
TREE_SUSPECT_STRIPE_FILL = '#FFFFFF'
TREE_LINK_COLOR = '#0563C1'          # standard hyperlink blue
SUMMARY_TITLE_FILL = '#107C41'
SUMMARY_HEADER_FILL = '#70AD47'
SUMMARY_TOTAL_FILL = '#A9D08E'
JAMA_PROJECT_ID = 47                 # must stay in sync with the JAMA feed script's PROJECT_ID

TREE_COLUMN_WIDTHS_PX = {
    # Column A only carries the outline-level numbers this sheet's
    # tree logic reads, so it's collapsed to zero width (hidden).
    'A': 0, 'B': 550, 'C': 110, 'D': 135, 'E': 135, 'F': 150,
}


def mark_feature_needs_change(current_value, assignee_nickname=None):
    """Return the (value, fill) to write into a Needs-Change cell that
    the logic has determined should be "Yes" -- or None if the cell
    must be left completely untouched.

    Refuses to write only when the cell holds something OTHER than the
    literal "Yes" this same logic would itself have written -- i.e. a
    team member typed their own name in to claim the task. Overwriting
    a real claim would erase it and leave no record of who's doing the
    work, so the caller must skip the write entirely (value AND fill)
    when this returns None. A blank cell, or a plain "Yes" left over
    from before assignee_nickname existed (or from a run where no
    single editor could be pinned down), is fair game to (re)write --
    otherwise an already-answered feature could never be upgraded from
    "Yes" to a name once its data resolves to one.

    If assignee_nickname is given -- meaning every item's own last
    editor in this module's section of the feature traces back to the
    same person (see _sole_editor_nickname) -- that
    nickname is written in place of the literal text "Yes", still with
    the same green fill, so an auto-attributed cell reads exactly like
    a manually claimed one. Once a cell holds an actual person's name,
    it stops being touched by this function (that name could be this
    logic's own past output or a manual claim; either way, treating it
    as claimed avoids flip-flopping the cell every run as the sole
    editor's identity shifts).
    """
    if current_value not in (None, '', NEED_CHANGE_YES_TEXT):
        return None
    return (assignee_nickname or NEED_CHANGE_YES_TEXT, NEED_CHANGE_YES_FILL)


def mark_feature_no_change():
    """Return the (value, fill) to write into a Needs-Change cell that
    the logic has determined should be "No": always blank with the gray
    fill. There's no assignee to protect for a "doesn't need change"
    result, so unlike mark_feature_needs_change this is unconditional.
    """
    return ('', NEED_CHANGE_NO_FILL)


def _read_need_change_snapshot(wb):
    """Return {feature_name: (c_val, d_val, e_val)} for whatever is
    currently in the Need Change Summary's C/D/E columns.

    Keyed by feature name rather than row number: the feature list is
    re-derived and re-sorted from the live report on every run, so as
    features get added or removed, everything below shifts rows. A
    snapshot keyed by row position would silently attach one feature's
    old answer/assignee to a completely different feature after a
    reshuffle -- keying by name is what makes carrying it forward safe.
    """
    result = wb('GET',
        f"worksheets('{SHEET_NAME}')/range(address='B{NEED_CHANGE_DATA_START_ROW}:"
        f"F{NEED_CHANGE_LAST_ALLOWED_ROW}')?$select=values")
    snapshot = {}
    for row_values in result.get('values', []):
        name = row_values[0] if row_values else None
        if not name:
            continue
        padded = list(row_values) + [None] * (5 - len(row_values))
        snapshot[name] = (padded[1], padded[2], padded[3], padded[4])
    return snapshot


def _compute_category_trouble(report):
    """Return {category: {'PRS': bool, 'SRS': bool, 'LVTC': bool,
    'VeTC': bool}} -- per-feature, per-module trouble flags feeding the
    Need Change Summary's cascade (see _need_change_flags).

    The domain relationship is PRS -> SRS -> LVTC (Product/Requirement
    Spec -> Software Requirement Spec -> test cases) plus a separate
    branch PRS -> VeTC: a problem in PRS can force changes anywhere
    downstream of it (SRS, LVTC, AND VeTC); a problem in SRS only
    forces changes further down ITS OWN chain (LVTC), never VeTC, since
    VeTC branches directly off PRS, not off SRS.

    A module is "trouble" for a category if ANY of its items in that
    category is suspect (has a non-empty 'Upstream Suspects' dict), OR
    if the module it's missing a link FROM has an item in the matching
    Unlinked Downstream group -- i.e. a missing link is attributed to
    the module that's missing the connection, not the one that failed
    to produce it:
      - PRS: is suspect. (No missing-link case attributes trouble TO
        PRS -- PRS is the top of the chain, nothing is ever missing
        "from" it in this sense.)
      - SRS: is suspect, OR a PRS item appears in the 'PRS Without SRS'
        Unlinked Downstream group (an SRS item is missing because its
        upstream PRS item never linked down to it).
      - LVTC: is suspect, OR an SRS item appears in 'SRS Without LVTC'
        (an LVTC item is missing because its upstream SRS item never
        linked down to it).
      - VeTC: is suspect, OR a PRS item appears in 'PRS Without VeTC'
        (a VeTC item is missing because its upstream PRS item never
        linked down to it). This check is entirely independent of SRS/
        LVTC's own trouble -- a PRS item can have a perfectly good SRS
        link and still be missing its VeTC link (and vice versa), so a
        category can be VeTC-trouble without being SRS-trouble at all,
        or the reverse. On an older report with no 'PRS Without VeTC'
        group at all, this half of the check is simply a no-op (the
        group is never found in 'Unlinked Downstream'), not an error.
    A single triggering item anywhere in a category is enough -- this
    is a per-feature flag, not a per-item one.
    """
    trouble = {}

    def flag(category, module):
        trouble.setdefault(category, {})[module] = True

    modules = report.get('Modules')
    if isinstance(modules, dict):
        for module, categories in modules.items():
            if module not in MODULE_ORDER or not isinstance(categories, dict):
                continue
            for category, items in categories.items():
                if not isinstance(items, dict):
                    continue
                for info in items.values():
                    if isinstance(info, dict) and info.get('Upstream Suspects'):
                        flag(category, module)

    unlinked = report.get('Unlinked Downstream')
    if isinstance(unlinked, dict):
        # Maps each Unlinked Downstream group to the module ON THE
        # RECEIVING END of the missing link (see the docstring above) --
        # NOT the module named first in the group label. Each mapping is
        # independent of the others: 'PRS Without VeTC' flags VeTC only,
        # never SRS/LVTC, even for the exact same PRS item that also
        # shows up (or doesn't) in 'PRS Without SRS'.
        group_to_module = {
            'PRS Without SRS': 'SRS',
            'SRS Without LVTC': 'LVTC',
            'PRS Without VeTC': 'VeTC',
        }
        for group_label, categories in unlinked.items():
            module = group_to_module.get(group_label)
            if module is None or not isinstance(categories, dict):
                continue
            for category, items in categories.items():
                if isinstance(items, dict) and items:
                    flag(category, module)

    return trouble


def _module_editors(report, category, module):
    """Return the set of every 'Last Editor' value on module's own items
    in category.

    Deliberately ignores each item's upstream suspects' own 'Last
    Editor' -- those belong to a *different* module (e.g. an SRS item's
    upstream suspects are PRS items), and whoever last touched THAT
    module's artifact has no bearing on who owns fixing THIS module's
    item. Every item under report['Modules'][module][category]
    necessarily has at least one upstream suspect already (see
    generate_json_report in the JAMA feed script -- items with none are
    excluded from the report entirely), so this is really "every
    flagged item's own editor," not a filtered subset.

    Reads only report['Modules'][module][category] -- the module's own
    section -- never a different module's, even when that module's Yes
    was forced by a cascade (e.g. PRS trouble forcing SRS to Yes): the
    question here is who last touched the SRS material itself, not who
    caused the PRS problem.
    """
    editors = set()
    modules = report.get('Modules')
    if not isinstance(modules, dict):
        return editors
    items = modules.get(module, {}).get(category, {})
    if not isinstance(items, dict):
        return editors
    for info in items.values():
        if isinstance(info, dict):
            editors.add(info.get('Last Editor'))
    return editors


def _sole_editor_nickname(report, category, module):
    """Return the nickname of the one person who last-edited every one
    of module's own items in category -- or None if there's no such
    data, more than one distinct editor is involved, or any editor
    value is missing/blank.

    None means "just write the plain Yes" -- the caller falls back to
    that whenever this can't pin the whole section on a single person.
    """
    editors = _module_editors(report, category, module)
    if len(editors) != 1:
        return None
    (editor,) = editors
    if not editor:
        return None
    return roster._assignee_nickname(editor)


def _need_change_flags(trouble):
    """Return (prs_yes, srs_yes, lvtc_yes, vetc_yes): the cascade that
    turns a category's raw per-module trouble dict (see
    _compute_category_trouble) into the Need Change Summary's four column
    decisions.

    PRS trouble forces every other column to Yes (SRS, LVTC, and VeTC
    all trace back to PRS). SRS trouble forces LVTC only (LVTC is
    downstream of SRS; VeTC is NOT -- it branches directly off PRS, so
    an SRS-only problem never affects it). VeTC trouble is local to
    VeTC. LVTC trouble is local to LVTC.
    """
    prs_trouble = trouble.get('PRS', False)
    srs_trouble = trouble.get('SRS', False)
    lvtc_trouble = trouble.get('LVTC', False)
    vetc_trouble = trouble.get('VeTC', False)

    prs_yes = prs_trouble
    srs_yes = prs_trouble or srs_trouble
    lvtc_yes = prs_trouble or srs_trouble or lvtc_trouble
    vetc_yes = prs_trouble or vetc_trouble
    return prs_yes, srs_yes, lvtc_yes, vetc_yes


def _write_need_change_map(wb, report):
    """Fully rebuild the Need Change Summary table every run.

    The feature list changes often (features get added/removed in
    JAMA), so this always clears and rewrites the whole table rather
    than trying to patch it incrementally. Each feature's C/D/E/F
    (PRS/SRS/LVTC/VeTC Needs Change) cells are decided by
    _compute_category_trouble + _need_change_flags, then written via
    mark_feature_needs_change / mark_feature_no_change:
      - A column decided "Yes" only overwrites an EMPTY cell -- an
        existing "Yes" or a team member's own name (claiming the task)
        is left alone (mark_feature_needs_change returns None for a
        non-empty cell).
      - A column decided "No" ALWAYS overwrites, even a claimed
        assignee's name -- once the computed logic says a module no
        longer needs a change, any prior answer (including a claim) is
        stale (mark_feature_no_change is unconditional; see its
        docstring).
    The prior snapshot is read first (via _read_need_change_snapshot)
    and matched by feature name so a feature that moves to a different
    row (because the alphabetical list shifted) keeps its own answer
    rather than inheriting whatever row it now lands on.
    """
    snapshot = _read_need_change_snapshot(wb)
    category_trouble = _compute_category_trouble(report)

    matrix, last_row = build_need_change_matrix(report)
    last_col_letter = 'F'  # B=feature, C/D/E/F=PRS/VeTC/SRS/LVTC Needs Change
    col_to_module = {3: 'PRS', 4: 'VeTC', 5: 'SRS', 6: 'LVTC'}

    cell_fill = {}
    for row in range(NEED_CHANGE_DATA_START_ROW, last_row + 1):
        feature = matrix.get(row, {}).get(2)
        prior = snapshot.get(feature, (None, None, None, None))
        trouble = category_trouble.get(feature, {})
        # _need_change_flags always returns (prs, srs, lvtc, vetc), in that
        # fixed logical order -- its computation is untouched. Reordered
        # here only, to line up with the C/D/E/F column order above
        # (PRS/VeTC/SRS/LVTC), which is a display-only choice.
        prs_yes, srs_yes, lvtc_yes, vetc_yes = _need_change_flags(trouble)
        computed_yes = (prs_yes, vetc_yes, srs_yes, lvtc_yes)

        for col, current_value, is_yes in zip((3, 4, 5, 6), prior, computed_yes):
            if is_yes:
                nickname = _sole_editor_nickname(report, feature, col_to_module[col])
                result = mark_feature_needs_change(current_value, nickname)
                if result is None:
                    # Already claimed by an actual person's name (not a
                    # plain "Yes" -- that's fair game, see
                    # mark_feature_needs_change) -- keep it, and keep it
                    # green.
                    matrix[row][col] = current_value
                    cell_fill[(row, col)] = NEED_CHANGE_YES_FILL
                    continue
            else:
                result = mark_feature_no_change()

            value, fill = result
            matrix[row][col] = value
            cell_fill[(row, col)] = fill

    # ── Full rebuild: clear the whole reserved block, then rewrite ──
    full_rng = f"B{NEED_CHANGE_TITLE_ROW}:{last_col_letter}{NEED_CHANGE_LAST_ALLOWED_ROW}"
    wb('POST', f"worksheets('{SHEET_NAME}')/range(address='{full_rng}')/clear",
       {'applyTo': 'All'})

    values = []
    for row in range(NEED_CHANGE_TITLE_ROW, last_row + 1):
        cells = matrix.get(row, {})
        values.append([cells.get(col) for col in range(2, 7)])
    wb('PATCH',
       f"worksheets('{SHEET_NAME}')/range(address='B{NEED_CHANGE_TITLE_ROW}:"
       f"{last_col_letter}{last_row}')",
       {'values': values})

    title_rng = f"B{NEED_CHANGE_TITLE_ROW}:{last_col_letter}{NEED_CHANGE_TITLE_ROW}"
    subtitle_rng = f"B{NEED_CHANGE_SUBTITLE_ROW}:{last_col_letter}{NEED_CHANGE_SUBTITLE_ROW}"
    header_rng = f"B{NEED_CHANGE_HEADER_ROW}:{last_col_letter}{NEED_CHANGE_HEADER_ROW}"
    data_rng = f"B{NEED_CHANGE_DATA_START_ROW}:{last_col_letter}{last_row}"

    # ── Title banner (mirrors the tree's title look) ──
    wb('POST', f"worksheets('{SHEET_NAME}')/range(address='{title_rng}')/merge",
       {'across': False})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{title_rng}')/format/font",
       {'bold': True, 'size': 16, 'color': '#FFFFFF'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{title_rng}')/format/fill",
       {'color': '#385899'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{title_rng}')/format",
       {'horizontalAlignment': 'Left', 'verticalAlignment': 'Center', 'rowHeight': 30})

    # ── Subtitle ──
    wb('POST', f"worksheets('{SHEET_NAME}')/range(address='{subtitle_rng}')/merge",
       {'across': False})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{subtitle_rng}')/format/font",
       {'italic': True, 'size': 9, 'color': '#666666'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{subtitle_rng}')/format",
       {'horizontalAlignment': 'Left', 'verticalAlignment': 'Center', 'rowHeight': 18})

    # ── Column header row ──
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{header_rng}')/format/font",
       {'bold': True, 'size': 10, 'color': '#FFFFFF'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{header_rng}')/format/fill",
       {'color': '#FFC000'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{header_rng}')/format",
       {'verticalAlignment': 'Center'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='B{NEED_CHANGE_HEADER_ROW}')/format",
       {'horizontalAlignment': 'Left'})
    wb('PATCH',
       f"worksheets('{SHEET_NAME}')/range(address='C{NEED_CHANGE_HEADER_ROW}:"
       f"{last_col_letter}{NEED_CHANGE_HEADER_ROW}')/format",
       {'horizontalAlignment': 'Center'})

    # ── Feature rows: blanket "undecided" look across the whole row
    # first (matches the tree's category-row look) ──
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{data_rng}')/format/font",
       {'bold': True, 'size': 11, 'color': '#FFFFFF'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{data_rng}')/format/fill",
       {'color': '#385899'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{data_rng}')/format",
       {'verticalAlignment': 'Center'})
    wb('PATCH',
       f"worksheets('{SHEET_NAME}')/range(address='B{NEED_CHANGE_DATA_START_ROW}:"
       f"B{last_row}')/format",
       {'horizontalAlignment': 'Left'})
    wb('PATCH',
       f"worksheets('{SHEET_NAME}')/range(address='C{NEED_CHANGE_DATA_START_ROW}:"
       f"{last_col_letter}{last_row}')/format",
       {'horizontalAlignment': 'Center'})

    # ── Then apply each cell's actual decided fill (green for a Yes or
    # a claimed assignee, gray for a No) on top of the blanket
    # "undecided" look computed above. Every feature/column gets a real
    # decision now (see the loop that built cell_fill), so nothing
    # should be left in that blanket look once this finishes. ──
    col_letters = {3: 'C', 4: 'D', 5: 'E', 6: 'F'}
    for (row, col), fill in cell_fill.items():
        wb('PATCH',
           f"worksheets('{SHEET_NAME}')/range(address='{col_letters[col]}{row}')/format/fill",
           {'color': fill})

    return last_row


def _labeled(key, name):
    """Return the column-B label 'key' as '[key] name' (or '[key]').

    'key' is the alphanumeric id (e.g. LEAHI-dFMEA-113); 'name' is the
    item/suspect title from the report.  The bracketed id keeps the id
    scannable now that the title follows it.
    """
    name = (name or '').strip()
    if name:
        return f'[{key}] {name}'
    return f'[{key}]'


# ═══════════════════════════════════════════════════════════════════
#  Row plan: pure data -> (matrix, row_level), no I/O
# ═══════════════════════════════════════════════════════════════════

def _leveled(base_level, is_missing):
    """Return the column-A level value: the plain int, or 'Nx' if missing.

    The formatting pass below reads the trailing 'x' to add a highlight
    on top of whatever formatting the base level (module/item/suspect)
    already gets -- the row is still that level, just also flagged as
    missing.
    """
    return f'{base_level}x' if is_missing else base_level


def _parse_level(raw_value):
    """Return (base_level, is_missing) for a column-A level value.

    Duplicates suspect_macro.vb's own ParseLevel (that copy still runs,
    for the row-outline-level pass this script can't do via the Graph
    API): a plain 1/2/3 for a normal module/item/suspect row, or
    '1x'/'2x'/'3x' for the same level flagged missing by _leveled().
    None or blank (category rows) parses to base_level 0.
    """
    if raw_value is None:
        return 0, False
    s = str(raw_value).strip()
    if not s:
        return 0, False
    is_missing = s[-1] in ('x', 'X')
    if is_missing:
        s = s[:-1]
    try:
        return int(s), is_missing
    except ValueError:
        return 0, is_missing


def _plan_tree_rows(tree, matrix, row_level):
    """Fill matrix/row_level for the tree portion. Returns last row used."""
    row = DATA_START_ROW
    for category, module_rows in tree:
        matrix.setdefault(row, {})[2] = INDENT_SPACES['category'] + category
        row += 1

        for module, item_rows in module_rows:
            module_links = sum(len(links) for _i, _n, _nm, _le, links, _m in item_rows)
            matrix.setdefault(row, {})[2] = INDENT_SPACES['module'] + module
            matrix.setdefault(row, {})[4] = module_links
            row_level[row] = LEVEL_MODULE
            row += 1

            for item_id, item_num, item_name, last_editor, links, is_missing in item_rows:
                matrix.setdefault(row, {})[2] = (
                    INDENT_SPACES['item'] + _labeled(item_id, item_name))
                matrix.setdefault(row, {})[3] = item_num
                matrix.setdefault(row, {})[5] = last_editor
                row_level[row] = _leveled(LEVEL_ITEM, is_missing)
                row += 1

                for up_id, up_num, up_name, up_last_editor in links:
                    matrix.setdefault(row, {})[2] = (
                        INDENT_SPACES['suspect'] + _labeled(up_id, up_name))
                    matrix.setdefault(row, {})[3] = up_num
                    matrix.setdefault(row, {})[5] = up_last_editor
                    row_level[row] = LEVEL_SUSPECT
                    row += 1

    return row - 1


def _plan_summary_rows(report, start_row, matrix):
    """Fill matrix for the module summary block (values only).

    Returns the block's last (Total) row, so the caller can confirm it
    didn't grow into the space reserved for whatever comes after it.
    """
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

    return total_row


def build_row_plan(report, tree):
    """Return (last_row, matrix, row_level, last_tree_row, summary_last_row).

    summary_last_row is None when WRITE_MODULE_SUMMARY is False -- see
    that flag's comment near the top of the file.
    """
    matrix = {}
    row_level = {}

    # Module Summary is planned first: it sits above the tree now, at a
    # fixed row, so a future table above it can rely on the tree always
    # starting at TREE_TITLE_ROW regardless of the summary's own size.
    summary_last_row = None
    if WRITE_MODULE_SUMMARY:
        summary_last_row = _plan_summary_rows(report, SUMMARY_TITLE_ROW, matrix)
        if summary_last_row >= TREE_TITLE_ROW:
            raise SuspectReportError(
                f"Module Summary block runs through row {summary_last_row}, which "
                f"reaches the fixed tree start at row {TREE_TITLE_ROW}. Move "
                "TREE_TITLE_ROW down or shrink the summary before writing."
            )

    relation_counts = report.get('Suspect Relation Counts', {})
    counts_text = ' / '.join(
        f'{name} {relation_counts.get(name, 0)}'
        for name in MODULE_ORDER if name in relation_counts)
    matrix.setdefault(TREE_TITLE_ROW, {})[2] = 'LEAHI - Upstream Suspect Links by Category'
    has_missing = isinstance(report.get('Unlinked Downstream'), dict) and any(
        report['Unlinked Downstream'].values())
    missing_note = "     Green-filled rows: no downstream link at all" if has_missing else ""
    matrix.setdefault(SUBTITLE_ROW, {})[2] = (
        f"Project ID: {report.get('Project ID', '')}     "
        f"Generated: {report.get('Generated At', '')}     "
        f"Suspect links: {counts_text}{missing_note}")

    headers = ('Category / Module / Item / Suspect', 'Numeric ID',
               'Suspect Links', 'Last Editor')
    for offset, label in enumerate(headers):
        matrix.setdefault(HEADER_ROW, {})[2 + offset] = label

    last_tree_row = _plan_tree_rows(tree, matrix, row_level)

    last_row = max(matrix)
    return last_row, matrix, row_level, last_tree_row, summary_last_row


# ═══════════════════════════════════════════════════════════════════
#  Graph API write: turns the plan into workbook-session calls
# ═══════════════════════════════════════════════════════════════════

def _ensure_sheet(wb):
    sheets = wb('GET', 'worksheets')['value']
    if not any(s['name'] == SHEET_NAME for s in sheets):
        wb('POST', 'worksheets/add', {'name': SHEET_NAME})


def _clear_sheet(wb, clear_through_row):
    # Never below FIRST_MANAGED_ROW: rows 1-62 belong to the Need Change
    # Map table, which has its own clear/write path in
    # _write_need_change_map and must not be touched here.
    wb('POST',
       f"worksheets('{SHEET_NAME}')/range(address='A{FIRST_MANAGED_ROW}:F{clear_through_row}')/clear",
       {'applyTo': 'All'})


def _write_values(wb, last_row, matrix):
    values = []
    for row in range(FIRST_MANAGED_ROW, last_row + 1):
        cells = matrix.get(row, {})
        values.append([cells.get(col) for col in range(1, 7)])
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='A{FIRST_MANAGED_ROW}:F{last_row}')",
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


# ═══════════════════════════════════════════════════════════════════
#  Tree + Module Summary formatting -- ported from what used to be
#  suspect_macro.vb's own formatting subs (see the constants block
#  above for the color values, each labeled with which VBA sub it came
#  from). suspect_macro.vb itself still exists; it now only sets row
#  outline levels (ApplyTreeGrouping), which the Graph API can't do
#  directly.
# ═══════════════════════════════════════════════════════════════════

def _format_tree_banner(wb):
    """Style the tree's title/subtitle/header rows. Ported from
    suspect_macro.vb's FormatBanner."""
    title_rng = f"B{TREE_TITLE_ROW}:E{TREE_TITLE_ROW}"
    subtitle_rng = f"B{SUBTITLE_ROW}:E{SUBTITLE_ROW}"
    header_rng = f"B{HEADER_ROW}:E{HEADER_ROW}"

    wb('POST', f"worksheets('{SHEET_NAME}')/range(address='{title_rng}')/merge",
       {'across': False})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{title_rng}')/format/fill",
       {'color': TREE_BANNER_FILL})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{title_rng}')/format/font",
       {'bold': True, 'size': 16, 'color': '#FFFFFF'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{title_rng}')/format",
       {'horizontalAlignment': 'Left', 'verticalAlignment': 'Center', 'rowHeight': 30})

    wb('POST', f"worksheets('{SHEET_NAME}')/range(address='{subtitle_rng}')/merge",
       {'across': False})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{subtitle_rng}')/format/font",
       {'italic': True, 'size': 9, 'color': TREE_SUBTITLE_COLOR})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{subtitle_rng}')/format",
       {'horizontalAlignment': 'Left', 'verticalAlignment': 'Center', 'rowHeight': 18})

    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{header_rng}')/format/fill",
       {'color': TREE_HEADER_FILL})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{header_rng}')/format/font",
       {'bold': True, 'size': 10, 'color': '#FFFFFF'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{header_rng}')/format",
       {'verticalAlignment': 'Center'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='B{HEADER_ROW}')/format",
       {'horizontalAlignment': 'Left'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='C{HEADER_ROW}:E{HEADER_ROW}')/format",
       {'horizontalAlignment': 'Center'})


def _contiguous_runs_from_list(rows):
    """Return maximal contiguous (start, end) row-number pairs from a
    sorted list of row numbers -- e.g. [74, 75, 76, 90] -> [(74, 76),
    (90, 90)]."""
    runs = []
    start = prev = None
    for row in rows:
        if start is None:
            start = prev = row
        elif row == prev + 1:
            prev = row
        else:
            runs.append((start, prev))
            start = prev = row
    if start is not None:
        runs.append((start, prev))
    return runs


def _hyperlink_formula_for_row(row, matrix):
    """Return the JAMA HYPERLINK() formula for row's column B, or None if
    it doesn't have a numeric id (column C) to link to."""
    numeric_id = matrix.get(row, {}).get(3)
    try:
        numeric_id = int(numeric_id)
    except (TypeError, ValueError):
        return None

    label = matrix.get(row, {}).get(2) or ''
    label_escaped = label.replace('"', '""')
    url = (f"https://diality-prod.jamacloud.com/perspective.req#/items/"
           f"{numeric_id}?projectId={JAMA_PROJECT_ID}")
    return f'=HYPERLINK("{url}","{label_escaped}")'


def _write_tree_hyperlink_formulas(wb, matrix, row_level, last_tree_row):
    """Write column B's ENTIRE cell content in one PATCH: a JAMA
    HYPERLINK() formula for item/suspect rows with a numeric id
    (_hyperlink_formula_for_row), plain label text everywhere else.
    Ported from suspect_macro.vb's AddJamaHyperlink, restructured for
    reliability (see _format_tree_rows). Returns the sorted list of rows
    that got a formula, so the caller can style just those cells
    afterward via _style_tree_hyperlinks.

    Deliberately called FIRST in _format_tree_rows, before any fill/font
    call -- writing cell content is not guaranteed to leave existing
    formatting alone, so doing it before any formatting runs means the
    formatting calls that follow are always the last word regardless.
    """
    formulas = []
    linked_rows = []
    for row in range(DATA_START_ROW, last_tree_row + 1):
        base_lvl, _ = _parse_level(row_level.get(row))
        formula = _hyperlink_formula_for_row(row, matrix) if base_lvl in (2, 3) else None
        if formula is not None:
            formulas.append([formula])
            linked_rows.append(row)
        else:
            formulas.append([matrix.get(row, {}).get(2) or ''])

    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='B{DATA_START_ROW}:B{last_tree_row}')",
       {'formulas': formulas})
    return linked_rows


def _style_tree_hyperlinks(wb, linked_rows):
    """Apply the link look (Calibri/10pt/blue/underline) to every
    contiguous run of linked rows' column-B cells, one PATCH per run
    instead of one per row. Deliberately called LAST in
    _format_tree_rows, after every other fill/font call -- this must be
    the final word on these specific cells' font, matching
    suspect_macro.vb's original AddJamaHyperlink, which always ran after
    (and so overrode) the per-row level/font branch's Font.Bold. Calling
    it any earlier would have this styling wiped out by the bulk font
    reset that follows (it explicitly sets color/underline back to
    black/none).
    """
    for start, end in _contiguous_runs_from_list(linked_rows):
        wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='B{start}:B{end}')/format/font",
           {'name': 'Calibri', 'size': 10, 'color': TREE_LINK_COLOR, 'underline': 'Single'})


def _format_tree_rows(wb, matrix, row_level, last_tree_row):
    """Apply the tree's per-row colors/fonts. Ported from
    suspect_macro.vb's bulk-defaults step plus its main per-row loop.
    Suspect rows (level 3) stripe alternating TREE_SUSPECT_STRIPE_FILL /
    no fill, restarting at each item's block; a level of 'Nx' (see
    _leveled) means the row is a normal level but was appended from the
    report's Unlinked Downstream section, so its fill is overridden to
    TREE_MISSING_FILL afterward regardless of level.

    Restructured for reliability, same as _write_tree_hyperlink_formulas:
    rather than one fill/font PATCH per row, every row's (fill, font)
    pair is computed first, then merged into contiguous runs of
    identical values before any Graph calls are issued -- a fill/font
    PATCH per row for a tree with hundreds of rows is slow enough to
    reliably hit Graph API read timeouts partway through a run. This
    produces identical final cell formatting to the original per-row
    version; only the number of Graph calls changes (fewer when the
    tree happens to have runs of identical-level rows back to back,
    e.g. several Unlinked Downstream items appended one after another).

    Call order matters here: _write_tree_hyperlink_formulas runs FIRST
    (writing column B's cell content is not guaranteed to leave existing
    formatting alone, so every fill/font call below needs to come after
    it to be the true last word), while _style_tree_hyperlinks runs LAST
    (it must override the per-row font branch's Bold, matching the
    original VBA precedence, and would itself get wiped out by the bulk
    font reset below if called any earlier).
    """
    linked_rows = _write_tree_hyperlink_formulas(wb, matrix, row_level, last_tree_row)

    band_all = f"B{DATA_START_ROW}:E{last_tree_row}"

    wb('POST', f"worksheets('{SHEET_NAME}')/range(address='{band_all}')/format/fill/clear")
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{band_all}')/format/font",
       {'bold': False, 'size': 10, 'color': '#000000', 'underline': 'None'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{band_all}')/format",
       {'verticalAlignment': 'Center'})

    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='B{DATA_START_ROW}:B{last_tree_row}')/format",
       {'horizontalAlignment': 'Left'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='C{DATA_START_ROW}:D{last_tree_row}')/format",
       {'horizontalAlignment': 'Center'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='E{DATA_START_ROW}:E{last_tree_row}')/format",
       {'horizontalAlignment': 'Left'})

    prev_base_lvl = -1
    stripe_on = False
    row_fill = {}
    row_font = {}     # values are hashable tuples so identical fonts compare equal
    row_missing = []
    for row in range(DATA_START_ROW, last_tree_row + 1):
        base_lvl, is_missing = _parse_level(row_level.get(row))
        fill = None
        font = None

        if base_lvl == 1:      # module
            fill = TREE_MODULE_FILL
            font = (('bold', True), ('color', '#FFFFFF'))
        elif base_lvl == 2:    # item (or missing item, appended at the end)
            fill = TREE_ITEM_FILL
            font = (('bold', True),)
        elif base_lvl == 3:    # suspect
            if prev_base_lvl != 3:
                stripe_on = False   # first suspect row under an item: no stripe
            if stripe_on:
                fill = TREE_SUSPECT_STRIPE_FILL
            stripe_on = not stripe_on
        else:                  # category (no level number in column A)
            fill = TREE_BANNER_FILL
            font = (('bold', True), ('color', '#FFFFFF'), ('size', 11))

        row_fill[row] = fill
        row_font[row] = font
        if is_missing:
            row_missing.append(row)
        prev_base_lvl = base_lvl

    def emit_runs(value_by_row, apply):
        """Merge consecutive rows sharing an identical (non-None) value
        into single (start, end, value) calls."""
        rows = sorted(value_by_row)
        i = 0
        while i < len(rows):
            j = i
            while (j + 1 < len(rows) and rows[j + 1] == rows[j] + 1
                   and value_by_row[rows[j + 1]] == value_by_row[rows[i]]):
                j += 1
            if value_by_row[rows[i]] is not None:
                apply(rows[i], rows[j], value_by_row[rows[i]])
            i = j + 1

    emit_runs(row_fill, lambda start, end, fill: wb(
        'PATCH', f"worksheets('{SHEET_NAME}')/range(address='B{start}:E{end}')/format/fill",
        {'color': fill}))
    emit_runs(row_font, lambda start, end, font: wb(
        'PATCH', f"worksheets('{SHEET_NAME}')/range(address='B{start}:E{end}')/format/font",
        dict(font)))

    for start, end in _contiguous_runs_from_list(row_missing):
        wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='B{start}:E{end}')/format/fill",
           {'color': TREE_MISSING_FILL})

    _style_tree_hyperlinks(wb, linked_rows)


def _format_summary_block(wb, total_row):
    """Style the Module Summary block above the tree. Ported from
    suspect_macro.vb's FormatSummary; total_row is summary_last_row
    from build_row_plan (the VBA version located it by scanning column
    B from the bottom -- unnecessary here since the row plan already
    knows it).
    """
    title_row = SUMMARY_TITLE_ROW
    header_row = title_row + 2

    title_rng = f"B{title_row}:F{title_row}"
    header_rng = f"B{header_row}:F{header_row}"
    total_rng = f"B{total_row}:F{total_row}"

    wb('POST', f"worksheets('{SHEET_NAME}')/range(address='{title_rng}')/merge",
       {'across': False})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{title_rng}')/format/fill",
       {'color': SUMMARY_TITLE_FILL})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{title_rng}')/format/font",
       {'bold': True, 'size': 16, 'color': '#FFFFFF'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{title_rng}')/format",
       {'horizontalAlignment': 'Left', 'verticalAlignment': 'Center'})

    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{header_rng}')/format/fill",
       {'color': SUMMARY_HEADER_FILL})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{header_rng}')/format/font",
       {'bold': True, 'size': 11, 'color': '#FFFFFF'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{header_rng}')/format",
       {'horizontalAlignment': 'Center', 'verticalAlignment': 'Center'})

    # Per-module data rows (absent if the summary had no modules).
    if total_row > header_row + 1:
        data_rng = f"B{header_row + 1}:F{total_row - 1}"
        wb('POST', f"worksheets('{SHEET_NAME}')/range(address='{data_rng}')/format/fill/clear")
        wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{data_rng}')/format/font",
           {'bold': False, 'size': 11, 'color': '#000000'})
        wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{data_rng}')/format",
           {'verticalAlignment': 'Center'})
        wb('PATCH',
           f"worksheets('{SHEET_NAME}')/range(address='B{header_row + 1}:B{total_row - 1}')/format",
           {'horizontalAlignment': 'Left'})
        wb('PATCH',
           f"worksheets('{SHEET_NAME}')/range(address='C{header_row + 1}:F{total_row - 1}')/format",
           {'horizontalAlignment': 'Center'})

    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{total_rng}')/format/fill",
       {'color': SUMMARY_TOTAL_FILL})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{total_rng}')/format/font",
       {'bold': True, 'size': 11, 'color': '#000000'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{total_rng}')/format",
       {'verticalAlignment': 'Center'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='B{total_row}')/format",
       {'horizontalAlignment': 'Left'})
    wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='C{total_row}:F{total_row}')/format",
       {'horizontalAlignment': 'Center'})


def _apply_tree_column_widths(wb):
    """Set fixed pixel widths for the tree's columns. Ported from
    suspect_macro.vb's ApplyColumnWidths / SetColumnPixelWidth -- the
    Graph API's columnWidth is already in points, so this converts once
    (1 px = 0.75 pt at 96 DPI) with none of the VBA version's iterative
    approximation (VBA's ColumnWidth is in characters, an affine, not
    proportional, relationship to points).
    """
    for col, px in TREE_COLUMN_WIDTHS_PX.items():
        wb('PATCH', f"worksheets('{SHEET_NAME}')/range(address='{col}:{col}')/format",
           {'columnWidth': px * 0.75})


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

    need_change_last_row = _write_need_change_map(wb, report)
    print(f"  Wrote {need_change_last_row} rows to the Need Change Summary "
          f"(B1:E{need_change_last_row})")

    last_row, matrix, row_level, last_tree_row, summary_last_row = build_row_plan(report, tree)

    existing = wb('GET', f"worksheets('{SHEET_NAME}')/usedRange?$select=rowCount")
    clear_through = max(last_row, existing.get('rowCount', 0), 200)
    _clear_sheet(wb, clear_through)

    _write_values(wb, last_row, matrix)
    _write_levels(wb, last_row, row_level)

    _format_tree_banner(wb)
    _format_tree_rows(wb, matrix, row_level, last_tree_row)
    if WRITE_MODULE_SUMMARY:
        _format_summary_block(wb, summary_last_row)
    _apply_tree_column_widths(wb)

    # Row grouping (outline levels) is NOT done here -- the Graph API has
    # no direct row-outline-level setter, only incremental group()/
    # ungroup() actions, so that one piece stays in suspect_macro.vb,
    # which reads BUILD_MARKER_CELL (just written below) to know a new
    # version is waiting and reads GROUPED_MARKER_CELL (which this
    # script must NOT write) to know whether it's already grouped it.
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
