"""Small, dependency-free helpers for reading Workbook API usedRange grids.

Shared by every module that reads a sheet's values grid (sprint_template,
sprint_retro, sprint_summary).
"""


def _gv(grid, row, col):
    """Read a 1-based cell from a usedRange values grid (0-based internally)."""
    r, c = row - 1, col - 1
    if 0 <= r < len(grid) and 0 <= c < len(grid[r]):
        v = grid[r][c]
        return v if v != "" else None
    return None


def _norm_sprint(v):
    # Normalize a sprint cell value to an integer, or None on failure.
    try:
        return int(v)
    except (TypeError, ValueError):
        return None
