import re

from openpyxl.utils import column_index_from_string

MAX_EXCEL_ROW = 1_048_576
MAX_EXCEL_COLUMN = 16_384
CELL_REFERENCE_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+)$")


def validate_grid_bounds(min_row: int, min_col: int, max_row: int, max_col: int) -> None:
    if not (1 <= min_row <= max_row <= MAX_EXCEL_ROW and 1 <= min_col <= max_col <= MAX_EXCEL_COLUMN):
        raise ValueError("Mutation must stay inside Excel's A1:XFD1048576 grid")


def validate_sheet_title(title: str, existing: list[str], *, current: str | None = None) -> None:
    try:
        units = len(title.encode("utf-16-le")) // 2 if isinstance(title, str) else 0
    except UnicodeEncodeError:
        units = 0
    if (not isinstance(title, str) or not title.strip() or not 1 <= units <= 31
            or any(ord(character) < 32 for character in title)
            or re.search(r"[\\/*?:\[\]]", title) or title.startswith("'") or title.endswith("'")):
        raise ValueError("Worksheet name must be 1-31 characters without \\ / * ? : [ ] or surrounding apostrophes")
    if any(name != current and name.casefold() == title.casefold() for name in existing):
        raise ValueError(f"Worksheet '{title}' already exists (names are case-insensitive)")


def _parse_cell_reference(cell_ref: str) -> tuple[int, int]:
    match = CELL_REFERENCE_RE.fullmatch(str(cell_ref).strip())
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")

    col_str, row_str = match.groups()
    row = int(row_str)
    try:
        col = column_index_from_string(col_str.upper())
    except ValueError as exc:
        raise ValueError(f"Invalid cell reference: {cell_ref}") from exc

    if not 1 <= row <= MAX_EXCEL_ROW or not 1 <= col <= MAX_EXCEL_COLUMN:
        raise ValueError(f"Cell reference outside Excel limits: {cell_ref}")
    return row, col

def parse_cell_range(
    cell_ref: str,
    end_ref: str | None = None
) -> tuple[int, int, int | None, int | None]:
    """Parse Excel cell reference into row and column indices."""
    if end_ref:
        start_cell = cell_ref
        end_cell = end_ref
    elif ":" in cell_ref:
        parts = cell_ref.split(":")
        if len(parts) != 2 or not all(parts):
            raise ValueError(f"Invalid cell range: {cell_ref}")
        start_cell, end_cell = parts
    else:
        start_cell = cell_ref
        end_cell = None

    start_row, start_col = _parse_cell_reference(start_cell)

    if end_cell:
        end_row, end_col = _parse_cell_reference(end_cell)
        if end_row < start_row:
            raise ValueError("End row cannot be before start row")
        if end_col < start_col:
            raise ValueError("End column cannot be before start column")
    else:
        end_row = None
        end_col = None

    return start_row, start_col, end_row, end_col

def validate_cell_reference(cell_ref: str) -> bool:
    """Validate Excel cell reference format (e.g., 'A1', 'BC123')"""
    try:
        _parse_cell_reference(cell_ref)
    except (TypeError, ValueError):
        return False
    return True
