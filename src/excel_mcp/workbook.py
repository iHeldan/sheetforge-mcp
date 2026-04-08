import logging
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from .exceptions import WorkbookError

logger = logging.getLogger(__name__)


def _get_sheet_usage(ws) -> tuple[int, int, str | None, bool]:
    """Return user-facing sheet dimensions instead of openpyxl's default 1x1 empty sheet."""
    is_empty = ws.max_row == 1 and ws.max_column == 1 and ws.cell(1, 1).value is None
    if is_empty:
        return 0, 0, None, True
    return ws.max_row, ws.max_column, f"A-{get_column_letter(ws.max_column)}", False


@contextmanager
def safe_workbook(filepath: str, save: bool = False, read_only: bool = False):
    """Context manager that ensures workbook is always closed.

    Args:
        filepath: Path to Excel file.
        save: If True, save the workbook before closing.
        read_only: If True, open in read-only mode.
    """
    wb = load_workbook(filepath, read_only=read_only)
    try:
        yield wb
    except Exception:
        raise
    else:
        if save and not read_only:
            wb.save(filepath)
    finally:
        wb.close()

def create_workbook(filepath: str, sheet_name: str = "Sheet1") -> dict[str, Any]:
    """Create a new Excel workbook with optional custom sheet name"""
    try:
        wb = Workbook()
        # Rename default sheet
        if "Sheet" in wb.sheetnames:
            sheet = wb["Sheet"]
            sheet.title = sheet_name
        else:
            wb.create_sheet(sheet_name)

        path = Path(filepath)
        path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(str(path))
        return {
            "message": f"Created workbook: {filepath}",
            "active_sheet": sheet_name,
            "workbook": wb
        }
    except Exception as e:
        logger.error(f"Failed to create workbook: {e}")
        raise WorkbookError(f"Failed to create workbook: {e!s}")

def get_or_create_workbook(filepath: str) -> Workbook:
    """Load an existing workbook. Raises FileNotFoundError if it doesn't exist."""
    return load_workbook(filepath)

def create_sheet(filepath: str, sheet_name: str) -> dict:
    """Create a new worksheet in the workbook if it doesn't exist."""
    try:
        with safe_workbook(filepath, save=True) as wb:
            # Check if sheet already exists
            if sheet_name in wb.sheetnames:
                raise WorkbookError(f"Sheet {sheet_name} already exists")

            # Create new sheet
            wb.create_sheet(sheet_name)
        return {"message": f"Sheet {sheet_name} created successfully"}
    except WorkbookError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to create sheet: {e}")
        raise WorkbookError(str(e))

def list_sheets(filepath: str) -> list[dict[str, Any]]:
    """List all sheets with basic metadata."""
    try:
        with safe_workbook(filepath) as wb:
            sheets = []
            for name in wb.sheetnames:
                ws = wb[name]
                rows, columns, column_range, is_empty = _get_sheet_usage(ws)
                sheets.append({
                    "name": name,
                    "rows": rows,
                    "columns": columns,
                    "column_range": column_range,
                    "is_empty": is_empty,
                })
            return sheets
    except Exception as e:
        logger.error(f"Failed to list sheets: {e}")
        raise WorkbookError(str(e))


def list_named_ranges(filepath: str) -> list[dict[str, Any]]:
    """List workbook-level and local defined names."""
    try:
        with safe_workbook(filepath) as wb:
            ranges = []
            for name, defined_name in wb.defined_names.items():
                destinations = []
                try:
                    destinations = [
                        {
                            "sheet_name": sheet_name,
                            "range": cell_range,
                        }
                        for sheet_name, cell_range in defined_name.destinations
                    ]
                except Exception:
                    destinations = []

                local_sheet = None
                if defined_name.localSheetId is not None:
                    try:
                        local_sheet = wb.sheetnames[defined_name.localSheetId]
                    except Exception:
                        local_sheet = None

                ranges.append(
                    {
                        "name": name,
                        "type": defined_name.type,
                        "value": defined_name.value,
                        "destinations": destinations,
                        "local_sheet": local_sheet,
                        "hidden": bool(getattr(defined_name, "hidden", False)),
                    }
                )

            return sorted(ranges, key=lambda item: item["name"].lower())
    except Exception as e:
        logger.error(f"Failed to list named ranges: {e}")
        raise WorkbookError(str(e))

def get_workbook_info(filepath: str, include_ranges: bool = False) -> dict[str, Any]:
    """Get metadata about workbook including sheets, ranges, etc."""
    try:
        path = Path(filepath)
        if not path.exists():
            raise WorkbookError(f"File not found: {filepath}")

        with safe_workbook(filepath) as wb:
            info = {
                "filename": path.name,
                "sheets": wb.sheetnames,
                "size": path.stat().st_size,
                "modified": path.stat().st_mtime
            }

            if include_ranges:
                # Add used ranges for each sheet
                ranges = {}
                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    if ws.max_row > 0 and ws.max_column > 0:
                        ranges[sheet_name] = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
                info["used_ranges"] = ranges

        return info

    except WorkbookError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to get workbook info: {e}")
        raise WorkbookError(str(e))
