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


def _get_used_range(ws) -> str | None:
    rows, columns, _, is_empty = _get_sheet_usage(ws)
    if is_empty or rows == 0 or columns == 0:
        return None
    return f"A1:{get_column_letter(columns)}{rows}"


def _freeze_panes_value(ws) -> str | None:
    value = ws.freeze_panes
    if value is None:
        return None
    if isinstance(value, str):
        return value
    return getattr(value, "coordinate", None)


def _sheet_type(ws: Any) -> str:
    return "chartsheet" if ws.__class__.__name__ == "Chartsheet" else "worksheet"


def _serialize_named_ranges(wb: Any) -> list[dict[str, Any]]:
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
                if _sheet_type(ws) == "chartsheet":
                    sheets.append(
                        {
                            "name": name,
                            "sheet_type": "chartsheet",
                            "rows": 0,
                            "columns": 0,
                            "column_range": None,
                            "is_empty": len(getattr(ws, "_charts", [])) == 0,
                        }
                    )
                    continue

                rows, columns, column_range, is_empty = _get_sheet_usage(ws)
                sheets.append({
                    "name": name,
                    "sheet_type": "worksheet",
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
            return _serialize_named_ranges(wb)
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
                    if _sheet_type(ws) == "chartsheet":
                        continue

                    used_range = _get_used_range(ws)
                    if used_range is not None:
                        ranges[sheet_name] = used_range
                info["used_ranges"] = ranges

        return info

    except WorkbookError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to get workbook info: {e}")
        raise WorkbookError(str(e))


def profile_workbook(filepath: str) -> dict[str, Any]:
    """Return a workbook-level inventory tuned for agent orientation."""
    try:
        path = Path(filepath)
        if not path.exists():
            raise WorkbookError(f"File not found: {filepath}")

        with safe_workbook(filepath) as wb:
            from .chart import (
                _chart_type_name,
                _extract_chart_anchor,
                _extract_chart_dimensions,
                _extract_title_text,
            )
            from .sheet import _sheet_protection_state
            from .tables import _build_table_metadata

            named_ranges = _serialize_named_ranges(wb)
            sheets: list[dict[str, Any]] = []
            total_tables = 0
            total_charts = 0

            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                charts = []
                for chart_index, chart in enumerate(getattr(ws, "_charts", []), start=1):
                    width, height = _extract_chart_dimensions(chart)
                    series = getattr(chart, "ser", None) or list(getattr(chart, "series", []))
                    charts.append(
                        {
                            "chart_index": chart_index,
                            "chart_type": _chart_type_name(chart),
                            "anchor": _extract_chart_anchor(chart),
                            "title": _extract_title_text(getattr(chart, "title", None)),
                            "width": width,
                            "height": height,
                            "series_count": len(series),
                        }
                    )

                if _sheet_type(ws) == "chartsheet":
                    total_charts += len(charts)
                    sheets.append(
                        {
                            "name": sheet_name,
                            "sheet_type": "chartsheet",
                            "visibility": ws.sheet_state,
                            "table_count": 0,
                            "chart_count": len(charts),
                            "tables": [],
                            "charts": charts,
                        }
                    )
                    continue

                rows, columns, column_range, is_empty = _get_sheet_usage(ws)
                tables = [
                    _build_table_metadata(sheet_name, ws, table)
                    for table in ws.tables.values()
                ]

                total_tables += len(tables)
                total_charts += len(charts)

                sheets.append(
                    {
                        "name": sheet_name,
                        "sheet_type": "worksheet",
                        "rows": rows,
                        "columns": columns,
                        "column_range": column_range,
                        "used_range": _get_used_range(ws),
                        "is_empty": is_empty,
                        "visibility": ws.sheet_state,
                        "freeze_panes": _freeze_panes_value(ws),
                        "has_autofilter": bool(ws.auto_filter.ref),
                        "autofilter_range": ws.auto_filter.ref or None,
                        "print_area": ws.print_area or None,
                        "print_title_rows": ws.print_title_rows or None,
                        "print_title_columns": ws.print_title_cols or None,
                        "merged_range_count": len(ws.merged_cells.ranges),
                        "table_count": len(tables),
                        "chart_count": len(charts),
                        "protection": {
                            "enabled": _sheet_protection_state(ws)["enabled"],
                            "password_protected": _sheet_protection_state(ws)["password_protected"],
                        },
                        "tables": tables,
                        "charts": charts,
                    }
                )

            return {
                "filename": path.name,
                "size": path.stat().st_size,
                "modified": path.stat().st_mtime,
                "sheet_count": len(wb.sheetnames),
                "named_range_count": len(named_ranges),
                "table_count": total_tables,
                "chart_count": total_charts,
                "sheets": sheets,
                "named_ranges": named_ranges,
            }

    except WorkbookError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to profile workbook: {e}")
        raise WorkbookError(str(e))
