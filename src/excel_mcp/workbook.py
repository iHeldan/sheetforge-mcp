import logging
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.worksheet import Worksheet

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


def _bounds_to_range(min_row: int, min_col: int, max_row: int, max_col: int) -> str:
    return f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"


def _bounds_intersect(
    first: tuple[int, int, int, int],
    second: tuple[int, int, int, int],
) -> bool:
    return not (
        first[2] < second[0]
        or second[2] < first[0]
        or first[3] < second[1]
        or second[3] < first[1]
    )


def _intersection_bounds(
    first: tuple[int, int, int, int],
    second: tuple[int, int, int, int],
) -> tuple[int, int, int, int] | None:
    if not _bounds_intersect(first, second):
        return None
    return (
        max(first[0], second[0]),
        max(first[1], second[1]),
        min(first[2], second[2]),
        min(first[3], second[3]),
    )


def _parse_range_reference(
    range_ref: str,
    *,
    worksheet: Worksheet | None = None,
    expected_sheet: str | None = None,
    error_cls: type[Exception] = WorkbookError,
) -> tuple[tuple[int, int, int, int], str]:
    cleaned = str(range_ref).strip()
    if not cleaned:
        raise error_cls("Range reference is required")

    if "!" in cleaned:
        range_sheet, cleaned = cleaned.rsplit("!", 1)
        normalized_sheet = range_sheet.strip().strip("'")
        if expected_sheet is not None and normalized_sheet != expected_sheet:
            raise error_cls(
                f"Range '{range_ref}' refers to sheet '{normalized_sheet}', expected '{expected_sheet}'"
            )

    cleaned = cleaned.replace("$", "")
    try:
        min_col, min_row, max_col, max_row = range_boundaries(cleaned)
    except ValueError as exc:
        raise error_cls(f"Invalid range reference: {range_ref}") from exc

    if None in (min_col, min_row, max_col, max_row):
        if worksheet is None:
            raise error_cls(f"Range reference requires worksheet bounds: {range_ref}")
        min_col = 1 if min_col is None else min_col
        min_row = 1 if min_row is None else min_row
        max_col = worksheet.max_column if max_col is None else max_col
        max_row = worksheet.max_row if max_row is None else max_row

    return (min_row, min_col, max_row, max_col), _bounds_to_range(min_row, min_col, max_row, max_col)


def _iter_range_references(
    range_ref: Any,
    *,
    worksheet: Worksheet | None = None,
    expected_sheet: str | None = None,
) -> list[tuple[tuple[int, int, int, int], str]]:
    if not range_ref:
        return []

    parts = range_ref if isinstance(range_ref, (list, tuple)) else str(range_ref).split(",")
    references: list[tuple[tuple[int, int, int, int], str]] = []
    for part in parts:
        cleaned = str(part).strip()
        if not cleaned:
            continue

        if "!" in cleaned:
            range_sheet, local_range = cleaned.rsplit("!", 1)
            normalized_sheet = range_sheet.strip().strip("'")
            if expected_sheet is not None and normalized_sheet != expected_sheet:
                continue
        else:
            local_range = cleaned

        local_range = local_range.replace("$", "")
        try:
            min_col, min_row, max_col, max_row = range_boundaries(local_range)
        except ValueError:
            continue
        if None in (min_col, min_row, max_col, max_row):
            if worksheet is None:
                continue
            min_col = 1 if min_col is None else min_col
            min_row = 1 if min_row is None else min_row
            max_col = worksheet.max_column if max_col is None else max_col
            max_row = worksheet.max_row if max_row is None else max_row
        references.append(
            ((min_row, min_col, max_row, max_col), _bounds_to_range(min_row, min_col, max_row, max_col))
        )
    return references


def _cell_is_within_bounds(
    sheet_name: str,
    row_index: int,
    column_index: int,
    target_sheet: str,
    target_bounds: tuple[int, int, int, int],
) -> bool:
    return (
        sheet_name == target_sheet
        and target_bounds[0] <= row_index <= target_bounds[2]
        and target_bounds[1] <= column_index <= target_bounds[3]
    )


def _extract_formula_dependencies(
    wb: Any,
    *,
    target_sheet: str,
    target_bounds: tuple[int, int, int, int],
) -> list[dict[str, Any]]:
    dependencies: list[dict[str, Any]] = []

    for formula_sheet_name in wb.sheetnames:
        formula_ws = wb[formula_sheet_name]
        if _sheet_type(formula_ws) == "chartsheet":
            continue

        for row in formula_ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or not cell.value.startswith("="):
                    continue
                if _cell_is_within_bounds(
                    formula_sheet_name,
                    cell.row,
                    cell.column,
                    target_sheet,
                    target_bounds,
                ):
                    continue

                matched_references: list[dict[str, Any]] = []
                try:
                    tokenizer = Tokenizer(cell.value)
                except Exception:
                    continue

                for token in tokenizer.items:
                    if token.type != "OPERAND" or token.subtype != "RANGE":
                        continue

                    token_value = str(token.value).strip()
                    if not token_value or "[" in token_value:
                        continue

                    reference_sheet_name = formula_sheet_name
                    local_reference = token_value
                    if "!" in token_value:
                        range_sheet, local_reference = token_value.rsplit("!", 1)
                        reference_sheet_name = range_sheet.strip().strip("'")

                    if reference_sheet_name != target_sheet or reference_sheet_name not in wb.sheetnames:
                        continue

                    reference_ws = wb[reference_sheet_name]
                    if _sheet_type(reference_ws) == "chartsheet":
                        continue

                    try:
                        reference_bounds, normalized_reference = _parse_range_reference(
                            local_reference,
                            worksheet=reference_ws,
                            error_cls=WorkbookError,
                        )
                    except WorkbookError:
                        continue

                    intersection = _intersection_bounds(target_bounds, reference_bounds)
                    if intersection is None:
                        continue

                    matched_references.append(
                        {
                            "reference": f"{reference_sheet_name}!{normalized_reference}",
                            "intersection_range": _bounds_to_range(*intersection),
                        }
                    )

                if matched_references:
                    dependencies.append(
                        {
                            "sheet_name": formula_sheet_name,
                            "cell": cell.coordinate,
                            "formula": cell.value,
                            "references": matched_references,
                        }
                    )

    return dependencies


def _freeze_panes_value(ws) -> str | None:
    value = ws.freeze_panes
    if value is None:
        return None
    if isinstance(value, str):
        return value
    return getattr(value, "coordinate", None)


def _sheet_type(ws: Any) -> str:
    return "chartsheet" if ws.__class__.__name__ == "Chartsheet" else "worksheet"


def require_worksheet(
    wb: Any,
    sheet_name: str,
    *,
    error_cls: type[Exception] = WorkbookError,
    operation: str = "worksheet operations",
) -> Worksheet:
    if sheet_name not in wb.sheetnames:
        raise error_cls(f"Sheet '{sheet_name}' not found")

    ws = wb[sheet_name]
    if not isinstance(ws, Worksheet):
        raise error_cls(
            f"Sheet '{sheet_name}' is a chartsheet and cannot be used for {operation}"
        )
    return ws


def first_worksheet(
    wb: Any,
    *,
    error_cls: type[Exception] = WorkbookError,
) -> tuple[str, Worksheet]:
    if not wb.sheetnames:
        raise error_cls("Workbook contains no sheets")

    worksheets = list(getattr(wb, "worksheets", []))
    if not worksheets:
        raise error_cls("Workbook contains no worksheets")

    worksheet = worksheets[0]
    return worksheet.title, worksheet


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
                _chart_occupied_range,
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
                    anchor = _extract_chart_anchor(chart)
                    series = getattr(chart, "ser", None) or list(getattr(chart, "series", []))
                    chart_info = {
                        "chart_index": chart_index,
                        "chart_type": _chart_type_name(chart),
                        "anchor": anchor,
                        "title": _extract_title_text(getattr(chart, "title", None)),
                        "width": width,
                        "height": height,
                        "series_count": len(series),
                    }
                    if anchor and width and height:
                        chart_info["occupied_range"] = _chart_occupied_range(
                            ws,
                            anchor,
                            width=width,
                            height=height,
                        )
                    charts.append(chart_info)

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


def analyze_range_impact(filepath: str, sheet_name: str, range_ref: str) -> dict[str, Any]:
    """Inspect workbook structures that overlap a worksheet range before mutation."""
    try:
        path = Path(filepath)
        if not path.exists():
            raise WorkbookError(f"File not found: {filepath}")

        with safe_workbook(filepath) as wb:
            ws = require_worksheet(
                wb,
                sheet_name,
                error_cls=WorkbookError,
                operation="range impact analysis",
            )
            target_bounds, normalized_range = _parse_range_reference(
                range_ref,
                worksheet=ws,
                expected_sheet=sheet_name,
                error_cls=WorkbookError,
            )

            from .chart import (
                DEFAULT_CHART_HEIGHT,
                DEFAULT_CHART_WIDTH,
                _chart_occupied_range,
                _chart_type_name,
                _extract_chart_anchor,
                _extract_chart_dimensions,
                _extract_title_text,
            )
            from .tables import _build_table_metadata

            tables: list[dict[str, Any]] = []
            for table in ws.tables.values():
                min_col, min_row, max_col, max_row = range_boundaries(table.ref)
                table_bounds = (min_row, min_col, max_row, max_col)
                intersection = _intersection_bounds(target_bounds, table_bounds)
                if intersection is None:
                    continue

                metadata = _build_table_metadata(sheet_name, ws, table)
                header_row_count = metadata["header_row_count"]
                totals_row_count = metadata["totals_row_count"]
                header_bounds = (
                    (min_row, min_col, min_row + header_row_count - 1, max_col)
                    if header_row_count > 0
                    else None
                )
                data_start_row = min_row + header_row_count
                data_end_row = max_row - totals_row_count
                data_bounds = (
                    (data_start_row, min_col, data_end_row, max_col)
                    if data_start_row <= data_end_row
                    else None
                )
                totals_bounds = (
                    (max_row - totals_row_count + 1, min_col, max_row, max_col)
                    if totals_row_count > 0
                    else None
                )
                tables.append(
                    {
                        "table_name": table.displayName,
                        "range": table.ref,
                        "intersection_range": _bounds_to_range(*intersection),
                        "covers_header": bool(
                            header_bounds and _bounds_intersect(target_bounds, header_bounds)
                        ),
                        "covers_data": bool(
                            data_bounds and _bounds_intersect(target_bounds, data_bounds)
                        ),
                        "covers_totals_row": bool(
                            totals_bounds and _bounds_intersect(target_bounds, totals_bounds)
                        ),
                    }
                )

            charts: list[dict[str, Any]] = []
            for chart_index, chart in enumerate(getattr(ws, "_charts", []), start=1):
                anchor = _extract_chart_anchor(chart)
                if not anchor:
                    continue
                width, height = _extract_chart_dimensions(chart)
                occupied_range = _chart_occupied_range(
                    ws,
                    anchor,
                    width=width or DEFAULT_CHART_WIDTH,
                    height=height or DEFAULT_CHART_HEIGHT,
                )
                chart_bounds, _ = _parse_range_reference(occupied_range, error_cls=WorkbookError)
                intersection = _intersection_bounds(target_bounds, chart_bounds)
                if intersection is None:
                    continue
                charts.append(
                    {
                        "chart_index": chart_index,
                        "chart_type": _chart_type_name(chart),
                        "title": _extract_title_text(getattr(chart, "title", None)),
                        "anchor": anchor,
                        "occupied_range": occupied_range,
                        "intersection_range": _bounds_to_range(*intersection),
                    }
                )

            merged_ranges: list[dict[str, Any]] = []
            for merged_range in ws.merged_cells.ranges:
                merged_bounds, merged_ref = _parse_range_reference(
                    str(merged_range),
                    worksheet=ws,
                    error_cls=WorkbookError,
                )
                intersection = _intersection_bounds(target_bounds, merged_bounds)
                if intersection is None:
                    continue
                merged_ranges.append(
                    {
                        "range": merged_ref,
                        "intersection_range": _bounds_to_range(*intersection),
                    }
                )

            named_ranges: list[dict[str, Any]] = []
            for named_range in _serialize_named_ranges(wb):
                matching_destinations = []
                for destination in named_range["destinations"]:
                    if destination["sheet_name"] != sheet_name:
                        continue
                    for destination_bounds, destination_ref in _iter_range_references(
                        destination["range"],
                        expected_sheet=sheet_name,
                    ):
                        intersection = _intersection_bounds(target_bounds, destination_bounds)
                        if intersection is None:
                            continue
                        matching_destinations.append(
                            {
                                "range": destination_ref,
                                "intersection_range": _bounds_to_range(*intersection),
                            }
                        )

                if matching_destinations:
                    named_ranges.append(
                        {
                            "name": named_range["name"],
                            "local_sheet": named_range["local_sheet"],
                            "hidden": named_range["hidden"],
                            "destinations": matching_destinations,
                        }
                    )

            autofilter = None
            if ws.auto_filter.ref:
                autofilter_bounds, autofilter_ref = _parse_range_reference(
                    ws.auto_filter.ref,
                    worksheet=ws,
                    error_cls=WorkbookError,
                )
                intersection = _intersection_bounds(target_bounds, autofilter_bounds)
                if intersection is not None:
                    autofilter = {
                        "range": autofilter_ref,
                        "intersection_range": _bounds_to_range(*intersection),
                    }

            print_area_matches = []
            for print_bounds, print_ref in _iter_range_references(
                ws.print_area,
                worksheet=ws,
                expected_sheet=sheet_name,
            ):
                intersection = _intersection_bounds(target_bounds, print_bounds)
                if intersection is None:
                    continue
                print_area_matches.append(
                    {
                        "range": print_ref,
                        "intersection_range": _bounds_to_range(*intersection),
                    }
                )

            formula_cells = []
            for row in ws.iter_rows(
                min_row=target_bounds[0],
                max_row=target_bounds[2],
                min_col=target_bounds[1],
                max_col=target_bounds[3],
            ):
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        formula_cells.append(cell.coordinate)

            dependent_formulas = _extract_formula_dependencies(
                wb,
                target_sheet=sheet_name,
                target_bounds=target_bounds,
            )

            impact_score = (
                len(tables) * 3
                + len(charts) * 3
                + len(merged_ranges) * 2
                + len(named_ranges) * 2
                + (1 if formula_cells else 0)
                + len(dependent_formulas) * 3
                + (1 if autofilter else 0)
                + (1 if print_area_matches else 0)
            )
            if impact_score >= 3:
                risk_level = "high"
            elif impact_score >= 1:
                risk_level = "medium"
            else:
                risk_level = "low"

            hints: list[str] = []
            if tables:
                hints.append("Selected range overlaps native Excel tables.")
                if any(table["covers_header"] for table in tables):
                    hints.append("Table header rows are inside the selected range.")
                if any(table["covers_totals_row"] for table in tables):
                    hints.append("Table totals rows are inside the selected range.")
            if charts:
                hints.append("Selected range overlaps embedded chart footprints.")
            if merged_ranges:
                hints.append("Selected range touches merged cells that may need to be unmerged first.")
            if named_ranges:
                hints.append("Named ranges point into the selected range.")
            if autofilter:
                hints.append("Selected range overlaps the worksheet autofilter.")
            if print_area_matches:
                hints.append("Selected range overlaps the worksheet print area.")
            if formula_cells:
                hints.append("Selected range contains formula cells that may recalculate or break.")
            if dependent_formulas:
                hints.append("Formulas elsewhere in the workbook reference the selected range.")
            if not hints:
                hints.append("No overlapping workbook structures detected for this range.")

            return {
                "sheet_name": sheet_name,
                "range": normalized_range,
                "used_range": _get_used_range(ws),
                "summary": {
                    "risk_level": risk_level,
                    "table_count": len(tables),
                    "chart_count": len(charts),
                    "merged_range_count": len(merged_ranges),
                    "named_range_count": len(named_ranges),
                    "formula_cell_count": len(formula_cells),
                    "dependent_formula_count": len(dependent_formulas),
                    "autofilter_overlap": autofilter is not None,
                    "print_area_overlap": bool(print_area_matches),
                },
                "tables": tables,
                "charts": charts,
                "merged_ranges": merged_ranges,
                "named_ranges": named_ranges,
                "formula_cells": {
                    "count": len(formula_cells),
                    "sample": formula_cells[:10],
                },
                "dependent_formulas": {
                    "count": len(dependent_formulas),
                    "sample": dependent_formulas[:10],
                },
                "worksheet_features": {
                    "autofilter": autofilter,
                    "print_area": print_area_matches,
                },
                "hints": hints,
            }

    except WorkbookError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to analyze range impact: {e}")
        raise WorkbookError(str(e))
