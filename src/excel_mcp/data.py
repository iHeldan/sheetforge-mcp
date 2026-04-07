from pathlib import Path
from typing import Any, Dict, List, Optional
import logging

from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import column_index_from_string, get_column_letter

from .exceptions import DataError
from .cell_utils import parse_cell_range
from .cell_validation import get_data_validation_for_cell
from .workbook import safe_workbook

logger = logging.getLogger(__name__)


def _cell_address(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


def _range_string(start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    return f"{_cell_address(start_row, start_col)}:{_cell_address(end_row, end_col)}"


def _get_header_map(ws: Worksheet, header_row: int) -> Dict[str, int]:
    header_map: Dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=header_row, column=col).value
        if value is None:
            continue
        header_name = str(value)
        if header_name in header_map:
            raise DataError(f"Duplicate header '{header_name}' found in row {header_row}")
        header_map[header_name] = col

    if not header_map:
        raise DataError(f"No headers found in row {header_row}")

    return header_map


def _find_last_data_row(ws: Worksheet, header_row: int, columns: List[int]) -> int:
    for row in range(ws.max_row, header_row, -1):
        if any(ws.cell(row=row, column=col).value is not None for col in columns):
            return row
    return header_row


def _build_cell_change(
    sheet_name: str,
    row: int,
    col: int,
    old_value: Any,
    new_value: Any,
    column_name: Optional[str] = None,
) -> Dict[str, Any]:
    change = {
        "sheet_name": sheet_name,
        "cell": _cell_address(row, col),
        "row": row,
        "column": col,
        "old_value": old_value,
        "new_value": new_value,
    }
    if column_name is not None:
        change["column_name"] = column_name
    return change

def read_excel_range(
    filepath: Path | str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: Optional[str] = None,
    preview_only: bool = False
) -> List[Dict[str, Any]]:
    """Read data from Excel range with optional preview mode"""
    try:
        with safe_workbook(str(filepath)) as wb:
            if sheet_name not in wb.sheetnames:
                raise DataError(f"Sheet '{sheet_name}' not found")

            ws = wb[sheet_name]

            # Parse start cell
            if ':' in start_cell:
                start_cell, end_cell = start_cell.split(':')

            # Get start coordinates
            try:
                start_coords = parse_cell_range(f"{start_cell}:{start_cell}")
                if not start_coords or not all(coord is not None for coord in start_coords[:2]):
                    raise DataError(f"Invalid start cell reference: {start_cell}")
                start_row, start_col = start_coords[0], start_coords[1]
            except ValueError as e:
                raise DataError(f"Invalid start cell format: {str(e)}")

            # Determine end coordinates
            if end_cell:
                try:
                    end_coords = parse_cell_range(f"{end_cell}:{end_cell}")
                    if not end_coords or not all(coord is not None for coord in end_coords[:2]):
                        raise DataError(f"Invalid end cell reference: {end_cell}")
                    end_row, end_col = end_coords[0], end_coords[1]
                except ValueError as e:
                    raise DataError(f"Invalid end cell format: {str(e)}")
            else:
                # If no end_cell, use the full data range of the sheet
                if ws.max_row == 1 and ws.max_column == 1 and ws.cell(1, 1).value is None:
                    # Handle empty sheet
                    end_row, end_col = start_row, start_col
                else:
                    end_row = ws.max_row
                    end_col = ws.max_column

            # Validate range bounds
            if start_row > ws.max_row or start_col > ws.max_column:
                logger.warning(
                    f"Start cell {start_cell} is outside the sheet's data boundary "
                    f"({get_column_letter(ws.min_column)}{ws.min_row}:{get_column_letter(ws.max_column)}{ws.max_row}). "
                    f"No data will be read."
                )
                return []

            data = []
            for row in range(start_row, end_row + 1):
                row_data = []
                for col in range(start_col, end_col + 1):
                    cell = ws.cell(row=row, column=col)
                    row_data.append(cell.value)
                if any(v is not None for v in row_data):
                    data.append(row_data)

            return data
    except DataError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to read Excel range: {e}")
        raise DataError(str(e))

def write_data(
    filepath: str,
    sheet_name: Optional[str],
    data: Optional[List[List]],
    start_cell: str = "A1",
    dry_run: bool = False,
) -> Dict[str, str]:
    """Write data to Excel sheet with workbook handling

    Headers are handled intelligently based on context.
    """
    try:
        if not data:
            raise DataError("No data provided to write")

        with safe_workbook(filepath, save=not dry_run) as wb:
            # If no sheet specified, use active sheet
            if not sheet_name:
                active_sheet = wb.active
                if active_sheet is None:
                    raise DataError("No active sheet found in workbook")
                sheet_name = active_sheet.title
                sheet_created = False
            else:
                sheet_created = sheet_name not in wb.sheetnames

            # Validate start cell
            try:
                start_coords = parse_cell_range(start_cell)
                if not start_coords or not all(coord is not None for coord in start_coords[:2]):
                    raise DataError(f"Invalid start cell reference: {start_cell}")
            except ValueError as e:
                raise DataError(f"Invalid start cell format: {str(e)}")

            start_row, start_col = start_coords[0], start_coords[1]
            total_cells = sum(len(row) for row in data)
            if data:
                max_cols = max(len(row) for row in data)
                end_row = start_row + len(data) - 1
                end_col = start_col + max_cols - 1
                target_range = _range_string(start_row, start_col, end_row, end_col)
            else:
                target_range = start_cell

            changes: List[Dict[str, Any]] = []
            existing_ws = wb[sheet_name] if not sheet_created else None
            for i, row in enumerate(data):
                for j, val in enumerate(row):
                    row_idx = start_row + i
                    col_idx = start_col + j
                    old_value = None if existing_ws is None else existing_ws.cell(row=row_idx, column=col_idx).value
                    if old_value != val:
                        changes.append(
                            _build_cell_change(
                                sheet_name=sheet_name,
                                row=row_idx,
                                col=col_idx,
                                old_value=old_value,
                                new_value=val,
                            )
                        )

            if len(data) > 0:
                if sheet_created:
                    ws = wb.create_sheet(sheet_name)
                else:
                    ws = wb[sheet_name]
                if not dry_run:
                    _write_data_to_worksheet(ws, data, start_cell)

        return {
            "message": f"{'Previewed' if dry_run else 'Wrote'} data to {sheet_name}",
            "active_sheet": sheet_name,
            "target_range": target_range,
            "sheet_created": sheet_created,
            "cells_written": total_cells,
            "changed_cells": len(changes),
            "changes": changes,
            "dry_run": dry_run,
        }
    except DataError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to write data: {e}")
        raise DataError(str(e))

def _write_data_to_worksheet(
    worksheet: Worksheet, 
    data: List[List], 
    start_cell: str = "A1",
) -> None:
    """Write data to worksheet with intelligent header handling"""
    try:
        if not data:
            raise DataError("No data provided to write")

        try:
            start_coords = parse_cell_range(start_cell)
            if not start_coords or not all(x is not None for x in start_coords[:2]):
                raise DataError(f"Invalid start cell reference: {start_cell}")
            start_row, start_col = start_coords[0], start_coords[1]
        except ValueError as e:
            raise DataError(f"Invalid start cell format: {str(e)}")

        # Write data
        for i, row in enumerate(data):
            for j, val in enumerate(row):
                worksheet.cell(row=start_row + i, column=start_col + j, value=val)
    except DataError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to write worksheet data: {e}")
        raise DataError(str(e))

def read_excel_range_with_metadata(
    filepath: Path | str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: Optional[str] = None,
    include_validation: bool = True
) -> Dict[str, Any]:
    """Read data from Excel range with cell metadata including validation rules.

    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_cell: Starting cell address
        end_cell: Ending cell address (optional)
        include_validation: Whether to include validation metadata

    Returns:
        Dictionary containing structured cell data with metadata
    """
    try:
        with safe_workbook(str(filepath)) as wb:
            if sheet_name not in wb.sheetnames:
                raise DataError(f"Sheet '{sheet_name}' not found")

            ws = wb[sheet_name]

            # Parse start cell
            if ':' in start_cell:
                start_cell, end_cell = start_cell.split(':')

            # Get start coordinates
            try:
                start_coords = parse_cell_range(f"{start_cell}:{start_cell}")
                if not start_coords or not all(coord is not None for coord in start_coords[:2]):
                    raise DataError(f"Invalid start cell reference: {start_cell}")
                start_row, start_col = start_coords[0], start_coords[1]
            except ValueError as e:
                raise DataError(f"Invalid start cell format: {str(e)}")

            # Determine end coordinates
            if end_cell:
                try:
                    end_coords = parse_cell_range(f"{end_cell}:{end_cell}")
                    if not end_coords or not all(coord is not None for coord in end_coords[:2]):
                        raise DataError(f"Invalid end cell reference: {end_cell}")
                    end_row, end_col = end_coords[0], end_coords[1]
                except ValueError as e:
                    raise DataError(f"Invalid end cell format: {str(e)}")
            else:
                # If no end_cell, use the full data range of the sheet
                if ws.max_row == 1 and ws.max_column == 1 and ws.cell(1, 1).value is None:
                    # Handle empty sheet
                    end_row, end_col = start_row, start_col
                else:
                    # Expand to sheet's max boundaries; start_row/start_col from caller are preserved
                    end_row = ws.max_row
                    end_col = ws.max_column

            # Validate range bounds
            if start_row > ws.max_row or start_col > ws.max_column:
                logger.warning(
                    f"Start cell {start_cell} is outside the sheet's data boundary "
                    f"({get_column_letter(ws.min_column)}{ws.min_row}:{get_column_letter(ws.max_column)}{ws.max_row}). "
                    f"No data will be read."
                )
                return {"range": f"{start_cell}:", "sheet_name": sheet_name, "cells": []}

            # Build structured cell data
            range_str = f"{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row}"
            range_data = {
                "range": range_str,
                "sheet_name": sheet_name,
                "cells": []
            }

            for row in range(start_row, end_row + 1):
                for col in range(start_col, end_col + 1):
                    cell = ws.cell(row=row, column=col)
                    cell_address = f"{get_column_letter(col)}{row}"

                    cell_data = {
                        "address": cell_address,
                        "value": cell.value,
                        "row": row,
                        "column": col
                    }

                    # Add validation metadata if requested
                    if include_validation:
                        validation_info = get_data_validation_for_cell(ws, cell_address)
                        if validation_info:
                            cell_data["validation"] = validation_info
                        else:
                            cell_data["validation"] = {"has_validation": False}

                    range_data["cells"].append(cell_data)

            return range_data

    except DataError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to read Excel range with metadata: {e}")
        raise DataError(str(e))


def read_as_table(
    filepath: str,
    sheet_name: str,
    header_row: int = 1,
    start_col: str = "A",
    end_col: Optional[str] = None,
    max_rows: Optional[int] = None,
) -> Dict[str, Any]:
    """Read Excel data as a compact table with headers.

    Returns a dict with headers, rows, total_rows, truncated.
    """
    try:
        with safe_workbook(str(filepath)) as wb:
            if sheet_name not in wb.sheetnames:
                raise DataError(f"Sheet '{sheet_name}' not found")

            ws = wb[sheet_name]

            from openpyxl.utils import column_index_from_string
            start_col_idx = column_index_from_string(start_col.upper())
            if end_col:
                end_col_idx = column_index_from_string(end_col.upper())
            else:
                end_col_idx = ws.max_column

            # Read headers
            headers = []
            for col in range(start_col_idx, end_col_idx + 1):
                val = ws.cell(row=header_row, column=col).value
                headers.append(val)

            # Read data rows
            total_rows = ws.max_row - header_row
            if total_rows < 0:
                total_rows = 0

            limit = max_rows if max_rows else total_rows
            rows = []
            for row_idx in range(header_row + 1, header_row + 1 + min(limit, total_rows)):
                row_data = []
                for col in range(start_col_idx, end_col_idx + 1):
                    row_data.append(ws.cell(row=row_idx, column=col).value)
                rows.append(row_data)

            return {
                "headers": headers,
                "rows": rows,
                "total_rows": total_rows,
                "truncated": max_rows is not None and total_rows > max_rows,
                "sheet_name": sheet_name,
            }
    except DataError:
        raise
    except Exception as e:
        logger.error(f"Failed to read as table: {e}")
        raise DataError(str(e))


def _matches_exact_query(cell_value: Any, query: Any) -> bool:
    """Match typed values while still supporting string queries from clients."""
    if cell_value == query:
        return True

    if isinstance(query, str) and not isinstance(cell_value, str):
        if isinstance(cell_value, bool):
            return str(cell_value).lower() == query.lower()
        return str(cell_value) == query

    return False


def search_cells(
    filepath: str,
    sheet_name: str,
    query: Any,
    exact: bool = True,
    max_results: int = 50,
) -> List[Dict[str, Any]]:
    """Search for cells matching a value."""
    try:
        with safe_workbook(str(filepath)) as wb:
            if sheet_name not in wb.sheetnames:
                raise DataError(f"Sheet '{sheet_name}' not found")

            ws = wb[sheet_name]
            results = []

            for row in range(1, ws.max_row + 1):
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row, column=col)
                    val = cell.value
                    if val is None:
                        continue

                    matched = False
                    if exact:
                        matched = _matches_exact_query(val, query)
                    else:
                        matched = str(query).lower() in str(val).lower()

                    if matched:
                        results.append({
                            "cell": f"{get_column_letter(col)}{row}",
                            "value": val,
                            "row": row,
                            "column": col,
                        })
                        if len(results) >= max_results:
                            return results

            return results
    except DataError:
        raise
    except Exception as e:
        logger.error(f"Failed to search cells: {e}")
        raise DataError(str(e))


def append_table_rows(
    filepath: str,
    sheet_name: str,
    rows: List[Dict[str, Any]],
    header_row: int = 1,
    dry_run: bool = False,
) -> Dict[str, Any]:
    """Append dictionary-shaped rows using the worksheet's header row."""
    try:
        if not rows:
            raise DataError("No rows provided to append")
        if not all(isinstance(row, dict) for row in rows):
            raise DataError("Rows must be a list of objects keyed by column name")

        with safe_workbook(str(filepath), save=not dry_run) as wb:
            if sheet_name not in wb.sheetnames:
                raise DataError(f"Sheet '{sheet_name}' not found")

            ws = wb[sheet_name]
            header_map = _get_header_map(ws, header_row)
            unknown_columns = sorted(
                {
                    key
                    for row in rows
                    for key in row.keys()
                    if key not in header_map
                }
            )
            if unknown_columns:
                raise DataError(f"Unknown columns for append: {', '.join(unknown_columns)}")

            ordered_columns = sorted(header_map.items(), key=lambda item: item[1])
            last_data_row = _find_last_data_row(ws, header_row, list(header_map.values()))
            next_row = last_data_row + 1

            changes: List[Dict[str, Any]] = []
            for row_offset, row_data in enumerate(rows):
                target_row = next_row + row_offset
                for column_name, col_idx in ordered_columns:
                    if column_name not in row_data:
                        continue
                    new_value = row_data[column_name]
                    old_value = ws.cell(row=target_row, column=col_idx).value
                    if old_value != new_value:
                        changes.append(
                            _build_cell_change(
                                sheet_name=sheet_name,
                                row=target_row,
                                col=col_idx,
                                old_value=old_value,
                                new_value=new_value,
                                column_name=column_name,
                            )
                        )
                    if not dry_run:
                        ws.cell(row=target_row, column=col_idx, value=new_value)

            end_row = next_row + len(rows) - 1
            target_range = _range_string(
                next_row,
                min(header_map.values()),
                end_row,
                max(header_map.values()),
            )

        return {
            "message": f"{'Previewed' if dry_run else 'Appended'} {len(rows)} row(s) to {sheet_name}",
            "sheet_name": sheet_name,
            "header_row": header_row,
            "rows_appended": len(rows),
            "start_row": next_row,
            "target_range": target_range,
            "changes": changes,
            "dry_run": dry_run,
        }
    except DataError:
        raise
    except Exception as e:
        logger.error(f"Failed to append table rows: {e}")
        raise DataError(str(e))


def update_rows_by_key(
    filepath: str,
    sheet_name: str,
    key_column: str,
    updates: List[Dict[str, Any]],
    header_row: int = 1,
    dry_run: bool = False,
) -> Dict[str, Any]:
    """Update existing rows by matching a named key column."""
    try:
        if not updates:
            raise DataError("No updates provided")
        if not all(isinstance(update, dict) for update in updates):
            raise DataError("Updates must be a list of objects keyed by column name")

        with safe_workbook(str(filepath), save=not dry_run) as wb:
            if sheet_name not in wb.sheetnames:
                raise DataError(f"Sheet '{sheet_name}' not found")

            ws = wb[sheet_name]
            header_map = _get_header_map(ws, header_row)
            if key_column not in header_map:
                raise DataError(f"Key column '{key_column}' not found")

            unknown_columns = sorted(
                {
                    key
                    for update in updates
                    for key in update.keys()
                    if key not in header_map
                }
            )
            if unknown_columns:
                raise DataError(f"Unknown columns for update: {', '.join(unknown_columns)}")

            key_col_idx = header_map[key_column]
            last_data_row = _find_last_data_row(ws, header_row, list(header_map.values()))
            row_lookup: Dict[Any, int] = {}
            for row_idx in range(header_row + 1, last_data_row + 1):
                key_value = ws.cell(row=row_idx, column=key_col_idx).value
                if key_value is None:
                    continue
                if key_value in row_lookup:
                    raise DataError(f"Duplicate key '{key_value}' found in column '{key_column}'")
                row_lookup[key_value] = row_idx

            seen_update_keys = set()
            changes: List[Dict[str, Any]] = []
            missing_keys: List[Any] = []
            matched_keys: List[Any] = []

            for update in updates:
                if key_column not in update:
                    raise DataError(f"Update row is missing key column '{key_column}'")

                key_value = update[key_column]
                if key_value in seen_update_keys:
                    raise DataError(f"Duplicate update key '{key_value}' provided")
                seen_update_keys.add(key_value)

                target_row = row_lookup.get(key_value)
                if target_row is None:
                    missing_keys.append(key_value)
                    continue

                matched_keys.append(key_value)
                for column_name, new_value in update.items():
                    if column_name == key_column:
                        continue

                    col_idx = header_map[column_name]
                    cell = ws.cell(row=target_row, column=col_idx)
                    old_value = cell.value
                    if old_value == new_value:
                        continue

                    changes.append(
                        _build_cell_change(
                            sheet_name=sheet_name,
                            row=target_row,
                            col=col_idx,
                            old_value=old_value,
                            new_value=new_value,
                            column_name=column_name,
                        )
                    )
                    if not dry_run:
                        cell.value = new_value

        updated_rows = len(matched_keys)
        message = (
            f"{'Previewed' if dry_run else 'Updated'} {updated_rows} row(s) in {sheet_name}"
        )
        if missing_keys:
            message += f"; {len(missing_keys)} key(s) not found"

        return {
            "message": message,
            "sheet_name": sheet_name,
            "key_column": key_column,
            "header_row": header_row,
            "updated_rows": updated_rows,
            "missing_keys": missing_keys,
            "changes": changes,
            "dry_run": dry_run,
        }
    except DataError:
        raise
    except Exception as e:
        logger.error(f"Failed to update rows by key: {e}")
        raise DataError(str(e))
