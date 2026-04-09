import uuid
import logging
from typing import Any, Optional

from openpyxl.utils import range_boundaries
from openpyxl.worksheet.table import Table, TableStyleInfo
from .data import augment_tabular_payload
from .exceptions import DataError
from .workbook import safe_workbook

logger = logging.getLogger(__name__)


def _get_table_style_details(table: Table) -> dict[str, Any]:
    style_name = None
    show_first_column = None
    show_last_column = None
    show_row_stripes = None
    show_column_stripes = None
    if table.tableStyleInfo is not None:
        style_name = table.tableStyleInfo.name
        show_first_column = table.tableStyleInfo.showFirstColumn
        show_last_column = table.tableStyleInfo.showLastColumn
        show_row_stripes = table.tableStyleInfo.showRowStripes
        show_column_stripes = table.tableStyleInfo.showColumnStripes

    return {
        "style": style_name,
        "show_first_column": show_first_column,
        "show_last_column": show_last_column,
        "show_row_stripes": show_row_stripes,
        "show_column_stripes": show_column_stripes,
    }


def _build_table_metadata(
    current_sheet_name: str,
    ws: Any,
    table: Table,
) -> dict[str, Any]:
    min_col, min_row, max_col, max_row = range_boundaries(table.ref)
    header_row_count = int(table.headerRowCount or 0)
    totals_row_count = int(table.totalsRowCount or 0)
    total_row_span = max_row - min_row + 1
    data_row_count = max(total_row_span - header_row_count - totals_row_count, 0)

    headers: list[Any] = []
    if header_row_count > 0:
        headers = [
            ws.cell(row=min_row, column=column_index).value
            for column_index in range(min_col, max_col + 1)
        ]

    metadata = {
        "sheet_name": current_sheet_name,
        "table_name": table.displayName,
        "range": table.ref,
        "headers": headers,
        "column_count": max_col - min_col + 1,
        "data_row_count": data_row_count,
        "header_row_count": header_row_count,
        "totals_row_count": totals_row_count,
        "totals_row_shown": bool(table.totalsRowShown),
    }
    metadata.update(_get_table_style_details(table))
    return metadata


def _find_table(
    wb: Any,
    table_name: str,
    sheet_name: Optional[str] = None,
) -> tuple[str, Any, Table]:
    if sheet_name is not None and sheet_name not in wb.sheetnames:
        raise DataError(f"Sheet '{sheet_name}' not found.")

    sheet_names = [sheet_name] if sheet_name is not None else list(wb.sheetnames)
    for current_sheet_name in sheet_names:
        ws = wb[current_sheet_name]
        for table in ws.tables.values():
            if table.displayName == table_name:
                return current_sheet_name, ws, table

    if sheet_name is not None:
        raise DataError(f"Table '{table_name}' not found in sheet '{sheet_name}'.")
    raise DataError(f"Table '{table_name}' not found.")

def create_excel_table(
    filepath: str,
    sheet_name: str,
    data_range: str,
    table_name: str | None = None,
    table_style: str = "TableStyleMedium9"
) -> dict:
    """Creates a native Excel table for the given data range.
    
    Args:
        filepath: Path to the Excel file.
        sheet_name: Name of the worksheet.
        data_range: The cell range for the table (e.g., "A1:D5").
        table_name: A unique name for the table. If not provided, a unique name is generated.
        table_style: The visual style to apply to the table.
        
    Returns:
        A dictionary with a success message and table details.
    """
    try:
        with safe_workbook(filepath, save=True) as wb:
            if sheet_name not in wb.sheetnames:
                raise DataError(f"Sheet '{sheet_name}' not found.")

            ws = wb[sheet_name]

            # If no table name is provided, generate a unique one
            if not table_name:
                table_name = f"Table_{uuid.uuid4().hex[:8]}"

            # Check if table name already exists
            if table_name in ws.parent.defined_names:
                raise DataError(f"Table name '{table_name}' already exists.")

            # Create the table
            table = Table(displayName=table_name, ref=data_range)

            # Apply style
            style = TableStyleInfo(
                name=table_style,
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False
            )
            table.tableStyleInfo = style

            ws.add_table(table)

        return {
            "message": f"Successfully created table '{table_name}' in sheet '{sheet_name}'.",
            "table_name": table_name,
            "range": data_range
        }

    except Exception as e:
        logger.error(f"Failed to create table: {e}")
        raise DataError(str(e))


def list_excel_tables(
    filepath: str,
    sheet_name: str | None = None,
) -> list[dict[str, Any]]:
    """List native Excel tables for one sheet or the whole workbook."""
    try:
        with safe_workbook(filepath) as wb:
            if sheet_name is not None and sheet_name not in wb.sheetnames:
                raise DataError(f"Sheet '{sheet_name}' not found.")

            sheet_names = [sheet_name] if sheet_name is not None else list(wb.sheetnames)
            tables: list[dict[str, Any]] = []

            for current_sheet_name in sheet_names:
                ws = wb[current_sheet_name]
                for table in ws.tables.values():
                    tables.append(_build_table_metadata(current_sheet_name, ws, table))

            return tables

    except DataError:
        raise
    except Exception as e:
        logger.error(f"Failed to list tables: {e}")
        raise DataError(str(e))


def read_excel_table(
    filepath: str,
    table_name: str,
    sheet_name: Optional[str] = None,
    max_rows: Optional[int] = None,
    compact: bool = False,
    row_mode: str = "arrays",
    infer_schema: bool = False,
) -> dict[str, Any]:
    """Read rows from a native Excel table by its table name.

    Supports the same array-vs-record row modes and lightweight inferred schema
    hints as the compact worksheet table readers.
    """
    try:
        with safe_workbook(filepath) as wb:
            current_sheet_name, ws, table = _find_table(wb, table_name, sheet_name=sheet_name)
            metadata = _build_table_metadata(current_sheet_name, ws, table)
            min_col, min_row, max_col, max_row = range_boundaries(table.ref)

            data_start_row = min_row + metadata["header_row_count"]
            data_end_row = max_row - metadata["totals_row_count"]
            total_rows = metadata["data_row_count"]
            row_limit = min(total_rows, max_rows) if max_rows is not None else total_rows

            rows: list[list[Any]] = []
            for row_index in range(data_start_row, data_start_row + row_limit):
                if row_index > data_end_row:
                    break
                rows.append(
                    [
                        ws.cell(row=row_index, column=column_index).value
                        for column_index in range(min_col, max_col + 1)
                    ]
                )

            result = {
                "sheet_name": current_sheet_name,
                "table_name": table.displayName,
                "range": table.ref,
                "style": metadata["style"],
                "headers": metadata["headers"],
                "rows": rows,
                "total_rows": total_rows,
                "truncated": max_rows is not None and total_rows > max_rows,
                "header_row_count": metadata["header_row_count"],
                "totals_row_count": metadata["totals_row_count"],
                "totals_row_shown": metadata["totals_row_shown"],
            }
            if compact:
                payload = {
                    "sheet_name": result["sheet_name"],
                    "table_name": result["table_name"],
                    "range": result["range"],
                    "headers": result["headers"],
                    "rows": result["rows"],
                }
                if result["truncated"]:
                    payload["total_rows"] = result["total_rows"]
                    payload["truncated"] = True
            else:
                payload = result

            return augment_tabular_payload(
                payload,
                headers=result["headers"],
                rows=result["rows"],
                row_mode=row_mode,
                infer_schema=infer_schema,
            )

    except DataError:
        raise
    except Exception as e:
        logger.error(f"Failed to read table: {e}")
        raise DataError(str(e))
