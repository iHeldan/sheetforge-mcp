import logging
import os
import json
from typing import Any, Callable, Dict, List, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

# Import exceptions
from excel_mcp.exceptions import (
    ValidationError,
    WorkbookError,
    SheetError,
    DataError,
    FormattingError,
    CalculationError,
    PivotError,
    ChartError
)

# Import from excel_mcp package with consistent _impl suffixes
from excel_mcp.validation import (
    validate_formula_in_cell_operation as validate_formula_impl,
    validate_range_in_sheet_operation as validate_range_impl
)
from excel_mcp.chart import (
    create_chart_from_series as create_chart_from_series_impl,
    create_chart_in_sheet as create_chart_impl,
    list_charts as list_charts_impl,
)
from excel_mcp.workbook import get_workbook_info, list_named_ranges as list_named_ranges_impl
from excel_mcp.data import (
    append_table_rows as append_table_rows_impl,
    quick_read as quick_read_impl,
    read_as_table,
    search_cells,
    update_rows_by_key as update_rows_by_key_impl,
    write_data,
)
from excel_mcp.pivot import create_pivot_table as create_pivot_table_impl
from excel_mcp.tables import (
    create_excel_table as create_table_impl,
    list_excel_tables as list_tables_impl,
    read_excel_table as read_excel_table_impl,
)
from excel_mcp.sheet import (
    autofit_columns as autofit_columns_impl,
    copy_sheet,
    delete_sheet,
    get_sheet_protection as get_sheet_protection_impl,
    set_print_area as set_print_area_impl,
    set_print_titles as set_print_titles_impl,
    set_sheet_protection as set_sheet_protection_impl,
    rename_sheet,
    set_sheet_visibility,
    merge_range,
    unmerge_range,
    get_merged_ranges,
    set_auto_filter,
    set_freeze_panes,
    set_column_widths as set_column_widths_impl,
    set_row_heights as set_row_heights_impl,
    insert_row,
    insert_cols,
    delete_rows,
    delete_cols,
)

# Get project root directory path for log file path.
# When using the stdio transmission method,
# relative paths may cause log files to fail to create
# due to the client's running location and permission issues,
# resulting in the program not being able to run.
# Thus using os.path.join(ROOT_DIR, "excel-mcp.log") instead.

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
LOG_FILE = os.path.join(ROOT_DIR, "excel-mcp.log")

# Initialize EXCEL_FILES_PATH variable without assigning a value
EXCEL_FILES_PATH = None

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        # Referring to https://github.com/modelcontextprotocol/python-sdk/issues/409#issuecomment-2816831318
        # The stdio mode server MUST NOT write anything to its stdout that is not a valid MCP message.
        logging.FileHandler(LOG_FILE)
    ],
)
logger = logging.getLogger("excel-mcp")
# Initialize FastMCP server
mcp = FastMCP(
    "sheetforge-mcp",
    host=os.environ.get("FASTMCP_HOST", "127.0.0.1"),
    port=int(os.environ.get("FASTMCP_PORT", "8017")),
    instructions="SheetForge MCP for manipulating Excel workbooks"
)

HANDLED_TOOL_ERRORS = (
    ValidationError,
    WorkbookError,
    SheetError,
    DataError,
    FormattingError,
    CalculationError,
    PivotError,
    ChartError,
    ValueError,
    FileNotFoundError,
)


def _extract_payload_parts(result: Any) -> tuple[Optional[str], Any, Dict[str, Any]]:
    if isinstance(result, dict):
        payload = dict(result)
        message = payload.pop("message", None)
        meta = {}
        for key in ("dry_run", "changes", "warnings", "preview"):
            if key in payload:
                meta[key] = payload.pop(key)
        data = payload or None
        return message, data, meta

    if isinstance(result, str):
        return result, None, {}

    return None, result, {}


def _success_response(
    operation: str,
    *,
    result: Any = None,
    message: Optional[str] = None,
    data: Any = None,
    **meta: Any,
) -> str:
    extracted_message, extracted_data, extracted_meta = _extract_payload_parts(result)
    if data is None:
        data = extracted_data

    payload = {
        "ok": True,
        "operation": operation,
        "message": message or extracted_message or f"{operation} completed",
        "data": data,
    }
    payload.update({key: value for key, value in extracted_meta.items() if value is not None})
    payload.update({key: value for key, value in meta.items() if value is not None})
    return json.dumps(payload, indent=2, default=str)


def _error_response(operation: str, error: Exception) -> str:
    return json.dumps(
        {
            "ok": False,
            "operation": operation,
            "error": {
                "type": type(error).__name__,
                "message": str(error),
            },
        },
        indent=2,
        default=str,
    )


def _run_tool(operation: str, action: Callable[[], Any], *, message: Optional[str] = None) -> str:
    try:
        return _success_response(operation, result=action(), message=message)
    except HANDLED_TOOL_ERRORS as e:
        logger.error("%s failed: %s", operation, e)
        return _error_response(operation, e)
    except Exception as e:
        logger.exception("Unhandled error in %s", operation)
        return _error_response(operation, e)

def get_excel_path(filename: str) -> str:
    """Get full path to Excel file.
    
    Args:
        filename: Name of Excel file
        
    Returns:
        Full path to Excel file
    """
    # If filename is already an absolute path, return it
    if os.path.isabs(filename):
        return filename

    # Check if in SSE mode (EXCEL_FILES_PATH is not None)
    if EXCEL_FILES_PATH is None:
        # Must use absolute path
        raise ValueError(f"Invalid filename: {filename}, must be an absolute path when not in SSE mode")

    # In SSE mode, if it's a relative path, resolve it based on EXCEL_FILES_PATH
    return os.path.join(EXCEL_FILES_PATH, filename)

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Apply Formula",
        destructiveHint=True,
    ),
)
def apply_formula(
    filepath: str,
    sheet_name: str,
    cell: str,
    formula: str,
) -> str:
    """
    Apply Excel formula to cell.
    Excel formula will write to cell with verification.
    """
    def action() -> Any:
        full_path = get_excel_path(filepath)
        validation = validate_formula_impl(full_path, sheet_name, cell, formula)
        if isinstance(validation, dict) and "error" in validation:
            raise ValidationError(validation["error"])

        from excel_mcp.calculations import apply_formula as apply_formula_impl

        return apply_formula_impl(full_path, sheet_name, cell, formula)

    return _run_tool("apply_formula", action)

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Validate Formula Syntax",
        readOnlyHint=True,
    ),
)
def validate_formula_syntax(
    filepath: str,
    sheet_name: str,
    cell: str,
    formula: str,
) -> str:
    """Validate Excel formula syntax without applying it."""
    return _run_tool(
        "validate_formula_syntax",
        lambda: validate_formula_impl(get_excel_path(filepath), sheet_name, cell, formula),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Format Range",
        destructiveHint=True,
    ),
)
def format_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: Optional[str] = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_size: Optional[int] = None,
    font_color: Optional[str] = None,
    bg_color: Optional[str] = None,
    border_style: Optional[str] = None,
    border_color: Optional[str] = None,
    number_format: Optional[str] = None,
    alignment: Optional[str] = None,
    wrap_text: bool = False,
    merge_cells: bool = False,
    protection: Optional[Dict[str, Any]] = None,
    conditional_format: Optional[Dict[str, Any]] = None,
    dry_run: bool = False,
    include_changes: Optional[bool] = None,
) -> str:
    """Apply formatting to a range of cells."""
    def action() -> Any:
        full_path = get_excel_path(filepath)
        from excel_mcp.formatting import format_range as format_range_func

        return format_range_func(
            filepath=full_path,
            sheet_name=sheet_name,
            start_cell=start_cell,
            end_cell=end_cell,
            bold=bold,
            italic=italic,
            underline=underline,
            font_size=font_size,
            font_color=font_color,
            bg_color=bg_color,
            border_style=border_style,
            border_color=border_color,
            number_format=number_format,
            alignment=alignment,
            wrap_text=wrap_text,
            merge_cells=merge_cells,
            protection=protection,
            conditional_format=conditional_format,
            dry_run=dry_run,
            include_changes=include_changes,
        )

    return _run_tool("format_range", action)


@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Format Multiple Ranges",
        destructiveHint=True,
    ),
)
def format_ranges(
    filepath: str,
    sheet_name: str,
    ranges: List[Dict[str, Any]],
    dry_run: bool = False,
    include_changes: Optional[bool] = None,
) -> str:
    """Apply formatting to multiple ranges in a single workbook pass."""
    def action() -> Any:
        full_path = get_excel_path(filepath)
        from excel_mcp.formatting import format_ranges as format_ranges_impl

        return format_ranges_impl(
            filepath=full_path,
            sheet_name=sheet_name,
            ranges=ranges,
            dry_run=dry_run,
            include_changes=include_changes,
        )

    return _run_tool("format_ranges", action)

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Read Data from Excel",
        readOnlyHint=True,
    ),
)
def read_data_from_excel(
    filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: Optional[str] = None,
    preview_only: bool = False,
    compact: bool = False,
) -> str:
    """
    Read data from Excel worksheet with cell metadata including validation rules.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_cell: Starting cell (default A1)
        end_cell: Ending cell (optional, auto-expands if not provided)
        preview_only: Whether to return preview only
        compact: Whether to omit default validation metadata for smaller responses
    
    Returns:  
    JSON string containing structured cell data with validation metadata.
    Each cell includes: address, value, row, column, and validation info (if any).
    """
    try:
        full_path = get_excel_path(filepath)
        from excel_mcp.data import read_excel_range_with_metadata
        result = read_excel_range_with_metadata(
            full_path,
            sheet_name,
            start_cell,
            end_cell,
            compact=compact,
        )
        if not result:
            result = {"range": f"{start_cell}:{end_cell}" if end_cell else start_cell, "sheet_name": sheet_name, "cells": []}

        if preview_only:
            total_cells = len(result["cells"])
            preview_row_limit = 10
            preview_cells = []
            preview_rows = set()

            for cell in result["cells"]:
                row = cell["row"]
                if row not in preview_rows and len(preview_rows) >= preview_row_limit:
                    break
                preview_rows.add(row)
                preview_cells.append(cell)

            result["cells"] = preview_cells
            result["preview_only"] = True
            result["truncated"] = len(preview_cells) < total_cells

        return _success_response(
            "read_data_from_excel",
            result=result,
            message=f"Read {len(result['cells'])} cell(s) from '{sheet_name}'",
        )

    except HANDLED_TOOL_ERRORS as e:
        logger.error("read_data_from_excel failed: %s", e)
        return _error_response("read_data_from_excel", e)
    except Exception as e:
        logger.exception("Unhandled error in read_data_from_excel")
        return _error_response("read_data_from_excel", e)

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Read Excel as Table",
        readOnlyHint=True,
    ),
)
def read_excel_as_table(
    filepath: str,
    sheet_name: str,
    header_row: int = 1,
    max_rows: Optional[int] = None,
    compact: bool = False,
) -> str:
    """
    Read Excel data as a compact table with headers and rows.
    Much more context-efficient than read_data_from_excel for structured data.
    """
    return _run_tool(
        "read_excel_as_table",
        lambda: read_as_table(
            get_excel_path(filepath),
            sheet_name,
            header_row=header_row,
            max_rows=max_rows,
            compact=compact,
        ),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Quick Read",
        readOnlyHint=True,
    ),
)
def quick_read(
    filepath: str,
    sheet_name: Optional[str] = None,
    header_row: int = 1,
    max_rows: Optional[int] = None,
) -> str:
    """Read a compact table from an explicit sheet or the first workbook sheet."""
    return _run_tool(
        "quick_read",
        lambda: quick_read_impl(
            get_excel_path(filepath),
            sheet_name=sheet_name,
            header_row=header_row,
            max_rows=max_rows,
        ),
    )


@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Read Excel Table",
        readOnlyHint=True,
    ),
)
def read_excel_table(
    filepath: str,
    table_name: str,
    sheet_name: Optional[str] = None,
    max_rows: Optional[int] = None,
    compact: bool = False,
) -> str:
    """Read a native Excel table by its table name."""
    return _run_tool(
        "read_excel_table",
        lambda: read_excel_table_impl(
            get_excel_path(filepath),
            table_name,
            sheet_name=sheet_name,
            max_rows=max_rows,
            compact=compact,
        ),
    )


@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Write Data to Excel",
        destructiveHint=True,
    ),
)
def write_data_to_excel(
    filepath: str,
    sheet_name: str,
    data: List[List],
    start_cell: str = "A1",
    dry_run: bool = False,
    include_changes: Optional[bool] = None,
) -> str:
    """
    Write data to Excel worksheet.
    Excel formula will write to cell without any verification.

    PARAMETERS:  
    filepath: Path to Excel file
    sheet_name: Name of worksheet to write to
    data: List of lists containing data to write to the worksheet, sublists are assumed to be rows
    start_cell: Cell to start writing to, default is "A1"
  
    """
    return _run_tool(
        "write_data_to_excel",
        lambda: write_data(
            get_excel_path(filepath),
            sheet_name,
            data,
            start_cell,
            dry_run=dry_run,
            include_changes=include_changes,
        ),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Create Workbook",
        destructiveHint=True,
    ),
)
def create_workbook(filepath: str) -> str:
    """Create new Excel workbook."""
    def action() -> Any:
        full_path = get_excel_path(filepath)
        from excel_mcp.workbook import create_workbook as create_workbook_impl

        result = create_workbook_impl(full_path)
        result.pop("workbook", None)
        result["filepath"] = full_path
        return result

    return _run_tool("create_workbook", action)

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Create Worksheet",
        destructiveHint=True,
    ),
)
def create_worksheet(filepath: str, sheet_name: str) -> str:
    """Create new worksheet in workbook."""
    def action() -> Any:
        full_path = get_excel_path(filepath)
        from excel_mcp.workbook import create_sheet as create_worksheet_impl

        result = create_worksheet_impl(full_path, sheet_name)
        result["sheet_name"] = sheet_name
        return result

    return _run_tool("create_worksheet", action)

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Create Chart",
        destructiveHint=True,
    ),
)
def create_chart(
    filepath: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = ""
) -> str:
    """Create chart in worksheet."""
    def action() -> Any:
        return create_chart_impl(
            filepath=get_excel_path(filepath),
            sheet_name=sheet_name,
            data_range=data_range,
            chart_type=chart_type,
            target_cell=target_cell,
            title=title,
            x_axis=x_axis,
            y_axis=y_axis
        )

    return _run_tool("create_chart", action)


@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Create Chart From Series",
        destructiveHint=True,
    ),
)
def create_chart_from_series(
    filepath: str,
    sheet_name: str,
    chart_type: str,
    target_cell: str,
    series: List[Dict[str, Any]],
    title: str = "",
    x_axis: str = "",
    y_axis: str = "",
    categories_range: Optional[str] = None,
    style: Optional[Dict[str, Any]] = None,
) -> str:
    """Create a chart from explicit series definitions for non-contiguous ranges."""
    def action() -> Any:
        return create_chart_from_series_impl(
            filepath=get_excel_path(filepath),
            sheet_name=sheet_name,
            chart_type=chart_type,
            target_cell=target_cell,
            series=series,
            title=title,
            x_axis=x_axis,
            y_axis=y_axis,
            categories_range=categories_range,
            style=style,
        )

    return _run_tool("create_chart_from_series", action)


@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="List Charts",
        readOnlyHint=True,
    ),
)
def list_charts(filepath: str, sheet_name: Optional[str] = None) -> str:
    """List embedded charts in a workbook or worksheet."""
    def action() -> Any:
        return {"charts": list_charts_impl(get_excel_path(filepath), sheet_name=sheet_name)}

    return _run_tool("list_charts", action)

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Create Pivot Table",
        destructiveHint=True,
    ),
)
def create_pivot_table(
    filepath: str,
    sheet_name: str,
    data_range: str,
    rows: List[str],
    values: List[str],
    columns: Optional[List[str]] = None,
    agg_func: str = "sum"
) -> str:
    """Create pivot table in worksheet."""
    return _run_tool(
        "create_pivot_table",
        lambda: create_pivot_table_impl(
            filepath=get_excel_path(filepath),
            sheet_name=sheet_name,
            data_range=data_range,
            rows=rows,
            values=values,
            columns=columns or [],
            agg_func=agg_func
        ),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Create Table",
        destructiveHint=True,
    ),
)
def create_table(
    filepath: str,
    sheet_name: str,
    data_range: str,
    table_name: Optional[str] = None,
    table_style: str = "TableStyleMedium9"
) -> str:
    """Creates a native Excel table from a specified range of data."""
    return _run_tool(
        "create_table",
        lambda: create_table_impl(
            filepath=get_excel_path(filepath),
            sheet_name=sheet_name,
            data_range=data_range,
            table_name=table_name,
            table_style=table_style
        ),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="List Excel Tables",
        readOnlyHint=True,
    ),
)
def list_tables(
    filepath: str,
    sheet_name: Optional[str] = None,
) -> str:
    """List native Excel tables for one worksheet or the whole workbook."""
    def action() -> Any:
        return {
            "sheet_name": sheet_name,
            "tables": list_tables_impl(get_excel_path(filepath), sheet_name=sheet_name),
        }

    return _run_tool("list_tables", action)

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Copy Worksheet",
        destructiveHint=True,
    ),
)
def copy_worksheet(
    filepath: str,
    source_sheet: str,
    target_sheet: str
) -> str:
    """Copy worksheet within workbook."""
    return _run_tool(
        "copy_worksheet",
        lambda: copy_sheet(get_excel_path(filepath), source_sheet, target_sheet),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Delete Worksheet",
        destructiveHint=True,
    ),
)
def delete_worksheet(
    filepath: str,
    sheet_name: str
) -> str:
    """Delete worksheet from workbook."""
    return _run_tool(
        "delete_worksheet",
        lambda: delete_sheet(get_excel_path(filepath), sheet_name),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Rename Worksheet",
        destructiveHint=True,
    ),
)
def rename_worksheet(
    filepath: str,
    old_name: str,
    new_name: str
) -> str:
    """Rename worksheet in workbook."""
    return _run_tool(
        "rename_worksheet",
        lambda: rename_sheet(get_excel_path(filepath), old_name, new_name),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Get Workbook Metadata",
        readOnlyHint=True,
    ),
)
def get_workbook_metadata(
    filepath: str,
    include_ranges: bool = False
) -> str:
    """Get metadata about workbook including sheets, ranges, etc."""
    return _run_tool(
        "get_workbook_metadata",
        lambda: get_workbook_info(get_excel_path(filepath), include_ranges=include_ranges),
    )


@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="List Named Ranges",
        readOnlyHint=True,
    ),
)
def list_named_ranges(filepath: str) -> str:
    """List workbook defined names and their destinations."""
    def action() -> Any:
        return {"named_ranges": list_named_ranges_impl(get_excel_path(filepath))}

    return _run_tool("list_named_ranges", action)

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Merge Cells",
        destructiveHint=True,
    ),
)
def merge_cells(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    dry_run: bool = False,
) -> str:
    """Merge a range of cells."""
    return _run_tool(
        "merge_cells",
        lambda: merge_range(get_excel_path(filepath), sheet_name, start_cell, end_cell, dry_run=dry_run),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Unmerge Cells",
        destructiveHint=True,
    ),
)
def unmerge_cells(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    dry_run: bool = False,
) -> str:
    """Unmerge a range of cells."""
    return _run_tool(
        "unmerge_cells",
        lambda: unmerge_range(get_excel_path(filepath), sheet_name, start_cell, end_cell, dry_run=dry_run),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Get Merged Cells",
        readOnlyHint=True,
    ),
)
def get_merged_cells(filepath: str, sheet_name: str) -> str:
    """Get merged cells in a worksheet."""
    def action() -> Any:
        return {"sheet_name": sheet_name, "ranges": get_merged_ranges(get_excel_path(filepath), sheet_name)}

    return _run_tool("get_merged_cells", action)


@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Freeze Panes",
        destructiveHint=True,
    ),
)
def freeze_panes(
    filepath: str,
    sheet_name: str,
    cell: Optional[str] = None,
    dry_run: bool = False,
) -> str:
    """Set or clear worksheet freeze panes."""
    return _run_tool(
        "freeze_panes",
        lambda: set_freeze_panes(get_excel_path(filepath), sheet_name, cell, dry_run=dry_run),
    )


@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Set Autofilter",
        destructiveHint=True,
    ),
)
def set_autofilter(
    filepath: str,
    sheet_name: str,
    range_ref: Optional[str] = None,
    dry_run: bool = False,
) -> str:
    """Set worksheet autofilter for an explicit or inferred range."""
    return _run_tool(
        "set_autofilter",
        lambda: set_auto_filter(get_excel_path(filepath), sheet_name, range_ref, dry_run=dry_run),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Set Worksheet Visibility",
        destructiveHint=True,
    ),
)
def set_worksheet_visibility(
    filepath: str,
    sheet_name: str,
    visibility: str,
    dry_run: bool = False,
) -> str:
    """Set worksheet visibility to visible, hidden, or veryHidden."""
    return _run_tool(
        "set_worksheet_visibility",
        lambda: set_sheet_visibility(
            get_excel_path(filepath),
            sheet_name,
            visibility,
            dry_run=dry_run,
        ),
    )


@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Get Worksheet Protection",
        readOnlyHint=True,
    ),
)
def get_worksheet_protection(filepath: str, sheet_name: str) -> str:
    """Get worksheet protection status and option flags."""
    return _run_tool(
        "get_worksheet_protection",
        lambda: get_sheet_protection_impl(get_excel_path(filepath), sheet_name),
    )


@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Set Worksheet Protection",
        destructiveHint=True,
    ),
)
def set_worksheet_protection(
    filepath: str,
    sheet_name: str,
    enabled: bool = True,
    password: Optional[str] = None,
    options: Optional[Dict[str, bool]] = None,
    dry_run: bool = False,
) -> str:
    """Enable or disable worksheet protection with optional capability flags."""
    return _run_tool(
        "set_worksheet_protection",
        lambda: set_sheet_protection_impl(
            get_excel_path(filepath),
            sheet_name,
            enabled=enabled,
            password=password,
            options=options,
            dry_run=dry_run,
        ),
    )


@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Set Print Area",
        destructiveHint=True,
    ),
)
def set_print_area(
    filepath: str,
    sheet_name: str,
    range_ref: Optional[str] = None,
    dry_run: bool = False,
) -> str:
    """Set or clear worksheet print area."""
    return _run_tool(
        "set_print_area",
        lambda: set_print_area_impl(
            get_excel_path(filepath),
            sheet_name,
            range_ref=range_ref,
            dry_run=dry_run,
        ),
    )


@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Set Print Titles",
        destructiveHint=True,
    ),
)
def set_print_titles(
    filepath: str,
    sheet_name: str,
    rows: Optional[str] = None,
    columns: Optional[str] = None,
    dry_run: bool = False,
) -> str:
    """Set, preserve, or clear repeating print title rows and columns."""
    return _run_tool(
        "set_print_titles",
        lambda: set_print_titles_impl(
            get_excel_path(filepath),
            sheet_name,
            rows=rows,
            columns=columns,
            dry_run=dry_run,
        ),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Set Column Widths",
        destructiveHint=True,
    ),
)
def set_column_widths(
    filepath: str,
    sheet_name: str,
    widths: Dict[str, float],
    dry_run: bool = False,
) -> str:
    """Set explicit widths for one or more worksheet columns."""
    return _run_tool(
        "set_column_widths",
        lambda: set_column_widths_impl(
            get_excel_path(filepath),
            sheet_name,
            widths,
            dry_run=dry_run,
        ),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Autofit Columns",
        destructiveHint=True,
    ),
)
def autofit_columns(
    filepath: str,
    sheet_name: str,
    columns: Optional[List[str]] = None,
    min_width: float = 8.43,
    max_width: Optional[float] = None,
    padding: float = 2.0,
    dry_run: bool = False,
) -> str:
    """Auto-fit worksheet columns based on content width."""
    return _run_tool(
        "autofit_columns",
        lambda: autofit_columns_impl(
            get_excel_path(filepath),
            sheet_name,
            columns=columns,
            min_width=min_width,
            max_width=max_width,
            padding=padding,
            dry_run=dry_run,
        ),
    )


@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Set Row Heights",
        destructiveHint=True,
    ),
)
def set_row_heights(
    filepath: str,
    sheet_name: str,
    heights: Dict[str, float],
    dry_run: bool = False,
) -> str:
    """Set explicit heights for one or more worksheet rows."""
    return _run_tool(
        "set_row_heights",
        lambda: set_row_heights_impl(
            get_excel_path(filepath),
            sheet_name,
            heights,
            dry_run=dry_run,
        ),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Copy Range",
        destructiveHint=True,
    ),
)
def copy_range(
    filepath: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: Optional[str] = None,
    dry_run: bool = False,
) -> str:
    """Copy a range of cells to another location."""
    def action() -> Any:
        from excel_mcp.sheet import copy_range_operation

        return copy_range_operation(
            get_excel_path(filepath),
            sheet_name,
            source_start,
            source_end,
            target_start,
            target_sheet or sheet_name,
            dry_run=dry_run,
        )

    return _run_tool("copy_range", action)

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Delete Range",
        destructiveHint=True,
    ),
)
def delete_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    shift_direction: str = "up",
    dry_run: bool = False,
) -> str:
    """Delete a range of cells and shift remaining cells."""
    def action() -> Any:
        from excel_mcp.sheet import delete_range_operation

        return delete_range_operation(
            get_excel_path(filepath),
            sheet_name,
            start_cell,
            end_cell,
            shift_direction,
            dry_run=dry_run,
        )

    return _run_tool("delete_range", action)

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Validate Excel Range",
        readOnlyHint=True,
    ),
)
def validate_excel_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: Optional[str] = None
) -> str:
    """Validate if a range exists and is properly formatted."""
    range_str = start_cell if not end_cell else f"{start_cell}:{end_cell}"
    return _run_tool(
        "validate_excel_range",
        lambda: validate_range_impl(get_excel_path(filepath), sheet_name, range_str),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Get Data Validation Info",
        readOnlyHint=True,
    ),
)
def get_data_validation_info(
    filepath: str,
    sheet_name: str
) -> str:
    """
    Get all data validation rules in a worksheet.
    
    This tool helps identify which cell ranges have validation rules
    and what types of validation are applied.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        
    Returns:
        JSON string containing all validation rules in the worksheet
    """
    try:
        full_path = get_excel_path(filepath)
        from excel_mcp.workbook import safe_workbook
        from excel_mcp.cell_validation import get_all_validation_ranges

        with safe_workbook(full_path) as wb:
            if sheet_name not in wb.sheetnames:
                raise SheetError(f"Sheet '{sheet_name}' not found")

            ws = wb[sheet_name]
            validations = get_all_validation_ranges(ws)

        return _success_response(
            "get_data_validation_info",
            data={
                "sheet_name": sheet_name,
                "validation_rules": validations,
            },
            message=(
                "No data validation rules found in this worksheet"
                if not validations
                else f"Found {len(validations)} validation rule(s) in '{sheet_name}'"
            ),
        )

    except HANDLED_TOOL_ERRORS as e:
        logger.error("get_data_validation_info failed: %s", e)
        return _error_response("get_data_validation_info", e)
    except Exception as e:
        logger.exception("Unhandled error in get_data_validation_info")
        return _error_response("get_data_validation_info", e)

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Insert Rows",
        destructiveHint=True,
    ),
)
def insert_rows(
    filepath: str,
    sheet_name: str,
    start_row: int,
    count: int = 1,
    dry_run: bool = False,
) -> str:
    """Insert one or more rows starting at the specified row."""
    return _run_tool(
        "insert_rows",
        lambda: insert_row(get_excel_path(filepath), sheet_name, start_row, count, dry_run=dry_run),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Insert Columns",
        destructiveHint=True,
    ),
)
def insert_columns(
    filepath: str,
    sheet_name: str,
    start_col: int,
    count: int = 1,
    dry_run: bool = False,
) -> str:
    """Insert one or more columns starting at the specified column."""
    return _run_tool(
        "insert_columns",
        lambda: insert_cols(get_excel_path(filepath), sheet_name, start_col, count, dry_run=dry_run),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Delete Rows",
        destructiveHint=True,
    ),
)
def delete_sheet_rows(
    filepath: str,
    sheet_name: str,
    start_row: int,
    count: int = 1,
    dry_run: bool = False,
) -> str:
    """Delete one or more rows starting at the specified row."""
    return _run_tool(
        "delete_sheet_rows",
        lambda: delete_rows(get_excel_path(filepath), sheet_name, start_row, count, dry_run=dry_run),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Delete Columns",
        destructiveHint=True,
    ),
)
def delete_sheet_columns(
    filepath: str,
    sheet_name: str,
    start_col: int,
    count: int = 1,
    dry_run: bool = False,
) -> str:
    """Delete one or more columns starting at the specified column."""
    return _run_tool(
        "delete_sheet_columns",
        lambda: delete_cols(get_excel_path(filepath), sheet_name, start_col, count, dry_run=dry_run),
    )

@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Search Cells",
        readOnlyHint=True,
    ),
)
def search_in_sheet(
    filepath: str,
    sheet_name: str,
    query: Any,
    exact: bool = True,
    max_results: int = 50,
) -> str:
    """
    Search for cells matching a value in a worksheet.
    """
    def action() -> Any:
        results = search_cells(get_excel_path(filepath), sheet_name, query, exact=exact, max_results=max_results)
        return {
            "sheet_name": sheet_name,
            "query": query,
            "exact": exact,
            "max_results": max_results,
            "matches": results,
        }

    return _run_tool("search_in_sheet", action)


@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="List Sheets",
        readOnlyHint=True,
    ),
)
def list_all_sheets(filepath: str) -> str:
    """
    List all sheets in a workbook with row/column counts.
    Quick overview before reading data.
    """
    def action() -> Any:
        from excel_mcp.workbook import list_sheets

        return {"sheets": list_sheets(get_excel_path(filepath))}

    return _run_tool("list_all_sheets", action)


@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Append Table Rows",
        destructiveHint=True,
    ),
)
def append_table_rows(
    filepath: str,
    sheet_name: str,
    rows: List[Dict[str, Any]],
    header_row: int = 1,
    dry_run: bool = False,
    include_changes: Optional[bool] = None,
) -> str:
    """Append dictionary-shaped rows by matching worksheet headers."""
    return _run_tool(
        "append_table_rows",
        lambda: append_table_rows_impl(
            get_excel_path(filepath),
            sheet_name,
            rows,
            header_row=header_row,
            dry_run=dry_run,
            include_changes=include_changes,
        ),
    )


@mcp.tool(
    structured_output=False,
    annotations=ToolAnnotations(
        title="Update Rows by Key",
        destructiveHint=True,
    ),
)
def update_rows_by_key(
    filepath: str,
    sheet_name: str,
    key_column: str,
    updates: List[Dict[str, Any]],
    header_row: int = 1,
    dry_run: bool = False,
    include_changes: Optional[bool] = None,
) -> str:
    """Update existing table rows using a named key column."""
    return _run_tool(
        "update_rows_by_key",
        lambda: update_rows_by_key_impl(
            get_excel_path(filepath),
            sheet_name,
            key_column,
            updates,
            header_row=header_row,
            dry_run=dry_run,
            include_changes=include_changes,
        ),
    )


def run_sse():
    """Run Excel MCP server in SSE mode."""
    # Assign value to EXCEL_FILES_PATH in SSE mode
    global EXCEL_FILES_PATH
    EXCEL_FILES_PATH = os.environ.get("EXCEL_FILES_PATH", "./excel_files")
    # Create directory if it doesn't exist
    os.makedirs(EXCEL_FILES_PATH, exist_ok=True)
    
    try:
        logger.info(f"Starting Excel MCP server with SSE transport (files directory: {EXCEL_FILES_PATH})")
        mcp.run(transport="sse")
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Server shutdown complete")

def run_streamable_http():
    """Run Excel MCP server in streamable HTTP mode."""
    # Assign value to EXCEL_FILES_PATH in streamable HTTP mode
    global EXCEL_FILES_PATH
    EXCEL_FILES_PATH = os.environ.get("EXCEL_FILES_PATH", "./excel_files")
    # Create directory if it doesn't exist
    os.makedirs(EXCEL_FILES_PATH, exist_ok=True)
    
    try:
        logger.info(f"Starting Excel MCP server with streamable HTTP transport (files directory: {EXCEL_FILES_PATH})")
        mcp.run(transport="streamable-http")
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Server shutdown complete")

def run_stdio():
    """Run Excel MCP server in stdio mode."""
    # No need to assign EXCEL_FILES_PATH in stdio mode

    try:
        logger.info("Starting Excel MCP server with stdio transport")
        mcp.run(transport="stdio")
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Server shutdown complete")
