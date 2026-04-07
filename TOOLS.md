# Excel MCP Server Tool Reference

This file documents the public MCP tool surface exposed by `src/excel_mcp/server.py`.

## Conventions

- In `stdio` mode, `filepath` must be an absolute path.
- In `streamable-http` and `sse` mode, relative `filepath` values resolve under `EXCEL_FILES_PATH`.
- Destructive tools update the workbook on disk in place.
- Most read tools return JSON strings. A few compatibility-oriented tools still return plain stringified Python data.

## Structured Output Shapes

### `list_all_sheets`

Returns a JSON array like:

```json
[
  {
    "name": "Sheet1",
    "rows": 12,
    "columns": 4,
    "column_range": "A-D",
    "is_empty": false
  }
]
```

### `read_excel_as_table`

Returns a compact JSON object like:

```json
{
  "headers": ["Name", "Age", "City"],
  "rows": [["Alice", 30, "Helsinki"]],
  "total_rows": 5,
  "truncated": false,
  "sheet_name": "Sheet1"
}
```

### `read_data_from_excel`

Returns a JSON object with cell metadata:

```json
{
  "range": "A1:C3",
  "sheet_name": "Sheet1",
  "cells": [
    {
      "address": "A1",
      "value": "Name",
      "row": 1,
      "column": 1,
      "validation": {
        "has_validation": false
      }
    }
  ]
}
```

If `preview_only=True`, the payload is limited to the first 10 rows from the selected range and includes `preview_only` and `truncated` flags.

### `search_in_sheet`

Returns a JSON array of matches:

```json
[
  {
    "cell": "B2",
    "value": 30,
    "row": 2,
    "column": 2
  }
]
```

## Workbook And Overview Tools

- `create_workbook(filepath: str) -> str`
  Creates a new workbook file on disk.
- `create_worksheet(filepath: str, sheet_name: str) -> str`
  Adds a worksheet to an existing workbook.
- `get_workbook_metadata(filepath: str, include_ranges: bool = False) -> str`
  Returns workbook metadata. The current response format is a plain stringified dictionary.
- `list_all_sheets(filepath: str) -> str`
  Returns JSON with one entry per worksheet, including `rows`, `columns`, `column_range`, and `is_empty`.

## Read, Search, And Write Tools

- `write_data_to_excel(filepath: str, sheet_name: str, data: List[List], start_cell: str = "A1") -> str`
  Writes tabular data starting at the given cell. Missing target sheets are created automatically.
- `read_data_from_excel(filepath: str, sheet_name: str, start_cell: str = "A1", end_cell: Optional[str] = None, preview_only: bool = False) -> str`
  Returns JSON for a cell range with row, column, address, value, and validation metadata.
- `read_excel_as_table(filepath: str, sheet_name: str, header_row: int = 1, max_rows: Optional[int] = None) -> str`
  Returns JSON with `headers`, `rows`, `total_rows`, `truncated`, and `sheet_name`.
- `search_in_sheet(filepath: str, sheet_name: str, query: Any, exact: bool = True, max_results: int = 50) -> str`
  Returns JSON matches for exact or partial value search across the worksheet.

## Formula And Validation Tools

- `apply_formula(filepath: str, sheet_name: str, cell: str, formula: str) -> str`
  Writes a validated Excel formula into the target cell.
- `validate_formula_syntax(filepath: str, sheet_name: str, cell: str, formula: str) -> str`
  Validates formula syntax without changing the workbook.
- `validate_excel_range(filepath: str, sheet_name: str, start_cell: str, end_cell: Optional[str] = None) -> str`
  Validates a cell or range reference against the worksheet.
- `get_data_validation_info(filepath: str, sheet_name: str) -> str`
  Returns JSON for data-validation rules defined in the worksheet.

## Formatting And Layout Tools

- `format_range(filepath: str, sheet_name: str, start_cell: str, end_cell: Optional[str] = None, bold: bool = False, italic: bool = False, underline: bool = False, font_size: Optional[int] = None, font_color: Optional[str] = None, bg_color: Optional[str] = None, border_style: Optional[str] = None, border_color: Optional[str] = None, number_format: Optional[str] = None, alignment: Optional[str] = None, wrap_text: bool = False, merge_cells: bool = False, protection: Optional[Dict[str, Any]] = None, conditional_format: Optional[Dict[str, Any]] = None) -> str`
  Applies formatting options to a cell or range.
- `merge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> str`
  Merges the selected range.
- `unmerge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> str`
  Unmerges the selected range.
- `get_merged_cells(filepath: str, sheet_name: str) -> str`
  Returns the worksheet's merged ranges as a plain stringified list.

## Worksheet And Range Mutation Tools

- `copy_worksheet(filepath: str, source_sheet: str, target_sheet: str) -> str`
  Copies a worksheet inside the same workbook.
- `delete_worksheet(filepath: str, sheet_name: str) -> str`
  Deletes a worksheet. The final remaining sheet cannot be deleted.
- `rename_worksheet(filepath: str, old_name: str, new_name: str) -> str`
  Renames a worksheet.
- `copy_range(filepath: str, sheet_name: str, source_start: str, source_end: str, target_start: str, target_sheet: Optional[str] = None) -> str`
  Copies a range to another location, optionally on a different sheet.
- `delete_range(filepath: str, sheet_name: str, start_cell: str, end_cell: str, shift_direction: str = "up") -> str`
  Deletes a range and shifts remaining cells up or left.
- `insert_rows(filepath: str, sheet_name: str, start_row: int, count: int = 1) -> str`
  Inserts one or more rows.
- `insert_columns(filepath: str, sheet_name: str, start_col: int, count: int = 1) -> str`
  Inserts one or more columns.
- `delete_sheet_rows(filepath: str, sheet_name: str, start_row: int, count: int = 1) -> str`
  Deletes one or more rows.
- `delete_sheet_columns(filepath: str, sheet_name: str, start_col: int, count: int = 1) -> str`
  Deletes one or more columns.

## Table, Chart, And Pivot Tools

- `create_table(filepath: str, sheet_name: str, data_range: str, table_name: Optional[str] = None, table_style: str = "TableStyleMedium9") -> str`
  Creates a native Excel table from an existing range.
- `create_chart(filepath: str, sheet_name: str, data_range: str, chart_type: str, target_cell: str, title: str = "", x_axis: str = "", y_axis: str = "") -> str`
  Creates a chart anchored at `target_cell`. Supported chart types are `line`, `bar`, `pie`, `scatter`, and `area`.
- `create_pivot_table(filepath: str, sheet_name: str, data_range: str, rows: List[str], values: List[str], columns: Optional[List[str]] = None, agg_func: str = "mean") -> str`
  Creates a pivot-style summary from the given data range.

## Practical Usage Notes

- Use `list_all_sheets` before reading unfamiliar workbooks.
- Use `read_excel_as_table` when the source data is tabular and headers matter.
- Use `read_data_from_excel` when you need cell addresses or validation metadata.
- Use `search_in_sheet` to find a value before mutating a workbook.
- Prefer `streamable-http` for long-running remote integrations.
- Prefer `stdio` for local desktop MCP clients.
