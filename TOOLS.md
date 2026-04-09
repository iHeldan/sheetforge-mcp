# SheetForge MCP Tool Reference

This file documents the public MCP tool surface exposed by `src/excel_mcp/server.py`.

## Conventions

- In `stdio` mode, `filepath` must be an absolute path.
- In `streamable-http` and `sse` mode, relative `filepath` values resolve under `EXCEL_FILES_PATH`.
- Destructive tools update the workbook on disk in place unless `dry_run=True`.
- Every tool returns a JSON envelope with `ok`, `operation`, `message`, and `data`.

## Structured Output Shapes

### Shared envelope

All tools return a response like:

```json
{
  "ok": true,
  "operation": "list_all_sheets",
  "message": "list_all_sheets completed",
  "data": {}
}
```

Dry-run responses may also include `dry_run` and `changes`:

```json
{
  "ok": true,
  "operation": "write_data_to_excel",
  "message": "Previewed data to Sheet1",
  "data": {
    "active_sheet": "Sheet1",
    "target_range": "A2:C2",
    "changed_cells": 3
  },
  "dry_run": true,
  "changes": [
    {
      "cell": "A2",
      "old_value": "Alice",
      "new_value": "Mallory"
    }
  ]
}
```

Committed write operations default to compact summaries. Pass `include_changes=True` on supported tools when you want detailed per-cell or per-range diffs.

### `list_all_sheets`

Returns workbook inventory under `data.sheets`:

```json
{
  "ok": true,
  "operation": "list_all_sheets",
  "message": "list_all_sheets completed",
  "data": {
    "sheets": [
      {
        "name": "Sheet1",
        "rows": 12,
        "columns": 4,
        "column_range": "A-D",
        "is_empty": false
      }
    ]
  }
}
```

### `read_excel_as_table`

Returns a compact table object under `data`:

```json
{
  "ok": true,
  "operation": "read_excel_as_table",
  "message": "read_excel_as_table completed",
  "data": {
    "headers": ["Name", "Age", "City"],
    "rows": [["Alice", 30, "Helsinki"]],
    "total_rows": 5,
    "truncated": false,
    "sheet_name": "Sheet1"
  }
}
```

If `compact=True`, the payload is reduced to `headers` and `rows` unless truncation metadata is required.

### `read_data_from_excel`

Returns cell metadata under `data`:

```json
{
  "ok": true,
  "operation": "read_data_from_excel",
  "message": "Read 9 cell(s) from 'Sheet1'",
  "data": {
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
}
```

If `preview_only=True`, the payload is limited to the first 10 rows from the selected range and includes `preview_only` and `truncated` flags.

If `compact=True`, cells without real validation rules omit the default `validation: {"has_validation": false}` stub.

### `search_in_sheet`

Returns matches under `data.matches`:

```json
{
  "ok": true,
  "operation": "search_in_sheet",
  "message": "search_in_sheet completed",
  "data": {
    "sheet_name": "Sheet1",
    "query": 30,
    "exact": true,
    "max_results": 50,
    "matches": [
      {
        "cell": "B2",
        "value": 30,
        "row": 2,
        "column": 2
      }
    ]
  }
}
```

## Workbook And Overview Tools

- `create_workbook(filepath: str) -> str`
  Creates a new workbook file on disk.
- `create_worksheet(filepath: str, sheet_name: str) -> str`
  Adds a worksheet to an existing workbook.
- `get_workbook_metadata(filepath: str, include_ranges: bool = False) -> str`
  Returns workbook metadata under the shared JSON envelope.
- `list_named_ranges(filepath: str) -> str`
  Returns workbook defined names, their values, and any sheet/range destinations.
- `list_all_sheets(filepath: str) -> str`
  Returns one entry per worksheet, including `rows`, `columns`, `column_range`, and `is_empty`.
- `list_tables(filepath: str, sheet_name: Optional[str] = None) -> str`
  Returns native Excel tables across the workbook or for one worksheet, including `sheet_name`, `table_name`, `range`, `style`, `headers`, row counts, and style flags.
- `read_excel_table(filepath: str, table_name: str, sheet_name: Optional[str] = None, max_rows: Optional[int] = None, compact: bool = False) -> str`
  Reads rows from a native Excel table by `table_name`, preserving the table's exact range instead of inferring worksheet bounds. With `compact=True`, the payload keeps `sheet_name`, `table_name`, `range`, `headers`, and `rows`, plus truncation metadata when needed.

## Read, Search, And Write Tools

- `write_data_to_excel(filepath: str, sheet_name: str, data: List[List], start_cell: str = "A1", dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Writes tabular data starting at the given cell. Missing target sheets are created automatically. Returns compact summaries by default on committed writes, and detailed `changes` during previews unless explicitly disabled.
- `read_data_from_excel(filepath: str, sheet_name: str, start_cell: str = "A1", end_cell: Optional[str] = None, preview_only: bool = False, compact: bool = False) -> str`
  Returns cell range data with row, column, address, value, and validation metadata under the shared envelope.
- `read_excel_as_table(filepath: str, sheet_name: str, header_row: int = 1, max_rows: Optional[int] = None, compact: bool = False) -> str`
  Returns `headers`, `rows`, `total_rows`, `truncated`, and `sheet_name`. With `compact=True`, only `headers` and `rows` are returned unless truncation metadata is needed.
- `quick_read(filepath: str, sheet_name: Optional[str] = None, header_row: int = 1, max_rows: Optional[int] = None) -> str`
  Returns a compact table from the requested sheet, or auto-selects the first workbook sheet when `sheet_name` is omitted.
- `search_in_sheet(filepath: str, sheet_name: str, query: Any, exact: bool = True, max_results: int = 50) -> str`
  Returns exact or partial value matches across the worksheet.
- `append_table_rows(filepath: str, sheet_name: str, rows: List[Dict[str, Any]], header_row: int = 1, dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Appends header-aware rows using dictionary keys that match worksheet headers. Returns `changed_cells` always, and detailed `changes` only when requested or during previews.
- `update_rows_by_key(filepath: str, sheet_name: str, key_column: str, updates: List[Dict[str, Any]], header_row: int = 1, dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Updates existing rows by matching a named key column, reports unmatched keys, and returns compact summaries by default.

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

- `format_range(filepath: str, sheet_name: str, start_cell: str, end_cell: Optional[str] = None, bold: bool = False, italic: bool = False, underline: bool = False, font_size: Optional[int] = None, font_color: Optional[str] = None, bg_color: Optional[str] = None, border_style: Optional[str] = None, border_color: Optional[str] = None, number_format: Optional[str] = None, alignment: Optional[str] = None, wrap_text: bool = False, merge_cells: bool = False, protection: Optional[Dict[str, Any]] = None, conditional_format: Optional[Dict[str, Any]] = None, dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Applies formatting options to a cell or range and supports preview mode. Returns compact summaries by default on committed writes.
- `format_ranges(filepath: str, sheet_name: str, ranges: List[Dict[str, Any]], dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Applies formatting to multiple ranges in one workbook pass. Each range object uses the same option keys as `format_range`, such as `start_cell`, `end_cell`, `bold`, `font_size`, `bg_color`, or `conditional_format`.
- `freeze_panes(filepath: str, sheet_name: str, cell: Optional[str] = None, dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Sets freeze panes at the given cell or clears them when `cell` is omitted or `A1`. Supports preview mode, and committed writes stay compact unless `include_changes=True`.
- `set_autofilter(filepath: str, sheet_name: str, range_ref: Optional[str] = None, dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Applies an autofilter to the given range or infers the used range automatically. Supports preview mode, and committed writes stay compact unless `include_changes=True`.
- `set_print_area(filepath: str, sheet_name: str, range_ref: Optional[str] = None, dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Sets a worksheet print area such as `A1:F40`, or clears the existing print area when `range_ref` is omitted. Committed writes stay compact unless `include_changes=True`.
- `set_print_titles(filepath: str, sheet_name: str, rows: Optional[str] = None, columns: Optional[str] = None, dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Sets repeating print title rows and columns. Use `rows=""` or `columns=""` to clear an existing setting while preserving the other dimension.
- `set_column_widths(filepath: str, sheet_name: str, widths: Dict[str, float], dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Sets explicit widths for one or more worksheet columns using a map keyed by column letter. Supports preview mode, and committed writes stay compact unless `include_changes=True`.
- `autofit_columns(filepath: str, sheet_name: str, columns: Optional[List[str]] = None, min_width: float = 8.43, max_width: Optional[float] = None, padding: float = 2.0, dry_run: bool = False) -> str`
  Auto-fits worksheet columns from the current content width. When `columns` is omitted, the tool scans the used worksheet range.
- `set_row_heights(filepath: str, sheet_name: str, heights: Dict[str, float], dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Sets explicit heights for one or more worksheet rows using a map keyed by row number. Supports preview mode, and committed writes stay compact unless `include_changes=True`.
- `merge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str, dry_run: bool = False) -> str`
  Merges the selected range and supports preview mode.
- `unmerge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str, dry_run: bool = False) -> str`
  Unmerges the selected range and supports preview mode.
- `get_merged_cells(filepath: str, sheet_name: str) -> str`
  Returns the worksheet's merged ranges under the shared JSON envelope.

## Worksheet And Range Mutation Tools

- `copy_worksheet(filepath: str, source_sheet: str, target_sheet: str) -> str`
  Copies a worksheet inside the same workbook.
- `delete_worksheet(filepath: str, sheet_name: str) -> str`
  Deletes a worksheet. The final remaining sheet cannot be deleted.
- `rename_worksheet(filepath: str, old_name: str, new_name: str) -> str`
  Renames a worksheet.
- `set_worksheet_visibility(filepath: str, sheet_name: str, visibility: str, dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Sets worksheet visibility to `visible`, `hidden`, or `veryHidden`, and supports preview mode. Committed writes stay compact unless `include_changes=True`.
- `get_worksheet_protection(filepath: str, sheet_name: str) -> str`
  Returns worksheet protection status, password presence, and the current option flags such as `selectUnlockedCells` or `formatCells`.
- `set_worksheet_protection(filepath: str, sheet_name: str, enabled: bool = True, password: Optional[str] = None, options: Optional[Dict[str, bool]] = None, dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Enables or disables worksheet protection and optionally overrides supported protection flags in one call. Committed writes stay compact unless `include_changes=True`.
- `copy_range(filepath: str, sheet_name: str, source_start: str, source_end: str, target_start: str, target_sheet: Optional[str] = None, dry_run: bool = False) -> str`
  Copies a range to another location, optionally on a different sheet, and supports preview mode.
- `delete_range(filepath: str, sheet_name: str, start_cell: str, end_cell: str, shift_direction: str = "up", dry_run: bool = False) -> str`
  Deletes a range and shifts remaining cells up or left. Supports preview mode.
- `insert_rows(filepath: str, sheet_name: str, start_row: int, count: int = 1, dry_run: bool = False) -> str`
  Inserts one or more rows and supports preview mode.
- `insert_columns(filepath: str, sheet_name: str, start_col: int, count: int = 1, dry_run: bool = False) -> str`
  Inserts one or more columns and supports preview mode.
- `delete_sheet_rows(filepath: str, sheet_name: str, start_row: int, count: int = 1, dry_run: bool = False) -> str`
  Deletes one or more rows and supports preview mode.
- `delete_sheet_columns(filepath: str, sheet_name: str, start_col: int, count: int = 1, dry_run: bool = False) -> str`
  Deletes one or more columns and supports preview mode.

## Table, Chart, And Pivot Tools

- `create_table(filepath: str, sheet_name: str, data_range: str, table_name: Optional[str] = None, table_style: str = "TableStyleMedium9") -> str`
  Creates a native Excel table from an existing range.
- `list_charts(filepath: str, sheet_name: Optional[str] = None) -> str`
  Lists embedded charts across the workbook or for one worksheet, including chart type, anchor, titles, and series references.
- `create_chart(filepath: str, sheet_name: str, data_range: str, chart_type: str, target_cell: str, title: str = "", x_axis: str = "", y_axis: str = "") -> str`
  Creates a chart anchored at `target_cell`. Supported chart types are `line`, `bar`, `pie`, `scatter`, and `area`.
- `create_chart_from_series(filepath: str, sheet_name: str, chart_type: str, target_cell: str, series: List[Dict[str, Any]], title: str = "", x_axis: str = "", y_axis: str = "", categories_range: Optional[str] = None, style: Optional[Dict[str, Any]] = None) -> str`
  Creates a chart from explicit series definitions, including non-contiguous ranges. Non-scatter series use `values_range` per series plus an optional shared `categories_range`; scatter series use `x_range` and `y_range`.
- `create_pivot_table(filepath: str, sheet_name: str, data_range: str, rows: List[str], values: List[str], columns: Optional[List[str]] = None, agg_func: str = "sum") -> str`
  Creates a pivot-style summary from the given data range. Supported aggregation functions: `sum`, `average`, `count`, `min`, `max`.

## Practical Usage Notes

- Use `list_all_sheets` before reading unfamiliar workbooks.
- Use `read_excel_table` when the workbook already contains native Excel tables and you want exact table semantics instead of a sheet-wide read.
- Use `read_excel_as_table` when the source data is tabular and headers matter.
- Use `read_data_from_excel` when you need cell addresses or validation metadata.
- Use `compact=True` on read tools when you want to minimize response size for agent workflows.
- Use `format_ranges` instead of repeated `format_range` calls when you're styling a report or dashboard in several places at once.
- Use `create_chart_from_series` when the chart data lives in non-adjacent columns or you want to control each series explicitly.
- Use `set_print_area` and `set_print_titles` when the workbook is meant for printing, export, or PDF generation.
- Use `get_worksheet_protection` before changing protection flags on an unfamiliar workbook.
- Use `autofit_columns` after writing or formatting tables when you want readable output without hand-tuning widths.
- Use `search_in_sheet` to find a value before mutating a workbook.
- Prefer `streamable-http` for long-running remote integrations.
- Prefer `stdio` for local desktop MCP clients.
