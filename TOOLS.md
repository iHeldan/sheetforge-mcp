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

Responses are serialized as compact JSON instead of pretty-printed JSON so large MCP payloads waste less context on whitespace.

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
        "sheet_type": "worksheet",
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

If `max_rows` is set, the payload is paged by row count and may include `total_rows`, `truncated`, `next_start_row`, and `next_start_cell`.

If `max_cols` is set, the payload is paged by column count and may include `total_cols`, `truncated`, `next_start_col`, and `next_column_start_cell`.

If the payload is truncated, `data.continuations` includes cursor tokens for the available continuation directions. Single-direction truncation also returns `next_cursor` for convenience.

If `compact=True`, cells without real validation rules omit the default `validation: {"has_validation": false}` stub.

If `values_only=True`, the payload returns `data.values` as a plain 2D array instead of `data.cells`, which is much smaller for large range reads that do not need per-cell metadata.

If a serialized response would exceed SheetForge's practical MCP payload limit, the tool returns `ResponseTooLargeError` with structured `error.hints` instead of depending on client-side truncation.

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
- `profile_workbook(filepath: str) -> str`
  Returns a workbook inventory with per-sheet summaries for visibility, freeze panes, autofilters, protection, tables, charts, print settings, and lightweight workbook-level counts. Grid-anchored worksheet charts include anchor, dimensions, and `occupied_range` for layout-aware follow-up steps.
- `describe_sheet_layout(filepath: str, sheet_name: str, sample_limit: int = 10, free_canvas_rows: int = 8, free_canvas_cols: int = 6, free_canvas_limit: int = 3) -> str`
  Returns a compact structural summary for one worksheet, including used range, freeze panes, autofilters, print settings, protection, merged ranges, tables, chart anchors, validation and conditional-format rule counts, custom column widths and row heights, plus a small `free_canvas_preview` for dashboard-safe follow-up placement.
- `audit_workbook(filepath: str, header_row: int = 1, sample_limit: int = 25) -> str`
  Audits workbook structure for high-signal issues that affect agent workflows. The response includes workbook summary counts, per-sheet assessments with recommended read tools, a sampled finding list grouped by severity/code, and deduplicated recommended actions.
- `plan_workbook_repairs(filepath: str, header_row: int = 1, sample_limit: int = 25) -> str`
  Turns workbook audit findings into prioritized next steps. The response includes ordered repair/inspection steps, suggested SheetForge tool calls, and a `quick_wins` list for issues that can be advanced entirely inside SheetForge.
- `apply_workbook_repairs(filepath: str, repair_types: Optional[List[str]] = None, sheet_names: Optional[List[str]] = None, header_row: int = 1, sample_limit: int = 25, dry_run: bool = True) -> str`
  Dry-runs or applies the safe workbook-repair subset currently supported inside SheetForge: broken named ranges, broken data validation rules, broken conditional formatting rules, and optional hidden-sheet reveals. The response includes planned/applied actions, audit summary before and after, and a structural diff.
- `diff_workbooks(before_filepath: str, after_filepath: str, sample_limit: int = 25, include_cell_changes: bool = True) -> str`
  Compares two workbook files and reports sheet/property changes, named-range changes, table/chart changes, validation and conditional-format changes, plus sampled cell-value diffs when `include_cell_changes=True`.
- `analyze_range_impact(filepath: str, sheet_name: str, range_ref: str) -> str`
  Inspects workbook structures that overlap a worksheet range before mutation. Reports intersections with native Excel tables, chart footprints, merged ranges, named ranges, worksheet data validations, conditional formatting rules, autofilters, print areas, formula cells inside the range, and downstream formulas or rule expressions elsewhere in the workbook that reference it directly or transitively, through named ranges, or through structured table references such as `Table1[Sales]`, plus a lightweight `risk_level`.
- `explain_formula_cell(filepath: str, sheet_name: str, cell: str, max_depth: int = 3) -> str`
  Explains a formula cell's direct references, upstream formula-chain cells, and downstream dependent formulas. Named ranges and structured references are resolved to concrete workbook ranges when possible.
- `detect_circular_dependencies(filepath: str, sample_limit: int = 25) -> str`
  Detects workbook formula cycles, including self-references and multi-cell dependency loops. Named ranges and structured references are resolved through the workbook graph before cycle analysis, and the response returns compact sampled cycle groups.
- `inspect_named_range(filepath: str, name: str, scope_sheet: Optional[str] = None) -> str`
  Inspects a defined name, including scope, destinations, hidden state, and whether it points at missing sheets or broken references.
- `list_named_ranges(filepath: str) -> str`
  Returns workbook defined names, their values, and any sheet/range destinations.
- `delete_named_range(filepath: str, name: str, scope_sheet: Optional[str] = None, dry_run: bool = False) -> str`
  Deletes a workbook-level or sheet-scoped named range. Use `dry_run=True` to preview which definition would be removed before committing.
- `list_all_sheets(filepath: str) -> str`
  Returns one entry per sheet, including `sheet_type`. Worksheets include `rows`, `columns`, `column_range`, and `is_empty`; chart sheets are reported with `sheet_type="chartsheet"` and zero grid dimensions.
- `list_tables(filepath: str, sheet_name: Optional[str] = None) -> str`
  Returns native Excel tables across the workbook or for one worksheet, including `sheet_name`, `table_name`, `range`, `style`, `headers`, row counts, and style flags.
- `suggest_read_strategy(filepath: str, goal: Optional[str] = None, sheet_name: Optional[str] = None, table_name: Optional[str] = None, header_row: int = 1, sample_rows: int = 25) -> str`
  Recommends the best next SheetForge read tool for the requested workbook target. The response identifies whether the target looks like a native Excel table, a worksheet-shaped dataset, a layout-heavy dashboard, or a chart sheet, then returns `recommended_tool`, `suggested_args`, and fallback `alternatives`.
- `describe_dataset(filepath: str, sheet_name: Optional[str] = None, table_name: Optional[str] = None, header_row: int = 1, sample_rows: int = 25) -> str`
  Returns a lightweight orientation summary for a worksheet or native Excel table, including `headers`, sample rows, inferred `schema`, header-quality signals, key-candidate guesses, structural observations, and a recommended follow-up read path. If the selected sheet is a chart sheet, the response explains that and points at chart-oriented tools instead of failing.
- `query_table(filepath: str, sheet_name: Optional[str] = None, table_name: Optional[str] = None, header_row: int = 1, select: Optional[List[str]] = None, filters: Optional[List[Dict[str, Any]]] = None, sort_by: Optional[str] = None, sort_desc: bool = False, limit: Optional[int] = None, row_mode: str = "arrays", infer_schema: bool = False) -> str`
  Filters, projects, sorts, and limits worksheet-shaped data or native Excel tables. `filters` accept operators such as `eq`, `neq`, `gt`, `gte`, `lt`, `lte`, `contains`, `starts_with`, `ends_with`, `in`, `not_in`, `is_blank`, and `not_blank`.
- `aggregate_table(filepath: str, sheet_name: Optional[str] = None, table_name: Optional[str] = None, header_row: int = 1, filters: Optional[List[Dict[str, Any]]] = None, group_by: Optional[List[str]] = None, metrics: Optional[List[Dict[str, Any]]] = None, sort_by: Optional[str] = None, sort_desc: bool = False, limit: Optional[int] = None, row_mode: str = "arrays", infer_schema: bool = False) -> str`
  Computes grouped summaries over worksheet-shaped data or native Excel tables. `metrics` support `count`, `count_non_null`, `count_distinct`, `sum`, `avg`, `min`, and `max`, each with an optional `as` alias.
- `bulk_aggregate_workbooks(filepaths: List[str], sheet_name: Optional[str] = None, table_name: Optional[str] = None, header_row: int = 1, filters: Optional[List[Dict[str, Any]]] = None, group_by: Optional[List[str]] = None, metrics: Optional[List[Dict[str, Any]]] = None, sort_by: Optional[str] = None, sort_desc: bool = False, limit: Optional[int] = None, schema_mode: str = "strict", source_sample_limit: int = 10, row_mode: str = "arrays", infer_schema: bool = False) -> str`
  Aggregates comparable worksheet or native-table data across multiple workbook files. `schema_mode="strict"` requires identical column sets, `intersect` requires every referenced column in every workbook, and `union` treats missing referenced columns as blank values. The response includes compact per-workbook sample metadata under `source_workbooks`.
- `bulk_filter_workbooks(filepaths: List[str], sheet_name: Optional[str] = None, table_name: Optional[str] = None, header_row: int = 1, select: Optional[List[str]] = None, filters: Optional[List[Dict[str, Any]]] = None, sort_by: Optional[str] = None, sort_desc: bool = False, limit: Optional[int] = None, schema_mode: str = "strict", source_sample_limit: int = 10, include_source_columns: bool = True, row_mode: str = "arrays", infer_schema: bool = False) -> str`
  Filters comparable worksheet or native-table data across multiple workbook files. By default, the output prepends `_source_file`, `_source_sheet`, and `_source_table` so matching rows keep their provenance. `schema_mode` follows the same `strict` / `intersect` / `union` contract as `bulk_aggregate_workbooks`.
- `read_excel_table(filepath: str, table_name: str, sheet_name: Optional[str] = None, start_row: int = 1, start_col: Optional[str] = None, end_col: Optional[str] = None, max_rows: Optional[int] = None, compact: bool = False, include_headers: bool = True, row_mode: str = "arrays", infer_schema: bool = False) -> str`
  Reads rows from a native Excel table by `table_name`, preserving the table's exact range instead of inferring worksheet bounds. Use `start_row` plus `max_rows` to paginate within table data rows, `start_col` / `end_col` to narrow the table width, and `include_headers=False` for follow-up pages after the first. Requested columns must fall inside the table range. When more rows remain, the response includes `next_start_row` for the next page.

## Read, Search, And Write Tools

- `write_data_to_excel(filepath: str, sheet_name: str, data: List[List], start_cell: str = "A1", dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Writes tabular data starting at the given cell. Missing target sheets are created automatically. Returns compact summaries by default on committed writes, and detailed `changes` during previews unless explicitly disabled.
- `read_data_from_excel(filepath: str, sheet_name: str, start_cell: str = "A1", end_cell: Optional[str] = None, max_rows: Optional[int] = None, max_cols: Optional[int] = None, cursor: Optional[str] = None, preview_only: bool = False, compact: bool = False, values_only: bool = False) -> str`
  Returns cell range data with row, column, address, value, and validation metadata under the shared envelope, or a plain 2D `values` matrix when `values_only=True`. Use `max_rows` to paginate tall rectangular ranges and `max_cols` to paginate wide ranges without manually recomputing the full rectangle. Follow-up pages can either use the explicit `next_start_row` / `next_start_cell` and `next_start_col` / `next_column_start_cell` fields or pass `cursor=...` from `continuations.down` / `continuations.right`.
- `read_excel_as_table(filepath: str, sheet_name: str, header_row: int = 1, start_row: Optional[int] = None, start_col: str = "A", end_col: Optional[str] = None, max_rows: Optional[int] = None, compact: bool = False, include_headers: bool = True, row_mode: str = "arrays", infer_schema: bool = False) -> str`
  Returns `headers`, `rows`, `total_rows`, `truncated`, and `sheet_name`. Use `start_row` plus `max_rows` to paginate into deeper worksheet sections without reading from the top first, `start_col` / `end_col` to limit the width of wide sheets, and `include_headers=False` on follow-up pages after the first. When more rows remain, the response includes `next_start_row` for the next page.
- `quick_read(filepath: str, sheet_name: Optional[str] = None, header_row: int = 1, start_row: Optional[int] = None, start_col: str = "A", end_col: Optional[str] = None, max_rows: Optional[int] = None, include_headers: bool = True, row_mode: str = "arrays", infer_schema: bool = False) -> str`
  Returns a compact table from the requested sheet, or auto-selects the first workbook sheet when `sheet_name` is omitted. Use `start_row` for large-sheet pagination, `start_col` / `end_col` for narrower column slices, and `include_headers=False` for follow-up pages. When more rows remain, the response includes `next_start_row` for the next page.
- `row_mode`
  Use `row_mode="arrays"` for the current `headers + rows` shape, or `row_mode="objects"` to receive `records` keyed by normalized field names such as `first_name`, `column_2`, or ASCII-safe transliterations like `nayttokerrat`.
- `infer_schema`
  When `infer_schema=True`, the response includes `schema` entries with `field`, `header`, `type`, and `nullable` hints inferred from the returned rows.
- `search_in_sheet(filepath: str, sheet_name: str, query: Any, exact: bool = True, max_results: int = 50) -> str`
  Returns exact or partial value matches across the worksheet.
- `append_table_rows(filepath: str, sheet_name: str, rows: List[Dict[str, Any]], header_row: int = 1, dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Appends header-aware rows using dictionary keys that match worksheet headers. Returns `changed_cells` always, and detailed `changes` only when requested or during previews.
- `upsert_excel_table_rows(filepath: str, table_name: str, key_column: str, rows: List[Dict[str, Any]], sheet_name: Optional[str] = None, dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Updates matching rows inside a native Excel table and appends missing keys in one call. Missing rows expand the table's `ref` automatically, and the tool refuses to grow the table into already occupied cells below it. Tables with an enabled totals row are update-only for now; append attempts are rejected.
- `update_rows_by_key(filepath: str, sheet_name: str, key_column: str, updates: List[Dict[str, Any]], header_row: int = 1, dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Updates existing rows by matching a named key column, reports unmatched keys, and returns compact summaries by default.

## Formula And Validation Tools

- `apply_formula(filepath: str, sheet_name: str, cell: str, formula: str) -> str`
  Writes a validated Excel formula into the target cell.
- `validate_formula_syntax(filepath: str, sheet_name: str, cell: str, formula: str) -> str`
  Validates formula syntax without changing the workbook.
- `inspect_formula(formula: str) -> str`
  Inspects a formula string without workbook context. Returns function inventory, reference token classification, literal-token counts, and flags for volatile or risky functions such as `INDIRECT`, `WEBSERVICE`, or `RTD`.
- `validate_excel_range(filepath: str, sheet_name: str, start_cell: str, end_cell: Optional[str] = None) -> str`
  Validates a cell or range reference against the worksheet.
- `get_data_validation_info(filepath: str, sheet_name: str) -> str`
  Returns JSON for data-validation rules defined in the worksheet. Chart sheets are rejected with a clear worksheet-only error.
- `inspect_data_validation_rules(filepath: str, sheet_name: str, broken_only: bool = False) -> str`
  Returns worksheet data validation rules with stable `rule_index` values plus `broken_reference` flags so agents can inspect and target specific rules safely.
- `remove_data_validation_rules(filepath: str, sheet_name: str, rule_indexes: Optional[List[int]] = None, broken_only: bool = False, dry_run: bool = False) -> str`
  Removes worksheet data validation rules by explicit index or removes all broken rules when `broken_only=True`. Use `dry_run=True` to preview the exact removals first.
- `inspect_conditional_format_rules(filepath: str, sheet_name: str, broken_only: bool = False) -> str`
  Returns worksheet conditional formatting rules with stable `rule_index` values plus `broken_reference` flags for broken workbook references and missing-sheet formulas.
- `remove_conditional_format_rules(filepath: str, sheet_name: str, rule_indexes: Optional[List[int]] = None, broken_only: bool = False, dry_run: bool = False) -> str`
  Removes worksheet conditional formatting rules by explicit index or removes all broken rules when `broken_only=True`. Use `dry_run=True` to preview the exact removals first.

## Formatting And Layout Tools

- `format_range(filepath: str, sheet_name: str, start_cell: str, end_cell: Optional[str] = None, bold: bool = False, italic: bool = False, underline: bool = False, font_size: Optional[int] = None, font_color: Optional[str] = None, bg_color: Optional[str] = None, border_style: Optional[str] = None, border_color: Optional[str] = None, number_format: Optional[str] = None, alignment: Optional[str] = None, wrap_text: bool = False, merge_cells: bool = False, protection: Optional[Dict[str, Any]] = None, conditional_format: Optional[Dict[str, Any]] = None, dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Applies formatting options to a cell or range and supports preview mode. Returns compact summaries by default on committed writes.
- `format_ranges(filepath: str, sheet_name: str, ranges: List[Dict[str, Any]], dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Applies formatting to multiple ranges in one workbook pass. Each range object uses the same option keys as `format_range`, such as `start_cell`, `end_cell`, `bold`, `font_size`, `bg_color`, or `conditional_format`. Invalid operations are reported under `errors` while successful ranges still apply.
- `read_range_formatting(filepath: str, sheet_name: str, range_ref: str, sample_limit: int = 10) -> str`
  Reads a compact formatting summary for a worksheet range. Instead of returning one style object per cell, the response groups the range into distinct `style_groups` with sample cells, reports overlapping merged ranges and conditional-format rules, and warns when the sampled style groups were truncated.
- `conditional_format` rule parameters
  You can pass rule parameters either under `conditional_format.params` or directly as top-level keys alongside `type`. Nested `params` win if both are present.
- `conditional_format` example
  `{"type": "data_bar", "params": {"start_type": "min", "end_type": "max", "color": "2E86C1"}}`
- `conditional_format` shorthand example
  `{"type": "data_bar", "start_type": "min", "end_type": "max", "color": "2E86C1"}`
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
- `merge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str, dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Merges the selected range and supports preview mode. Committed writes stay compact unless `include_changes=True`.
- `unmerge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str, dry_run: bool = False, include_changes: Optional[bool] = None) -> str`
  Unmerges the selected range and supports preview mode. Committed writes stay compact unless `include_changes=True`.
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
  Creates a native Excel table from an existing range. Requires a worksheet target; chart sheets cannot host native Excel tables.
- `list_charts(filepath: str, sheet_name: Optional[str] = None) -> str`
  Lists embedded charts across the workbook or for one worksheet, including chart type, anchor, titles, dimensions, occupied worksheet range, and series references.
- `find_free_canvas(filepath: str, sheet_name: str, width: Optional[float] = None, height: Optional[float] = None, min_rows: Optional[int] = None, min_cols: Optional[int] = None, limit: int = 5, origin_cell: str = "A1", search_rows: Optional[int] = None, search_columns: Optional[int] = None, padding_rows: int = 0, padding_columns: int = 0) -> str`
  Suggests non-overlapping worksheet slots for charts or dashboard blocks. Use `width` plus `height` for chart-sized slots in centimeters, or `min_rows` plus `min_cols` for cell-grid blocks. If you omit both sizing styles, SheetForge assumes the default chart size (`15 x 7.5 cm`).
- `create_chart(filepath: str, sheet_name: str, chart_type: str, target_cell: Optional[str] = None, data_range: Optional[str] = None, title: str = "", x_axis: str = "", y_axis: str = "", style: Optional[Dict[str, Any]] = None, series: Optional[List[Dict[str, Any]]] = None, categories_range: Optional[str] = None, width: Optional[float] = None, height: Optional[float] = None, placement: Optional[Dict[str, Any]] = None) -> str`
  Creates a chart anchored at `target_cell`, or auto-places it when you provide `placement` instead. Supported chart types are `line`, `bar`, `pie`, `scatter`, and `area`. Use either a contiguous `data_range` or explicit `series` definitions; non-scatter series use `values_range` plus an optional shared `categories_range`, while scatter series use `x_range` and `y_range`. `width` and `height` are measured in centimeters; defaults are `15` and `7.5`. `placement.relative_to="free_canvas"` scans for the first chart-sized gap that does not overlap current cells, tables, merges, or charts.
- `create_chart_from_series(filepath: str, sheet_name: str, chart_type: str, target_cell: Optional[str] = None, series: Optional[List[Dict[str, Any]]] = None, title: str = "", x_axis: str = "", y_axis: str = "", categories_range: Optional[str] = None, style: Optional[Dict[str, Any]] = None, width: Optional[float] = None, height: Optional[float] = None, placement: Optional[Dict[str, Any]] = None) -> str`
  Creates a chart from explicit series definitions, including non-contiguous ranges. Non-scatter series use `values_range` per series plus an optional shared `categories_range`; scatter series use `x_range` and `y_range`. `width` and `height` are measured in centimeters; defaults are `15` and `7.5`. You can omit `target_cell` if `placement` is provided, including the `placement.relative_to="free_canvas"` layout scan.
- `placement`
  Placement objects support `direction` (`right` or `below`), optional `padding_columns` / `padding_rows`, and `relative_to` values of `content`, `used_range`, `data_range`, `free_canvas`, `table:<name>`, or a worksheet range like `A1:C10`. `free_canvas` also accepts `origin_cell`, `search_rows`, and `search_columns`.
- `create_pivot_table(filepath: str, sheet_name: str, data_range: str, rows: List[str], values: List[str], columns: Optional[List[str]] = None, agg_func: str = "sum") -> str`
  Creates a pivot-style summary from the given data range. Supported aggregation functions: `sum`, `average`, `count`, `min`, `max`.

## Practical Usage Notes

- Use `list_all_sheets` before reading unfamiliar workbooks.
- Use `profile_workbook` when you need a one-call inventory of sheets, tables, charts, filters, freeze panes, and print/protection state before deciding what to mutate.
- Use `analyze_range_impact` before deleting, overwriting, or restructuring an important worksheet region; it gives a fast blast-radius summary without mutating the workbook, including downstream formula chains, validation rules, and conditional formatting expressions that point at the range from elsewhere, through named ranges, or through structured table references.
- Use workbook inventory tools to discover `sheet_type` first when the workbook may contain chart sheets; cell-grid tools reject chart sheets with clear worksheet-only errors.
- Use `read_excel_table` when the workbook already contains native Excel tables and you want exact table semantics instead of a sheet-wide read.
- Use `row_mode="objects"` when the agent benefits more from named fields than the smallest possible payload.
- Use `infer_schema=True` when downstream steps need lightweight type hints without doing a second pass over the rows.
- Use `read_excel_as_table` when the source data is tabular and headers matter.
- Use `read_data_from_excel` when you need cell addresses or validation metadata.
- Use `compact=True` on read tools when you want to minimize response size for agent workflows.
- Use `upsert_excel_table_rows` when the workbook already has a native Excel table and your workflow is naturally key-based.
- Use `format_ranges` instead of repeated `format_range` calls when you're styling a report or dashboard in several places at once.
- Use `create_chart` as the default chart authoring entry point. Pass `data_range` for contiguous source data, or `series` plus optional `categories_range` when the chart data lives in non-adjacent columns.
- Use chart `placement` when dashboard composition matters more than a hardcoded anchor. `relative_to="content"` avoids overlapping existing charts by accounting for both used cells and current chart footprints.
- Use `find_free_canvas` when you want several safe candidate slots before mutating the workbook, or `placement.relative_to="free_canvas"` when one chart should auto-land in the first gap.
- Use `create_chart_from_series` when you want the older explicit-series entry point or need to preserve existing automation prompts unchanged.
- Use top-level `width` and `height` for chart sizing. The legacy `style.width` and `style.height` keys still work as a compatibility fallback.
- Use `set_print_area` and `set_print_titles` when the workbook is meant for printing, export, or PDF generation.
- Use `get_worksheet_protection` before changing protection flags on an unfamiliar workbook.
- Use `autofit_columns` after writing or formatting tables when you want readable output without hand-tuning widths.
- Use `search_in_sheet` to find a value before mutating a workbook.
- Prefer `streamable-http` for long-running remote integrations.
- Prefer `stdio` for local desktop MCP clients.
