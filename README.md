# SheetForge MCP

SheetForge MCP is an Excel MCP server for `.xlsx` automation over the Model Context Protocol. It helps AI agents, MCP clients, and automation workflows read, write, search, format, chart, validate, and restructure Excel workbooks with Python and `openpyxl`, without launching Microsoft Excel or LibreOffice.

If you are looking for an Excel MCP server for spreadsheet automation, workbook inspection, Excel report generation, dashboard authoring, or `.xlsx` editing from AI tools, SheetForge MCP is built for that workflow.

Package name: `sheetforge-mcp`
CLI command: `sheetforge-mcp`
Published package release: `0.7.0`
Repository docs track the current main-branch tool surface, which currently exposes `75` MCP tools.

## Excel MCP Server Features

- workbook creation and metadata
- worksheet creation, renaming, copying, deletion, and visibility
- structured reads, compact table reads, declarative table queries, grouped aggregates, and cell search
- row, column, and range mutations
- formulas and validation checks
- formatting, freezes, autofilters, merges, and conditional formatting
- native Excel tables, charts, and pivot summaries
- `stdio`, `streamable-http`, and deprecated `sse` transports

## Common Use Cases

- AI agents that need safe, structured Excel workbook access through MCP
- spreadsheet automation workflows that read and update `.xlsx` reports
- Excel dashboard generation with formatting, tables, charts, freeze panes, and print setup
- workbook QA and inspection flows that need metadata, named ranges, tables, charts, and protection state
- data extraction from native Excel tables or worksheet-shaped datasets without hand-written `openpyxl` scripts

## Requirements

- Python `3.10+`
- `.xlsx` workbooks
- either `uvx` or a local package install

## Quick Start

Install and run directly from PyPI with `uvx`, or install the package locally in your Python environment.

### Stdio

Use `stdio` when the MCP client starts the server locally.

```bash
uvx sheetforge-mcp stdio
```

```json
{
  "mcpServers": {
    "excel": {
      "command": "uvx",
      "args": ["sheetforge-mcp", "stdio"]
    }
  }
}
```

### Streamable HTTP

Use `streamable-http` when you want a long-running local or remote server process.

```bash
EXCEL_FILES_PATH=/path/to/excel-files uvx sheetforge-mcp streamable-http
```

Default endpoint:

```text
http://127.0.0.1:8017/mcp
```

Example client config:

```json
{
  "mcpServers": {
    "excel": {
      "url": "http://127.0.0.1:8017/mcp"
    }
  }
}
```

### SSE

SSE is kept for compatibility, but new integrations should prefer `streamable-http`.

```bash
EXCEL_FILES_PATH=/path/to/excel-files uvx sheetforge-mcp sse
```

Default endpoint:

```text
http://127.0.0.1:8017/sse
```

## File Path Rules

- In `stdio` mode, `filepath` values must be absolute paths.
- In `streamable-http` and `sse` mode, relative paths are resolved under `EXCEL_FILES_PATH`.
- Absolute paths are accepted in every transport.
- In `streamable-http` and `sse` mode, the server creates `EXCEL_FILES_PATH` automatically if it does not exist.

## Environment Variables

| Variable | Default | Used by | Purpose |
| --- | --- | --- | --- |
| `FASTMCP_HOST` | `127.0.0.1` | HTTP and SSE | Bind address for the server process |
| `FASTMCP_PORT` | `8017` | HTTP and SSE | Port for the server process |
| `EXCEL_FILES_PATH` | `./excel_files` | HTTP and SSE | Base directory for relative workbook paths |

## Tooling Overview

The server currently registers 75 MCP tools across these groups:

- workbook overview: `create_workbook`, `create_worksheet`, `get_workbook_metadata`, `profile_workbook`, `describe_sheet_layout`, `audit_workbook`, `plan_workbook_repairs`, `apply_workbook_repairs`, `diff_workbooks`, `analyze_range_impact`, `explain_formula_cell`, `detect_circular_dependencies`, `inspect_named_range`, `list_named_ranges`, `delete_named_range`, `list_all_sheets`, `list_tables`
- data access: `suggest_read_strategy`, `describe_dataset`, `query_table`, `aggregate_table`, `bulk_aggregate_workbooks`, `bulk_filter_workbooks`, `union_tables`, `cross_workbook_lookup`, `quick_read`, `read_excel_table`, `read_data_from_excel`, `read_excel_as_table`, `search_in_sheet`, `write_data_to_excel`, `append_table_rows`, `append_excel_table_rows`, `upsert_excel_table_rows`, `update_rows_by_key`
- worksheet and range changes: `copy_worksheet`, `delete_worksheet`, `rename_worksheet`, `set_worksheet_visibility`, `get_worksheet_protection`, `set_worksheet_protection`, `copy_range`, `delete_range`, `insert_rows`, `insert_columns`, `delete_sheet_rows`, `delete_sheet_columns`
- formatting and layout: `format_range`, `format_ranges`, `read_range_formatting`, `freeze_panes`, `set_autofilter`, `set_print_area`, `set_print_titles`, `set_column_widths`, `autofit_columns`, `set_row_heights`, `merge_cells`, `unmerge_cells`, `get_merged_cells`
- formulas and validation: `apply_formula`, `validate_formula_syntax`, `inspect_formula`, `validate_excel_range`, `get_data_validation_info`, `inspect_data_validation_rules`, `remove_data_validation_rules`, `inspect_conditional_format_rules`, `remove_conditional_format_rules`
- analysis and structure: `create_table`, `list_charts`, `find_free_canvas`, `create_chart`, `create_chart_from_series`, `create_pivot_table`

For chart authoring, prefer `create_chart` as the primary entry point:

- use `data_range` for the simple contiguous-data path
- use explicit `series` plus optional `categories_range` for non-contiguous or hand-authored charts
- use top-level `width` and `height` to control chart size in centimeters; defaults are `15 x 7.5`
- use `placement` when you want SheetForge to position the chart relative to worksheet content, a source range, or a named table instead of guessing `target_cell` manually
- use `placement={"relative_to": "free_canvas"}` when a busy dashboard needs the first non-overlapping chart slot instead of a simple right/below placement rule
- keep `create_chart_from_series` for backward compatibility or existing prompts that already rely on it

The most agent-friendly read tools are:

- `suggest_read_strategy`: recommends the best next read tool for a workbook target, including whether SheetForge should treat it as a native Excel table, a clean worksheet dataset, a layout-heavy dashboard sheet, or a chart sheet
- `describe_dataset`: samples a worksheet or native Excel table and returns headers, schema hints, key-candidate guesses, structural signals, and a recommended follow-up read path
- `query_table`: filters, projects, sorts, and limits worksheet-shaped data or native Excel tables with a declarative JSON query instead of ad hoc cell loops
- `aggregate_table`: computes grouped metrics such as `count`, `sum`, `avg`, `min`, and `max` over worksheet-shaped data or native Excel tables
- `bulk_aggregate_workbooks`: computes the same grouped metrics across many workbook files in one call, with explicit schema handling via `strict`, `intersect`, or `union`
- `bulk_filter_workbooks`: returns matching rows across many workbook files with optional source provenance columns, so recurring cross-file QA and reporting checks no longer need one-tool-call-per-file loops
- `union_tables`: combines comparable worksheet or native-table rows across many workbook files, with optional deduplication keys and explicit schema handling for workbook collections that drift over time
- `cross_workbook_lookup`: enriches one workbook dataset from one or more lookup workbooks with left-join style matching, optional duplicate-match handling, and compact per-row provenance for matched lookup rows
- `profile_workbook`: one-call inventory for sheets, tables, charts, named ranges, and key layout/protection state, including chart `occupied_range` for grid-anchored worksheet charts
- `describe_sheet_layout`: worksheet-level structural summary for safe dashboard edits, including freeze panes, print settings, merges, chart anchors, table metadata, conditional-format and validation counts, custom row/column sizing, and a small free-canvas preview
- `audit_workbook`: workbook-level audit for high-signal problems such as broken `#REF!` formulas, error cells, hidden sheets, header-quality issues, layout-heavy sheets, and named ranges that reference missing sheets
- `plan_workbook_repairs`: converts workbook audit findings into prioritized next steps, including suggested SheetForge tool calls for inspection, safe dry runs, and repair workflows
- `apply_workbook_repairs`: dry-runs or applies the safe repair subset from those plans, including broken named ranges, broken validation rules, broken conditional formats, and optional hidden-sheet reveals
- `diff_workbooks`: compares two workbook files and reports structural changes plus sampled cell-value diffs, which is useful for before/after verification in agent workflows
- `analyze_range_impact`: preflight blast-radius check for a worksheet range, including overlaps with tables, chart footprints, merged cells, named ranges, data validations, conditional formats, autofilters, print areas, formula cells inside the range, and formulas or rule expressions elsewhere that depend on it directly or transitively, through named ranges, or through structured table references such as `Table1[Sales]`
- `explain_formula_cell`: resolves a formula cell's direct references, shows upstream formula-chain cells, returns a compact `formula_chain` summary with depth layers and sampled paths, and reports downstream dependents so agents can debug workbook logic without manual tracing
- `detect_circular_dependencies`: scans workbook formula graphs, including named-range-driven edges, and reports self-references plus multi-cell circular dependency groups before they surprise downstream automation
- `inspect_formula`: inspects a formula string without workbook context, listing functions, reference token types, volatile functions, and risky functions such as `INDIRECT`
- `inspect_named_range`: inspects one defined name, including its scope, destinations, and whether it points at missing sheets or broken references
- `quick_read`: single-call compact table read that auto-selects the first sheet when needed, now with `start_row` pagination and `start_col` / `end_col` column windowing for large sheets
- `read_excel_table`: read a native Excel table by `table_name` without guessing worksheet bounds, now with `start_row` pagination and optional `start_col` / `end_col` table column windowing
- `list_all_sheets`: quick workbook inventory with sheet sizes, emptiness flags, and `sheet_type` for worksheets versus chart sheets
- `read_excel_as_table`: compact `headers + rows` output for structured datasets, with `compact=True` for the smallest payload, `start_row` for page-like reads, and `start_col` / `end_col` for narrower column slices
- `read_data_from_excel`: cell-address-aware range reader that supports `max_rows` and `max_cols` windowing for large non-tabular ranges, `values_only=True` for smaller 2D payloads, and cursor-based continuations for multi-step 2D traversal
- `read_range_formatting`: compact formatting readback for a worksheet range, grouped by distinct style signatures instead of noisy per-cell dumps, with merged-range and conditional-format overlap summaries
- `search_in_sheet`: exact or partial value search across a worksheet

Workbook inventory tools such as `list_all_sheets`, `profile_workbook`, and `list_charts` surface both worksheets and chart sheets. Grid-oriented tools such as `quick_read`, `read_excel_table`, `create_table`, formatting, formulas, and validation require a real worksheet and return a clear chartsheet error if you target the wrong sheet type.

The most agent-friendly write helpers for structured data are:

- `upsert_excel_table_rows`: update matching rows in a native Excel table and append missing keys in one call
  Note: totals-row tables are update-only for now; append attempts are rejected rather than shifting unrelated rows.
- `append_excel_table_rows`: append rows to a native Excel table when you want the table `ref` to grow with the new records
- `append_table_rows`: append header-aware rows to worksheet-shaped data when you do not have a native Excel table
- `update_rows_by_key`: update worksheet-shaped data by a named key column without appending missing keys

For the compact table readers (`quick_read`, `read_excel_as_table`, `read_excel_table`):

- `row_mode="arrays"` keeps the smallest `headers + rows` shape
- `row_mode="objects"` returns `records` keyed by normalized field names such as `first_name`
- normalized field names are ASCII-safe transliterations, so headers like `Näyttökerrat` become `nayttokerrat`
- `infer_schema=True` adds lightweight `schema` hints inferred from the returned rows
- `start_col` / `end_col` let you slice wide worksheets or native Excel tables down to just the columns you need before pagination or schema inference
- truncated pages now include `next_start_row`, which you can pass back to the same tool for the next page
- non-tabular range reads can also return `continuations.down` and `continuations.right` cursor tokens so agents can continue large 2D windows without recomputing coordinates
- `suggest_read_strategy` helps agents choose between table-aware, worksheet-aware, range-aware, and workbook-orientation reads before they spend context on the wrong path
- `describe_dataset` provides a lighter-weight dataset summary than a full read, including sample rows, header quality, key candidates, and recommended next tool
- `query_table` is the lightest way to pull just the matching rows and columns you need from a worksheet dataset or native Excel table
- `query_table` and `bulk_filter_workbooks` accept `ne` as a shorthand for `neq`, and membership filters can use either `values` or the shorter `value` list form
- `aggregate_table` lets agents compute grouped summaries directly in SheetForge instead of over-reading the full dataset into context first
- `bulk_aggregate_workbooks` extends that pattern across many workbook files when a recurring reporting workflow would otherwise need ad hoc Python or repeated per-file tool calls
- aggregate metrics accept both the canonical `{"op": "sum", "field": "Sales", "as": "total_sales"}` shape and the more guessable alias form `{"agg": "sum", "column": "Sales", "as": "total_sales"}`
- `bulk_filter_workbooks` does the same for row-level inspection, while keeping workbook provenance visible by default
- `union_tables` is the fastest way to normalize many comparable workbook datasets into one combined tabular payload before downstream QA, export, or further aggregation
- `cross_workbook_lookup` is the fastest way to enrich one workbook from another without writing an ad hoc merge script, especially for master-data lookups, status enrichment, and cross-file QA workflows
- `append_excel_table_rows` is the right append path for native Excel tables when you do not need key-based upsert behavior
- `append_table_rows` now refuses to write directly under an adjacent native Excel table and points you at `append_excel_table_rows` instead of silently leaving the table range stale
- formatting color inputs accept `RRGGBB`, `#RRGGBB`, `AARRGGBB`, or `#AARRGGBB`, so prompts do not need to strip CSS-style `#` prefixes first
- `audit_workbook` is the fastest workbook-wide preflight when you need to know whether a spreadsheet is safe and predictable enough for autonomous editing
- `plan_workbook_repairs` is the fastest way to turn those audit findings into an actual action queue instead of manually deciding the next tool call for every problem
- `apply_workbook_repairs` lets agents preview or apply the safe subset of those repairs without having to orchestrate each broken workbook artifact manually
- `diff_workbooks` is the quickest before/after QA pass when an agent has touched workbook structure and wants proof of what actually changed

## Recommended Agent Workflows

1. Unfamiliar workbook -> safe mutation -> verification
   Start with `list_all_sheets` or `profile_workbook`, inspect layout-heavy tabs with `describe_sheet_layout`, run `analyze_range_impact` before overwriting an important range, then confirm the before/after result with `diff_workbooks`.
2. Workbook repair loop
   Use `audit_workbook` to find high-signal issues, `plan_workbook_repairs` to turn them into an action queue, `apply_workbook_repairs(..., dry_run=True)` to preview the safe subset, then rerun `audit_workbook` after applying repairs to confirm the workbook is back to a low-risk state.
3. Multi-workbook reporting
   Use `bulk_aggregate_workbooks`, `bulk_filter_workbooks`, `union_tables`, or `cross_workbook_lookup` to build the reporting dataset first, then write the summarized rows into a fresh workbook tab and finish the presentation layer with `format_ranges`, `find_free_canvas`, `create_chart`, and `autofit_columns`.

See [TOOLS.md](TOOLS.md) for the full reference.
Release notes live in [CHANGELOG.md](CHANGELOG.md).

## Response Format

Every tool now returns a JSON envelope with a consistent top-level shape:

```json
{
  "ok": true,
  "operation": "read_excel_as_table",
  "message": "read_excel_as_table completed",
  "data": {}
}
```

Error responses follow the same contract:

```json
{
  "ok": false,
  "operation": "write_data_to_excel",
  "error": {
    "type": "DataError",
    "message": "No data provided to write"
  }
}
```

For destructive tools that support preview mode, the envelope may also include `dry_run` and `changes`.
Committed write operations now default to compact summaries; pass `include_changes=True` when you want per-cell, per-range, or per-operation detail.

## Development

Install dependencies:

```bash
uv sync --extra dev
```

Run tests:

```bash
uv run --extra dev pytest -q
```

Run lint checks:

```bash
uv run --extra dev ruff check src tests
```

Run the package locally:

```bash
uv run sheetforge-mcp stdio
```

Build distributions locally:

```bash
uv build
```

## Release Flow

- Update `pyproject.toml`, `manifest.json`, and the tracked `.mcpb` bundle together for each release.
- Keep the tracked bundle filename in sync with the package version, for example `sheetforge-mcp-<version>.mcpb`.
- GitHub releases run a build verification workflow only.
- PyPI publishing is a separate manual workflow, so releases do not create a failing deployment before Trusted Publisher is configured for the package.

## Repository Layout

- `src/excel_mcp/server.py`: MCP server, transport setup, and tool registration
- `src/excel_mcp/workbook.py`: workbook lifecycle helpers and workbook metadata
- `src/excel_mcp/data.py`: read, write, table, and search helpers
- `src/excel_mcp/sheet.py`: worksheet and range mutations
- `tests/`: regression tests covering data, layout, charts, pivots, formatting, tables, and resource safety
- `manifest.json`: packaged MCP bundle metadata
- `docs/index.html`: static project landing page

## Why SheetForge MCP

- Excel-first MCP surface: the toolset is focused on real `.xlsx` workbook operations, not generic file I/O
- agent-friendly responses: consistent JSON envelopes, compact writes, and `dry_run` previews reduce context waste
- workbook introspection: `profile_workbook`, `list_all_sheets`, `list_tables`, and `list_charts` make unfamiliar spreadsheets easier to navigate
- safer edits: `analyze_range_impact` gives agents a read-only preflight before overwriting, deleting, or restructuring an important range, including downstream formula chains plus validation-rule and conditional-format references elsewhere in the workbook even when formulas point at the range through named ranges or structured table references
- layout planning: `find_free_canvas` suggests safe empty slots for charts or dashboard blocks before you place them, defaulting to the standard chart footprint when you omit explicit sizing
- practical Excel output: formatting, print setup, worksheet protection, table upserts, chart authoring, and autofit helpers cover real reporting workflows
- Python ecosystem fit: built on `openpyxl`, packaged for `uvx`, and easy to run locally over `stdio` or remotely over HTTP

## Notes For Integrators

- `stdio` mode is careful not to write non-protocol text to `stdout`.
- All tools return structured JSON envelopes, which makes client-side parsing predictable.
- Tool responses now use compact JSON serialization to reduce MCP payload size while keeping the same envelope shape.
- `read_data_from_excel(..., preview_only=True)` limits the response to the first 10 rows in the selected range and marks the payload as truncated when applicable.
- `read_data_from_excel(..., compact=True)` omits default validation stubs for cells that do not have validation rules.
- `read_data_from_excel(..., values_only=True)` returns a plain 2D `values` array for range reads that do not need per-cell addresses or validation metadata.
- `read_data_from_excel(..., max_rows=...)` paginates tall rectangular ranges and returns `next_start_row` plus `next_start_cell` when more rows remain.
- `read_data_from_excel(..., max_cols=...)` paginates wide rectangular ranges and returns `next_start_col` plus `next_column_start_cell` when more columns remain.
- `read_data_from_excel(..., cursor=...)` resumes from a continuation token so agents can keep paging without recomputing the next window manually; 2D windows expose directional continuations under `continuations.down` and `continuations.right`
- `read_excel_as_table(..., compact=True)` returns only `headers` and `rows` unless truncation metadata is needed.
- `quick_read(..., start_row=...)` and `read_excel_as_table(..., start_row=...)` let agents paginate deep worksheets without first reading from the top.
- `quick_read(..., start_col=..., end_col=...)` and `read_excel_as_table(..., start_col=..., end_col=...)` let agents request only the relevant columns from wide worksheets instead of pulling every column into context.
- `read_excel_table(..., start_col=..., end_col=...)` now supports the same narrower column slices for native Excel tables, as long as the requested columns fall inside the table range.
- `quick_read(..., include_headers=False)`, `read_excel_as_table(..., include_headers=False)`, and `read_excel_table(..., include_headers=False)` let follow-up pages omit repeated header payload once the first page already established the schema.
- `read_excel_table(..., start_row=...)` now supports deeper pagination into native Excel tables instead of always reading from the top.
- Truncated tabular reads now return `next_start_row` so agents can continue paging without recalculating offsets.
- Oversized read responses now fail early with `ResponseTooLargeError` plus structured `hints`, so agents can retry with smaller ranges or pagination before the client truncates the payload.
- `quick_read`, `read_excel_as_table`, and `read_excel_table` can now return `records` plus inferred `schema` hints when you opt into `row_mode="objects"` and `infer_schema=True`.
- Read tools do not recalculate Excel formulas; formula cells surface as formula text such as `=B2*C2`, and inferred schema labels formula-backed columns as `formula` so agents do not mistake them for fresh numeric values.
- `profile_workbook` provides a single-call workbook inventory with sheet-level table, chart, protection, print, and filter metadata for faster agent orientation, and now includes chart `occupied_range` alongside anchors and dimensions for grid-anchored worksheet charts.
- Core mutation tools now default to compact responses on committed writes, including data writes, formatting, worksheet layout helpers, and merge/unmerge helpers. Use `include_changes=True` for detailed diffs.
- `format_ranges` batches multiple formatting operations into one workbook pass, and now reports per-range `errors` without discarding successful ranges in the same batch.
- `autofit_columns` estimates practical column widths from the current cell contents, with optional column filters and min/max bounds.
- `list_charts` now reports chart `width` and `height` in centimeters in addition to anchor, type, and series metadata.
- `get_worksheet_protection` and `set_worksheet_protection` add a safe worksheet-level wrapper around Excel protection flags.
- `set_print_area` and `set_print_titles` make report/export setup scriptable without dropping into raw openpyxl workbook internals.
- `list_tables` now returns lightweight schema metadata such as headers, row counts, and stripe settings in addition to table names and ranges.
- `upsert_excel_table_rows` expands native Excel table ranges automatically when it appends missing keys, refuses to grow a table into already occupied cells, and rejects append attempts when the target table has an enabled totals row.
- Core mutation tools support `dry_run=True` so clients can preview changes before saving a workbook.

## License

MIT. See [LICENSE](LICENSE).
