# SheetForge MCP

SheetForge MCP is an Excel MCP server for `.xlsx` automation over the Model Context Protocol. It lets AI agents, MCP clients, and automation workflows read, write, search, format, chart, and restructure Excel workbooks with Python and `openpyxl`, without launching Microsoft Excel or LibreOffice.

If you are looking for an MCP server for Excel, spreadsheet automation, workbook inspection, Excel report generation, or `.xlsx` editing from AI tools, SheetForge MCP is built for that workflow.

Package name: `sheetforge-mcp`
CLI command: `sheetforge-mcp`
Current release: `0.4.2`

## Excel MCP Server Features

- workbook creation and metadata
- worksheet creation, renaming, copying, deletion, and visibility
- structured reads, compact table reads, and cell search
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

The server currently registers 49 MCP tools across these groups:

- workbook overview: `create_workbook`, `create_worksheet`, `get_workbook_metadata`, `profile_workbook`, `list_named_ranges`, `list_all_sheets`, `list_tables`
- data access: `quick_read`, `read_excel_table`, `read_data_from_excel`, `read_excel_as_table`, `search_in_sheet`, `write_data_to_excel`, `append_table_rows`, `upsert_excel_table_rows`, `update_rows_by_key`
- worksheet and range changes: `copy_worksheet`, `delete_worksheet`, `rename_worksheet`, `set_worksheet_visibility`, `get_worksheet_protection`, `set_worksheet_protection`, `copy_range`, `delete_range`, `insert_rows`, `insert_columns`, `delete_sheet_rows`, `delete_sheet_columns`
- formatting and layout: `format_range`, `format_ranges`, `freeze_panes`, `set_autofilter`, `set_print_area`, `set_print_titles`, `set_column_widths`, `autofit_columns`, `set_row_heights`, `merge_cells`, `unmerge_cells`, `get_merged_cells`
- formulas and validation: `apply_formula`, `validate_formula_syntax`, `validate_excel_range`, `get_data_validation_info`
- analysis and structure: `create_table`, `list_charts`, `create_chart`, `create_chart_from_series`, `create_pivot_table`

For chart authoring, prefer `create_chart` as the primary entry point:

- use `data_range` for the simple contiguous-data path
- use explicit `series` plus optional `categories_range` for non-contiguous or hand-authored charts
- use top-level `width` and `height` to control chart size in centimeters; defaults are `15 x 7.5`
- use `placement` when you want SheetForge to position the chart relative to worksheet content, a source range, or a named table instead of guessing `target_cell` manually
- keep `create_chart_from_series` for backward compatibility or existing prompts that already rely on it

The most agent-friendly read tools are:

- `profile_workbook`: one-call inventory for sheets, tables, charts, named ranges, and key layout/protection state, including chart `occupied_range` for grid-anchored worksheet charts
- `quick_read`: single-call compact table read that auto-selects the first sheet when needed, now with `start_row` pagination for large sheets
- `read_excel_table`: read a native Excel table by `table_name` without guessing worksheet bounds
- `list_all_sheets`: quick workbook inventory with sheet sizes, emptiness flags, and `sheet_type` for worksheets versus chart sheets
- `read_excel_as_table`: compact `headers + rows` output for structured datasets, with `compact=True` for the smallest payload and `start_row` for page-like reads
- `search_in_sheet`: exact or partial value search across a worksheet

Workbook inventory tools such as `list_all_sheets`, `profile_workbook`, and `list_charts` surface both worksheets and chart sheets. Grid-oriented tools such as `quick_read`, `read_excel_table`, `create_table`, formatting, formulas, and validation require a real worksheet and return a clear chartsheet error if you target the wrong sheet type.

The most agent-friendly write helpers for structured data are:

- `upsert_excel_table_rows`: update matching rows in a native Excel table and append missing keys in one call
  Note: totals-row tables are update-only for now; append attempts are rejected rather than shifting unrelated rows.
- `append_table_rows`: append header-aware rows to worksheet-shaped data when you do not have a native Excel table
- `update_rows_by_key`: update worksheet-shaped data by a named key column without appending missing keys

For the compact table readers (`quick_read`, `read_excel_as_table`, `read_excel_table`):

- `row_mode="arrays"` keeps the smallest `headers + rows` shape
- `row_mode="objects"` returns `records` keyed by normalized field names such as `first_name`
- normalized field names are ASCII-safe transliterations, so headers like `Näyttökerrat` become `nayttokerrat`
- `infer_schema=True` adds lightweight `schema` hints inferred from the returned rows

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
- practical Excel output: formatting, print setup, worksheet protection, table upserts, chart authoring, and autofit helpers cover real reporting workflows
- Python ecosystem fit: built on `openpyxl`, packaged for `uvx`, and easy to run locally over `stdio` or remotely over HTTP

## Notes For Integrators

- `stdio` mode is careful not to write non-protocol text to `stdout`.
- All tools return structured JSON envelopes, which makes client-side parsing predictable.
- Tool responses now use compact JSON serialization to reduce MCP payload size while keeping the same envelope shape.
- `read_data_from_excel(..., preview_only=True)` limits the response to the first 10 rows in the selected range and marks the payload as truncated when applicable.
- `read_data_from_excel(..., compact=True)` omits default validation stubs for cells that do not have validation rules.
- `read_data_from_excel(..., values_only=True)` returns a plain 2D `values` array for range reads that do not need per-cell addresses or validation metadata.
- `read_excel_as_table(..., compact=True)` returns only `headers` and `rows` unless truncation metadata is needed.
- `quick_read(..., start_row=...)` and `read_excel_as_table(..., start_row=...)` let agents paginate deep worksheets without first reading from the top.
- Oversized read responses now fail early with `ResponseTooLargeError` plus structured `hints`, so agents can retry with smaller ranges or pagination before the client truncates the payload.
- `quick_read`, `read_excel_as_table`, and `read_excel_table` can now return `records` plus inferred `schema` hints when you opt into `row_mode="objects"` and `infer_schema=True`.
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
