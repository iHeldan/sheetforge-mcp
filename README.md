# SheetForge MCP

SheetForge MCP exposes `.xlsx` workbook operations over the Model Context Protocol. It uses `openpyxl` under the hood, so MCP clients can inspect and modify Excel files without launching Microsoft Excel or LibreOffice.

Package name: `sheetforge-mcp`
CLI command: `sheetforge-mcp`
Current release: `0.4.1`

## What This Project Covers

- workbook creation and metadata
- worksheet creation, renaming, copying, deletion, and visibility
- structured reads, compact table reads, and cell search
- row, column, and range mutations
- formulas and validation checks
- formatting, freezes, autofilters, merges, and conditional formatting
- native Excel tables, charts, and pivot summaries
- `stdio`, `streamable-http`, and deprecated `sse` transports

## Requirements

- Python `3.10+`
- `.xlsx` workbooks
- either `uvx` or a local package install

## Quick Start

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
- keep `create_chart_from_series` for backward compatibility or existing prompts that already rely on it

The most agent-friendly read tools are:

- `profile_workbook`: one-call inventory for sheets, tables, charts, named ranges, and key layout/protection state
- `quick_read`: single-call compact table read that auto-selects the first sheet when needed
- `read_excel_table`: read a native Excel table by `table_name` without guessing worksheet bounds
- `list_all_sheets`: quick workbook inventory with sheet sizes and emptiness flags
- `read_excel_as_table`: compact `headers + rows` output for structured datasets, with `compact=True` for the smallest payload
- `search_in_sheet`: exact or partial value search across a worksheet

The most agent-friendly write helpers for structured data are:

- `upsert_excel_table_rows`: update matching rows in a native Excel table and append missing keys in one call
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

## Notes For Integrators

- `stdio` mode is careful not to write non-protocol text to `stdout`.
- All tools return structured JSON envelopes, which makes client-side parsing predictable.
- `read_data_from_excel(..., preview_only=True)` limits the response to the first 10 rows in the selected range and marks the payload as truncated when applicable.
- `read_data_from_excel(..., compact=True)` omits default validation stubs for cells that do not have validation rules.
- `read_excel_as_table(..., compact=True)` returns only `headers` and `rows` unless truncation metadata is needed.
- `quick_read`, `read_excel_as_table`, and `read_excel_table` can now return `records` plus inferred `schema` hints when you opt into `row_mode="objects"` and `infer_schema=True`.
- `profile_workbook` provides a single-call workbook inventory with sheet-level table, chart, protection, print, and filter metadata for faster agent orientation.
- Core mutation tools now default to compact responses on committed writes, including data writes, formatting, worksheet layout helpers, and merge/unmerge helpers. Use `include_changes=True` for detailed diffs.
- `format_ranges` batches multiple formatting operations into one workbook pass, and now reports per-range `errors` without discarding successful ranges in the same batch.
- `autofit_columns` estimates practical column widths from the current cell contents, with optional column filters and min/max bounds.
- `list_charts` now reports chart `width` and `height` in centimeters in addition to anchor, type, and series metadata.
- `get_worksheet_protection` and `set_worksheet_protection` add a safe worksheet-level wrapper around Excel protection flags.
- `set_print_area` and `set_print_titles` make report/export setup scriptable without dropping into raw openpyxl workbook internals.
- `list_tables` now returns lightweight schema metadata such as headers, row counts, and stripe settings in addition to table names and ranges.
- `upsert_excel_table_rows` expands native Excel table ranges automatically when it appends missing keys, and refuses to grow a table into already occupied cells.
- Core mutation tools support `dry_run=True` so clients can preview changes before saving a workbook.

## License

MIT. See [LICENSE](LICENSE).
