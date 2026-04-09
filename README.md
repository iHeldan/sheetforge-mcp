# SheetForge MCP

SheetForge MCP exposes `.xlsx` workbook operations over the Model Context Protocol. It uses `openpyxl` under the hood, so MCP clients can inspect and modify Excel files without launching Microsoft Excel or LibreOffice.

Package name: `sheetforge-mcp`
CLI command: `sheetforge-mcp`

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

The server currently registers 45 MCP tools across these groups:

- workbook overview: `create_workbook`, `create_worksheet`, `get_workbook_metadata`, `list_named_ranges`, `list_all_sheets`, `list_tables`
- data access: `quick_read`, `read_data_from_excel`, `read_excel_as_table`, `search_in_sheet`, `write_data_to_excel`, `append_table_rows`, `update_rows_by_key`
- worksheet and range changes: `copy_worksheet`, `delete_worksheet`, `rename_worksheet`, `set_worksheet_visibility`, `get_worksheet_protection`, `set_worksheet_protection`, `copy_range`, `delete_range`, `insert_rows`, `insert_columns`, `delete_sheet_rows`, `delete_sheet_columns`
- formatting and layout: `format_range`, `format_ranges`, `freeze_panes`, `set_autofilter`, `set_print_area`, `set_print_titles`, `set_column_widths`, `autofit_columns`, `set_row_heights`, `merge_cells`, `unmerge_cells`, `get_merged_cells`
- formulas and validation: `apply_formula`, `validate_formula_syntax`, `validate_excel_range`, `get_data_validation_info`
- analysis and structure: `create_table`, `list_charts`, `create_chart`, `create_pivot_table`

The most agent-friendly read tools are:

- `quick_read`: single-call compact table read that auto-selects the first sheet when needed
- `list_all_sheets`: quick workbook inventory with sheet sizes and emptiness flags
- `read_excel_as_table`: compact `headers + rows` output for structured datasets, with `compact=True` for the smallest payload
- `search_in_sheet`: exact or partial value search across a worksheet

See [TOOLS.md](TOOLS.md) for the full reference.

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
Committed write operations now default to compact summaries; pass `include_changes=True` when you want per-cell or per-range detail.

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

## Release Flow

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
- `write_data_to_excel`, `append_table_rows`, `update_rows_by_key`, and `format_range` now default to compact responses on committed writes. Use `include_changes=True` for detailed diffs.
- `format_ranges` batches multiple formatting operations into one workbook pass, which is much more efficient for report polish than repeated single-range calls.
- `autofit_columns` estimates practical column widths from the current cell contents, with optional column filters and min/max bounds.
- `get_worksheet_protection` and `set_worksheet_protection` add a safe worksheet-level wrapper around Excel protection flags.
- `set_print_area` and `set_print_titles` make report/export setup scriptable without dropping into raw openpyxl workbook internals.
- `list_tables` now returns lightweight schema metadata such as headers, row counts, and stripe settings in addition to table names and ranges.
- Core mutation tools support `dry_run=True` so clients can preview changes before saving a workbook.

## License

MIT. See [LICENSE](LICENSE).
