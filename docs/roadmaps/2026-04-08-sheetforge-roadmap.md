# SheetForge MCP Roadmap

Date: 2026-04-08
Status: Active

## Current Focus

The next milestone is about making SheetForge more reliable for MCP clients and much easier for agents to use safely.

### 1. Uniform JSON Responses

Goal:
Every MCP tool should return the same top-level response envelope instead of mixing plain strings, raw lists, and `str(dict)` payloads.

Why it matters:
- Clients can parse every tool result the same way.
- Errors become machine-readable.
- Future features like previews and audit trails can reuse one response contract.

Acceptance criteria:
- Every tool returns JSON.
- Success responses include `ok`, `operation`, `message`, and `data`.
- Error responses include `ok`, `operation`, and structured `error`.
- Existing manifest and tests stay in sync with the runtime tool list.

### 2. Dry Run and Change Previews

Goal:
Destructive tools should be able to describe intended changes without saving them.

Why it matters:
- Agents can inspect impact before mutating a workbook.
- Human users can confirm what will happen.
- Safer automation becomes possible without external diff tooling.

Acceptance criteria:
- Core write operations accept `dry_run`.
- Dry-run responses clearly state that no file was saved.
- Responses expose `changes` or `preview` metadata when available.
- Tests verify that dry-run paths do not persist workbook edits.

### 3. Table-Oriented Editing Tools

Goal:
Add higher-level table operations so clients can work with headers and keys instead of raw cell coordinates.

Planned tools:
- `append_table_rows`
- `update_rows_by_key`

Why it matters:
- Real workflows are usually row-oriented, not cell-oriented.
- Agents should not need to calculate the next empty row manually.
- Updating a row by primary key is safer than patching arbitrary coordinates.

Acceptance criteria:
- Header-aware append works with dictionaries keyed by column name.
- Row updates work by matching a named key column.
- Responses report inserted or changed cells.
- Dry-run mode is supported for both tools.

## Next Wave

### 4. Workbook Change Diffing
- Add a `preview_changes` or `diff_workbook` tool for before/after summaries.

### 5. Better Table Detection
- Add automatic table region detection and schema inference.

### 6. More Excel-Native Controls
- Named ranges
- Data validation authoring
- Freeze panes
- AutoFilter
- Column widths and row heights

### 7. Concurrency and Safety
- Workbook locking for simultaneous writers
- Stronger error boundaries
- Workbook health checks

### 8. Product Experience
- Better usage recipes in docs
- Client examples for common MCP hosts
- A clearer “best tools for this job” guide

## Definition of Success

This milestone is complete when:
- SheetForge returns predictable machine-readable results.
- Risky operations can be previewed before save.
- Common table edits no longer require low-level cell math.
