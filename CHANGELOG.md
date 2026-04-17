# Changelog

## Unreleased

### Added

- Added `next_start_row` to truncated tabular read responses so agents can continue pagination without recalculating offsets.

## 0.4.2 - 2026-04-11

This patch release improves workbook profiling coverage, safer native-table behavior, and package discoverability for SheetForge MCP.

### Added

- Added `profile_workbook` as a one-call workbook inventory for sheets, tables, charts, named ranges, and key layout/protection state.
- Added `upsert_excel_table_rows` for key-based updates plus append-missing behavior directly on native Excel tables.
- Added `start_row` pagination support to `quick_read` and `read_excel_as_table` for large worksheet reads that need to start below the top of the sheet.
- Added `values_only=True` support to `read_data_from_excel` for compact 2D range reads without per-cell metadata overhead.
- Added `include_headers=False` to tabular read tools so follow-up pages can omit repeated header payload.
- Added `start_row` pagination support to `read_excel_table` for deeper native-table reads.

### Quality

- Added regression coverage for workbook profiling and native-table upserts, including refusal to expand a table into occupied cells below it.
- Hardened `upsert_excel_table_rows` by rejecting append attempts on totals-row tables until row-shift semantics can be modeled safely.
- Hardened `profile_workbook` so workbooks with chart sheets no longer crash inventory reads.
- Refreshed package metadata, README messaging, and landing-page SEO copy to reflect the current Excel MCP surface for AI agents and automation workflows.
- Added ignore rules for local workspace notes such as `CONTEXT.md` and `LOCAL_*.md` so private planning files are less likely to be committed accidentally.
- Tightened tabular reads so selected-column windows no longer over-read trailing rows caused by unrelated data outside the requested column range.
- Switched JSON envelopes to compact serialization to reduce MCP payload size without changing the response schema.
- Added `ResponseTooLargeError` with structured hints so oversized read responses fail early with recovery guidance instead of relying on client-side truncation.

## 0.4.1 - 2026-04-09

This patch release tightens chart behavior for real workbook authoring and visual verification workflows.

### Fixes

- Fixed empty `x_axis` and `y_axis` inputs so Excel no longer renders visible `None` axis titles.
- Added top-level `width` and `height` chart parameters in centimeters to `create_chart` and `create_chart_from_series`.
- Kept `style.width` and `style.height` working as a backward-compatible sizing fallback for older prompts and automations.
- Extended `list_charts` to report the actual persisted chart `width` and `height` from drawing anchors, not just in-memory defaults.
- Expanded regression coverage to 177 passing tests.

## 0.4.0 - 2026-04-09

This release turns SheetForge from an early MCP workbook helper into a production-ready Excel automation tool for agent workflows.

### Highlights

- Added compact write responses across mutation tools, with `include_changes=True` available for detailed diffs when needed.
- Added `quick_read`, `read_excel_as_table`, and `read_excel_table` support for `row_mode="objects"` and `infer_schema=True`.
- Added ASCII-safe transliteration for object-mode field names, so headers like `Näyttökerrat` become `nayttokerrat`.
- Added `list_tables`, `set_worksheet_visibility`, `set_column_widths`, `set_row_heights`, `autofit_columns`, worksheet protection helpers, and print setup helpers.
- Added `format_ranges` for batch formatting and improved conditional formatting DX.
- Unified chart authoring so `create_chart` now supports both contiguous `data_range` inputs and explicit `series` definitions.
- Fixed pivot grouping bugs around `columns` handling and case-insensitive field resolution.
- Fixed contiguous scatter chart creation so the first Y data point is no longer dropped.
- Expanded regression coverage from the original baseline to 172 passing tests.

### Packaging and Docs

- Added Ruff linting and build smoke checks to CI.
- Synced the manifest, README, and static landing page with the current 47-tool surface.
- Added this changelog for future release notes.
