# Changelog

## Unreleased

### Added

- Added `profile_workbook` as a one-call workbook inventory for sheets, tables, charts, named ranges, and key layout/protection state.
- Added `upsert_excel_table_rows` for key-based updates plus append-missing behavior directly on native Excel tables.

### Quality

- Added regression coverage for workbook profiling and native-table upserts, including refusal to expand a table into occupied cells below it.

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
