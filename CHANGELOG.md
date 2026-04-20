# Changelog

## Unreleased

### Added

- Added `create_named_range` so agents can create workbook-level or sheet-scoped named ranges with `dry_run` previews and same-scope `replace=True` support instead of dropping to ad hoc Python for a common workbook-structure task.
- Added semantic dataset identity metadata to `describe_dataset`, `quick_read`, `read_excel_as_table`, and `read_excel_table`: each response now includes `structure_token`, `content_token`, and `snapshot_metadata` for optimistic-concurrency style follow-up writes.

### Changed

- Changed `rename_worksheet` so it now updates formula cells in addition to chart references and named ranges, and it also renames the default sibling pivot sheet (`Data_pivot` -> `Revenue_pivot`) when that move is conflict-free.
- Changed worksheet-shaped dataset reads and row-mutation helpers to stop at the main contiguous data block after the header, so sparse footer notes or distant outlier rows no longer inflate `quick_read` / `read_excel_as_table` row counts or confuse `append_table_rows` and `update_rows_by_key`; `describe_dataset` now surfaces `data_end_row` and `ignored_trailing_row_count` when later rows are treated as a separate block.
- Changed `append_table_rows`, `append_excel_table_rows`, `upsert_excel_table_rows`, and `update_rows_by_key` so they can enforce `expected_structure_token` preconditions. Append-style writes now require explicit `allow_structure_change=True` when the caller is intentionally growing a dataset or native table.
- Changed workbook persistence through `safe_workbook(..., save=True)` to use temp-file save, `fsync`, atomic replace, and reopen verification instead of writing directly over the original path.
- Fixed structured-write `dry_run` previews so simulated `new_*_token` values no longer reuse live-file `snapshot_metadata`; previews now report `token_basis="dry_run_preview"` with `source_file_*` metadata instead.
- Fixed the atomic save path so a post-replace verify failure rolls back to the original workbook when possible instead of raising after the destination file has already been irreversibly changed.
- Changed `audit_workbook` so dominant native-table sheets are judged by the table's own headers when nearby dashboard artifacts extend the used range, reducing false `blank_headers` or `duplicate_headers` findings on mixed sheets.
- Refreshed README, package metadata, manifest copy, and landing-page positioning to better highlight SheetForge's current local-first, agent-friendly workbook reading, introspection, and safer mutation strengths without promising unreleased concurrency features.

## 0.7.0 - 2026-04-19

### Added

- Added `append_excel_table_rows` so agents can append rows to native Excel tables without forcing a key-based upsert flow, while still expanding the table `ref` safely and respecting occupied-cell and totals-row guardrails.

### Changed

- Changed `append_table_rows` to reject writes that would land directly below an adjacent native Excel table, and to point callers at `append_excel_table_rows` instead of silently leaving the table range stale.

## 0.6.1 - 2026-04-19

### Fixed

- Fixed Python packaging so built wheels now include the `excel_mcp` module files instead of publishing metadata-only wheels.
- Fixed `query_table`, `aggregate_table`, and the multi-workbook filter helpers so mixed-type rows such as totals formulas no longer abort `gt` / `gte` / `lt` / `lte` filters; incompatible rows are treated as non-matches instead.
- Fixed filter DX so `ne` now aliases `neq`, and `in` / `not_in` filters now accept the shorthand `value: [...]` in addition to `values: [...]`.
- Fixed formatting color parsing so `format_range` and `format_ranges` now accept CSS-style `#RRGGBB` and `#AARRGGBB` inputs, with clearer error guidance when a color token is invalid.

## 0.6.0 - 2026-04-19

This minor release expands SheetForge MCP into a stronger workbook-audit, repair, query, and multi-workbook reporting surface, while tightening rename/copy semantics, response guidance, release integrity, and end-to-end agent workflows.

### Added

- Added `cross_workbook_lookup` so agents can enrich one worksheet or native Excel table from matching rows in one or more lookup workbooks, with left or inner join behavior, duplicate-match strategies (`first`, `all`, `error`), lookup-side sort precedence, and explicit `strict`, `intersect`, and `union` schema handling across lookup files.
- Added `union_tables` so agents can combine comparable worksheet or native-table rows across many workbook files in one call, with optional `dedupe_on` keys for "latest row per ID" style union workflows and the same `strict`, `intersect`, and `union` schema controls as the other multi-workbook read helpers.
- Added `bulk_filter_workbooks` as the row-level companion to `bulk_aggregate_workbooks`, so agents can pull matching rows across many workbook files with optional source provenance columns and explicit `strict`, `intersect`, and `union` schema modes.
- Added `bulk_aggregate_workbooks` as the first multi-workbook workflow so agents can roll up comparable workbook files with the same filter/group/metric contract as `aggregate_table`, plus explicit `strict`, `intersect`, and `union` schema modes.
- Added `inspect_formula` so agents can classify formula strings before writing them, including function inventory, reference token types, and flags for volatile or risky functions such as `INDIRECT`.
- Added `detect_circular_dependencies` so agents can scan workbook formula graphs for self-references and multi-cell circular dependency groups, including loops introduced through named ranges.
- Added `read_range_formatting` so agents can inspect worksheet look-and-feel by grouped style signatures, merged-range overlap, and conditional-format overlap instead of pulling noisy per-cell style dumps into context.
- Added `describe_sheet_layout` as a worksheet-level structural summary for dashboard-safe edits, including freeze panes, print settings, merges, tables, chart anchors, validation and conditional-format counts, custom dimensions, and a compact free-canvas preview.
- Added `suggest_read_strategy` so agents can ask SheetForge which read path best fits a workbook target before spending context on the wrong tool.
- Added `describe_dataset` as a lightweight worksheet/native-table summary with sample rows, inferred schema, header-quality signals, key-candidate guesses, and recommended follow-up reads.
- Added `query_table` for declarative filtering, projection, sorting, and limiting over worksheet-shaped data or native Excel tables.
- Added `aggregate_table` for grouped metrics such as `count`, `sum`, `avg`, `min`, and `max` without forcing agents to over-read whole datasets into context.
- Added `audit_workbook` as a workbook-level preflight that surfaces high-signal issues such as broken formula references, error cells, hidden sheets, layout-heavy tabs, header-quality problems, and missing-sheet named ranges.
- Added `plan_workbook_repairs` so agents can turn workbook audit findings into a prioritized SheetForge action queue instead of manually mapping each issue to the next tool call.
- Added `apply_workbook_repairs` so agents can dry-run or apply the safe repair subset for broken named ranges, broken validation rules, broken conditional formatting rules, and optional hidden-sheet reveals.
- Added `diff_workbooks` so agents can compare workbook versions with structural changes plus sampled cell-level before/after diffs.
- Added `explain_formula_cell` so agents can resolve formula inputs through named ranges and structured references, inspect upstream formula chains, and see downstream dependents in one call.
- Added `inspect_named_range`, `delete_named_range`, `inspect_data_validation_rules`, `remove_data_validation_rules`, `inspect_conditional_format_rules`, and `remove_conditional_format_rules` to turn workbook-repair plans into inspect → dry-run → apply workflows.
- Added public-doc regression checks so `manifest.json`, `TOOLS.md`, `README.md`, and the landing page stay aligned with the currently registered MCP tool surface.
- Added release-artifact regression checks so `pyproject.toml`, `manifest.json`, the published-version doc markers, and the tracked `.mcpb` bundle stay synchronized.
- Refreshed the manifest, README, and landing-page copy for the current 74-tool surface and the newer workbook-repair, workbook-diff, formula-lineage, formula-inspection, circular-dependency, layout-introspection, and multi-workbook read helpers.

### Changed

- Expanded `explain_formula_cell` with a compact `formula_chain` summary so agents can see chain depth layers, sampled formula edges, leaf precedents, root-to-leaf path samples, and whether `max_depth` truncated a deeper upstream chain.

### Fixed

- Fixed inferred read schemas so formula-backed columns are now labeled as `formula` instead of misleadingly appearing as plain `string` columns when SheetForge returns formula text from workbook reads.
- Fixed `rename_worksheet` so chart series references now follow the renamed sheet instead of keeping stale source formulas that still point at the old worksheet name.
- Fixed `rename_worksheet` so workbook-level and sheet-scoped named ranges now follow the renamed sheet instead of drifting to missing-sheet destinations.
- Fixed `copy_worksheet` so sheet-scoped named ranges are now duplicated onto the copied sheet with their references rewritten to the new sheet.
- Fixed oversized `read_excel_table` errors so the recovery hints now mention `compact=True`, matching the tool's real payload-trimming option.
- Fixed aggregation DX so `aggregate_table` and `bulk_aggregate_workbooks` now accept the more intuitive metric aliases `agg` and `column` in addition to the canonical `op` and `field` keys.
- Fixed `copy_range` so copied formulas now translate their relative references to the target cells instead of being pasted back as unchanged formula text.
- Fixed range-read continuation cursors so follow-up pages now preserve the original payload mode, including `values_only=True` and compact cell-metadata reads, instead of silently falling back to the default cell-metadata shape.
- Fixed `plan_workbook_repairs` so small `sample_limit` values no longer drop whole repair classes from the plan when the underlying audit contains more findings than fit in the sampled audit payload.
- Fixed `bulk_aggregate_workbooks` provenance so the top-level `auto_selected_sheet` flag now turns true whenever any source workbook had to auto-select its first worksheet.
- Fixed worksheet visibility guards so visible chart sheets now count correctly when preventing a workbook from ending up with zero visible sheets.
- Fixed `copy_range_operation` so overlapping same-sheet copies now use a stable source snapshot instead of reading already-overwritten cells mid-copy.
- Fixed `delete_range_operation` so upward and leftward shifts now move only the selected columns or rows instead of deleting whole worksheet rows or columns outside the requested range, and so `changes` previews match the actual shifted cells.
- Fixed workbook audit broken-formula detection so formulas with missing structured table references are now reported instead of only formulas with literal `#REF!` text.

## 0.5.0 - 2026-04-18

This minor release expands SheetForge MCP from workbook CRUD into a stronger workbook-analysis and large-read workflow for AI agents and automation pipelines.

### Added

- Added `next_start_row` to truncated tabular read responses so agents can continue pagination without recalculating offsets.
- Added `max_rows` pagination plus `next_start_row` / `next_start_cell` continuation hints to `read_data_from_excel` for large non-tabular range reads.
- Added `max_cols` windowing plus `next_start_col` / `next_column_start_cell` continuation hints to `read_data_from_excel` so wide non-tabular ranges can be paged horizontally as well.
- Added cursor-based range continuations to `read_data_from_excel`, including directional `continuations.down` / `continuations.right` tokens for 2D window traversal.
- Added a complex workbook regression fixture covering chartsheets, dashboard formulas, native tables, validations, conditional formats, and cursor-driven reads in one realistic test workbook.
- Added `start_col` / `end_col` support to `quick_read` and `read_excel_as_table` so wide worksheet reads can request a narrower column slice before pagination.
- Added `start_col` / `end_col` support to `read_excel_table` so native Excel tables can now use the same narrower column slices as the worksheet table readers.
- Added `find_free_canvas` plus `placement.relative_to="free_canvas"` so chart and dashboard workflows can discover non-overlapping layout slots automatically.
- Added `analyze_range_impact` as a read-only preflight that reports tables, chart footprints, merges, named ranges, autofilter, print area, and formulas touched by a worksheet range.
- Extended `analyze_range_impact` to report downstream formulas elsewhere in the workbook that reference the selected range.
- Extended `analyze_range_impact` again to catch downstream formula dependencies that reach the selected range through named ranges.
- Extended `analyze_range_impact` to catch downstream dependencies that reach the selected range through structured table references such as `Table1[Sales]` and `Table1[@Sales]`.
- Extended `analyze_range_impact` to report overlapping data validations and conditional formatting rules, plus downstream validation and conditional-format expressions that reference the selected range.
- Extended `analyze_range_impact` to follow transitive formula chains, so second- and third-hop workbook formulas are now surfaced with dependency depth and predecessor cells.
- Refreshed README, package metadata, and landing-page copy for the current 51-tool surface and the newer large-read plus impact-analysis capabilities.

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
