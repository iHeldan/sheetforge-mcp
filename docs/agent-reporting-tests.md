# Recurring Reporting Acceptance

These synthetic tasks exercise repeatable reporting rather than isolated tool
calls. No customer data or external model access is required for the pytest suite.

## Task Families

| Workflow | Repeated input changes | Required outcome |
| --- | --- | --- |
| Weekly sales | 1,200 rows, growth to 1,500 with reordered columns, then 180 rows in one group; cancellations, refunds and zero amounts | Refresh the same approved Contract/Recipe report; exclude declared totals and clear obsolete output |
| Monthly consolidation | Two native-table exports, 750 combined rows then 120; first tabs are chartsheets, columns reorder and an unrelated column appears | Aggregate approved common fields, preserve both inputs, refresh the existing report without stale rows/columns |
| Closing variance | 240 rows versus a revised version, with additions, removals, status/value changes and group moves; repeat with a duplicate key on an excluded row | Reconcile signed source contributions, or explicitly refuse invalid source identity without changing either file |

The output rectangle is explicitly owned: `Report!D4:H40`. Permission to replace
it includes any values within that rectangle. It is not inferred ownership or a
general authorization to overwrite another part of a workbook.

Weekly reports use the supplied logical-field contract unchanged. Monthly
consolidation includes only posted lines and computes the average across all
included rows, not the average of per-file averages. The task explicitly rounds
Revenue to two decimal places and Average to four, using half-even rounding.
This reporting decision is not an automatic coercion policy for ordinary writes.
See [write value safety](value-safety.md).

## Recovery Without Weakening Checks

- `row_mode` is `arrays` or `objects`, as enumerated in tabular tool schemas;
  `rows`, `columns` and `dicts` are not additional modes.
- Set `filepath` on the outer ChangeSet only. Nested operations inherit that
  file; omit `filepath`, `dry_run` and `include_changes` from their `args`.
- Choose a reporting precision explicitly before copying computed aggregates to
  cells. Do not turn a rejected number into text or a formula just to bypass a
  guard. Exact identifiers and explicit Decimal inputs are not rounded for you.
- After any source or approved-definition change, request a fresh preview and
  use its current credentials. An old successful run does not authorize a new one.
- Duplicate keys remain invalid even on cancelled/filter-excluded rows. Report
  the source issue instead of dropping rows or weakening the definition.

## Independent Acceptance

`tests/reporting_cases.py` generates fixtures, task prompts and expected results
without calling SheetForge calculation functions. Its grader checks:

- Every output group, sum, count and average against a separate Decimal calculation.
- Duplicate/stale groups, gaps, unused output rows and unused output columns.
- Source cells/formulas/types, native table metadata and non-output cell styles.
- Full chart XML, chart anchors, freezes, print areas, merged ranges and dimensions.
- Byte and modification-time identity for read-only source files.
- Signed added/removed/updated contributions and all group deltas in a variance report.

Numeric comparisons use an absolute `1e-9` tolerance for computed floating results;
this is not exact-decimal accounting. Source preservation checks are separate from
numeric comparisons. The scripted client also verifies each commit's snapshot
against the complete pre-mutation bytes and reads back the saved report over MCP.

The grader is tested with deliberately wrong totals, missing or stale groups,
gapped rows, stale columns, changed source cells, lost styling, changed chart
titles and invented variance components. A correct-result control must pass too;
a checker that rejects everything is not evidence of safety.

## Scripted Checks

```bash
uv sync --locked --extra dev
uv run pytest -q tests/test_reporting_workflows.py tests/test_value_safety.py
```

`tests/test_reporting_workflows.py` runs the three task families over real stdio
MCP. Each family reuses its report across sequential phases. The server resolves
the refreshed source on every call. Per-response and total response-size bounds
are asserted. These clients are deterministic programs, not language models.

## Running an Actual Agent

Use `build_case(folder, workflow, phase)` and `case.task()` from
`tests.reporting_cases` to supply the task to a fresh agent. Keep the case object
and expected results outside the agent's accessible files. After completion, call
`grade_case(case)` for saved reports, or `grade_case(case, answer)` for the requested
variance JSON. Run phases in order in the same fixture folder; each phase updates
the input while retaining the previous report.

Use a separately installed wheel, verify its source identity, and expose only
the public SheetForge MCP tools with paths confined to synthetic fixtures. Do
not expose repository/private notes, the grading oracle, shell/code execution,
unrelated MCP servers, hooks or memory. For every attempt, retain the exact task,
source/build identity, resolved model, effort, CLI version, calls, usage, timing,
errors and independent grade. Set explicit time, call and spend bounds. Inspect
successful readback and baseline-snapshot evidence, not just the agent's text.

Compare baseline and candidate using the same tasks, fixture definitions and
model/settings. Preserve failures and label all reruns; do not silently replace
a failed attempt with a successful one. Report cached input, uncached input and
output tokens separately. Response characters are not token counts.

## Manual Pilot (2026-11-27)

A bounded pilot used Claude Code 2.1.332, resolved model `claude-sonnet-5-5`,
medium effort and fresh sessions with only the 83 public SheetForge tools. Each
cohort ran the seven sequential task phases without human task interventions:

| Installed build | Independently accepted tasks | MCP tool calls |
| --- | --- | --- |
| Baseline | 6 / 7 | 49 |
| Unreleased float-identity correction | 7 / 7 | 42 |
| Correction plus row-mode/schema and ChangeSet guidance | 7 / 7 | 33 |

The baseline failure was a safe but incorrect refusal to write `71753.65`.
The corrected builds completed that report; all cohorts retained the expected
duplicate-key refusal. The final cohort had no unexpected tool-error responses.
Task text was identical after replacing temporary file roots, the resolved model
and CLI were unchanged, and installed source hashes were checked. Later phases
retain each cohort's earlier report, so a prior failure can affect later starting
output. The local candidate is not the published `0.15.4` package.

These observations do not establish a general speed/token advantage. Model
choices and cache state vary, and one attempt per phase is not a reliability
estimate. Raw transcripts, temporary workbooks and usage receipts are not shipped
in packages; public fixtures, prompts, graders and scripted controls are reusable.

These are small synthetic acceptance cases, not a statistically representative
success rate, a cross-model benchmark, customer acceptance or native Excel visual
validation. They do not establish cloud locking, universal XLSX fidelity or the
business correctness of the supplied report definitions. Variance results explain
recomputed source contributions, not saved historical report cells or real-world
causality. See [Verified Explanations](verified-explanations.md).
