# XLSX Fidelity Checks

Available since SheetForge MCP `0.14.1`.

An XLSX file is a ZIP package. A workbook can reopen successfully and satisfy
cell assertions even when a writer discarded package content. Guarded writes
check the staged save against the original package separately from cell changes.

## What to check in preview

`package_fidelity.status` is `passed`, `blocked`, or `unverified`. Check
`ready_to_commit` too: a passed fidelity check does not replace the other
ChangeSet preconditions. Issues report a code, severity, count and bounded
sample, without embedding custom XML contents or cached values.

| Finding | Behavior |
| --- | --- |
| Recognized custom XML part removed, changed, or no longer associated with the same relationships or content type | Blocks commit |
| Stored result lost or changed at a formula cell outside effective cell edits | Blocks commit |
| Calculation-chain part removed | Warning only; spreadsheet software can rebuild the chain |
| Other package member names removed | Counted as unclassified; neither proof of lost content nor proof of a harmless rewrite |
| Targeted comparison cannot finish (for example, unsupported layout or an inspection limit) | Blocks commit as unverified |

Part checks recognize conventional custom XML paths, custom XML relationship
types, and the custom XML properties content type. Formula results are compared
by worksheet name and cell, not by a worksheet's ZIP member name. The check
covers stored values on cells with a formula element. An effective cell edit
may intentionally replace a formula or its old result; writing the same formula
does not waive its check.

Commit recomputes the check. If it fails, the original remains unchanged and
no snapshot is created by that attempt. Read the report in
`error.details.package_fidelity`: the error code is
`changeset_package_fidelity_failed` for ChangeSet or
`write_package_fidelity_failed` for standalone mutations. These paths reject
observed candidate drift before replacement.

For `write_data_to_excel`, `dry_run=true` saves and checks an isolated
candidate without changing the source. Read `data.ready_to_commit` and
`data.package_fidelity`; a preview is not a lock or an exact-file pin between
calls. Use ChangeSet if the commit must match a specific previewed SHA-256.
Within ChangeSet, an omitted `sheet_name` on a cell write selects the active
worksheet (an active chart sheet is rejected).

### Post-save Recovery (Since 0.15.1)

Standalone saves also check source identity and bytes before replacement and
compare the reopened file with the saved candidate. An observed outside change
blocks automatic rollback. `write_recovery_required` includes
`error.details.recovery_backup` (null for a new file) and `retry_safe=false`.
Inspect the current file and retained backup; do not blindly retry or restore.
This concurrency protection is separate from the targeted package-fidelity
guard below; neither check establishes complete XLSX feature fidelity.

ChangeSet's final source SHA-256 must match the inspected candidate, not just
describe a readable workbook. If verification fails, rollback checks again
that the current source is still that candidate before restoring the baseline.
If it observes a different file or content, it skips rollback and retains the
baseline rather than knowingly overwriting an outside save.

The structured error is `changeset_recovery_required`.
`error.details.recovery_baseline` points to the retained original;
`rollback_status="skipped_source_changed"` and `retry_safe=false` mean that the
current file and baseline require inspection before another mutation. The
current file may contain the attempted edit, external changes, or both. Do not
automatically copy the baseline over it or blindly repeat the commit. Existing
verified snapshots are also retained. Other rollback failures retain the
baseline and report that recovery could not be verified.

This applies to standard and preserving ChangeSets and their backed workflows.
It is still optimistic concurrency with a cooperative same-host lock, not an
atomic compare-and-swap against Excel, a cloud sync client, or another outside
writer. An external change after a check remains possible.

## Scope and recovery

- Since `0.14.1`, the guard runs in `write_data_to_excel`,
  `apply_workbook_changeset` and ChangeSet-backed writes (Verified Recipes,
  Rebase and Scenarios). Since `0.15.1`, it also runs on ordinary standalone
  mutations, including formatting, charts, formulas, sheet operations and
  native-table writes. Their dry runs serialize and inspect an isolated
  candidate too. Unlike direct cell previews, which return `ready_to_commit`
  and the report, other standalone tools return a structured guard error on
  refusal. `diff_workbooks` is read-only and does not perform this comparison.
- ChangeSet intermediate saves defer the package comparison to its final
  original-versus-candidate check; they do not weaken that final check.
- Editable mutation loads, including standard ChangeSet staging, use openpyxl's
  [rich-text preservation](https://openpyxl.readthedocs.io/en/3.1/rich_text.html).
  Ordinary text writes that repeat the same non-formula text leave existing
  inline formatting intact; replacing a cell with new text can replace its
  formatting runs. Read APIs and cell-value assertions still use plain text.
- An ordinary mutation may intentionally replace or remove a formula's value
  at its existing worksheet/cell coordinate. Its old result is then exempted;
  an unchanged formula is not. Moving or renaming cached formulas can still
  be blocked when the check cannot establish their preservation. No automatic
  fallback to a less-protected writer is performed.
- This does not preserve or recalculate formulas, validate the correctness of
  cached values, or guarantee complete Excel round-trip fidelity. It does not
  establish whether unclassified changes such as comment or drawing member
  renames are harmless.
- Inspection is bounded per archive: 20,000 ZIP entries, 50,000 relationships,
  256 MiB of scanned uncompressed content, 4 MiB per metadata XML part, and
  250,000 stored formula results per worksheet. Issue samples are capped at
  five entries and may be smaller based on the requested sample limit.
- If a workbook is blocked, keep the original. For supported literal edits,
  try a fresh preview with the preserving mode below when available. Otherwise,
  use Excel or another writer that preserves the needed features. A snapshot or
  repeated commit does not resolve the fidelity problem.

For the role of the optional calculation-chain part, see
[Working with the calculation chain](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-calculation-chain).

## Preserving Save Mode

Available since SheetForge MCP `0.15.0`.

`apply_workbook_changeset` can use `save_mode="preserve"` to patch only the
target cells in worksheet XML. It checks the ZIP member list and SHA-256 of
every uncompressed member against the intended result, then reopens the
candidate and checks the written values. Unedited worksheet XML, recognized
custom XML and its relationships, and supported formula caches retain their
original bytes. The ZIP may be recompressed. This is not an assertion that
preserved custom XML has no business-level dependency on the edited values.

The first version is intentionally narrow:

- Only `write_data_to_excel` operations with plain nonempty strings, booleans,
  finite Excel-safe numbers (at most 15 significant digits), or `null` to clear
  a cell. No formulas, error values, dates, rich or shared-string targets, or
  overlapping edits. An unrecorded cell can be filled only on an existing row
  with a cell body, inside the declared worksheet dimension and any row spans.
- Every existing formula must be verifiably independent of the edit. This
  version accepts only numeric constant arithmetic such as `=6*9`; any cell
  reference, function, or unsupported formula encoding blocks the whole plan.
  SheetForge does not recalculate or validate the correctness of stored results.
- Only a restricted, unsigned `.xlsx` package with ordinary worksheets. Tables,
  charts, pivots, sheet protection, merges, validations, defined names, external
  relationships, and unknown package features are outside this profile.
- At most 1,000 target cells, 250,000 row and cell elements combined, 68 MiB on
  disk, 64 MiB uncompressed, and 16 MiB per worksheet XML part. These limits
  apply to the input and proposed output. Metadata and package-inspection
  limits still apply.

Example preview arguments:

```json
{
  "filepath": "/path/to/workbook.xlsx",
  "mode": "preview",
  "save_mode": "preserve",
  "operations": [
    {"tool": "write_data_to_excel", "args": {"sheet_name": "Inputs", "start_cell": "B2", "data": [[120]]}}
  ],
  "assertions": [{"type": "cell_equals", "sheet_name": "Inputs", "cell": "B2", "expected": 120}]
}
```

Review `preservation` for the changed-part count and sample, unchanged-part
count, and number of verified constant formulas. Commit only if
`ready_to_commit=true`: repeat the same plan and `save_mode` with
`mode="commit"`, `expected_workbook_sha256`, and `changeset_token` from preview.
The standard fidelity check, exact-file stale-write check, snapshot, and
verified rollback still apply. This mode does not change direct writes or
ChangeSet-backed recipes, rebase, or scenarios.

On `changeset_preserve_unsupported`, inspect `error.details.reason`. The
original workbook is unchanged. Do not retry in `standard` just to bypass the
rejection; use a compatible edit or a writer that supports the workbook.
