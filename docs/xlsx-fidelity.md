# XLSX Fidelity Checks

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
`write_package_fidelity_failed` for `write_data_to_excel`. Both paths reject
observed candidate drift before replacement.

For `write_data_to_excel`, `dry_run=true` saves and checks an isolated
candidate without changing the source. Read `data.ready_to_commit` and
`data.package_fidelity`; a preview is not a lock or an exact-file pin between
calls. Use ChangeSet if the commit must match a specific previewed SHA-256.
Within ChangeSet, an omitted `sheet_name` on a cell write selects the active
worksheet (an active chart sheet is rejected).

## Scope and recovery

- The guard runs in `write_data_to_excel`, `apply_workbook_changeset` and
  ChangeSet-backed writes (Verified Recipes, Rebase and Scenarios). It does not
  run in other standalone mutation tools or in `diff_workbooks`.
- Direct cell writes and all ChangeSet staging saves load with openpyxl's
  [rich-text preservation](https://openpyxl.readthedocs.io/en/3.1/rich_text.html).
  Ordinary text writes that repeat the same non-formula text leave existing
  inline formatting intact; replacing a cell with new text can replace its
  formatting runs. Read APIs and cell-value assertions still use plain text.
- This does not preserve or recalculate formulas, validate the correctness of
  cached values, or guarantee complete Excel round-trip fidelity. It does not
  establish whether unclassified changes such as comment or drawing member
  renames are harmless.
- Inspection is bounded per archive: 20,000 ZIP entries, 50,000 relationships,
  256 MiB of scanned uncompressed content, 4 MiB per metadata XML part, and
  250,000 stored formula results per worksheet. Issue samples are capped at
  five entries and may be smaller based on the requested sample limit.
- If a workbook is blocked, keep the original and use Excel or another writer
  that preserves the needed features. A snapshot, a repeated commit, or
  switching to a different mutation tool does not resolve the fidelity problem.

For the role of the optional calculation-chain part, see
[Working with the calculation chain](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-calculation-chain).
