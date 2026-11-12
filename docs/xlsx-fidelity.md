# ChangeSet XLSX Fidelity Checks

An XLSX file is a ZIP package. A workbook can reopen successfully and satisfy
cell assertions even when a writer discarded package content. ChangeSet checks
the staged save against the original package separately from its workbook diff.

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
no snapshot is created by that attempt. The error code is
`changeset_package_fidelity_failed`, with the report in
`error.details.package_fidelity`. Commit also checks the candidate's bytes
again before replacing the original, so a candidate that changed after
inspection cannot be committed.

## Scope and recovery

- The guard runs in `apply_workbook_changeset` and in writes routed through it,
  including Verified Recipes, Rebase and Scenarios. It is not a guard on
  standalone mutation tools or on the public `diff_workbooks` tool.
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
  switching to a standalone mutation does not resolve the fidelity problem.

For the role of the optional calculation-chain part, see
[Working with the calculation chain](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-calculation-chain).
