# Verified Rebase v1

Status: development source, not included in the published `0.12.0` package.

`apply_table_patch` rebinds explicit keyed updates against a current native Excel
table without replacing it with an old copy. It compares a retained base snapshot,
current values and desired values. Only selected literal target cells are written.
It does not infer the meaning of a change or discover what an agent relied on.

## Prepare an intent

Retain a base using `create_workbook_snapshot` **before** preparing changes, and
retain its SHA-256. Pass its path as `base_filepath`, the live file as `filepath`,
and its digest as `expected_base_sha256` on every call. The files must be distinct
regular `.xlsx` files, not symlink/hardlink aliases. Each is captured independently;
no hidden workbook metadata or server session history is used.

The [example patch](examples/product-copy.patch.json) requires a native table
`Products` with `SKU`, `Name`, `Features` and `Description` columns:

```json
{
  "version": 1,
  "source": {"table_name": "Products"},
  "key_fields": ["SKU"],
  "read_fields": ["Name", "Features"],
  "updates": [
    {"key": {"SKU": "SKU-123"}, "values": {"Description": "Updated product copy"}}
  ]
}
```

This declares that the proposed description relied on `Name` and `Features` **on
that same row**. If either changed, the proposal is stale even if Description is
unchanged. Declare all relevant row-local dependencies. An explicit `read_fields:
[]` means no additional dependencies; it does not request automatic inference.
Cross-row totals, other sheets, external research and unlisted fields are not
verified. The caller remains responsible for those premises and business intent.

Unknown properties are rejected. `source` accepts `table_name` and an optional
`sheet_name`. Exact case-sensitive headers are required. Both versions must have
the same complete header set; reordering is supported, renaming/addition/removal
requires a new contract/base. No fuzzy matches are attempted. Referenced calculated
table columns are rejected even if the cells contain literal overrides.

Keys must be unique and nonblank across **all** data rows in each table. Composite
keys are supported and every update supplies every component. Numeric 1 and 1.0
identify the same record; text "1" and boolean true are distinct. Text identities
are exact, not trimmed/case-folded. Duplicate update keys are errors. Every target
must exist in the base; the tool never inserts a new record. Totals rows are excluded.
Write fields cannot be key fields or declared read fields.

## Preview, resolve, commit

1. Call with `mode="preview"` (default). Neither source is modified.
2. Inspect `ready_to_commit`, complete `counts`, bounded `sample`, `bindings`,
   assertions and diff. Samples prioritize conflicts and mark truncation.
3. If ready, repeat identical paths, patch and base digest with `mode="commit"`,
   `rebase_token` and `expected_workbook_sha256` from that preview.
4. Any intervening save invalidates commit. Request a fresh preview; do not drop
   preconditions or retry the stale commit without reconsidering conflicts.

Counts are **target cells**, not rows. Classification checks declared read fields
before considering a target already satisfied:

| Status | Meaning |
| --- | --- |
| `apply` | Current target still equals base; desired differs. Eligible, not written yet. |
| `already_satisfied` | Current target equals desired and declared read fields still match. |
| `write_conflict` | Current target differs from both base and desired. |
| `stale_precondition` | At least one declared read field changed. |
| `missing_row` | The base key no longer exists in current. |

One blocking cell prevents **every** write in the patch. Preview succeeds with
`ready_to_commit=false`; commit raises a precondition error. There is no force,
last-writer-wins or automatic partial apply. Resolve conflicting intent explicitly
or prepare a separately reviewed smaller patch. Removing a read field changes the
contract, not just the preview token; it must not be used to conceal stale premises.

When all targets are already satisfied, no write or snapshot is created, and
`persisted=false` even in commit mode. This describes the captured state, not who
wrote it or proof that an earlier operation executed exactly once.

## Persistence and evidence

Only eligible changed cells are compiled into contiguous one-column spans, with
complete value assertions. Compilation uses current coordinates and current file
bytes, never an old workbook as the replacement candidate. The existing ChangeSet
layer checks the exact current hash under its lock, creates a verified snapshot of
the **current** workbook, saves/reopens/asserts, and rolls back those exact current
bytes if post-replace verification fails. The original retained base is not modified.
Normal openpyxl compatibility limits still apply; this is not a byte-preservation
promise for every unsupported Excel feature or an exhaustive presentation diff.

The rebase token binds engine version, canonical base/current paths, both file
digests, the full patch and compiled ChangeSet plan. Sample size is presentation
only. It is an integrity fingerprint, not authorization or a durable retry receipt.
Same-host locks coordinate cooperating SheetForge writers, not cloud/desktop
editors on other machines; there is no distributed or cross-file atomicity guarantee.

Referenced cells must be finite JSON scalars, not formulas, Excel errors or dates.
Unreferenced formulas are not evaluated or rewritten. Null clears a target. Empty
output strings (use null), formula/error-like strings, illegal control characters,
text longer than 32,767 characters, and numbers that openpyxl cannot preserve
exactly are rejected. Output types may intentionally change. Protected current
sheets and merged targets fail closed.

Limits: 50,000 table cells including headers/totals and 64 columns per source;
128,000 UTF-8 bytes per patch; 1,000 target cells; 50 contiguous column write spans;
sample_limit 1-10. At most 8,000 ASCII-escaped response characters are added to
ChangeSet's pre-commit response budget. Full verification is bounded, never sampled;
only examples are shortened. Oversized or unsupported patches fail before writing.
Workbook loading is not a streaming or whole-file memory guarantee.

## Workflow validation

`tests/test_rebase_workflows.py` uses the real stdio MCP boundary with 120 keyed
updates, reversed rows/reordered columns, an intervening human comment, stale commit
rejection, fresh preview/commit and independent full-table readback. It also checks
that changed premises block already-satisfied values and malformed requests cannot
enable force merging. This is a synthetic integration test, not customer acceptance
or evidence that a model will always declare all relevant dependencies.
