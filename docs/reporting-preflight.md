# Reporting Preflight v1

Unreleased. Main registers three new read-only tools (86 total); the published
0.15.5 MCPB and package still expose 83. No new dependency is required.

These tools help decide whether to join, which declared filter rule to approve,
and whether an existing report needs attention. They do not infer business truth,
choose definitions, authorize writes or run arbitrary code.

## Join Rehearsal

```json
{
  "left_filepath": "/reports/invoices.xlsx",
  "right_filepath": "/reports/shipments.xlsx",
  "left_source": {"table_name": "Invoices"},
  "right_source": {"table_name": "Shipments"},
  "left_key": "InvoiceId",
  "right_key": "InvoiceId",
  "measure_field": "Amount",
  "join_type": "inner"
}
```

Send these arguments to `rehearse_table_join`. Sources are explicit native Excel
tables, optionally scoped by `sheet_name`. Only `left` and `inner` all-match joins
are modeled. One numeric measure on the **left** is followed through each row's
projected multiplicity. This is not the lookup tool's `match_mode="first"`.

For two 120-unit invoices, if one has two shipment matches and the other none,
an inner join still totals 240. The tool separately reports one multiplied row,
one dropped row, `extra_signed: "120"`, `lost_signed: "120"` and `delta: "0"`.
Absolute amounts expose cancellation among positive and negative measures too.
Counts are complete; physical-row examples and right matching rows are sampled.
Measure evidence uses exact decimal strings derived from the loaded numbers,
not formatted Excel displays or a newly materialized joined workbook.

Keys use exact typed equality: numeric `1` and `1.0` match, but text `"1"` and
boolean `true` do not. No case folding or whitespace trimming. Null, empty and
whitespace-only keys never match each other. Null measures contribute zero and
are counted explicitly. Formula/error/date inputs and mixed field types are
rejected, not coerced. Declared totals are excluded; hidden/filtered rows remain.

Exact numeric keys follow executed Python/query equality, not decimal spelling:
integer `10**23` and float `1e23` do not match, while `int(1e23)` and that float do.
This key policy is separate from the decimal strings used for measure arithmetic.

`requires_decision` flags multiplied, dropped or unmatched left rows. It does
not declare them errors. Many shipments to one unique invoice can legitimately
preserve a shipment-count measure; allocating invoice revenue among shipment
types still requires an owner-supplied allocation rule. Neither net conservation
nor a clean preflight proves correct grain, attribution or missing-source coverage.

## Rule Probe

`probe_recipe_rules(filepath, left_recipe, right_recipe)` requires two Recipe v1
objects identical **except for filters**. Up to eight sum/count metrics, no checks
or Workbook Contracts in this first version. Both source definitions are validated
and evaluated on the complete bounded actual table before seeking a witness.
Output rectangles are neither read nor written by this probe.

If the historical table has only positive sales, both "all entries" and
"Amount >= 0" can reproduce its report. A hypothetical negative entry can expose
their difference. A returned `witness` contains a single record and both computed
results. `kind` clearly distinguishes an observed row from a hypothetical case;
`source_row` is null for a hypothetical record. This is an example question for
the data owner, not proof that the example exists or satisfies business constraints.

The deterministic search tries up to 128 actual singleton rows and then up to
128 hypothetical rows with at most four referenced fields and 16 values per
field. Values come from bounded observed values, explicit predicates, boundary
neighbors and type-appropriate examples. Known column types are taken from the
complete loaded source, even when its first eight values are null; example values
remain sampled. An entirely null column has no observed type. Expected validation
or calculation errors on hypothetical rows are counted as invalid synthetic cases,
not failures of the real comparison. Actual-source errors and unexpected engine
errors still propagate; no general exception suppression is performed.
`real_results_differ` can be true even if no compact singleton witness is found.
Results expose hashes, counts, bounded samples and search coverage. No witness
means **no difference found in these tests**, never equivalence. Numeric group
identities are exact; metric comparisons use Recipe's documented tolerances.
Actual aggregates are compared by keyed groups, not zipped presentation rows.
Equivalent `1`/`1.0` group representations may sort differently without implying
a different aggregate. Presentation ordering and raw result hashes are unchanged.
The client may retain an owner-approved witness as a test case; the server does
not store approval or automatically change the recipe.

## Report Receipts

Set `include_report_receipt=true` on both preview and commit of
`run_workbook_recipe`. Keep the returned `report_receipt` outside the workbook.
The default remains false. A passing preview constructs and budgets it but does
**not** issue it. Commit returns it only after existing ChangeSet assertions have verified
the exact saved output and source. No additional fallible receipt construction
occurs after persistence. Oversized receipts fail before any mutation. If Recipe
checks already failed, preview/commit report those checks without attempting to
construct an optional receipt. No receipt is returned in that failure path.

New receipts use **receipt format 3** (the Recipe itself stays version 1). Their
checksum and source/output/group fingerprints use exact integer/float ratios,
matching executable numeric equality rather than rounded decimal spellings.
JSON containers/scalars remain tagged. Equivalent `240.0`/`240`, `-0.0`/`0` and
value-preserving exponent spellings survive reserialization. But integer `10**23`
and float `1e23` are **not** equal: a JavaScript roundtrip between them changes
an executable filter and is rejected. Booleans, strings and numbers remain distinct.
This is not a guarantee that JavaScript can preserve every Excel number or
identifier; retain original receipt JSON when the client cannot preserve it.
Other Recipe/ChangeSet tokens and unrelated fingerprint algorithms are unchanged.

Formats 2/3 store `output_headers` from the actual compiled report. This preserves
existing writer behavior: ordinary Recipe metric aliases are trimmed by the query
engine; Contract-backed output retains requested aliases. The checker and its
comparison labels follow those effective headers, without modifying the request
or granting permission to rename a report column.

Legacy formats 1/2 retain their original checksum and value-fingerprint algorithms
only for numerically unambiguous domains. Format 1 still requires its original
JSON numeric spelling. Legacy integers not exactly representable as floats, and
floats whose integral decimal spelling differs from their exact value, cannot
be certified safely. Such numbers in the receipt, referenced source fields or
reserved saved output fail closed with `legacy_receipt_numeric_ambiguity`,
`requires_new_receipt=true`, and `suggested_next_tool="run_workbook_recipe"`.
Obtain a new format-3 receipt through an authorized preview/verified commit;
there is no automatic upgrade or checksum repair. Unsupported versions fail closed.

The receipt holds the normalized recipe and optional contract, the recipe engine
version, a fingerprint of referenced source values, bounded aggregate result rows,
and a fingerprint of the entire reserved output rectangle including padding.
Raw source records are not embedded, but group labels and aggregate metric values
are retained to support historical comparisons. Definitions can also contain
sensitive literals, and low-entropy hashes can be guessed: **treat the
receipt as private**. Its checksum detects accidental edits, not an adversary
who can replace it and recompute hashes. It is not a signature, authentication,
durable delivery acknowledgment or exactly-once retry credential.

```python
from excel_mcp.report_receipts import check_report_receipt

checked = check_report_receipt(filepath, report_receipt)
```

Use the same arguments with the MCP tool. It recomputes the receipt's exact
definition against a fresh stable capture, including complete Contract/check
validation, and compares it with stored output values:

| Observation | Status / action |
| --- | --- |
| Stored values match current computed values | `current`; no numeric refresh needed |
| Referenced source changed but aggregate is unchanged | `current`, `source_changed=true` |
| Source result changed, saved output unchanged | `refresh_needed`; obtain a fresh preview |
| Saved output changed and differs from current result | `conflict`; obtain an owner decision |
| Output was reordered/repositioned and the computed result also changed | `conflict` with `conflict_reason="output_positions_changed"`; no implicit permission to discard the observed ordering |
| Source/contract invalid or unsupported | Error; no claim that the report is current |

`source_changed` covers referenced fields and **all** source rows in physical
order, excluding declared totals. It does not mean every workbook byte or every
unreferenced column was compared. `source_sha256` separately identifies the full
current capture. Logical groups and metrics are matched by hashes rather than
output row position; reordering output groups alone is not a metric conflict,
but `output_changed` still reports the full-grid change. Extra values in padding,
changed headers, formulas and duplicate output groups cannot silently certify a
matching result. Values are compared exactly after numeric identity normalization
(`1` equals `1.0`, not `true`). Styles/layout/causes/authorship are not certified.
`output_values_changed` compares stored group/metric values with the receipt,
and is null when output cannot be interpreted. `output_positions_only_changed`
identifies a changed grid whose keyed values still match the receipt (including
row order or blank-row placement). It does not identify who changed it. A combined
source refresh deliberately remains conservative; the separate `conflict_reason`
distinguishes positions, values and unparseable output structure. An already-current
reordered report still needs no numeric refresh. These advisory fields never write.
`comparison_sample` returns bounded per-group `previous`, `computed` and `stored`
metric values. Missing groups have null entries, not invented zero values.
If headers/formulas/duplicate groups prevent interpretation of saved output,
`stored_output_parseable=false`; no per-group stored interpretation is certified.

Checking a receipt never writes. A later Recipe commit still needs its existing
fresh preview credentials and output overwrite permission. The receipt does not
automatically guard direct writers or resolve a changed output definition. Reuse
the approved filepath; this portable receipt is not evidence of file identity.

## Bounds And Acceptance

Each source is limited to 50,000 table cells/64 columns and literal finite JSON
values in referenced fields. Each file is captured stably under existing path,
lock and input controls; two captures are **not** a distributed/cross-file
transaction. This is not streaming or a memory bound on unrelated workbook data.
All three tools use `sample_limit=3` by default, accepting integers 1-5. Evidence
responses are capped at 12,000 escaped JSON characters. Receipts are capped at
6,000 characters when issued and share Recipe's existing 8,000-character pre-commit
summary budget. Checking accepts up to 24,000 characters for bounded JSON notation
growth (such as exponent-to-integer formatting), with unchanged integrity checks.
This is not a larger issuance or write budget. Oversized input errors advise using
the original issued JSON or obtaining a new receipt through an authorized refresh,
not reducing a sample size that cannot shrink an existing receipt.
Large definitions/groups can exceed issuance limits; omit the optional
receipt rather than weakening mutation checks. No receipt history is persisted.
An entirely null result row is indistinguishable from blank output padding and
is refused when requesting a receipt; ordinary recipes retain their behavior.

```bash
uv run pytest -q tests/test_reporting_preflight.py tests/test_preflight_audit_regressions.py tests/test_receipt_identity_regressions.py
```

Additional exact-key, operator-guard, retained-receipt and real CLI regression
workflows are in `tests/test_reporting_guard_regressions.py` and
`tests/test_guard_workflows.py`. Rebase controls verify blocked commits leave both
files unchanged and create no snapshot; valid equal-value commits still persist
and preserve their baseline snapshot. These are synthetic automated checks.

Acceptance includes independent nested-loop join arithmetic, cancelling losses
and duplication, typed/blank keys, legitimate many-to-one counts, absent refunds,
identical-rule inconclusiveness, receipt refresh/conflict/no-op cases, Contract
rebinding, tampering/budget/transaction failures, byte/mtime preservation and real
stdio tool discovery plus preview/commit/readback. These are synthetic scripted
acceptance tests, not an actual-model success-rate or token-saving benchmark,
native Excel visual acceptance or a guarantee of business correctness.
The audit regressions include optional actual Node JSON and Node-plus-MCP checks
(skipped if Node is absent), alongside unconditional Python numeric-identity,
exact unsafe-integer/float identities, legacy ambiguity refusal, notation-growth
budgets, check-failure priority, known-type, synthetic-error and conflict-policy
controls. Node is not a runtime dependency of SheetForge.
