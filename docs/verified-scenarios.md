# Verified Scenarios v1

Available since SheetForge MCP `0.13.0`.

`evaluate_workbook_scenarios` compares explicit keyed alternatives against one
captured native Excel table. It recomputes a shared sum/count recipe, checks exact
metric bounds, and can commit one explicitly selected, still-current alternative.
It does not choose changes, optimize a business objective, forecast outcomes or
evaluate Excel formulas. No external model or service is called.

## First comparison

The [example scenario set](examples/budget.scenarios.json) expects this native
table named `Budget` (the source worksheet name is optional):

| Id | Unit | Budget | Note |
| --- | --- | --- | --- |
| 1 | North | 15000 | one |
| 2 | North | 25000 | two |
| 3 | South | 30000 | three |
| 4 | East | 30000 | four |

The overall maximum is 100000, and every included unit must retain at least 10000.
All three alternatives start from the same 100000 baseline, not from the result
of the previous alternative:

| Alternative | Overall result | Smallest unit | Guards |
| --- | --- | --- | --- |
| A | 103000 | 30000 | Fail: overall maximum |
| B | 100000 | 25000 | Pass |
| C | 98000 | 9500 | Fail: unit minimum |

```python
from excel_mcp.scenarios import evaluate_workbook_scenarios

# scenario_set is the parsed example JSON, not a path or executable script.
comparison = evaluate_workbook_scenarios(filepath, scenario_set)
selected = next(s for s in comparison["scenarios"] if s["id"] == "B")
if selected["ready_to_commit"]:
    committed = evaluate_workbook_scenarios(
        filepath, scenario_set, mode="commit", selected_scenario="B",
        expected_workbook_sha256=comparison["expected_workbook_sha256"],
        scenario_token=selected["scenario_token"],
    )
```

MCP wraps results in the normal `ok`/`data` envelope. Comparison creates private
temporary candidates, but changes neither the source bytes/mtime nor its sheets,
and creates no persistent snapshot. A guarded failure is an ordinary comparison
result. Committing a failed alternative raises a precondition error without writes.

## Closed input contract

`scenario_set` requires exactly these fields:

- `version`: integer `1`.
- `recipe`: a full [Recipe v1](verified-recipes.md), with only `sum` and `count`
  metrics (at most eight). Output settings remain validated and fingerprinted but
  are **not written or inspected**. Saved report cells and their capacity are not
  evaluated. Filters, recipe checks and invariance probes run normally.
- `key_fields`: nonempty exact header names; composite keys are supported.
- `guards`: one to twenty bounds, as described below.
- `scenarios`: one to five `{id, updates}` alternatives. IDs must be unique,
  case-sensitive ASCII letters/digits/underscore/hyphen, 1-64 characters. Each
  nonempty `updates` list uses `{key: {all key fields}, values: {field: literal}}`.

Unknown properties, duplicate updates, missing keys and unsupported inputs fail
the whole request. Only guard violations produce per-alternative rejection; an
unverifiable calculation is not presented as a verified partial comparison.
Source keys must be nonblank and unique over every row, including filter-excluded
rows. Totals rows are excluded. Headers are exact and unambiguous; table metadata
must agree with worksheet headers. A fresh comparison follows reordered rows and
columns by keys/headers. It does not infer renames or intended field meanings.

Updates address existing rows only and cannot change key fields. Referenced values
must be homogeneous finite JSON scalars; no formulas, Excel errors, dates or native
calculated columns, including literal overrides in such columns. Unreferenced
formulas are preserved, not evaluated. Outputs use [Rebase's literal restrictions](verified-rebase.md):
null clears; empty text, formula/error-like text, invalid control/Unicode strings,
overlong text and numbers that cannot roundtrip exactly are rejected. Protected
worksheets and merged targets are rejected. Only explicitly compiled cells are
written; recipe output rectangles are untouched.

Comparison uses the current capture, not an earlier retained base. If proposed
values were prepared from a previous read, pass its full workbook SHA-256 as
`expected_workbook_sha256` even in compare mode. Without that pin, SheetForge
cannot detect stale assumptions behind the caller's proposed values.

## Guard semantics

A guard is `{metric, scope, min?, max?}`. `metric` is an exact metric alias, `scope`
is `overall` or `each_group`, and at least one inclusive finite numeric bound is
required. Booleans, numeric strings and null bounds are rejected; min cannot exceed
max. Bounds have only the units assigned by the caller; no unit/currency inference.

- `overall` checks the sum across all recomputed groups. It is zero for an empty
  filtered input. Add a count minimum if an empty result must be rejected.
- `each_group` requires nonempty recipe `group_by`. It checks the union of groups
  present in the baseline and candidate. A missing baseline group fails rather than
  evading a minimum. New groups are checked too. Empty group coverage fails. Null
  group labels are valid identities, not missing groups.

Only the candidate must satisfy the guards: a scenario may repair a baseline that
violates them. Both versions must still pass recipe checks, input validation and
independent reconciliation. Every supported row/group is checked, not sampled.

Guards use exact rational arithmetic on the decimal representations of loaded
source numbers. There is no tolerance in a guard comparison. Recipe arithmetic
is independently reconciled using the existing integer-exact / floating
relative `1e-12`, absolute `1e-9` policy. An arithmetic discrepancy fails closed.
This verifies the supplied model, not the accuracy of source data, the meaning of
the metric, or a real-world consequence such as increased sales.

## Results and evidence

Each alternative reports `passes_guards`, `ready_to_commit`, an integrity token,
overall metric before/after/delta, group counts, exact observed min/max per guard,
complete evaluated/failed guard counts, bounded failure examples and write samples.
Numerical metric/guard evidence uses decimal strings, not binary float JSON values.
Numeric group labels in failure samples carry explicit type plus decimal text.
Very long sampled source text is represented by preview/length/digest objects.

`passes_guards` alone is not write readiness. Eligible nonempty alternatives also
run a real ChangeSet persistence preview, which checks serialized values, outcome
assertions and snapshot availability. A snapshot conflict can make
`ready_to_commit=false` even when all metric bounds pass; `snapshot` reports the
conflicting path and reason without overwriting it. No automatic ranking or
winner is returned. Sample size affects presentation only, not acceptance/tokens.

Tokens bind canonical target path, exact source digest, the full normalized scenario
set (including unselected alternatives), selected ID, engine/numeric policy versions,
compiled operations/assertions and complete referenced-input/result digests. They
are not authorization or an exactly-once delivery receipt. The service is stateless.

## Commit and failure safety

Commit requires the identical scenario set, one `selected_scenario`, its token and
the comparison's exact `expected_workbook_sha256`. Any intervening change to file
bytes invalidates the approval. Compare again; do not remove guards to bypass rejection. A fresh
comparison includes changes to **unpatched** rows in all relevant metric checks.

The selected alternative is recomputed and compiled into a ChangeSet. Besides
exact target-value assertions, `table_scenario_matches` recomputes the baseline
and saved candidate's metrics, independently reconciles them, checks all guards
and compares the complete referenced-input/result digest. It runs both before
replacement and after reopening the saved destination within ChangeSet rollback.
An unexpected post-replace guard failure or result drift restores the exact
pre-commit source bytes, even if a changed total happens to remain within bounds.

All-already-satisfied alternatives still require guard checks, but do not save or
create a snapshot: `persisted=false`. This is a captured-state observation, not
proof of earlier execution. Successful commit credentials become stale; a fresh
comparison may show that the requested values are already satisfied.

The underlying snapshot, same-host locks, exact-file concurrency checks and atomic
replace/rollback remain in force. They do not provide distributed cloud locking,
cross-file transactions or perfect preservation of unsupported Excel features.

Advanced ChangeSet clients can use the same assertion explicitly:
`{type: "table_scenario_matches", recipe, key_fields, guards, expected_result_sha256}`.
Use the selected alternative's `result_sha256`; changing baseline, referenced
inputs, metrics or guard definitions invalidates it. This assertion does not itself
restrict a direct ChangeSet's mutation allowlist; the scenario tool additionally
enforces existing-key/literal-only updates.

## Limits and validation

Limits: 256000 UTF-8 bytes per scenario set; five alternatives; eight sum/count
metrics; twenty guards; 50000 native-table cells/64 columns; 1000 target cells and
49 contiguous one-column write spans per alternative. The extra outcome assertion
uses ChangeSet's fiftieth assertion slot. Each generated patch also has Rebase's
128000-byte limit. Recipe source/result/JSON limits and ChangeSet limits still apply.
Source loading is not streaming or a whole-workbook memory guarantee.

`sample_limit` is 1-5 (default 3). Comparison responses have a 40000-character budget;
commit additions have an 8000-character budget alongside ChangeSet's 40000 budget.
Counts include ASCII JSON escaping. Budget failures occur before original-file
replacement. Unsupported or over-limit input does not trigger partial writes.

`tests/test_scenario_workflows.py` exercises the actual stdio transport on 120
synthetic records, three alternatives, a competing editor, stale approval rejection,
complete Decimal-based readback, and guard drift on an unpatched row even when all
proposed cells are already satisfied. Domain tests inject post-replace guard/result
failures and verify exact-byte rollback. These are synthetic integration checks,
not customer acceptance or evidence that an unassisted model always supplies the
right alternatives and constraints.
