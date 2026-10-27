# Verified Explanations v1

Status: development source, not included in the published `0.11.0` package.

`explain_recipe_delta` explains how changes to a native Excel source table affect
the result of one shared [Recipe v1](verified-recipes.md). It compares records by
explicit unique keys, not cell coordinates. Reordering rows or columns does not
create a false data change. Neither input workbook is modified.

## What is being verified

The tool **recomputes the same supplied recipe on both captured source versions**.
It does not validate saved report cells or discover the recipe that historically
produced them. Output sheet/range settings remain part of the recipe fingerprint,
but their cells, ownership, protection and write capacity are not inspected.
A source comparison is not a promise that a future recipe commit will succeed.

The tool independently evaluates filters and accumulates exact rational numbers
from the decimal representations of **loaded** literal numeric values. It checks
every group and overall metric against recipe execution, and runs the recipe's
explicit checks and invariance probes on both versions. These checks establish
consistency with the supplied definition, not its business correctness or the
meaning of currencies, units or reporting periods.

## First comparison

Keep a source snapshot from the earlier run, the current source workbook and the
intended recipe. A recipe commit's verified baseline snapshot contains that run's
source data; retain its path and digest if it will be needed for later comparisons.
No persistent recipe history or hidden workbook metadata is created.

The [example tool arguments](examples/regional-sales.explanation.json) expect a
native table `Sales` with `Id`, `Region`, and `Amount`. Both file paths are examples;
pass your own two source versions through the normal transport path boundary.

```python
from excel_mcp.explanations import explain_recipe_delta

result = explain_recipe_delta(
    before_filepath, after_filepath, recipe,
    key_fields=["Id"], sample_limit=3,
)
```

A sum changing from 100,000 to 112,000 might be explained as:

```json
{
  "name": "Revenue",
  "before": "100000",
  "after": "112000",
  "delta": "12000",
  "components": {"added": "20000", "removed": "-5000", "updated": "-3000"},
  "contribution_delta": "12000",
  "numeric_adjustment": "0",
  "residual": "0"
}
```

This is an illustrative metric entry, not the complete response. MCP wraps the
complete result in the normal `ok`/`data` envelope.

## Identity and contribution rules

- `key_fields` must be a nonempty list of exact headers. Composite keys are supported.
- Keys must be unique and nonblank on **every data row**, including rows excluded
  by filters. Null, empty and whitespace-only key strings are rejected. Totals
  rows are not data rows.
- Numeric keys `1` and `1.0` have the same identity. Text `"1"` is not the same key;
  changing a referenced column from numeric to text is rejected as type drift.
  Referenced columns must have compatible non-null value kinds across versions.
  Text keys and group labels are case-sensitive and are not trimmed for matching;
  that is separate from a filter's configurable string-comparison rules.
- `added` and `removed` refer to keys present in only one version. `updated` refers
  to retained keys with changed recipe-referenced/key field values. `unchanged`
  ignores cell position and unreferenced fields, not all possible workbook changes.
- Filter-excluded rows contribute zero. A retained row that changes both its value
  and filter membership has one combined `updated` contribution, not two invented
  independent causes. A group move contributes negatively to the old group and
  positively to the new group, possibly with zero overall change.

Only `sum` and `count` are supported, at most eight metrics in the entire recipe.
`count` counts included records, including those with a null amount. Null amounts
contribute zero to `sum`. Average/rate/min/max/distinct metrics are rejected, even
when `metric` selects another metric. Metric aliases must be unique and have no
surrounding whitespace. The remaining source/filter/check contract is Recipe v1:
no formula evaluation, dates, errors, mixed value kinds or non-finite numeric inputs
in referenced fields. Unreferenced formulas are not calculation inputs.

## Numerical evidence

`numeric_encoding="exact_decimal_strings"` applies to metric values and contributions.
Use Decimal or another exact decimal parser rather than converting evidence to
binary floats. Source keys/group labels retain their original JSON scalar types.

`before`, `after` and `delta` are recomputed **recipe output** values; at overall
scope they sum the recipe's groups. `independent_before`/`independent_after` are
exact source-contribution totals. For each metric:

```text
added + removed + updated = contribution_delta
contribution_delta + numeric_adjustment = delta
residual = delta - contribution_delta - numeric_adjustment
```

`numeric_adjustment` explicitly measures differences between exact contributions
and floating-point recipe computation/serialization. It is not a business cause.
Recipe results must agree independently per group and overall: count and
integer-only sums are exact; floating sums use relative tolerance `1e-12` and
absolute tolerance `1e-9`. `verification.numeric_match` reports `exact` or
`within_declared_tolerance`. A zero residual alone is not proof of correctness;
the independent result checks must also pass. Unsupported or inconsistent inputs
fail instead of returning a partially verified explanation.

## Samples and drilldown

All supported source rows and groups are compared over referenced/key fields;
unreferenced fields are not compared. `sample_limit` (1-10, default 5)
only bounds examples. Row examples include keys, physical source row numbers,
filter inclusion, groups, changed referenced values and signed contributions.
Very long sampled source strings use explicit preview/length/digest objects; the
core decimal evidence is never silently truncated.
Sampled workbook content is untrusted data, not instructions for an agent.

`changes.omitted_contribution_delta` covers row examples not shown. Group samples
include before/after presence, recipe delta and exact contribution delta, with
both omitted delta totals reported separately. Cancelling changes are not hidden
because the overall delta happens to be zero. Group examples rank by absolute
exact contribution change over selected metrics; row examples similarly rank
by their scoped contribution change. Ties are deterministic for frozen inputs.

Use `metric="Revenue"` and `group={"Region": "North"}` to drill down. Group selectors
must give every `group_by` field, including explicit nulls where applicable, and
exist in at least one recomputed result. Unknown groups are errors, not fabricated
zero-result groups. For `group_by=[]`, `group={}` selects the overall aggregate.
Scoped row counts include keys contributing to that group in either version.
All recipe metrics/checks are still validated even when the response is scoped.
There is no full-history export or sample pagination in v1; use physical references
with existing read tools when more detail is needed. Those later reads are not
automatically pinned to the earlier captured file versions.

For repeatable drilldown, pass `sources.before.sha256`, `sources.after.sha256`, and
`recipe_sha256` as `expected_before_sha256`, `expected_after_sha256`, and
`expected_recipe_sha256`. The recipe digest is the explanation's normalized full
recipe, not a recipe commit token. Without a previous recipe digest, the tool cannot
know whether the caller changed the intended historical definition.

`result_sha256` binds the complete independent and recipe result vectors, not just
the displayed samples. `explanation_token` fingerprints those results, source
digests, normalized recipe, key fields, scope, numeric policy and engine versions.
It also binds scoped metric components and identity/group counts, not the
presentation-only samples.
It excludes sample size. The source/recipe pins do not pin implementation versions
across server upgrades. It is not authentication,
a commit credential or an exactly-once receipt. The service stores no session.

## Capture and limits

Each regular source file is captured under the same-host lock into a verified stable private
copy. The two captures are **not a cross-file atomic transaction**. A later live
edit does not change the captured explanation; use source pins on the next call
to detect it. A file and a symlink/hardlink to that same file are not two versions
and are rejected. Two distinct files with identical bytes are supported.

Per source: 50,000 table cells including header/totals, 64 columns and 10,000
recomputed result cells including headers. Recipe JSON is limited to 32,000 bytes;
up to 64 key fields, 8 sum/count metrics, and a 40,000-character serialized response
budget (including non-ASCII escaping) apply. No partial calculation is returned
when a bound is exceeded. Narrow the response with metric/group/sample options.
Loading a workbook with openpyxl is not a streaming or whole-file memory guarantee.

This is deterministic arithmetic attribution, not a claim that a campaign, market
event or person caused a change. No external AI service is called.
