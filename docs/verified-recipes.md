# Verified Recipes v1

Available since SheetForge MCP `0.11.0`.

Unreleased on main: optional `workbook_contract` enforces explicit logical bindings
and full bounded source checks in preview/commit. See [Workbook Contracts](workbook-contracts.md).
Use inspection's `recipe_input_fields` (logical binding keys), not physical Excel
headers. Contract field errors provide bounded request-location corrections;
apply only exact input-reference fixes, preserve the approved contract and request
a fresh preview. See [field guidance and safe recovery](workbook-contracts.md#field-guidance-and-safe-recovery).
Without that argument, the existing exact-physical-header contract below is unchanged.

A recipe is a versioned JSON document describing how to filter and aggregate one
native Excel table and materialize its result in a reserved worksheet rectangle.
SheetForge passes intermediate values directly into a verified ChangeSet. The
agent receives a bounded sample, counts, checks, lineage samples and the diff.

## First recipe

The [regional sales example](examples/regional-sales.recipe.json) expects a native
table named `Sales` with `Id`, `Region` and `Amount` headers. It groups nonnegative
amounts by region and writes `Region` and `Revenue` to `Report!A1:B101` in the same
workbook. It deliberately permits replacing **all values in that rectangle**.

Keep the recipe JSON under version control separately from your workbook data.
Send its parsed object as `recipe`; the tool does not execute scripts or load a
recipe path. The workbook path is a separate tool argument, so the same recipe can
be reused for another workbook with the expected table/headers.

1. Call `run_workbook_recipe(filepath, recipe, mode="preview")`.
2. Inspect `ready_to_commit`, `verification`, `source.bindings`, `rows_sample` and `diff`.
3. Repeat the identical recipe with `mode="commit"`, `recipe_token` and
   `expected_workbook_sha256` from the preview.
4. After the source changes, start with a new preview. Old credentials are rejected,
   including after an already-successful commit whose response was lost.

Example using the Python API, with the example document already loaded as `recipe`:

```python
from excel_mcp.recipes import run_workbook_recipe

preview = run_workbook_recipe(filepath, recipe)
if preview["ready_to_commit"]:
    result = run_workbook_recipe(
        filepath,
        recipe,
        mode="commit",
        expected_workbook_sha256=preview["expected_workbook_sha256"],
        recipe_token=preview["recipe_token"],
    )
```

MCP responses wrap this result in the normal `ok`/`data` envelope. A successful
preview with `ready_to_commit=false` has not changed the workbook.

## Recipe contract

| Field | Meaning |
| --- | --- |
| `version` | Integer `1`. |
| `source` | `{table_name, optional sheet_name}` within `filepath`. |
| `filters` | Optional list of AND-combined query filters. |
| `group_by` | List of exact header names; `[]` produces a single overall aggregate. |
| `metrics` | Nonempty list of `{op, field, as}`; `count` omits `field`. |
| `output` | `{sheet_name, range_ref, allow_overwrite?}`; overwrite defaults to false. |
| `checks` | Optional list of explicit source/output checks. |

Unknown properties are rejected. Header names must be exact, including case.
Blank or duplicate source headers (also case/whitespace variants) are rejected.
Column positions are resolved on each run, so a fresh preview can follow reordered
columns. A missing or renamed header requires an explicit recipe change.

Supported metrics: `sum`, `avg`, `min`, `max`, `count`, `count_non_null`,
`count_distinct`. Each metric needs a unique explicit `as` label. Metrics use
`op` and `field`, not the query API's `agg`/`column` aliases. Nulls are omitted by
numeric aggregates; `count` counts matched rows. `group_by=[]` over no matched rows
produces one aggregate row (sum/count zero, average/min/max null); other empty
groupings produce zero rows. Output group ordering is deterministic by serialized
group key, not locale-aware or numeric ranking. No limit silently truncates the
materialized result.

Filter operators: `eq`, `neq` (`ne` alias), `gt`, `gte`, `lt`, `lte`,
`contains`, `starts_with`, `ends_with`, `is_blank`, `not_blank`, `in`, `not_in`.
Use `value` for scalar comparisons, `values` for membership, and neither for blank
checks. Membership also accepts a scalar `value`. String comparisons default to
case-insensitive; `case_sensitive` must be a boolean. Operand types must match
non-null input types. Numeric-looking strings are not coerced into numbers.

Referenced fields must contain homogeneous JSON scalar values, allowing mixed
integer/float numbers and nulls. Formula cells, Excel errors, dates/times and
non-finite numbers in those fields are rejected. Unreferenced fields are not used
as calculation inputs. A native table's declared totals row is excluded. Table
filters and hidden rows do not filter the input; only recipe filters do.

## Output ownership

`range_ref` is an explicitly reserved capacity, including the header row. The
complete rectangle is written: unused rows and columns become blank. Results
that do not fit are rejected before mutation. Choose a larger range and preview
again when growth is intentional.

Missing output worksheets are created. Existing values require
`allow_overwrite=true`, even if a previous recipe wrote them. This flag authorizes
the entire rectangle, including values placed there by a person. No hidden marker
asserts ownership. Cells outside it remain outside the recipe's write operations.
Output cannot overlap native tables or merged cells or target a protected worksheet
or chartsheet. Existing styles are retained. The output is ordinary worksheet
cells, not a newly created/resized native Excel table. Formula-like output strings
and Excel error literals are rejected rather than converted into formulas/errors.

## Checks, probes and evidence

- `unique_key`: `fields` identifies a composite key checked over **all source rows**
  before filtering. Null/empty-string keys and duplicates fail.
- `row_count`: `min` and `max` bound the number of output data rows, excluding headers.
- `sum_matches`: compares `source_field` over matched rows with `output_field` over
  all output groups. Both must be numeric; null values are excluded.

Every run also re-executes aggregation over in-memory copies with columns reversed
and rows reversed. The normalized results must agree. These probes cover the full
bounded input/result, not a sample. Integer comparisons are exact; floating-point
comparisons use relative tolerance `1e-12` and absolute tolerance `1e-9`.
Floating-point metric outputs are normalized to openpyxl's numeric serialization
before checks and preview, so binary tails such as `0.30000000000000004` do not
disagree with the saved `0.3`. The ChangeSet still asserts the exact normalized
values after reopening. This is ordinary floating-point arithmetic, not exact
decimal accounting or a guarantee of precision beyond Excel's number model.

These probes exercise field binding and order sensitivity. They are not automatic
Excel calculation, a proof of the business definition, or a substitute for choosing
the right metric (for example, a weighted rate instead of an average of rates).
They are also not an independent arithmetic oracle: two input orders can share
the same error. Floating-point sums and averages use compensated summation, and
the acceptance workflow separately checks them against a Decimal calculation.
Overflow is rejected rather than returning an infinite result.

`source.bindings` identifies the actual header cells. `source_sha256` identifies
the exact input workbook. `verification.lineage_sample` identifies groups, their
full contributing row counts, and bounded **physical worksheet row numbers**.
Truncation flags describe samples only; the calculations and checks are complete
within the stated limits. This version does not persist a row-level history or
provide a cross-run "why did this value change" debugger itself. Since `0.12.0`,
[Verified Explanations](verified-explanations.md) provides a separate read-only
comparison of two source versions under one shared sum/count recipe.

The generated ChangeSet verifies the complete reserved output values and confirms
the complete source table's values/formula text are unchanged. It provides a
structural diff and sampled value changes, not an exhaustive presentation diff.

## Persistence and limits

Compilation reads an isolated, stable byte copy. Tokens bind the recipe, recipe
engine version, canonical workbook path, source digest and compiled operations.
ChangeSet independently checks that same source digest before preparing and
committing the result. Every recipe commit uses a verified snapshot, same-host
locking, atomic replacement and post-replace verification with rollback.

The service holds no recipe session state. Tokens are integrity fingerprints,
not authentication. Same-host locks do not coordinate other computers or guarantee
cloud-sync consistency. Replaying a recipe on newer data always requires a fresh
preview. There is no cross-workbook transaction or exactly-once retry service.

| Bound | v1 limit |
| --- | --- |
| Recipe JSON | 32,000 UTF-8 bytes |
| Source table | 50,000 cells including header/totals, 64 columns |
| Reserved output | 10,000 cells including header and blank padding |
| Group fields / metrics | At most 64 each |
| Filters / checks | At most 50 / 20 |
| `sample_limit` | 1-10, default 5 |
| Recipe response additions | 8,000 serialized characters before commit |

ChangeSet's 2 MB compiled-plan and 40,000-character pre-commit response budgets
also apply. Counts include JSON escaping so non-ASCII text cannot bypass the wire
budget. Exceeding a bound fails before replacement; no partial result is committed.
The implementation loads the bounded workbook with openpyxl; these table limits
do not imply streaming or a memory bound for unrelated content in a very large file.

## Reproducible workflow validation

`tests/test_recipe_workflows.py` drives the real stdio MCP protocol with a
synthetic sales workbook and checks the complete saved report against an
independent Decimal oracle. It runs the same recipe at 1,000 source rows, after
growth to 1,200 rows with reordered columns, and after shrinking to one region.
It verifies source values/formulas, totals, snapshot bytes, unrelated report
cells, header styling and chart references/layout remain intact.

The scripted client uses 10 tool calls across those three runs, including initial
orientation and a complete output readback after each commit. The test enforces
the per-response 50,000-character limit and a combined 100,000-character bound.
It does not read the full source through MCP. These are reproducible synthetic
workflow checks, not a production benchmark, customer acceptance, or proof that
an unassisted model will always choose the right recipe.
