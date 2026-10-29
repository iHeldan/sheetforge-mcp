# Workbook Contracts v1

Implemented on main, not yet published. Published 0.13.0 does not include this API.

A contract is client-owned JSON declaring one native table's logical fields,
approved physical headers, grain and enforceable data rules. It is passed as an
object, never loaded/executed from a path and never stored in hidden workbook
metadata. Keep it separately versioned, with business definitions approved by
the data owner. A contract is not authentication or proof of business truth.

## First workflow

Use the [net-sales example](examples/net-sales.contract.json). It accepts either
`Id` or `OrderId`, and either `Net` or `NetSales`, with exact case-sensitive
matching. Exactly one header per logical field must exist. Having both aliases
is ambiguous, even if their values happen to agree. No preferred-first or fuzzy
fallback exists. Extra source columns are permitted but cannot be referenced by
a contract-backed Recipe unless declared.

1. Call `inspect_workbook_contract(filepath, contract)`. Inspect `status`,
   `unresolved`, `checked`, `declared` and `bindings`.
2. Call `run_workbook_recipe(filepath, recipe, workbook_contract=contract)`.
3. Commit with the identical recipe/contract and preview's `recipe_token` and
   `expected_workbook_sha256`. Source/contract changes require a fresh preview.

For the example contract, a Recipe can be:

```json
{
  "version": 1,
  "source": {"table_name": "Sales"},
  "group_by": ["currency"],
  "metrics": [{"op": "sum", "field": "net", "as": "Revenue"}],
  "output": {"sheet_name": "Report", "range_ref": "A1:B5"}
}
```

Contract filters are mandatory AND predicates, combined with any Recipe filters.
The example excludes `cancelled` orders, but an invalid currency or duplicate key
on a cancelled row still fails: contract checks run over **all source data rows
before filtering**. Native declared totals are excluded; hidden rows and worksheet
filters do not exclude rows. The output uses logical group headers (`currency`)
and unchanged metric aliases (`Revenue`); source/lineage bindings may show physical
headers. Output reservation, overwrite permission and numeric policy remain those
of [Recipes](verified-recipes.md).

## Closed grammar

Required top-level keys: `version: 1`, `name`, `source`, `row_grain`, `key_fields`,
`fields`. Optional: `min_rows` (default 1), `filters` (default []). Unknown keys fail.
`source` is `{table_name, optional sheet_name}` and must exactly equal Recipe source.

Each field requires `name` (logical name), `headers` (approved exact alternatives),
and `type`: `string`, `number`, `integer`, or `boolean`. Optional rules:

- `nullable`: boolean, default false. Allowed nulls skip enum/bound checks.
- `allowed_values`: 1-32 non-null values matching the declared type.
- `min`, `max`: inclusive finite numeric bounds matching the numeric type.
- `meaning`: a short declaration, not an executable condition.

Names, grain and meanings are nonblank strings up to 128 characters. Logical
names and each alias list are unique; aliases cannot be shared between fields.
No trimming, case folding or numeric/text coercion occurs in field binding.
`integer` means a Python integer read from Excel, excluding booleans; `number`
accepts finite integers/floats, not numeric-looking strings. Excel can serialize
`1.0` as `1`, so contracts validate the loaded representation, not its old display.
Composite keys must reference declared fields and be non-null/non-whitespace-blank
and unique over the complete source. Numeric `1` and `1.0` are the same key.

Filters use the [Recipe operators](verified-recipes.md), with logical field names.
String filters default to case-insensitive unless `case_sensitive: true` is set.
The contract's field binding and enum comparisons remain case-sensitive.

## Evidence and failures

Responses use the normal MCP `ok`/`data` envelope. `ok: true` means inspection
completed, **not** that the contract passed. Check `data.passed`/`data.status`:

- `valid`: all bindings resolve and every supported data condition passed.
- `unresolved`: a required header is missing or multiple approved aliases exist;
  no full value checks are claimed.
- `violated`: a supported rule failed; complete counts and sampled physical row
  numbers identify failures. `failed_rows` for keys counts blank/later-duplicate
  rows, not every member of a duplicate set. Examples omit raw cell values.

Malformed contracts, missing tables, chartsheets, mixed source kinds, formulas,
Excel errors, dates and other unsupported referenced cells return tool errors,
not a successful partial scan. All declared fields are referenced, even if a
particular Recipe does not use them. Other fields remain outside contract checks.
`row_grain` and `meaning` remain under `declared`: uniqueness does not prove a row
really is an order, and a `EUR` column does not prove an amount was denominated
correctly. Currency/units cannot be inferred from formatting alone.

`contract_sha256` covers normalized JSON, including declarations and filters.
`source_sha256` identifies the stable captured file. `binding_sha256` additionally
binds engine version, actual table bounds, headers and resolved header cells.
Fingerprints are integrity checks, not approval signatures. The server has no
contract session state and cannot authenticate who authored a definition.

## Persistence and scope

Inspection does not modify source bytes or mtime. A contract-backed Recipe checks
the same stable source before compilation and pins it through existing ChangeSet
persistence. Its generated `table_contract_matches` assertion repeats conditions
and binding identity on the candidate and reopened destination under rollback.
The assertion accepts `contract` and `expected_binding_sha256`; it checks data and
binding, not historical source bytes. Recipes additionally preserve source values
and assert the complete reserved output. Failed checks cannot partially commit.

V1 integrates **Recipes only**. Explanations, Scenarios, Rebase and low-level writes
are not implicitly gated by a prior inspection. There is no coercion, conversion,
formula engine, remote source, join, automatic semantic inference or global policy
enforcement. Same-host locks do not provide a distributed/cloud transaction.

Limits: 32,000 UTF-8 bytes per contract; 32 fields, 8 aliases per field, 8 key
fields, 20 contract predicates; 50,000 source cells including header/totals and
64 physical columns. `sample_limit` 1-5 affects only inspection row examples.
Contract plus Recipe filters share the Recipe 50-filter cap. Inspection's complete
escaped summary is capped at 12K characters; Recipe evidence shares its existing
8K metadata cap. Over-budget calls fail before mutation, even within other limits.
Source checks are complete within these bounds, not inferred from samples. As with
Recipes, openpyxl loads the workbook; table caps do not bound unrelated file content.
