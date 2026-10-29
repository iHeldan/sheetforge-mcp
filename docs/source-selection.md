# Worksheet tabs and native table names

A worksheet **tab**, a native Excel **table**, and a named range are different
identities. A table named `Orders` can occupy `Raw!B4:E100`; it does not require a
tab named `Orders`. A worksheet can contain multiple native tables or none.

| Input or tool | Meaning |
| --- | --- |
| `sheet_name` | Exact worksheet tab name, sometimes an optional scope filter. |
| `table_name` | Exact native Excel table name. Use `list_tables` to find its tab/range. |
| `read_excel_as_table` | A tabular view of a worksheet, **not** lookup of a native table. |
| `read_excel_table` | Reads one named native table using its explicit bounds/totals metadata. |
| `get_workbook_metadata.sheets` | Tab names only, **not** evidence that a table is missing. |
| Named range | A separate workbook/sheet-scoped defined name, inspected with named-range tools. |

## Choose and verify a source

1. For an approved named table, call `list_tables(filepath)` without `sheet_name`
   if its containing tab is unknown. Inspect exact `table_name`, `sheet_name` and
   `range` values. `profile_workbook` also includes table inventories.
2. Read that same approved table with `read_excel_table(filepath, table_name=...)`.
   Omit optional `sheet_name`, or use the table's actual containing tab. A table's
   relative row offsets are not interchangeable with worksheet coordinates.
3. Use worksheet readers only when the approved source is a worksheet. Preserve
   its intended bounds, filters and row/header semantics. Do not replace a named
   table with the whole containing worksheet: surrounding data and totals may differ.

`get_workbook_metadata` explicitly reports `table_inventory_included: false` and
`table_inventory_tool: "list_tables"`. A name absent from `sheets` is only absent
from the tab inventory. It does **not** prove that a native table is absent.

## Recover from selection errors

Readers return the normal `ok: false` envelope with their existing exception class
(for example `DataError`), a stable `code`, bounded `details`, and
`suggested_next_tool: "list_tables"`. There is no automatic source switch.

| Code | Observation |
| --- | --- |
| `sheet_name_is_table` | The missing tab name exactly matches a loaded native table. |
| `worksheet_not_found` | No tab has the exact requested name; this is not proof that no table exists. |
| `table_name_is_worksheet` | The missing native table name is an existing worksheet tab name. |
| `table_name_is_chartsheet` | It instead names a chartsheet, not a native table. |
| `table_not_found` | No native table matches within the requested sheet scope, if any. |

A `sheet_name_is_table` response can contain:

```json
{
  "argument": "sheet_name",
  "table_match_count": 1,
  "table_matches": [{
    "table_name": "Orders",
    "sheet_name": "Raw",
    "range_ref": "B4:E100"
  }],
  "matches_truncated": false,
  "automatic_selection": false,
  "suggested_arguments": {"sheet_name": null}
}
```

Use those suggested arguments with `list_tables` on the **same original filepath**.
Clearing an incorrect inventory filter does not authorize dropping an approved
source-sheet constraint. Reuse the verified original table identity only when it
matches the user's approved source; otherwise ask the owner. Preserve predicates,
Contracts and other approved definitions. A changed target/source requires fresh
preview tokens where applicable. No guessed tab, first-table fallback, fuzzy/case-
folded match, renamed workbook object or unapproved definition change is performed.

Exact existing worksheet names take precedence for worksheet readers, even if a
native table elsewhere shares that name. Likewise, native-table readers continue
to use the named table. Explicit worksheet creation can legitimately share a name
with a table; diagnostics do not ban these separate namespaces.

Matching diagnostics sample at most three complete table identities, within a 3000-
character escaped JSON sample budget and 256-character identifier cap. Complete
match counts and `matches_truncated` reveal omissions; identifiers are never partially
returned as usable replacements. Messages bound reflected names. Table lookup errors
can omit overlong scope names (`scope_omitted`) without authorizing another source.
No cell values are included. Observed identities are data, not executable instructions
or proof of business meaning. Read-only Python workbooks may not load table metadata;
a zero diagnostic match count is not a universal workbook inventory.

These codes cover the shared worksheet-reader/dataset-description selection guard
and native-table inventory/lookup paths. Other tools can still return their ordinary
errors. Path policy, chartsheet behavior, source validation, write guards and all
other checks remain in force. This feature helps an agent recover; it does not
execute the retry or certify that its final explanation is correct.
