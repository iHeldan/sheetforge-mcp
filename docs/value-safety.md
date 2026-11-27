# Standard Write Value Safety

Available since `0.15.1`. Applies to cell writes, worksheet append/update, native-table
append/upsert, formula application, range copies, generated pivot cells, and
standard ChangeSet cell operations.

Inputs are validated before saving. The serialized candidate's requested cells
are read back before replacement, even without optional ChangeSet assertions.
If a value would change, the original remains unchanged. This checks requested
cell values, not complete XLSX feature fidelity or calculated formula results.

- `null` clears a cell; booleans retain their type.
- Numbers must be finite and round-trip through the standard writer's numeric
  serialization. Precision-losing integers and decimals fail; they are never
  automatically converted to text. Supply an identifier as a string explicitly
  when its digits must be preserved. Binary tails such as `0.1 + 0.2` may also
  require the caller to provide the intended representable number explicitly.
- Unreleased correction: float identity is checked after decoding the serialized decimal into a float.
  A value such as `71753.65` can have a longer XML spelling yet reload as the
  identical IEEE-754 value; that is accepted, not rounded to a nearby float.
  Integers and explicit Python `Decimal` inputs still require exact decimal
  digits. This does not make binary floating-point amounts exact decimal money.
- Numerically equal values such as `1` and `1.0` may reload as the same integer.
  Boolean `true` is not equal to numeric `1`, including for keyed mutations.
- Text must be nonempty, at most 32767 UTF-16 code units, and valid XML text.
  No truncation is applied. Use `null` rather than an empty string to clear.
- Standard writes reject carriage returns because their XML round-trip can
  normalize them. Supply LF text explicitly, or use preserving mode for CR text
  when that mode's restricted workbook profile supports the edit.
- Python date/time inputs must be timezone-naive and millisecond-aligned.
  Dates may reopen as midnight datetimes. A retained date number format can
  reinterpret a numeric replacement; the candidate check blocks that mismatch.
  Change the number format explicitly if the cell is meant to become numeric.
- Formula strings remain formulas and are not evaluated. Formula text is limited
  to 8192 UTF-16 code units, including the leading `=`. Existing supported rich
  text is retained when repeating its plain text.

Range copies transfer the source style, including the default General style,
and check saved values. Since `0.15.2`, copying also retains and checks the explicit
source cell type: a literal `=A2` or `#N/A` string stays text rather than becoming
a formula or error. Only real formula cells are translated or rewritten by
copy, rename and structural edits. Literal text retains the 32767-unit text
limit, not the shorter formula limit. Copies refuse malformed/out-of-grid targets, intersecting
merged cells and failed relative-formula translation rather than saving an
untranslated formula. The source rectangle is limited to 250,000 cells.
Copy targets cannot intersect array/data-table formula extents. Sheet rename
and copying an array-bearing sheet refuse unsupported non-scalar dependency
rewriting. Full-axis edits refuse unverifiable cross-sheet non-scalar formulas
and local array inputs that would need rewriting, as well as affected extents.

Pivot summaries validate and read back generated labels and aggregates too.
Exact integer results that cannot round-trip are refused, never implicitly
converted to text or float. Computed floating-point aggregates are explicitly
normalized to openpyxl's standard numeric serialization precision (16 significant
digits) before validation and readback. For example, the computed sum of `0.1`
and `0.2` is materialized as `0.3`. This is a declared output policy for generated
floating results, not permission to round exact integer results or caller inputs.

Row-write previews serialize an isolated candidate too. Their new tokens refer
to that simulated state and are marked unpersisted. Committed row-write tokens
describe the serialized, verified candidate, not unsaved Python types. Tokens
remain sampled semantic identities, not full content hashes or locks.

## Response and Retry Safety

Cell and row writes, formula application, range copy/delete, and pivot summaries
check the serialized response budget before replacement.
An oversized detailed response fails without writing; omit `include_changes`
or reduce the batch. Native upsert key lists are included in this budget too.
Post-save filesystem metadata is advisory: if unavailable, the response keeps
the committed tokens and reports `snapshot_metadata_unavailable` instead of
claiming the write failed.

For other ordinary saves, a late response-encoding/size failure returns a compact
success receipt with `persisted=true`, `result_available=false`, `retry_safe=false`
and the verified saved SHA-256. If a subsequent action step itself fails after a
verified save, `write_committed_followup_failed` explicitly reports the persisted
state and unsafe retry rather than claiming full action success. Staging files
do not count as commits to the requested source.

These checks do not guarantee exactly-once delivery. A broken connection or
terminated process after commit can prevent any receipt reaching the caller.
Read back the intended workbook before repeating an uncertain mutation.
