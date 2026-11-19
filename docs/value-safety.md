# Standard Write Value Safety

Unreleased. Applies to cell writes, worksheet append/update, native-table
append/upsert, and standard ChangeSet cell operations.

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
- Formula strings remain formulas and are not evaluated. Existing supported
  rich text is retained when repeating its plain text.

Row-write previews serialize an isolated candidate too. Their new tokens refer
to that simulated state and are marked unpersisted. Committed row-write tokens
describe the serialized, verified candidate, not unsaved Python types. Tokens
remain sampled semantic identities, not full content hashes or locks.
