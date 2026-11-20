# XLSX Input Limits

Available since `0.15.1`. Ordinary workbook loads inspect ZIP metadata before openpyxl
allocates archive-entry objects or expands XML. The same checked file descriptor
is used for loading, avoiding a separate unchecked path reopen. Read-only
workbooks retain that descriptor until closed.

| Server environment setting | Default |
| --- | --- |
| `SHEETFORGE_XLSX_MAX_FILE_BYTES` | 134217728 (128 MiB compressed file) |
| `SHEETFORGE_XLSX_MAX_TOTAL_BYTES` | 268435456 (256 MiB total uncompressed members) |
| `SHEETFORGE_XLSX_MAX_MEMBER_BYTES` | 67108864 (64 MiB per uncompressed member) |
| `SHEETFORGE_XLSX_MAX_DIRECTORY_BYTES` | 16777216 (16 MiB ZIP directory) |
| `SHEETFORGE_XLSX_MAX_MEMBERS` | 20000 |

Values must be positive decimal integers, at most `2^63 - 1`. Invalid settings
fail closed; zero does not disable a limit. Settings belong to the trusted
server operator, not tool arguments. Review a trusted large workbook before
raising the relevant limits; a client should not guess settings or bypass a
refusal. Existing path confinement remains enforced independently.

An input refusal reports `xlsx_input_rejected` with a bounded reason and, for
size/count failures, observed and allowed values. It does not disclose the server
root or workbook cell contents. ZIP64 directory metadata is supported within
the same budgets. Duplicate members, unsafe member paths, encrypted entries,
unsupported compression (anything except stored/deflated), multidisk archives
and inconsistent directory metadata are refused before expansion. Invalid ZIP
variants that the Python ZIP reader cannot handle are refused too.

Checks apply to source loads and serialized candidates. ChangeSet assertions
and ordinary save verification use the checked loader. Preserving mode also
checks its initial package, then applies its own stricter profile and budgets;
raising these general limits does not override preserving-mode restrictions.

`max_rows` and pagination bound returned data, not the work required to load an
XLSX shared-string table. These input budgets provide a separate early boundary.
They are byte/count limits, not a hard process-RAM, CPU-time, concurrency or
expanded-XML-entity limit. Dense workbooks below them may still be expensive.
Use an appropriately isolated server process for untrusted workloads.
