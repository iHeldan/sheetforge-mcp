# Workbook paths and safe recovery

Paths refer to the **server's filesystem**, not the MCP client's working directory.
Every workbook-path parameter advertises the same rule in its MCP schema, including
source/destination paths, optional snapshot paths and multi-workbook list items.
The short field hint remains self-contained: reuse the supplied server path, never
prefix client cwd, allow relative paths on rooted servers, and require an absolute
path (ask if missing) in unrooted stdio. Server initialization supplies the full
policy once, rather than repeating it in every field. Structured error recovery and
all path guards remain unchanged; shortening advisory text is not a new permission.

## Choose a path

- Reuse the path supplied for the task. **Never prepend the client's cwd** to a
  relative workbook path or guess another absolute prefix.
- Normal `stdio` has no configured file root and requires an absolute server-side
  path. If only a relative name is supplied, ask the owner for the absolute path.
  The server does not discover a user's intended workbook from its basename.
- Rooted servers (`streamable-http`/`sse`, using `EXCEL_FILES_PATH`) accept relative
  paths under that root. For example, use `reports/monthly.xlsx` unchanged, rather
  than constructing `/client/project/reports/monthly.xlsx`.
- An absolute path on a rooted server must also remain inside its canonical root.
  Parent-directory and symlink escapes are rejected. A valid path can refer to a
  new file for creation tools; a path-policy pass does not prove a file exists.
- An embedding that explicitly configures a root uses rooted policy even over
  stdio. Merely setting the environment variable does not enable a root in normal
  `sheetforge-mcp stdio`; its absolute-path policy is unchanged.

## Recover from a path-policy error

Path-policy failures use the ordinary `ok: false` JSON envelope with error type
`WorkbookPathError`, a stable `code`, `details`, and `suggested_next_tool` naming
the **same failed tool**, not another way to access the rejected file.

| Code | Next step |
| --- | --- |
| `filepath_invalid` | Ask for a nonblank string identifying the intended workbook. |
| `filepath_absolute_required` | Obtain the intended server-side absolute path; do not manufacture one from client cwd. |
| `filepath_outside_root` | If the task originally supplied an approved root-relative path, resend that original path unchanged. Otherwise ask for an approved path. |

A rooted rejection includes advice of this shape:

```json
{
  "ok": false,
  "operation": "profile_workbook",
  "error": {
    "type": "WorkbookPathError",
    "code": "filepath_outside_root",
    "suggested_next_tool": "profile_workbook",
    "details": {
      "path_mode": "rooted",
      "relative_to": "EXCEL_FILES_PATH",
      "action": "reuse_supplied_relative_path_or_ask",
      "automatic_relocation": false
    }
  }
}
```

The actual response also contains a message and recovery instruction. It does not
echo the rejected path or the configured root directory. No replacement file is
guessed, searched for, selected or opened by the error-recovery code.

Do not strip a rejected path to its basename: two different directories may contain
different workbooks with the same name. Do not retry the same rejected path across
multiple tools or disable confinement. If the original path itself escapes the root
or the intended target is unclear, ask the owner rather than modifying that path.
Apply these rules separately to every source, destination and snapshot argument.

After obtaining the intended path, retry the same tool. Preserve the task's other
approved arguments. For token-guarded writes, a **different workbook target needs a
fresh preview and tokens**; path advice does not authorize reusing a preview from
another file, overwriting a file or weakening Workbook Contracts. Existing file,
permission, format and concurrency checks still apply after path validation.

Direct Python calls to `get_excel_path` raise `WorkbookPathError`, which also remains
a `ValueError` and `ValidationError` for compatibility. MCP clients should use
`error.code` for path-policy recovery. Malformed MCP argument types can instead be
rejected by SDK validation before the function runs; unrelated filesystem errors
are not reclassified as path-policy failures. These rules are guidance and local
confinement, not authentication, filesystem-race isolation or a distributed sandbox.
