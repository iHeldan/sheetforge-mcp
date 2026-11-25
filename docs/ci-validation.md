# CI and platform validation

The CI matrix keeps the Ubuntu/Python 3.10, 3.11 and 3.12 jobs and adds one
Windows/Python 3.12 job. This is four jobs, not a full OS-by-Python cross product.
Every job installs the audited lock, runs Ruff and the complete pytest suite,
builds distributions and verifies release artifacts. Windows failures are not
optional; jobs have a ten-minute timeout and pytest reports skip reasons.
Existing Linux check names are retained; Windows has its own distinct check name.

## Filesystem safety coverage

- `test_resource_safety.py` checks thread/process serialization, OS-lock timeout
  and reacquisition, safe-save rollback/recovery and temporary-file cleanup.
  Concurrent writers also use a Unicode filename. Unicode directory/workbook/
  snapshot tests verify readback, byte identity and refusal to overwrite.
- `test_windows_safety.py` checks the one-byte, offset-zero `msvcrt` contract with
  a simulated backend on every platform. Native Windows tests require the real
  backend, prove contention between file handles and check that an open target
  cannot be silently replaced or damaged. Saving must work after it closes.
- `test_path_guidance.py`, `test_server_security.py` and other existing suites
  check rooted paths, traversal, symlink escapes and refusal to auto-relocate.
  Native Windows cases also reject drive-relative and root-relative paths in
  unrooted stdio; Unicode names are tested through actual MCP tool wrappers.
- The lock-directory symlink test runs on Windows too; only the POSIX ownership
  portion of the implementation is platform-specific. File and directory
  symlink creation is required in Windows CI, not silently skipped.

Windows symlinks require Developer Mode or the relevant Windows privilege.
The path-guidance test may skip the specific missing-privilege error (1314)
outside GitHub Actions; it re-raises in CI, as do the native capability tests.
Other errors are not converted into a successful confinement check. Do not
disable path confinement, remove lock checks or use `continue-on-error` to make
an unsupported runner green. See the [Python symlink documentation](https://docs.python.org/3.12/library/os.html#os.symlink)
and [GitHub-hosted runner permissions](https://docs.github.com/en/actions/reference/runners/github-hosted-runners#administrative-privileges).

## Local checks and evidence limits

```bash
uv sync --locked --extra dev
uv run ruff check src tests
uv run pytest -q -ra
uv build
python scripts/verify_release_artifacts.py dist
```

A passing macOS/Linux suite validates its native locking and save behavior, not
Windows `msvcrt`, NTFS sharing/replace semantics or native Windows paths. Those
tests are explicitly skipped on non-Windows hosts. A simulated backend tests
call shape and cleanup only; it is not Windows acceptance. Verify the actual
Windows job at the intended commit before claiming Windows CI has passed.

These checks cover local filesystem operations, not native Excel/Claude Desktop
installation, OneDrive synchronization, network filesystems or distributed locks.
Published packages and existing release tags are not modified by a CI-only patch.
