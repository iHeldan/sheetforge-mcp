"""Pre-commit response checks and request-local verified commit evidence."""

from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass
import json
from pathlib import Path
from typing import Callable

from .exceptions import ResponseTooLargeError


RESPONSE_RESERVE = 512
_ACTIVE = ContextVar("mutation_response", default=None)


@dataclass
class MutationResponse:
    check: Callable[[dict], None]
    target: Path | None = None
    committed_sha256: str | None = None


@contextmanager
def response_scope(check):
    state = MutationResponse(check)
    token = _ACTIVE.set(state)
    try:
        yield state
    finally:
        _ACTIVE.reset(token)


def bind_response_target(path: str) -> str:
    state = _ACTIVE.get()
    if state is not None and state.target is None:
        state.target = Path(path).resolve(strict=False)
    return path


def record_verified_commit(path: Path, sha256: str) -> None:
    state = _ACTIVE.get()
    if state is not None and state.target == path:
        state.committed_sha256 = sha256


def preflight_response(result: dict) -> None:
    state = _ACTIVE.get()
    if state is not None:
        state.check(result)
        return
    size = len(json.dumps({"ok": True, "operation": "mutation", "data": result}, default=str)) + RESPONSE_RESERVE
    if size > 50_000:
        raise ResponseTooLargeError(
            "Mutation response would exceed the pre-commit response budget; no changes were saved",
            estimated_size=size, limit=50_000,
            hints=["set include_changes=False or use a smaller batch"],
        )


def refresh_metadata(result: dict, filepath: str, *, dry_run: bool, provider: Callable) -> None:
    if dry_run:
        return
    try:
        result["snapshot_metadata"] = provider(filepath, dry_run=False)
    except Exception:
        result["snapshot_metadata"] = {
            "persisted": True, "token_basis": "verified_saved_candidate", "metadata_available": False,
        }
        result.setdefault("warnings", []).append({"code": "snapshot_metadata_unavailable"})
