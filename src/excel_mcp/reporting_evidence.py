"""Shared bounded, read-only capture and exact identities for reporting probes."""

from contextlib import contextmanager
import hashlib
import json

from . import changeset, recipes, workbook
from .exceptions import ResponseTooLargeError, ValidationError
from .explanations import _identity, _text, _number


def digest(value):
    return hashlib.sha256(changeset._canonical_json(value).encode("utf-8")).hexdigest()


def bounded(value, limit=12_000):
    size = len(
        json.dumps(value, ensure_ascii=True, allow_nan=False, separators=(",", ":"))
    )
    if size > limit:
        raise ResponseTooLargeError(
            "Reporting evidence exceeds its response budget",
            estimated_size=size,
            limit=limit,
            hints=["Reduce sample_limit or the requested definition/output size"],
        )
    return value


def sample_size(value):
    if type(value) is not int or not 1 <= value <= 5:
        raise ValidationError("sample_limit must be an integer from 1 to 5")
    return value


def load_fields(wb, selector, fields):
    selector = changeset._json_copy(selector, location="source")
    recipes._object(
        selector,
        allowed={"table_name", "sheet_name"},
        required={"table_name"},
        location="source",
    )
    for value in selector.values():
        recipes._name(value, "source name")
    return recipes._load_table(
        wb,
        {
            "source": selector,
            "group_by": list(dict.fromkeys(fields)),
            "metrics": [],
            "filters": [],
            "checks": [],
        },
    )


@contextmanager
def capture(paths):
    """Each input is a stable local snapshot, not a cross-file transaction."""
    with recipes._recipe_workspace() as folder:
        captures = {}
        ordered = []
        for requested in paths:
            path = workbook._canonical_workbook_path(requested, must_exist=True)
            if path.suffix.lower() != ".xlsx":
                raise ValidationError("filepath must end with .xlsx")
            if path not in captures:
                baseline = folder / f"source-{len(captures)}.xlsx"
                with workbook._exclusive_workbook_lock(path):
                    sha = changeset._stage_workbook(path, baseline, None)
                captures[path] = (baseline, sha)
            ordered.append(captures[path])
        yield ordered


def identity(value):
    return _identity(value)


def number(value):
    return _number(value)


def text(value):
    return _text(value)
