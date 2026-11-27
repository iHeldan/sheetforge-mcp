"""Read-only multiplicity and measure-conservation evidence for explicit joins."""

from collections import Counter
from fractions import Fraction

from . import recipes, workbook
from . import reporting_evidence as evidence
from .exceptions import ValidationError


def _key(value):
    # Blank keys never match; text is exact (no trimming or case folding).
    return (
        None
        if value is None or isinstance(value, str) and not value.strip()
        else evidence.identity(value)
    )


def rehearse_table_join(
    left_filepath: str,
    right_filepath: str,
    left_source: dict,
    right_source: dict,
    left_key: str,
    right_key: str,
    measure_field: str,
    join_type: str = "left",
    sample_limit: int = 3,
) -> dict:
    evidence.sample_size(sample_limit)
    if join_type not in ("left", "inner"):
        raise ValidationError("join_type must be left or inner")
    for name in (left_key, right_key, measure_field):
        recipes._name(name, "field")
    with evidence.capture([left_filepath, right_filepath]) as captures:
        loaded = []
        for (path, sha), selector, fields in zip(
            captures,
            (left_source, right_source),
            ([left_key, measure_field], [right_key]),
        ):
            with workbook.safe_workbook(str(path)) as wb:
                source = evidence.load_fields(wb, selector, fields)
            loaded.append((source, sha))
        (left, left_sha), (right, right_sha) = loaded
        li, ri = left["headers"].index(left_key), right["headers"].index(right_key)
        mi = left["headers"].index(measure_field)
        right_counts = Counter(
            _key(row[ri]) for row in right["rows"] if _key(row[ri]) is not None
        )
        left_counts = Counter(
            _key(row[li]) for row in left["rows"] if _key(row[li]) is not None
        )
        right_examples = {}
        for offset, row in enumerate(right["rows"]):
            key = _key(row[ri])
            if key is not None:
                members = right_examples.setdefault(key, [])
                if len(members) < sample_limit:
                    members.append(right["data_start_row"] + offset)
        original = projected = extra = lost = extra_abs = lost_abs = Fraction(0)
        output_rows = unmatched = duplicated = removed = null_measures = 0
        samples = {"multiplied": [], "dropped": [], "unmatched": []}
        for offset, row in enumerate(left["rows"]):
            matches = right_counts.get(_key(row[li]), 0)
            multiplier = matches if join_type == "inner" else max(1, matches)
            value = evidence.number(row[mi])
            null_measures += row[mi] is None
            original += value
            projected += multiplier * value
            output_rows += multiplier
            kinds = []
            if not matches:
                unmatched += 1
                kinds.append("unmatched")
            if multiplier > 1:
                duplicated += 1
                extra += (multiplier - 1) * value
                extra_abs += (multiplier - 1) * abs(value)
                kinds.append("multiplied")
            if not multiplier:
                removed += 1
                lost += value
                lost_abs += abs(value)
                kinds.append("dropped")
            for kind in kinds:
                if len(samples[kind]) < sample_limit:
                    samples[kind].append(
                        {
                            "left_row": left["data_start_row"] + offset,
                            "right_match_count": matches,
                            "output_copies": multiplier,
                            "right_rows": right_examples.get(_key(row[li]), []),
                            "right_rows_truncated": matches > sample_limit,
                            "measure": evidence.text(value),
                            "contribution_delta": evidence.text(
                                (multiplier - 1) * value
                            ),
                        }
                    )
        return evidence.bounded(
            {
                "version": 1,
                "read_only": True,
                "join_type": join_type,
                "left_source_sha256": left_sha,
                "right_source_sha256": right_sha,
                "left_source": {
                    k: left[k]
                    for k in (
                        "sheet_name",
                        "table_name",
                        "range_ref",
                        "totals_rows_excluded",
                    )
                },
                "right_source": {
                    k: right[k]
                    for k in (
                        "sheet_name",
                        "table_name",
                        "range_ref",
                        "totals_rows_excluded",
                    )
                },
                "left_rows": len(left["rows"]),
                "right_rows": len(right["rows"]),
                "projected_rows": output_rows,
                "unmatched_left_rows": unmatched,
                "multiplied_left_rows": duplicated,
                "dropped_left_rows": removed,
                "left_duplicate_key_count": sum(n > 1 for n in left_counts.values()),
                "right_duplicate_key_count": sum(n > 1 for n in right_counts.values()),
                "null_measure_rows": null_measures,
                "measure": {
                    "field": measure_field,
                    "before": evidence.text(original),
                    "after": evidence.text(projected),
                    "delta": evidence.text(projected - original),
                    "extra_signed": evidence.text(extra),
                    "lost_signed": evidence.text(lost),
                    "extra_absolute": evidence.text(extra_abs),
                    "lost_absolute": evidence.text(lost_abs),
                },
                "samples": samples,
                "samples_truncated": {
                    k: n > sample_limit
                    for k, n in (
                        ("multiplied", duplicated),
                        ("dropped", removed),
                        ("unmatched", unmatched),
                    )
                },
                "requires_decision": bool(duplicated or removed or unmatched),
                "semantics": "all-match join; blank keys never match; exact typed keys; no coercion or case folding",
                "limits": "50000 cells/64 columns per table; stable independent captures, not a cross-file transaction",
                "next_action": "Review multiplicity and missing matches against the intended row grain; no join or allocation has been selected or applied",
            }
        )
