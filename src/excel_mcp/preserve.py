"""Narrow SpreadsheetML cell patches for the preserving ChangeSet writer."""

from __future__ import annotations

import ast
from collections import defaultdict, deque
from contextlib import nullcontext
from copy import copy
from dataclasses import dataclass
import hashlib
import math
from pathlib import Path
import re
import stat
from xml.parsers import expat
from xml.sax.saxutils import escape
from zipfile import BadZipFile, ZIP_DEFLATED, ZIP_STORED, ZipFile
from zlib import error as ZlibError

from openpyxl import load_workbook
from openpyxl.cell.cell import ERROR_CODES
from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter, range_boundaries

from .exceptions import PreconditionFailedError
from .package_fidelity import (
    _Archive,
    _CELL_REF,
    _OFFICE_REL,
    _SHEET_NS,
    _Unverified,
    _inventory,
)

PRESERVE_VERSION = 1
MAX_PACKAGE_BYTES = 64 * 1024 * 1024
MAX_SOURCE_BYTES = MAX_PACKAGE_BYTES + 4 * 1024 * 1024
MAX_SHEET_BYTES = 16 * 1024 * 1024
MAX_EDITS = 1_000
_REL_KINDS = {
    _OFFICE_REL + "/" + name
    for name in (
        "officeDocument",
        "worksheet",
        "styles",
        "theme",
        "sharedStrings",
        "calcChain",
        "customXml",
        "customXmlProps",
        "extended-properties",
        "custom-properties",
    )
} | {
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
}
_PART_TYPES = {
    "application/vnd.openxmlformats-package.relationships+xml",
    "application/vnd.openxmlformats-package.core-properties+xml",
    "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    "application/vnd.openxmlformats-officedocument.custom-properties+xml",
    "application/vnd.openxmlformats-officedocument.theme+xml",
    *(
        "application/vnd.openxmlformats-officedocument.spreadsheetml." + name
        for name in (
            "sheet.main+xml",
            "worksheet+xml",
            "styles+xml",
            "sharedStrings+xml",
            "calcChain+xml",
        )
    ),
}
_BOOK_ELEMENTS = {
    "fileVersion",
    "workbookPr",
    "workbookProtection",
    "bookViews",
    "sheets",
    "definedNames",
    "calcPr",
}
_SHEET_ELEMENTS = {
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
}
_BOOK_TAGS = {_SHEET_NS + tag for tag in _BOOK_ELEMENTS}
_SHEET_TAGS = {_SHEET_NS + tag for tag in _SHEET_ELEMENTS}
_TEXT_ESCAPE = re.compile(r"_[xX][0-9a-fA-F]{4}_")
_CONSTANT_CHARS = re.compile(r"[0-9eE.+*/() \t\r\n-]{1,256}", re.ASCII)
_ROW = (_SHEET_NS + "worksheet", _SHEET_NS + "sheetData", _SHEET_NS + "row")
_CELL = (*_ROW, _SHEET_NS + "c")
_FORMULA = (*_CELL, _SHEET_NS + "f")
_MAIN_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
)
_PARSED_PART_ROOTS = {
    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml": _SHEET_NS
    + "styleSheet",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml": _SHEET_NS
    + "sst",
}


def _unsupported(reason: str) -> None:
    raise PreconditionFailedError(
        "Preserving save could not verify this edit; the original workbook was not modified",
        code="changeset_preserve_unsupported",
        details={"save_mode": "preserve", "reason": reason},
    )


def check_source_size(size: int) -> None:
    if size > MAX_SOURCE_BYTES:
        _unsupported("package_limit")


def _constant_formula(formula: str) -> bool:
    if not _CONSTANT_CHARS.fullmatch(formula):
        return False
    try:
        tree = ast.parse(formula.strip(), mode="eval")
    except (SyntaxError, ValueError):
        return False
    allowed = (
        ast.Expression,
        ast.Constant,
        ast.BinOp,
        ast.UnaryOp,
        ast.Add,
        ast.Sub,
        ast.Mult,
        ast.Div,
        ast.UAdd,
        ast.USub,
    )
    return all(
        isinstance(node, allowed)
        and (not isinstance(node, ast.Constant) or type(node.value) in {int, float})
        for node in ast.walk(tree)
    )


def _literal(value):
    if value is None:
        return "", ""
    if type(value) is bool:
        return ' t="b"', f"<v>{int(value)}</v>"
    if type(value) is int and abs(value) <= 999_999_999_999_999:
        return ' t="n"', f"<v>{value}</v>"
    if type(value) is float and math.isfinite(value):
        number = format(value, ".15g")
        if float(number) == value and (
            value == 0 or 2.2250738585072014e-308 <= abs(value) <= 9.99999999999999e307
        ):
            return ' t="n"', f"<v>{number}</v>"
    if type(value) is str:
        if (
            not value
            or value.startswith("=")
            or value in ERROR_CODES
            or _TEXT_ESCAPE.search(value)
            or any(
                (ord(ch) < 32 and ch not in "\t\n\r")
                or 0xD800 <= ord(ch) <= 0xDFFF
                or ord(ch) in {0xFFFE, 0xFFFF}
                for ch in value
            )
            or len(value.encode("utf-16-le")) > 65_534
        ):
            _unsupported("unsupported_literal")
        text = escape(value).replace("\r", "&#13;")
        return ' t="inlineStr"', f'<is><t xml:space="preserve">{text}</t></is>'
    _unsupported("unsupported_literal")


def _equal(left, right) -> bool:
    if isinstance(left, bool) or isinstance(right, bool):
        return type(left) is type(right) and left == right
    return left == right


@dataclass(frozen=True)
class _Edit:
    coord: str
    before: object
    after: object
    kind: str
    body: str

    @property
    def changed(self):
        return not _equal(self.before, self.after)

    def xml(self, style=None):
        style_attr = f' s="{style}"' if style is not None else ""
        return (
            f'<c xmlns="{_SHEET_NS[:-1]}" r="{self.coord}"{style_attr}{self.kind}>'
            f"{self.body}</c>"
        ).encode("utf-8")


def _package_profile(archive: _Archive):
    entries = archive.archive.infolist()
    if sum(info.file_size for info in entries) > MAX_PACKAGE_BYTES:
        _unsupported("package_limit")
    seen = set()
    for entry in entries:
        name = entry.filename.rstrip("/")
        if (
            not name
            or name in seen
            or any(p in {"", ".", ".."} for p in name.split("/"))
            or any(ch in name for ch in "\\%?#:")
            or any(ord(ch) < 32 for ch in name)
            or entry.flag_bits & 1
            or stat.S_ISLNK(entry.external_attr >> 16)
            or entry.compress_type not in {ZIP_STORED, ZIP_DEFLATED}
            or (entry.is_dir() and entry.file_size)
        ):
            _unsupported("unsupported_zip_member")
        seen.add(name)

    inventory = _inventory(archive, strict_types=True)
    if (
        not inventory.sheets
        or not inventory.custom_parts <= archive.files
        or {p for p, t in inventory.content_types.items() if t == _MAIN_TYPE}
        != {inventory.workbook_part}
    ):
        _unsupported("unsupported_package_layout")
    for part, content_type in inventory.content_types.items():
        if (
            part not in {"[Content_Types].xml", "_rels/.rels"}
            and content_type not in _PART_TYPES
            and not (
                part in inventory.custom_parts
                and content_type
                in {
                    "application/xml",
                    "text/xml",
                    "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
                }
            )
        ):
            _unsupported("unsupported_package_part")
    for link in inventory.relationships:
        if (
            link.kind not in _REL_KINDS
            or link.mode != "Internal"
            or link.part() not in archive.files
        ):
            _unsupported("unsupported_package_relationship")
    if set(inventory.sheets.values()) != {
        part
        for part, kind in inventory.content_types.items()
        if kind
        == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    } or len(set(inventory.sheets.values())) != len(inventory.sheets):
        _unsupported("unsupported_package_layout")

    seen_book_tags = set()

    def workbook_xml(path, attrs):
        if (
            (
                len(path) == 2
                and (path[-1] not in _BOOK_TAGS or path[-1] in seen_book_tags)
            )
            or (len(path) >= 3 and path[1] == _SHEET_NS + "definedNames")
            or (
                len(path) >= 2
                and path[1] == _SHEET_NS + "workbookProtection"
                and (len(path) > 2 or attrs)
            )
            or path[-1].rsplit("}", 1)[-1] in {"AlternateContent", "extLst"}
        ):
            _unsupported("unsupported_workbook_feature")
        if len(path) == 2:
            seen_book_tags.add(path[-1])

    archive.xml(inventory.workbook_part, _SHEET_NS + "workbook", workbook_xml)
    for part, content_type in inventory.content_types.items():
        if content_type in _PARSED_PART_ROOTS:
            archive.xml(part, _PARSED_PART_ROOTS[content_type], lambda *_: None)
    return inventory


def _collect_edits(book, operations):
    edits = defaultdict(dict)
    results = []
    changes = []
    target_count = 0
    for index, operation in enumerate(operations, 1):
        args = operation["args"]
        selected = args.get("sheet_name")
        sheet = (
            book[selected]
            if selected in book.sheetnames
            else book.active
            if not selected
            else None
        )
        if sheet is None or sheet.title not in book.sheetnames:
            _unsupported("target_sheet_missing")
        start_cell = args.get("start_cell", "A1")
        if not isinstance(start_cell, str) or not _CELL_REF.fullmatch(
            start_cell.upper()
        ):
            _unsupported("unsupported_target_cell")
        row, col = coordinate_to_tuple(start_cell.upper())
        count = changed = 0
        for r, values in enumerate(args["data"], row):
            for c, value in enumerate(values, col):
                target_count += 1
                if target_count > MAX_EDITS:
                    _unsupported("edit_limit")
                if r > 1_048_576 or c > 16_384:
                    _unsupported("unsupported_target_cell")
                coord = f"{get_column_letter(c)}{r}"
                if (r, c) in edits[sheet.title]:
                    _unsupported("overlapping_edits")
                cell = sheet[coord]
                old = cell.value
                if type(old) not in {
                    type(None),
                    int,
                    float,
                    bool,
                    str,
                } or cell.data_type in {"f", "e"}:
                    _unsupported("unsupported_target_cell_value")
                kind, body = _literal(value)
                edit = _Edit(coord, old, value, kind, body)
                edits[sheet.title][r, c] = edit
                count += 1
                if edit.changed:
                    changed += 1
                    changes.append(
                        {
                            "sheet_name": sheet.title,
                            "cell": coord,
                            "before": old,
                            "after": value,
                        }
                    )
        results.append(
            {
                "index": index,
                "tool": operation["tool"],
                "result": {
                    "sheet_name": sheet.title,
                    "cells_written": count,
                    "changed_cells": changed,
                },
            }
        )
    if not target_count:
        _unsupported("no_target_cells")
    return edits, results, changes


def _tag_end(xml, start):
    if xml[start : start + 1] != b"<":
        _unsupported("unsupported_xml_span")
    quote = 0
    for offset in range(start + 1, min(len(xml), start + 8192)):
        byte = xml[offset]
        if quote:
            if byte == quote:
                quote = 0
        elif byte in (34, 39):
            quote = byte
        elif byte == 62:
            return offset + 1
    _unsupported("unsupported_xml_span")


def _patch_sheet(xml: bytes, edits):
    """Locate exact cell spans with a namespace-aware parser; never serialize the tree."""
    if b"\0" in xml:
        _unsupported("unsupported_sheet_encoding")
    parser = expat.ParserCreate("UTF-8", namespace_separator="}")
    tags = []
    seen_sheet_tags = set()
    per_row = defaultdict(list)
    for (row, col), edit in sorted(edits.items()):
        per_row[row].append((col, edit))
    pending = deque()
    patches = []
    dimension = row_spans = None
    row_start = current_row = previous_col = 0
    target = cell_start = style = None
    formula_text = []
    formula_length = formula_count = 0

    def insert_at(index, col, edit):
        if dimension is None:
            _unsupported("insertion_outside_sheet_dimension")
        c1, r1, c2, r2 = dimension
        if not (c1 <= col <= c2 and r1 <= current_row <= r2):
            _unsupported("insertion_outside_sheet_dimension")
        if row_spans and not any(lo <= col <= hi for lo, hi in row_spans):
            _unsupported("insertion_outside_row_span")
        if edit.changed:
            patches.append((index, index, edit.xml()))

    def start(tag, attrs):
        nonlocal dimension, row_start, current_row, previous_col, pending, row_spans
        nonlocal target, cell_start, style, formula_length
        if len(tags) >= 128:
            _unsupported("xml_depth_limit")
        tags.append(tag)
        depth = len(tags)
        if depth == 1 and tag != _ROW[0]:
            _unsupported("unsupported_worksheet")
        if (
            depth == 2
            and (tag not in _SHEET_TAGS or tag in seen_sheet_tags)
            or tag.rsplit("}", 1)[-1] in {"AlternateContent", "extLst"}
        ):
            _unsupported("unsupported_worksheet_feature")
        if depth == 2:
            seen_sheet_tags.add(tag)
        if depth == 2 and tag == _SHEET_NS + "dimension":
            dimension = range_boundaries(attrs["ref"])
            if not all(dimension):
                _unsupported("unsupported_sheet_dimension")
        if tuple(tags[:2]) == _ROW[:2] and depth in {3, 4}:
            if tuple(tags) == _ROW:
                row = attrs.get("r", "")
                if (
                    not re.fullmatch(r"[1-9][0-9]{0,6}", row)
                    or not current_row < int(row) <= 1_048_576
                ):
                    _unsupported("ambiguous_row")
                current_row = int(row)
                row_start = parser.CurrentByteIndex
                previous_col = 0
                pending = deque(per_row.pop(current_row, []))
                row_spans = None
                if pending and "spans" in attrs:
                    spans = attrs["spans"]
                    if not re.fullmatch(
                        r"[0-9]{1,5}:[0-9]{1,5}( [0-9]{1,5}:[0-9]{1,5})*", spans
                    ):
                        _unsupported("unsupported_row_span")
                    row_spans = [
                        tuple(map(int, span.split(":"))) for span in spans.split()
                    ]
            elif tuple(tags) == _CELL:
                ref = attrs.get("r", "")
                if not _CELL_REF.fullmatch(ref) or "cm" in attrs or "vm" in attrs:
                    _unsupported("unsupported_cell_layout")
                r, col = coordinate_to_tuple(ref)
                if r != current_row or not previous_col < col <= 16_384:
                    _unsupported("ambiguous_cell")
                previous_col = col
                cell_start = parser.CurrentByteIndex
                target = None
                while pending and pending[0][0] < col:
                    insert_at(cell_start, *pending.popleft())
                if pending and pending[0][0] == col:
                    target = pending.popleft()[1]
                    style = attrs.get("s")
                    if (
                        set(attrs) - {"r", "s", "t"}
                        or attrs.get("t", "n") not in {"n", "b", "inlineStr"}
                        or (
                            style is not None and not re.fullmatch(r"[0-9]{1,9}", style)
                        )
                    ):
                        _unsupported("unsupported_target_cell_encoding")
            else:
                _unsupported("unsupported_sheet_data")
        if tag == _FORMULA[-1]:
            if tuple(tags) != _FORMULA or attrs or target is not None:
                _unsupported("unsupported_formula")
            formula_text.clear()
            formula_length = 0
        elif depth > 5 and tuple(tags[:5]) == _FORMULA:
            _unsupported("unsupported_formula")
        if (
            depth == 5
            and tuple(tags[:4]) == _CELL
            and tag
            not in {
                _SHEET_NS + "v",
                _SHEET_NS + "is",
                _FORMULA[-1],
            }
        ):
            _unsupported("unsupported_cell_layout")
        if target is not None and tuple(tags[:4]) == _CELL and depth > 4:
            if tuple(tags[4:]) not in {
                (_SHEET_NS + "v",),
                (_SHEET_NS + "is",),
                (_SHEET_NS + "is", _SHEET_NS + "t"),
            }:
                _unsupported("unsupported_target_cell_encoding")

    def characters(text):
        nonlocal formula_length
        if len(tags) == 5 and tuple(tags) == _FORMULA:
            formula_length += len(text)
            if formula_length > 256:
                _unsupported("unsupported_formula")
            formula_text.append(text)

    def end(_tag):
        nonlocal target, formula_count
        path = tuple(tags)
        index = parser.CurrentByteIndex
        if path == _FORMULA:
            if not _constant_formula("".join(formula_text)):
                _unsupported("unsupported_formula")
            formula_count += 1
        elif path == _CELL and target is not None:
            if target.changed:
                opening_end = _tag_end(xml, cell_start)
                end_at = opening_end
                if xml[opening_end - 2 : opening_end] != b"/>":
                    if index < opening_end or xml[index : index + 2] != b"</":
                        _unsupported("unsupported_xml_span")
                    end_at = _tag_end(xml, index)
                patches.append((cell_start, end_at, target.xml(style)))
            target = None
        elif path == _ROW and pending:
            opening_end = _tag_end(xml, row_start)
            if xml[opening_end - 2 : opening_end] == b"/>":
                _unsupported("insertion_requires_existing_row_body")
            if index < opening_end or xml[index : index + 2] != b"</":
                _unsupported("unsupported_xml_span")
            for col, edit in pending:
                insert_at(index, col, edit)
        tags.pop()

    def no_dtd(*_args):
        _unsupported("unsupported_package_xml")

    def encoding(_version, name, _standalone):
        if name and name.lower() != "utf-8":
            _unsupported("unsupported_sheet_encoding")

    parser.StartDoctypeDeclHandler = no_dtd
    parser.XmlDeclHandler = encoding
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.CharacterDataHandler = characters
    parser.Parse(xml, True)
    if _ROW[1] not in seen_sheet_tags:
        _unsupported("unsupported_sheet_data")
    if per_row:
        _unsupported("target_row_missing")
    if not patches:
        return xml, formula_count
    chunks, previous = [], 0
    for start_at, end_at, replacement in patches:
        if not previous <= start_at <= end_at <= len(xml):
            _unsupported("ambiguous_xml_spans")
        chunks.extend((xml[previous:start_at], replacement))
        previous = end_at
    chunks.append(xml[previous:])
    return b"".join(chunks), formula_count


def build_preserving_candidate(baseline: Path, candidate: Path, operations):
    """Patch a staged copy and verify every package member plus cell readback."""
    if any(op["tool"] != "write_data_to_excel" for op in operations):
        _unsupported("unsupported_operation")
    try:
        with ZipFile(baseline) as source:
            scan = _Archive(source)
            inventory = _package_profile(scan)
            originals = {}
            formula_count = 0
            for part in inventory.sheets.values():
                original = b"".join(scan.chunks(part, limit=MAX_SHEET_BYTES))
                _, formulas = _patch_sheet(original, {})
                originals[part] = original
                formula_count += formulas
            book = load_workbook(baseline, rich_text=True)
            try:
                if set(book.sheetnames) != set(inventory.sheets):
                    _unsupported("unsupported_package_layout")
                edits, results, changes = _collect_edits(book, operations)
            finally:
                book.close()

            patched = {}
            for name, targets in edits.items():
                part = inventory.sheets[name]
                original = originals[part]
                updated, _ = _patch_sheet(original, targets)
                if updated != original:
                    patched[part] = updated

            expected = {}
            original_parts = _Archive(source)
            # A no-op keeps the staged baseline copy, including its ZIP encoding.
            with ZipFile(candidate, "w") if patched else nullcontext() as saved:
                if saved is not None:
                    saved.comment = source.comment
                for entry in source.infolist():
                    name = entry.filename
                    content = patched.get(name)
                    if content is None:
                        content = originals.get(name)
                    if content is None:
                        content = (
                            b""
                            if entry.is_dir()
                            else b"".join(original_parts.chunks(name))
                        )
                    expected[name] = hashlib.sha256(content).digest()
                    if saved is not None:
                        saved.writestr(copy(entry), content)

        with ZipFile(candidate) as saved:
            if saved.namelist() != list(expected):
                _unsupported("candidate_package_mismatch")
            inspected = _Archive(saved)
            for name, digest in expected.items():
                if inspected.digest(name) != digest:
                    _unsupported("candidate_package_mismatch")

        book = load_workbook(candidate, rich_text=True)
        try:
            for name, targets in edits.items():
                for edit in targets.values():
                    if not _equal(book[name][edit.coord].value, edit.after):
                        _unsupported("candidate_value_mismatch")
        finally:
            book.close()
    except _Unverified as exc:
        _unsupported(str(exc))
    except (
        BadZipFile,
        expat.ExpatError,
        ZlibError,
        OSError,
        IndexError,
        KeyError,
        OverflowError,
        TypeError,
        ValueError,
        RuntimeError,
    ):
        _unsupported("unreadable_package")
    sample = sorted(patched)[:5]
    return (
        results,
        changes,
        {
            "changed_part_count": len(patched),
            "changed_parts_sample": [name[:160] for name in sample],
            "changed_parts_truncated": len(patched) > len(sample),
            "unchanged_parts_verified": len(expected) - len(patched),
            "constant_formula_count": formula_count,
        },
    )
