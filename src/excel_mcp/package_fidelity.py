"""Targeted XLSX round-trip checks for staged ChangeSet writes."""

from __future__ import annotations

from collections.abc import Collection, Iterator
from dataclasses import dataclass
import hashlib
from pathlib import Path
import posixpath
import re
from urllib.parse import unquote, urlsplit
from xml.parsers import expat
from zipfile import BadZipFile, ZipFile
from zlib import error as ZlibError

MAX_MEMBERS = 20_000
MAX_RELATIONSHIPS = 50_000
MAX_SCAN_BYTES = 256 * 1024 * 1024  # Per archive, uncompressed.
MAX_METADATA_PART_BYTES = 4 * 1024 * 1024
MAX_CACHES_PER_SHEET = 250_000
MAX_SAMPLES = 5

_SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships}"
_TYPE_NS = "http://schemas.openxmlformats.org/package/2006/content-types}"
_OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_CUSTOM_REL_TYPES = {_OFFICE_REL + "/customXml", _OFFICE_REL + "/customXmlProps"}
_CUSTOM_PROPS_TYPE = (
    "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
)
_CALC_CHAIN_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"
)
_CELL_REF = re.compile(r"[A-Z]{1,3}[1-9][0-9]{0,6}", re.ASCII)


class _Unverified(Exception):
    """A known condition that prevents a complete targeted comparison."""


class _Archive:
    def __init__(self, archive: ZipFile):
        entries = archive.infolist()
        if len(entries) > MAX_MEMBERS:
            raise _Unverified("inspection_limit")
        self.files = {entry.filename for entry in entries if not entry.is_dir()}
        if len(self.files) != sum(not entry.is_dir() for entry in entries):
            raise _Unverified("ambiguous_package_members")
        self.archive = archive
        self.bytes_left = MAX_SCAN_BYTES

    def chunks(self, part: str, *, limit: int = MAX_SCAN_BYTES) -> Iterator[bytes]:
        entry = self.archive.getinfo(part)
        if entry.file_size > min(limit, self.bytes_left):
            raise _Unverified("inspection_limit")
        self.bytes_left -= entry.file_size
        with self.archive.open(entry) as stream:
            while chunk := stream.read(64 * 1024):
                yield chunk

    def digest(self, part: str) -> bytes:
        digest = hashlib.sha256()
        for chunk in self.chunks(part):
            digest.update(chunk)
        return digest.digest()

    def xml(self, part, root, visit, *, leave=None, characters=None, metadata=True):
        path = []
        parser = expat.ParserCreate(namespace_separator="}")

        def reject_dtd(*_args):
            raise _Unverified("unsupported_package_xml")

        def start(tag, attrs):
            if not path and tag != root:
                raise _Unverified("unsupported_package_xml")
            if len(path) >= 128:
                raise _Unverified("inspection_limit")
            path.append(tag)
            visit(path, attrs)

        def end(_tag):
            if leave is not None:
                leave(path)
            path.pop()

        parser.StartDoctypeDeclHandler = reject_dtd
        parser.StartElementHandler = start
        parser.EndElementHandler = end
        parser.CharacterDataHandler = characters
        limit = MAX_METADATA_PART_BYTES if metadata else MAX_SCAN_BYTES
        for chunk in self.chunks(part, limit=limit):
            parser.Parse(chunk, False)
        parser.Parse(b"", True)


def _resolve_target(owner: str, target: str) -> str:
    address = urlsplit(target)
    part = unquote(address.path)
    if address.scheme or address.netloc or not part or "\\" in part or "\0" in part:
        raise _Unverified("unsupported_relationship_target")
    resolved = posixpath.normpath(
        part.lstrip("/")
        if part.startswith("/")
        else posixpath.join(posixpath.dirname(owner), part)
    )
    if resolved in {".", ".."} or resolved.startswith("../"):
        raise _Unverified("unsupported_relationship_target")
    return resolved


def _custom_path(part: str) -> bool:
    return part.casefold().startswith(("customxml/", "xl/customxml/"))


@dataclass(frozen=True)
class _Relationship:
    owner: str
    identifier: str
    kind: str
    mode: str
    target: str

    def part(self) -> str | None:
        return (
            _resolve_target(self.owner, self.target)
            if self.mode == "Internal"
            else None
        )

    def identity(self) -> tuple[str, str, str, str, str, str, str]:
        part = self.part()
        if part is None:
            return (
                self.owner,
                self.identifier,
                self.kind,
                self.mode,
                self.target,
                "",
                "",
            )
        address = urlsplit(self.target)
        return (
            self.owner,
            self.identifier,
            self.kind,
            self.mode,
            part,
            address.query,
            address.fragment,
        )


def _relationships(archive: _Archive) -> list[_Relationship]:
    relationships = []
    for part in sorted(archive.files):
        directory, filename = posixpath.split(part)
        if posixpath.basename(directory) != "_rels" or not filename.endswith(".rels"):
            continue
        owner = posixpath.join(posixpath.dirname(directory), filename[:-5])
        ids = set()

        def visit(path, attrs):
            if len(path) != 2 or path[-1] != _REL_NS + "Relationship":
                return
            identifier = attrs["Id"]
            if identifier in ids:
                raise _Unverified("ambiguous_package_relationships")
            if len(relationships) >= MAX_RELATIONSHIPS:
                raise _Unverified("inspection_limit")
            ids.add(identifier)
            mode = attrs.get("TargetMode", "Internal")
            if mode not in {"Internal", "External"}:
                raise _Unverified("unsupported_relationship_target")
            relationships.append(
                _Relationship(owner, identifier, attrs["Type"], mode, attrs["Target"])
            )

        archive.xml(part, _REL_NS + "Relationships", visit)
    return relationships


@dataclass
class _Inventory:
    sheets: dict[str, str]
    content_types: dict[str, str | None]
    custom_parts: set[str]
    custom_links: set[tuple[str, str, str, str, str, str, str]]
    calculation_chains: set[str]


def _inventory(archive: _Archive) -> _Inventory:
    defaults, overrides = {}, {}

    def content_type(path, attrs):
        if len(path) == 2 and path[-1] == _TYPE_NS + "Default":
            defaults[attrs["Extension"].casefold()] = attrs["ContentType"]
        elif len(path) == 2 and path[-1] == _TYPE_NS + "Override":
            overrides[_resolve_target("", attrs["PartName"])] = attrs["ContentType"]

    archive.xml("[Content_Types].xml", _TYPE_NS + "Types", content_type)

    def type_for(part):
        return overrides.get(
            part, defaults.get(posixpath.splitext(part)[1][1:].casefold())
        )

    types = {name: type_for(name) for name in archive.files}
    custom = {
        name
        for name, media_type in types.items()
        if (_custom_path(name) and not name.endswith(".rels"))
        or media_type == _CUSTOM_PROPS_TYPE
    }
    chains = {
        name
        for name, media_type in types.items()
        if name == "xl/calcChain.xml" or media_type == _CALC_CHAIN_TYPE
    }
    relationships = _relationships(archive)
    for link in relationships:
        if link.kind in _CUSTOM_REL_TYPES:
            if (target := link.part()) is not None:
                custom.add(target)
        elif (
            link.kind == _OFFICE_REL + "/calcChain"
            and (target := link.part()) is not None
        ):
            chains.add(target)
    links = {
        link.identity()
        for link in relationships
        if link.owner in custom or link.kind in _CUSTOM_REL_TYPES
    }
    workbooks = [
        link.part()
        for link in relationships
        if not link.owner and link.kind == _OFFICE_REL + "/officeDocument"
    ]
    if len(workbooks) != 1 or workbooks[0] is None:
        raise _Unverified("unsupported_workbook_layout")
    workbook = workbooks[0]
    workbook_links = {
        link.identifier: link for link in relationships if link.owner == workbook
    }
    sheets: dict[str, str] = {}

    def sheet(path, attrs):
        if len(path) != 3 or path[1:] != [_SHEET_NS + "sheets", _SHEET_NS + "sheet"]:
            return
        link = workbook_links[attrs[_OFFICE_REL + "}id"]]
        if link.kind == _OFFICE_REL + "/worksheet":
            name, part = attrs["name"], link.part()
            if part not in archive.files or name in sheets:
                raise _Unverified("unsupported_workbook_layout")
            sheets[name] = part
        elif link.kind != _OFFICE_REL + "/chartsheet":
            raise _Unverified("unsupported_workbook_layout")

    archive.xml(workbook, _SHEET_NS + "workbook", sheet)
    return _Inventory(sheets, types, custom, links, chains)


def _formula_caches(
    archive: _Archive, part: str, *, wanted: Collection[str] | None = None
) -> dict[str, tuple[str, bytes]]:
    cell_path = [_SHEET_NS + tag for tag in ("worksheet", "sheetData", "row", "c")]
    value_tag = _SHEET_NS + "v"
    caches = {}
    coord = None
    cell_type = "n"
    formula = in_value = has_value_text = False
    digest = None

    def visit(path, attrs):
        nonlocal coord, cell_type, formula, digest, in_value, has_value_text
        if path == cell_path:
            coord, cell_type = attrs.get("r"), attrs.get("t", "n")
            formula = in_value = has_value_text = False
            digest = None
        elif (
            len(path) == 5
            and path[:4] == cell_path
            and (wanted is None or coord in wanted)
        ):
            if path[-1] == _SHEET_NS + "f":
                formula = True
            elif path[-1] == value_tag and formula:
                digest = hashlib.sha256()
                in_value = True

    def leave(path):
        nonlocal in_value
        if len(path) == 5 and path[-1] == value_tag and path[:4] == cell_path:
            in_value = False
        if (
            path != cell_path
            or digest is None
            or not (has_value_text or cell_type == "str")
        ):
            return
        if not coord or not _CELL_REF.fullmatch(coord) or coord in caches:
            raise _Unverified("ambiguous_formula_cache_location")
        if len(caches) >= MAX_CACHES_PER_SHEET:
            raise _Unverified("inspection_limit")
        caches[coord] = (cell_type, digest.digest())

    def characters(text):
        nonlocal has_value_text
        if in_value:
            has_value_text = True
            digest.update(text.encode("utf-8"))

    archive.xml(
        part, cell_path[0], visit, leave=leave, characters=characters, metadata=False
    )
    return caches


def _label(name: str) -> str:
    return name if len(name) <= 100 else name[:97] + "..."


def inspect_package_fidelity(
    baseline: Path,
    candidate: Path,
    *,
    changed_cells: set[tuple[str, str]],
    sample_limit: int,
) -> dict:
    """Report targeted losses, not general XLSX equivalence or cache correctness."""
    limit = min(sample_limit, MAX_SAMPLES)
    issues = []

    def issue(code, count, samples, *, blocking=True):
        if count:
            issues.append(
                {
                    "code": code,
                    "severity": "blocking" if blocking else "warning",
                    "count": count,
                    "sample": samples,
                    "truncated": count > len(samples),
                }
            )

    try:
        with ZipFile(baseline) as old_zip, ZipFile(candidate) as new_zip:
            old, new = _Archive(old_zip), _Archive(new_zip)
            before, after = _inventory(old), _inventory(new)
            if not before.custom_parts <= old.files:
                raise _Unverified("missing_custom_xml_part")
            removed = sorted(before.custom_parts - new.files)
            modified = [
                part
                for part in sorted(before.custom_parts & new.files)
                if before.content_types[part] != after.content_types[part]
                or old.digest(part) != new.digest(part)
            ]
            issue(
                "custom_xml_parts_removed",
                len(removed),
                [_label(p) for p in removed[:limit]],
            )
            issue(
                "custom_xml_parts_changed",
                len(modified),
                [_label(p) for p in modified[:limit]],
            )
            links = before.custom_links - after.custom_links
            sample_links = sorted((link[0], link[1]) for link in links)[:limit]
            issue(
                "custom_xml_relationships_lost_or_changed",
                len(links),
                [
                    {"owner_part": _label(owner), "relationship_id": _label(identifier)}
                    for owner, identifier in sample_links
                ],
            )
            chains = sorted(before.calculation_chains - after.calculation_chains)
            issue(
                "calculation_chain_removed",
                len(chains),
                [_label(p) for p in chains[:limit]],
                blocking=False,
            )

            cache_count = 0
            cache_sample = []
            for sheet, part in before.sheets.items():
                cached = {
                    coord: value
                    for coord, value in _formula_caches(old, part).items()
                    if (sheet, coord) not in changed_cells
                }
                if not cached:
                    continue
                new_part = after.sheets.get(sheet)
                actual = (
                    _formula_caches(new, new_part, wanted=cached) if new_part else {}
                )
                for coord, value in cached.items():
                    if actual.get(coord) != value:
                        cache_count += 1
                        if len(cache_sample) < limit:
                            cache_sample.append(
                                {
                                    "sheet_name": _label(sheet),
                                    "cell": coord,
                                    "change": "changed"
                                    if coord in actual
                                    else "missing",
                                }
                            )
            issue("formula_caches_lost_or_changed", cache_count, cache_sample)
            unclassified = sum(
                not _custom_path(part)
                for part in old.files - new.files - set(removed) - set(chains)
            )
    except _Unverified as exc:
        return {"status": "unverified", "reason": str(exc), "issues": issues}
    except (
        BadZipFile,
        expat.ExpatError,
        ZlibError,
        OSError,
        KeyError,
        ValueError,
        RuntimeError,
        NotImplementedError,
    ):
        return {
            "status": "unverified",
            "reason": "unreadable_package",
            "issues": issues,
        }

    result = {
        "status": "blocked"
        if any(item["severity"] == "blocking" for item in issues)
        else "passed",
        "issues": issues,
    }
    if unclassified:
        result["unclassified_removed_part_count"] = unclassified
    return result
