import hashlib
import json
from pathlib import Path
import shutil
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.worksheet.table import Table

from excel_mcp import (
    changeset,
    data,
    package_fidelity,
    rebase,
    recipes,
    scenarios,
    server,
)
from excel_mcp.exceptions import PreconditionFailedError

_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
_DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
_TYPE = "http://schemas.openxmlformats.org/package/2006/content-types"
_WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
_CHAIN = "xl/calcChain.xml"
_CUSTOM = "customXml/item1.xml"
_CUSTOM_PROPS = "customXml/itemProps1.xml"
_CUSTOM_RELS = "customXml/_rels/item1.xml.rels"
_FORMULA_SHEET = "xl/worksheets/sheet2.xml"

_EDIT = [
    {
        "tool": "write_data_to_excel",
        "args": {
            "sheet_name": "Data",
            "start_cell": "E1",
            "data": [["intended"]],
        },
    },
]
_ASSERTIONS = [
    {"type": "cell_equals", "sheet_name": "Data", "cell": "E1", "expected": "intended"},
    {"type": "range_values_unchanged", "sheet_name": "Formulas", "range_ref": "A1:A3"},
]


def _rewrite(path, edit, compression=ZIP_DEFLATED):
    with ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    edit(parts)
    with ZipFile(path, "w", compression=compression) as archive:
        for name, content in parts.items():
            archive.writestr(name, content)


def _xml(parts, name, edit):
    root = ET.fromstring(parts[name])
    edit(root)
    parts[name] = ET.tostring(root)


def _add_link(root, number, kind, target):
    ET.SubElement(
        root,
        f"{{{_REL}}}Relationship",
        Id=f"rId{number}",
        Type=_DOC_REL + kind,
        Target=target,
    )


def _book(path, *, custom=False, chain=False, caches=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Id", "Unit", "Amount"])
    ws.append([1, "North", 10])
    ws.append([2, "South", 20])
    ws.add_table(Table(displayName="Sales", ref="A1:C3"))
    formulas = wb.create_sheet("Formulas")
    formulas["A1"] = "=6*9"
    formulas["A2"] = '=IF(TRUE,"synthetic-result","")'
    formulas["A3"] = "=0"
    wb.save(path)
    wb.close()

    def add_parts(parts):
        def links(root):
            if custom:
                _add_link(root, 101, "customXml", "../customXml/item1.xml")
            if chain:
                _add_link(root, 102, "calcChain", "calcChain.xml")

        def content_types(root):
            if custom:
                ET.SubElement(
                    root,
                    f"{{{_TYPE}}}Override",
                    PartName="/" + _CUSTOM_PROPS,
                    ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
                )
            if chain:
                ET.SubElement(
                    root,
                    f"{{{_TYPE}}}Override",
                    PartName="/" + _CHAIN,
                    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
                )

        def values(root):
            for coord, (kind, value) in (caches or {}).items():
                cell = root.find(f".//{{{_MAIN}}}c[@r='{coord}']")
                cell.set("t", kind)
                cell.find(f"{{{_MAIN}}}v").text = value

        if custom:
            parts[_CUSTOM] = b'<test xmlns="urn:synthetic:sample">not-for-output</test>'
            parts[_CUSTOM_PROPS] = (
                b'<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" '
                b'ds:itemID="{66E39937-AC3B-4F87-91CD-8649A7C8E7A2}"><ds:schemaRefs/></ds:datastoreItem>'
            )
            rels = ET.Element(f"{{{_REL}}}Relationships")
            _add_link(rels, 1, "customXmlProps", "itemProps1.xml")
            parts[_CUSTOM_RELS] = ET.tostring(rels)
        if chain:
            parts[_CHAIN] = (
                f'<calcChain xmlns="{_MAIN}"><c r="A1" i="2"/></calcChain>'.encode()
            )
        _xml(parts, _WORKBOOK_RELS, links)
        _xml(parts, "[Content_Types].xml", content_types)
        _xml(parts, _FORMULA_SHEET, values)

    _rewrite(path, add_parts)
    return path


def _commit(path, preview, ops=_EDIT, assertions=None):
    return changeset.apply_workbook_changeset(
        str(path),
        ops,
        assertions,
        mode="commit",
        expected_workbook_sha256=preview["expected_workbook_sha256"],
        changeset_token=preview["changeset_token"],
    )


def _compare(before, after):
    return package_fidelity.inspect_package_fidelity(
        before, after, changed_cells=set(), sample_limit=5
    )


def _issues(result):
    return {item["code"]: item for item in result["issues"]}


def test_direct_write_previews_and_blocks_same_loss_as_changeset(tmp_path):
    path = _book(
        tmp_path / "direct.xlsx", custom=True, chain=True, caches={"A1": ("n", "54")}
    )
    frozen, mtime = path.read_bytes(), path.stat().st_mtime_ns
    args = {"filepath": str(path), **_EDIT[0]["args"]}
    direct = json.loads(
        server.write_data_to_excel(**args, dry_run=True, include_changes=False)
    )
    staged = changeset.apply_workbook_changeset(str(path), _EDIT)
    assert direct["ok"] and direct["dry_run"] and not direct["data"]["ready_to_commit"]
    report = direct["data"]["package_fidelity"]
    assert report == staged["package_fidelity"]
    assert report["status"] == "blocked"

    denied = json.loads(server.write_data_to_excel(**args, include_changes=False))
    assert not denied["ok"]
    assert denied["error"]["code"] == "write_package_fidelity_failed"
    assert denied["error"]["details"]["package_fidelity"] == report
    assert path.read_bytes() == frozen and path.stat().st_mtime_ns == mtime
    assert list(tmp_path.iterdir()) == [path]


def test_direct_write_may_replace_a_formula_in_its_target_cell(tmp_path):
    path = _book(tmp_path / "direct.xlsx", chain=True, caches={"A1": ("n", "54")})
    preview = data.write_data(str(path), "Formulas", [[99]], "A1", dry_run=True)
    assert preview["ready_to_commit"]
    report = preview["package_fidelity"]
    assert set(_issues(report)) == {"calculation_chain_removed"}
    assert (
        data.write_data(str(path), "Formulas", [[99]], "A1")["package_fidelity"]
        == report
    )
    wb = load_workbook(path, data_only=True)
    try:
        assert wb["Formulas"]["A1"].value == 99
    finally:
        wb.close()


def test_direct_write_blocks_when_inspection_cannot_finish(tmp_path, monkeypatch):
    path = _book(tmp_path / "direct.xlsx")
    frozen = path.read_bytes()
    monkeypatch.setattr(package_fidelity, "MAX_MEMBERS", 1)
    preview = data.write_data(str(path), "Data", [["new"]], "E1", dry_run=True)
    assert preview["package_fidelity"]["status"] == "unverified"
    assert not preview["ready_to_commit"]
    with pytest.raises(PreconditionFailedError) as rejected:
        data.write_data(str(path), "Data", [["new"]], "E1")
    assert rejected.value.code == "write_package_fidelity_failed"
    assert path.read_bytes() == frozen and list(tmp_path.iterdir()) == [path]


def test_direct_write_refuses_source_drift_after_inspection(tmp_path, monkeypatch):
    path = _book(tmp_path / "direct.xlsx")
    original_copy = shutil.copy2

    def intervening_save(source, target, **kwargs):
        copied = original_copy(source, target, **kwargs)
        with ZipFile(source, "a") as archive:
            archive.comment = b"saved by another editor"
        return copied

    monkeypatch.setattr(shutil, "copy2", intervening_save)
    with pytest.raises(PreconditionFailedError) as rejected:
        data.write_data(str(path), "Data", [["new"]], "E1")
    assert rejected.value.code == "stale_workbook"
    with ZipFile(path) as archive:
        assert archive.comment == b"saved by another editor"
    wb = load_workbook(path)
    try:
        assert wb["Data"]["E1"].value is None
    finally:
        wb.close()
    assert list(tmp_path.iterdir()) == [path]


def test_direct_write_refuses_candidate_drift_after_inspection(tmp_path, monkeypatch):
    path = _book(tmp_path / "direct.xlsx")
    frozen = path.read_bytes()
    inspect = data.inspect_package_fidelity

    def changed_candidate(baseline, candidate, **kwargs):
        report = inspect(baseline, candidate, **kwargs)
        assert report["status"] == "passed"
        with ZipFile(candidate, "a") as archive:
            archive.comment = b"changed after inspection"
        return report

    monkeypatch.setattr(data, "inspect_package_fidelity", changed_candidate)
    with pytest.raises(PreconditionFailedError) as rejected:
        data.write_data(str(path), "Data", [["new"]], "E1")
    assert rejected.value.code == "write_candidate_changed"
    assert path.read_bytes() == frozen and list(tmp_path.iterdir()) == [path]


def _rich_book(path, storage):
    wb = Workbook()
    wb.active.title = "First"
    active = wb.create_sheet("Selected")
    active["A1"] = CellRichText(
        TextBlock(InlineFont(b=True), "bold"),
        " and ",
        TextBlock(InlineFont(i=True), "italic"),
    )
    wb.active = active
    wb.save(path)
    wb.close()
    if storage == "inline":
        return

    def to_shared(parts):
        strings = ET.Element(f"{{{_MAIN}}}sst", count="1", uniqueCount="1")

        def edit_cell(root):
            cell = root.find(f".//{{{_MAIN}}}c[@r='A1']")
            inline = cell.find(f"{{{_MAIN}}}is")
            cell.remove(inline)
            cell.set("t", "s")
            ET.SubElement(cell, f"{{{_MAIN}}}v").text = "0"
            inline.tag = f"{{{_MAIN}}}si"
            strings.append(inline)

        _xml(parts, _FORMULA_SHEET, edit_cell)
        parts["xl/sharedStrings.xml"] = ET.tostring(strings)
        _xml(
            parts,
            _WORKBOOK_RELS,
            lambda root: _add_link(root, 111, "sharedStrings", "sharedStrings.xml"),
        )
        _xml(
            parts,
            "[Content_Types].xml",
            lambda root: ET.SubElement(
                root,
                f"{{{_TYPE}}}Override",
                PartName="/xl/sharedStrings.xml",
                ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
            ),
        )

    _rewrite(path, to_shared)


@pytest.mark.parametrize("storage", ["inline", "shared"])
@pytest.mark.parametrize("route", ["changeset", "direct"])
def test_writes_preserve_rich_text_and_changeset_uses_active_sheet(
    tmp_path, storage, route
):
    path = tmp_path / "rich.xlsx"
    _rich_book(path, storage)
    frozen = path.read_bytes()
    wb = load_workbook(path, rich_text=True)
    try:
        original = wb["Selected"]["A1"].value
        assert isinstance(original, CellRichText)
    finally:
        wb.close()
    if route == "changeset":
        ops = [
            {"tool": "create_worksheet", "args": {"sheet_name": "Report"}},
            {
                "tool": "write_data_to_excel",
                "args": {"start_cell": "C1", "data": [["new"]]},
            },
            {
                "tool": "format_range",
                "args": {"sheet_name": "Selected", "start_cell": "C1", "bold": True},
            },
        ]
        preview = changeset.apply_workbook_changeset(str(path), ops)
        assert preview["ready_to_commit"] and path.read_bytes() == frozen
        assert _commit(path, preview, ops)["persisted"]
    else:
        preview = data.write_data(str(path), "Selected", [["new"]], "C1", dry_run=True)
        assert preview["ready_to_commit"] and path.read_bytes() == frozen
        data.write_data(str(path), "Selected", [["new"]], "C1")

    wb = load_workbook(path, rich_text=True)
    try:
        assert wb["First"]["C1"].value is None and wb["Selected"]["C1"].value == "new"
        actual = wb["Selected"]["A1"].value
        assert isinstance(actual, CellRichText) and actual == original
    finally:
        wb.close()


def test_same_visible_text_does_not_flatten_target_rich_text(tmp_path):
    path = tmp_path / "rich.xlsx"
    _rich_book(path, "inline")
    unchanged = data.write_data(str(path), "Selected", [["bold and italic"]], "A1")
    assert unchanged["changed_cells"] == 0
    wb = load_workbook(path, rich_text=True)
    try:
        assert isinstance(wb["Selected"]["A1"].value, CellRichText)
    finally:
        wb.close()
    args = dict(
        filepath=str(path),
        sheet_name="Selected",
        data=[["replacement"]],
        include_changes=True,
    )
    replaced = json.loads(server.write_data_to_excel(**args))
    assert replaced["ok"]
    assert replaced["changes"][0]["old_value"] == "bold and italic"


def test_changeset_checks_only_the_final_candidate(tmp_path):
    path = _book(tmp_path / "final.xlsx", caches={"A1": ("n", "54")})
    ops = [
        *_EDIT,
        {
            "tool": "write_data_to_excel",
            "args": {
                "sheet_name": "Formulas",
                "start_cell": "A1",
                "data": [[99]],
            },
        },
    ]
    preview = changeset.apply_workbook_changeset(str(path), ops)
    assert (
        preview["ready_to_commit"] and preview["package_fidelity"]["status"] == "passed"
    )
    assert _commit(path, preview, ops)["persisted"]


def test_changeset_mcp_identifies_silent_content_loss_and_refuses_commit(tmp_path):
    path = _book(
        tmp_path / "synthetic.xlsx",
        custom=True,
        chain=True,
        caches={
            "A1": ("n", "54"),
            "A2": ("str", "synthetic-result"),
            "A3": ("n", "0"),
        },
    )
    frozen, mtime = path.read_bytes(), path.stat().st_mtime_ns
    wb = load_workbook(path, data_only=True)
    assert wb["Formulas"]["A1"].value == 54
    wb.close()

    preview = json.loads(
        server.apply_workbook_changeset(str(path), _EDIT, _ASSERTIONS, sample_limit=1)
    )
    assert preview["ok"], preview
    data = preview["data"]
    assert data["assertions_passed"] == 2 and data["ready_to_commit"] is False
    assert data["diff"]["cell_changes"]["count"] == 1
    report = data["package_fidelity"]
    assert report["status"] == "blocked"
    issues = _issues(report)
    assert issues["custom_xml_parts_removed"]["sample"] == [_CUSTOM]
    assert issues["custom_xml_parts_removed"]["count"] == 2
    assert issues["custom_xml_relationships_lost_or_changed"]["count"] == 2
    assert issues["calculation_chain_removed"]["severity"] == "warning"
    assert issues["formula_caches_lost_or_changed"] == {
        "code": "formula_caches_lost_or_changed",
        "severity": "blocking",
        "count": 3,
        "sample": [{"sheet_name": "Formulas", "cell": "A1", "change": "missing"}],
        "truncated": True,
    }
    assert "not-for-output" not in json.dumps(
        report
    ) and "synthetic-result" not in json.dumps(report)

    rejected = json.loads(
        server.apply_workbook_changeset(
            str(path),
            _EDIT,
            _ASSERTIONS,
            mode="commit",
            sample_limit=1,
            expected_workbook_sha256=data["expected_workbook_sha256"],
            changeset_token=data["changeset_token"],
        )
    )
    assert not rejected["ok"]
    assert rejected["error"]["code"] == "changeset_package_fidelity_failed"
    assert rejected["error"]["details"]["package_fidelity"] == report
    assert path.read_bytes() == frozen and path.stat().st_mtime_ns == mtime
    assert list(tmp_path.glob("*.xlsx")) == [path]


def test_calc_chain_removal_alone_warns_and_can_commit(tmp_path):
    path = _book(tmp_path / "chain.xlsx", chain=True)
    preview = changeset.apply_workbook_changeset(str(path), _EDIT)
    assert preview["ready_to_commit"]
    assert preview["package_fidelity"] == {
        "status": "passed",
        "issues": [
            {
                "code": "calculation_chain_removed",
                "severity": "warning",
                "count": 1,
                "sample": [_CHAIN],
                "truncated": False,
            }
        ],
    }
    result = _commit(path, preview)
    assert (
        result["persisted"]
        and result["package_fidelity"] == preview["package_fidelity"]
    )
    with ZipFile(path) as archive:
        assert _CHAIN not in archive.namelist()


@pytest.mark.parametrize(("value", "ready"), [(100, True), ("=6*9", False)])
def test_only_effective_explicit_cell_changes_exempt_old_cache(tmp_path, value, ready):
    path = _book(tmp_path / "formula.xlsx", caches={"A1": ("n", "54")})
    ops = [
        {
            "tool": "write_data_to_excel",
            "args": {
                "sheet_name": "Formulas",
                "start_cell": "A1",
                "data": [[value]],
            },
        }
    ]
    preview = changeset.apply_workbook_changeset(str(path), ops)
    assert preview["ready_to_commit"] is ready
    if ready:
        assert _commit(path, preview, ops)["persisted"]
    else:
        assert "formula_caches_lost_or_changed" in _issues(preview["package_fidelity"])


def test_repacking_and_equivalent_relationship_targets_do_not_block(tmp_path):
    before = _book(tmp_path / "before.xlsx", custom=True, caches={"A1": ("n", "54")})
    after = tmp_path / "after.xlsx"
    shutil.copyfile(before, after)

    def repack(parts):
        _xml(
            parts,
            _WORKBOOK_RELS,
            lambda root: root.find(f"{{{_REL}}}Relationship[@Id='rId101']").set(
                "Target", "/customXml/item1.xml"
            ),
        )
        parts["xl/worksheets/alternate.xml"] = parts.pop(_FORMULA_SHEET)
        _xml(
            parts,
            _WORKBOOK_RELS,
            lambda root: root.find(
                f"{{{_REL}}}Relationship[@Type='{_DOC_REL}worksheet'][@Id='rId2']"
            ).set("Target", "/xl/worksheets/alternate.xml"),
        )
        _xml(
            parts,
            "[Content_Types].xml",
            lambda root: root.find(
                f"{{{_TYPE}}}Override[@PartName='/{_FORMULA_SHEET}']"
            ).set("PartName", "/xl/worksheets/alternate.xml"),
        )

    _rewrite(after, repack, compression=ZIP_STORED)
    assert after.read_bytes() != before.read_bytes()
    assert _compare(before, after) == {
        "status": "passed",
        "issues": [],
        "unclassified_removed_part_count": 1,
    }


def test_protected_part_and_cache_changes_are_reported_without_contents(tmp_path):
    before = _book(tmp_path / "before.xlsx", custom=True, caches={"A1": ("n", "54")})
    after = tmp_path / "after.xlsx"
    shutil.copyfile(before, after)

    def edit(parts):
        parts[_CUSTOM] = b"<modified/>"
        _xml(parts, _CUSTOM_RELS, lambda root: root[0].set("Id", "rId99"))
        _xml(
            parts,
            _FORMULA_SHEET,
            lambda root: setattr(
                root.find(f".//{{{_MAIN}}}c[@r='A1']/{{{_MAIN}}}v"), "text", "55"
            ),
        )

    _rewrite(after, edit)
    report = _compare(before, after)
    assert report["status"] == "blocked"
    issues = _issues(report)
    assert issues["custom_xml_parts_changed"]["sample"] == [_CUSTOM]
    assert issues["custom_xml_relationships_lost_or_changed"]["sample"] == [
        {"owner_part": _CUSTOM, "relationship_id": "rId1"}
    ]
    assert issues["formula_caches_lost_or_changed"]["sample"] == [
        {"sheet_name": "Formulas", "cell": "A1", "change": "changed"}
    ]


def test_guard_limit_fails_closed_before_snapshot_or_replace(tmp_path, monkeypatch):
    path = _book(tmp_path / "limit.xlsx")
    frozen = path.read_bytes()
    monkeypatch.setattr(package_fidelity, "MAX_MEMBERS", 1)
    preview = changeset.apply_workbook_changeset(str(path), _EDIT)
    assert not preview["ready_to_commit"]
    assert preview["package_fidelity"] == {
        "status": "unverified",
        "reason": "inspection_limit",
        "issues": [],
    }
    with pytest.raises(PreconditionFailedError) as error:
        _commit(path, preview)
    assert error.value.code == "changeset_package_fidelity_failed"
    assert path.read_bytes() == frozen and list(tmp_path.glob("*.xlsx")) == [path]


def test_inspector_rejects_doctype_without_returning_xml_content(tmp_path):
    before = _book(tmp_path / "before.xlsx")
    after = tmp_path / "after.xlsx"
    shutil.copyfile(before, after)
    _rewrite(
        after,
        lambda parts: parts.update(
            {
                _FORMULA_SHEET: b'<!DOCTYPE worksheet [<!ENTITY ignored "test">]>'
                + parts[_FORMULA_SHEET],
            }
        ),
    )
    assert _compare(after, before) == {
        "status": "unverified",
        "reason": "unsupported_package_xml",
        "issues": [],
    }


def test_replacement_must_use_the_inspected_candidate(tmp_path, monkeypatch):
    path = _book(tmp_path / "source.xlsx")
    frozen = path.read_bytes()
    preview = changeset.apply_workbook_changeset(str(path), _EDIT)
    diff = changeset._preview_diff

    def drift(baseline, candidate, **kwargs):
        result = diff(baseline, candidate, **kwargs)
        with ZipFile(candidate, "a") as archive:
            archive.comment = b"changed after inspection"
        return result

    monkeypatch.setattr(changeset, "_preview_diff", drift)
    with pytest.raises(PreconditionFailedError) as error:
        _commit(path, preview)
    assert error.value.code == "changeset_candidate_changed"
    assert path.read_bytes() == frozen and list(tmp_path.glob("*.xlsx")) == [path]


@pytest.mark.parametrize("kind", ["recipe", "rebase", "scenario"])
def test_verified_workflows_surface_the_same_guard(tmp_path, kind):
    source = _book(tmp_path / "source.xlsx", caches={"A1": ("n", "54")})
    recipe = {
        "version": 1,
        "source": {"table_name": "Sales"},
        "group_by": ["Unit"],
        "metrics": [{"op": "sum", "field": "Amount", "as": "Total"}],
        "output": {"sheet_name": "Report", "range_ref": "A1:B10"},
    }
    updates = [{"key": {"Id": 1}, "values": {"Amount": 12}}]
    frozen = source.read_bytes()
    if kind == "recipe":

        def invoke(**kwargs):
            return recipes.run_workbook_recipe(str(source), recipe, **kwargs)

        token = "recipe_token"
        selected = {}
    elif kind == "rebase":
        base = _book(tmp_path / "base.xlsx")
        patch = {
            "version": 1,
            "source": recipe["source"],
            "key_fields": ["Id"],
            "read_fields": ["Unit"],
            "updates": updates,
        }

        def invoke(**kwargs):
            return rebase.apply_table_patch(
                str(base),
                str(source),
                patch,
                hashlib.sha256(base.read_bytes()).hexdigest(),
                **kwargs,
            )

        token = "rebase_token"
        selected = {}
    else:
        spec = {
            "version": 1,
            "recipe": recipe,
            "key_fields": ["Id"],
            "guards": [{"scope": "overall", "metric": "Total", "max": 40}],
            "scenarios": [{"id": "A", "updates": updates}],
        }

        def invoke(**kwargs):
            return scenarios.evaluate_workbook_scenarios(str(source), spec, **kwargs)

        token = "scenario_token"
        selected = {"selected_scenario": "A"}

    preview = invoke()
    plan = preview["scenarios"][0] if kind == "scenario" else preview
    assert not plan["ready_to_commit"]
    assert plan["package_fidelity"]["status"] == "blocked"
    assert "completed and verified" not in preview["message"]
    with pytest.raises(PreconditionFailedError) as error:
        invoke(
            mode="commit",
            expected_workbook_sha256=preview["expected_workbook_sha256"],
            **{token: plan[token]},
            **selected,
        )
    assert error.value.code == "changeset_package_fidelity_failed"
    assert source.read_bytes() == frozen
    assert not any("snapshot" in path.name for path in Path(tmp_path).glob("*.xlsx"))
