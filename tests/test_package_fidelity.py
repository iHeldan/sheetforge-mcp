import hashlib
import json
from pathlib import Path
import shutil
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table

from excel_mcp import changeset, package_fidelity, rebase, recipes, scenarios, server
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
