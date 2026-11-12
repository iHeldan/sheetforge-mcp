import json
from pathlib import Path
import re
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Font
from openpyxl.worksheet.datavalidation import DataValidation

from excel_mcp import changeset, preserve, server
from excel_mcp.exceptions import PreconditionFailedError, WorkbookError

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
INPUT_PART = "xl/worksheets/sheet1.xml"
FORMULA_PART = "xl/worksheets/sheet2.xml"
OPERATIONS = [
    {
        "tool": "write_data_to_excel",
        "args": {
            "start_cell": "B2",
            "data": [["Fresh <&> text"]],
        },
    }
]


def _members(path):
    with ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _write_members(path, members):
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        for name, value in members.items():
            archive.writestr(name, value)


def _book(path, *, formula="=6*9", custom=True, feature=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Inputs"
    ws["A2"] = 6
    ws["C2"] = "end of row"
    ws["C3"] = CellRichText("left ", TextBlock(InlineFont(b=True), "bold"))
    ws["A3"].font = Font(italic=True)  # Explicit styled blank.
    ws["B3"] = "old text"
    wb.create_sheet("Results")["A1"] = formula
    if feature == "validation":
        ws.add_data_validation(DataValidation(type="whole", formula1=1, sqref="B2"))
    elif feature == "protected":
        ws.protection.sheet = True
    wb.save(path)
    wb.close()
    parts = _members(path)
    if formula:
        # A synthetic cache, not a SheetForge calculation.
        root = ET.fromstring(parts[FORMULA_PART])
        root.find(f".//{{{NS}}}c[@r='A1']/{{{NS}}}v").text = "54"
        parts[FORMULA_PART] = ET.tostring(root)
    if custom:
        rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
        for index, kind, target in [
            (99, "customXml", "../customXml/item1.xml"),
            (100, "calcChain", "calcChain.xml"),
        ]:
            ET.SubElement(
                rels,
                f"{{{REL}}}Relationship",
                Id=f"rId{index}",
                Type=OFFICE + kind,
                Target=target,
            )
        parts["xl/_rels/workbook.xml.rels"] = ET.tostring(rels)
        types = ET.fromstring(parts["[Content_Types].xml"])
        ET.SubElement(
            types,
            f"{{{TYPES}}}Override",
            PartName="/xl/calcChain.xml",
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
        )
        parts["[Content_Types].xml"] = ET.tostring(types)
        parts["customXml/item1.xml"] = (
            b'<synthetic xmlns="urn:sheetforge:test">retained</synthetic>'
        )
        parts["xl/calcChain.xml"] = (
            f'<calcChain xmlns="{NS}"><c r="A1" i="2"/></calcChain>'.encode()
        )
    _write_members(path, parts)
    return path


def _call(path, operations=OPERATIONS, **kwargs):
    return changeset.apply_workbook_changeset(
        str(path), operations, save_mode="preserve", **kwargs
    )


def _commit(path, preview, operations=OPERATIONS, **kwargs):
    return _call(
        path,
        operations,
        mode="commit",
        changeset_token=preview["changeset_token"],
        expected_workbook_sha256=preview["expected_workbook_sha256"],
        **kwargs,
    )


def test_preserve_commit_keeps_all_other_xml_and_formula_cache(tmp_path, monkeypatch):
    path = _book(tmp_path / "synthetic.xlsx")
    frozen, mtime = path.read_bytes(), path.stat().st_mtime_ns
    before = _members(path)
    standard = changeset.apply_workbook_changeset(str(path), OPERATIONS)
    assert not standard["ready_to_commit"]
    monkeypatch.setattr(
        changeset,
        "_apply_operations",
        lambda *_a, **_kw: pytest.fail("standard fallback"),
    )
    assertions = [
        {
            "type": "cell_equals",
            "sheet_name": "Inputs",
            "cell": "B2",
            "expected": "Fresh <&> text",
        },
        {
            "type": "range_values_unchanged",
            "sheet_name": "Results",
            "range_ref": "A1:A1",
        },
    ]
    preview = _call(path, assertions=assertions)
    assert preview["ready_to_commit"] and preview["save_mode"] == "preserve"
    assert preview["preservation"] == {
        "changed_part_count": 1,
        "changed_parts_sample": [INPUT_PART],
        "changed_parts_truncated": False,
        "unchanged_parts_verified": len(before) - 1,
        "constant_formula_count": 1,
    }
    assert preview["diff"]["cell_changes"]["count"] == 1
    assert preview["assertions_passed"] == 2
    assert path.read_bytes() == frozen and path.stat().st_mtime_ns == mtime
    assert list(tmp_path.iterdir()) == [path]
    result = _commit(path, preview, assertions=assertions)
    assert result["persisted"] and result["package_fidelity"]["status"] == "passed"
    assert result["preservation"] == preview["preservation"]
    assert Path(result["receipt"]["snapshot"]["path"]).read_bytes() == frozen
    after = _members(path)
    assert before.keys() == after.keys()
    assert {name for name in before if before[name] != after[name]} == {INPUT_PART}
    inserted = re.compile(rb'<c xmlns="[^"]+" r="B2" t="inlineStr">.*?</c>')
    restored, matches = inserted.subn(b"", after[INPUT_PART])
    assert matches == 1 and restored == before[INPUT_PART]
    wb = load_workbook(path, data_only=True)
    assert wb["Results"]["A1"].value == 54
    assert wb["Inputs"]["B2"].value == "Fresh <&> text"
    wb.close()
    wb = load_workbook(path, rich_text=True)
    assert wb["Results"]["A1"].value == "=6*9"
    assert isinstance(wb["Inputs"]["C3"].value, CellRichText)
    wb.close()


@pytest.mark.parametrize(
    "formula", ["=Inputs!A2*9", '=INDIRECT("A1")', "=RAND()", "=1^2"]
)
def test_unverified_formula_blocks_even_if_cache_exists(tmp_path, formula):
    path = _book(tmp_path / "dependent.xlsx", formula=formula)
    frozen = path.read_bytes()
    ops = [{"tool": "write_data_to_excel", "args": {"start_cell": "A2", "data": [[8]]}}]
    with pytest.raises(PreconditionFailedError) as error:
        _call(path, ops)
    assert error.value.code == "changeset_preserve_unsupported"
    assert error.value.details["reason"] == "unsupported_formula"
    assert path.read_bytes() == frozen and list(tmp_path.iterdir()) == [path]


def test_preserving_supported_types_style_and_replace_order(tmp_path):
    path = _book(tmp_path / "literals.xlsx")
    ops = [
        {
            "tool": "write_data_to_excel",
            "args": {
                "start_cell": "A2",
                "data": [[False, 12.25], [" tabs\tline\nnext\r & ", None]],
            },
        }
    ]
    preview = _call(path, ops)
    assert preview["ready_to_commit"] and preview["diff"]["cell_changes"]["count"] == 4
    assert _commit(path, preview, ops)["persisted"]
    wb = load_workbook(path)
    ws = wb["Inputs"]
    assert ws["A2"].value is False and ws["B2"].value == 12.25
    assert ws["A3"].value == " tabs\tline\nnext\r & " and ws["A3"].font.italic
    assert ws["B3"].value is None and ws["C2"].value == "end of row"
    wb.close()


@pytest.mark.parametrize(
    ("address", "reason"),
    [
        ("D2", "insertion_outside_sheet_dimension"),
        ("A4", "target_row_missing"),
        ("C3", "unsupported_target_cell_value"),
    ],
)
def test_refuses_to_grow_layout_or_replace_rich_cell(tmp_path, address, reason):
    path = _book(tmp_path / "layout.xlsx")
    frozen = path.read_bytes()
    ops = [
        {
            "tool": "write_data_to_excel",
            "args": {"start_cell": address, "data": [["new"]]},
        }
    ]
    with pytest.raises(PreconditionFailedError) as error:
        _call(path, ops)
    assert error.value.details["reason"] == reason
    assert path.read_bytes() == frozen


@pytest.mark.parametrize(
    "value",
    [
        "=1+2",
        "#REF!",
        "",
        "literal _x000D_",
        "\x01",
        1_000_000_000_000_000,
        1.234567891234567,
    ],
)
def test_refuses_unsupported_literals(tmp_path, value):
    path = _book(tmp_path / "literals.xlsx")
    ops = [
        {"tool": "write_data_to_excel", "args": {"start_cell": "B2", "data": [[value]]}}
    ]
    with pytest.raises(PreconditionFailedError) as error:
        _call(path, ops)
    assert error.value.details["reason"] == "unsupported_literal"


@pytest.mark.parametrize("feature", ["validation", "protected"])
def test_refuses_unverified_sheet_features(tmp_path, feature):
    path = _book(tmp_path / "features.xlsx", feature=feature)
    with pytest.raises(PreconditionFailedError) as error:
        _call(path)
    assert error.value.details["reason"] == "unsupported_worksheet_feature"


def test_refuses_extra_parts_and_operations(tmp_path):
    path = _book(tmp_path / "unsupported.xlsx")
    ops = [{"tool": "create_worksheet", "args": {"sheet_name": "new"}}]
    with pytest.raises(PreconditionFailedError) as error:
        _call(path, ops)
    assert error.value.details["reason"] == "unsupported_operation"
    with ZipFile(path, "a") as archive:
        archive.writestr("_xmlsignatures/unsupported.xml", b"<synthetic/>")
    with pytest.raises(PreconditionFailedError) as error:
        _call(path)
    assert error.value.details["reason"] == "unsupported_package_part"


def test_commit_cannot_switch_save_mode_or_accept_source_drift(tmp_path):
    path = _book(tmp_path / "token.xlsx")
    preview = _call(path)
    common = dict(
        mode="commit",
        changeset_token=preview["changeset_token"],
        expected_workbook_sha256=preview["expected_workbook_sha256"],
    )
    with pytest.raises(PreconditionFailedError) as error:
        changeset.apply_workbook_changeset(str(path), OPERATIONS, **common)
    assert error.value.code == "changeset_token_mismatch"
    with ZipFile(path, "a") as archive:
        archive.comment = b"outside save"
    outside = path.read_bytes()
    with pytest.raises(PreconditionFailedError) as error:
        _call(path, **common)
    assert error.value.code == "stale_workbook"
    assert path.read_bytes() == outside and list(tmp_path.iterdir()) == [path]


def test_mcp_preview_and_bounded_inspection(tmp_path, monkeypatch):
    path = _book(tmp_path / "mcp.xlsx")
    reply = json.loads(
        server.apply_workbook_changeset(str(path), OPERATIONS, save_mode="preserve")
    )
    assert reply["ok"] and reply["data"]["ready_to_commit"]
    monkeypatch.setattr(preserve, "MAX_PACKAGE_BYTES", 1)
    reply = json.loads(
        server.apply_workbook_changeset(str(path), OPERATIONS, save_mode="preserve")
    )
    assert not reply["ok"]
    assert reply["error"]["code"] == "changeset_preserve_unsupported"
    assert reply["error"]["details"]["reason"] == "package_limit"


def test_preflight_rejects_unsupported_package_before_openpyxl(tmp_path, monkeypatch):
    path = _book(tmp_path / "bounds.xlsx")
    monkeypatch.setattr(preserve, "MAX_SOURCE_BYTES", 1)

    def unexpected_open(*_args, **_kwargs):
        pytest.fail("openpyxl should not open an unadmitted package")

    monkeypatch.setattr(
        changeset.workbook_module, "_verify_saved_workbook", unexpected_open
    )
    monkeypatch.setattr(preserve, "load_workbook", unexpected_open)
    with pytest.raises(PreconditionFailedError) as error:
        _call(path)
    assert error.value.details["reason"] == "package_limit"
    monkeypatch.setattr(preserve, "MAX_SOURCE_BYTES", path.stat().st_size)
    monkeypatch.setattr(preserve, "MAX_SHEET_BYTES", 1)
    with pytest.raises(PreconditionFailedError) as error:
        _call(path)
    assert error.value.details["reason"] == "inspection_limit"
    assert list(tmp_path.iterdir()) == [path]


def test_rejects_ambiguous_package_content_types(tmp_path):
    path = _book(tmp_path / "ambiguous.xlsx")
    members = _members(path)
    types = ET.fromstring(members["[Content_Types].xml"])
    original = types.find(f"{{{TYPES}}}Override[@PartName='/xl/workbook.xml']")
    ET.SubElement(types, original.tag, original.attrib)
    members["[Content_Types].xml"] = ET.tostring(types)
    _write_members(path, members)
    with pytest.raises(PreconditionFailedError) as error:
        _call(path)
    assert error.value.details["reason"] == "ambiguous_content_types"


def test_no_op_keeps_exact_archive_bytes(tmp_path):
    path = _book(tmp_path / "unchanged.xlsx", formula=None)
    with ZipFile(path, "a") as archive:
        archive.comment = b"Retain this archive encoding"
    frozen = path.read_bytes()
    ops = [{"tool": "write_data_to_excel", "args": {"start_cell": "A2", "data": [[6]]}}]
    preview = _call(path, ops)
    assert preview["ready_to_commit"]
    assert preview["preservation"]["changed_part_count"] == 0
    result = _commit(path, preview, ops)
    assert result["receipt"]["before_sha256"] == result["receipt"]["after_sha256"]
    assert path.read_bytes() == frozen


def test_rollback_restores_exact_original_package(tmp_path, monkeypatch):
    path = _book(tmp_path / "rollback.xlsx")
    frozen = path.read_bytes()
    preview = _call(path, create_snapshot=False)
    verify = changeset.workbook_module._verify_saved_workbook
    failed_after_replace = []

    def fail_first_live_verify(filepath):
        if Path(filepath) == path and not failed_after_replace:
            assert path.read_bytes() != frozen
            failed_after_replace.append(True)
            raise WorkbookError("Injected reopen failure")
        verify(filepath)

    monkeypatch.setattr(
        changeset.workbook_module, "_verify_saved_workbook", fail_first_live_verify
    )
    with pytest.raises(WorkbookError, match="original workbook was restored"):
        _commit(path, preview, create_snapshot=False)
    assert failed_after_replace
    assert path.read_bytes() == frozen and list(tmp_path.iterdir()) == [path]


@pytest.mark.parametrize(
    "limit", ["MAX_PACKAGE_BYTES", "MAX_SHEET_BYTES", "MAX_SOURCE_BYTES"]
)
def test_candidate_growth_respects_profile_limits(tmp_path, monkeypatch, limit):
    path = _book(tmp_path / "bounded.xlsx")
    preview = _call(path)
    before = path.read_bytes()
    parts = _members(path)
    limits = {
        "MAX_PACKAGE_BYTES": sum(map(len, parts.values())),
        "MAX_SHEET_BYTES": len(parts[INPUT_PART]),
        "MAX_SOURCE_BYTES": len(before),
    }
    monkeypatch.setattr(preserve, limit, limits[limit])
    with pytest.raises(PreconditionFailedError) as error:
        _commit(path, preview)
    assert error.value.details["reason"] == "candidate_package_limit"
    assert path.read_bytes() == before and list(tmp_path.iterdir()) == [path]


def test_workbook_wide_grid_limit_covers_baseline_and_insertions(tmp_path, monkeypatch):
    path = _book(tmp_path / "grid-limit.xlsx")
    before = path.read_bytes()
    original = _members(path)
    grid_count = sum(
        elem.tag in {f"{{{NS}}}row", f"{{{NS}}}c"}
        for name in (INPUT_PART, FORMULA_PART)
        for elem in ET.fromstring(original[name]).iter()
    )
    for limit in (grid_count - 1, grid_count):
        monkeypatch.setattr(preserve, "MAX_GRID_ELEMENTS", limit)
        with pytest.raises(PreconditionFailedError) as error:
            _call(path)
        assert error.value.details["reason"] == "grid_limit"
    assert path.read_bytes() == before


def test_rejects_case_aliased_package_parts(tmp_path):
    path = _book(tmp_path / "aliased.xlsx")
    with ZipFile(path, "a") as archive:
        archive.writestr("customXml/ITEM1.xml", b"<second/>")
    before = path.read_bytes()
    with pytest.raises(PreconditionFailedError) as error:
        _call(path)
    assert error.value.details["reason"] == "unsupported_zip_member"
    assert path.read_bytes() == before


@pytest.mark.parametrize(
    "new_cell_body",
    [b"<v>6</v><v>7</v>", b'<v xml:space="preserve">6</v>', b"metadata<v>6</v>"],
)
def test_rejects_noncanonical_target_payload(tmp_path, new_cell_body):
    path = _book(tmp_path / "ambiguous-cell.xlsx")
    parts = _members(path)
    cell = b'<c r="A2" t="n">'
    assert cell + b"<v>6</v>" in parts[INPUT_PART]
    parts[INPUT_PART] = parts[INPUT_PART].replace(
        cell + b"<v>6</v>", cell + new_cell_body
    )
    _write_members(path, parts)
    before = path.read_bytes()
    ops = [{"tool": "write_data_to_excel", "args": {"start_cell": "A2", "data": [[8]]}}]
    with pytest.raises(PreconditionFailedError) as error:
        _call(path, ops)
    assert error.value.details["reason"] == "unsupported_target_cell_encoding"
    assert path.read_bytes() == before
