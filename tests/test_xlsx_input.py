import os
from io import BytesIO
import struct
from zipfile import ZIP_BZIP2, ZIP_DEFLATED, ZipFile

import pytest
from openpyxl import Workbook

from excel_mcp import server, xlsx_input
from excel_mcp.data import quick_read, write_data
from excel_mcp.xlsx_input import WorkbookInputError, inspect_xlsx, load_workbook_checked


def make_input(tmp_path):
    path = tmp_path / "input.xlsx"
    book = Workbook()
    book.active.append(["Id", "Value"])
    book.active.append([1, "safe"])
    book.save(path)
    book.close()
    return path


@pytest.mark.parametrize("extension", [".xlsm", ".xltm", ".xltx", ".xls", ".XLSM"])
@pytest.mark.parametrize("operation", ["read", "write", "format", "preview"])
def test_only_xlsx_inputs_are_admitted_without_touching_unsupported_files(tmp_path, monkeypatch, extension, operation):
    source = make_input(tmp_path)
    path = tmp_path / ("unsupported" + extension)
    path.write_bytes(source.read_bytes())
    before = path.read_bytes()

    def never_load(*args, **kwargs):
        pytest.fail("Unsupported extensions must be refused before openpyxl")

    monkeypatch.setattr(xlsx_input, "_load_workbook", never_load)
    calls = {
        "read": lambda: server.quick_read(str(path)),
        "write": lambda: server.write_data_to_excel(str(path), "Sheet", [[2]]),
        "format": lambda: server.format_range(str(path), "Sheet", "A1", bold=True),
        "preview": lambda: server.write_data_to_excel(str(path), "Sheet", [[2]], dry_run=True),
    }
    import json
    payload = json.loads(calls[operation]())
    assert not payload["ok"] and payload["error"]["details"]["reason"] == "unsupported_extension", payload
    assert path.read_bytes() == before


def test_uppercase_xlsx_extension_stays_supported(tmp_path):
    source = make_input(tmp_path)
    path = tmp_path / "input.XLSX"
    path.write_bytes(source.read_bytes())
    assert quick_read(str(path))["rows"] == [[1, "safe"]]


@pytest.mark.parametrize("loader", [load_workbook_checked, inspect_xlsx])
def test_checked_input_cannot_hide_unsupported_target_extension_behind_symlink(tmp_path, loader):
    source = make_input(tmp_path)
    unsupported = source.with_suffix(".xlsm")
    source.rename(unsupported)
    source.symlink_to(unsupported)
    before = unsupported.read_bytes()
    with pytest.raises(WorkbookInputError) as error:
        result = loader(source)
        if result is not None:
            result.close()
    assert error.value.details["reason"] == "unsupported_extension"
    assert unsupported.read_bytes() == before


@pytest.mark.parametrize("setting,value,reason", [
    ("FILE_BYTES", "100", "file_bytes"), ("MEMBERS", "1", "member_count"),
    ("DIRECTORY_BYTES", "10", "directory_bytes"), ("TOTAL_BYTES", "100", "total_bytes"),
    ("MEMBER_BYTES", "100", "member_bytes"),
])
def test_limits_apply_before_any_xml_expansion(tmp_path, monkeypatch, setting, value, reason):
    path = make_input(tmp_path)
    original = path.read_bytes()
    monkeypatch.setenv(f"SHEETFORGE_XLSX_MAX_{setting}", value)

    def never_load(*args, **kwargs):
        pytest.fail("openpyxl must not run before admission")

    monkeypatch.setattr(xlsx_input, "_load_workbook", never_load)
    with pytest.raises(WorkbookInputError) as error:
        quick_read(str(path), max_rows=1)
    assert error.value.code == "xlsx_input_rejected"
    assert error.value.details["reason"] == reason
    with pytest.raises(WorkbookInputError):
        write_data(str(path), "Sheet", [["new"]])
    assert path.read_bytes() == original


def test_shared_string_size_is_checked_without_opening_the_xml_part(tmp_path, monkeypatch):
    base = make_input(tmp_path)
    with ZipFile(base) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        b"</Types>", b'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
    )
    with ZipFile(base, "w", ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
        archive.writestr("xl/sharedStrings.xml", b"x" * (2 * 1024 * 1024))
    monkeypatch.setenv("SHEETFORGE_XLSX_MAX_MEMBER_BYTES", str(1024 * 1024))
    original_open = ZipFile.open

    def no_shared_string_expansion(self, name, *args, **kwargs):
        assert getattr(name, "filename", name) != "xl/sharedStrings.xml"
        return original_open(self, name, *args, **kwargs)

    monkeypatch.setattr(ZipFile, "open", no_shared_string_expansion)
    with pytest.raises(WorkbookInputError) as error:
        quick_read(str(base), max_rows=1)
    assert error.value.details["reason"] == "member_bytes"


def test_forged_small_eocd_count_does_not_bypass_preallocation_count_limit(tmp_path, monkeypatch):
    path = make_input(tmp_path)
    content = bytearray(path.read_bytes())
    offset = content.rfind(b"PK\x05\x06")
    struct.pack_into("<HH", content, offset + 8, 1, 1)
    path.write_bytes(content)
    monkeypatch.setenv("SHEETFORGE_XLSX_MAX_MEMBERS", "2")

    def never_construct(*args, **kwargs):
        pytest.fail("ZipFile must not allocate unadmitted entry objects")

    monkeypatch.setattr(xlsx_input, "ZipFile", never_construct)
    with pytest.raises(WorkbookInputError) as error:
        inspect_xlsx(path)
    assert error.value.details["reason"] == "member_count"


def _zip64_input(*, sentinel=False, conflict=None):
    stream = BytesIO()
    with ZipFile(stream, "w") as archive:
        for index in range(20):
            archive.writestr(f"item-{index}.bin", b"")
    raw = stream.getvalue()
    end = raw.rfind(b"PK\x05\x06")
    _, _, _, _, count, size, offset, _ = struct.unpack_from("<4s4H2LH", raw, end)
    record = struct.pack("<4sQ2H2L4Q", b"PK\x06\x06", 44, 45, 45, 0, 0, count, count, size, offset)
    locator = struct.pack("<4sLQL", b"PK\x06\x07", 0, end, 1)
    classic_count, classic_size, classic_offset = count, size, offset
    if sentinel:
        classic_count, classic_size, classic_offset = 0xFFFF, 0xFFFFFFFF, 0xFFFFFFFF
    if conflict == "count":
        classic_count = 1
    elif conflict == "size":
        classic_size = 46 + sum(struct.unpack_from("<3H", raw, offset + 28))
    elif conflict == "offset":
        classic_offset = 0
    classic = struct.pack("<4s4H2LH", b"PK\x05\x06", 0, 0, classic_count, classic_count, classic_size, classic_offset, 0)
    return raw[:end] + record + locator + classic


@pytest.mark.parametrize("conflict", ["count", "size", "offset"])
def test_non_sentinel_zip64_conflicts_are_rejected_before_zipfile_allocation(tmp_path, monkeypatch, conflict):
    path = tmp_path / "conflict.xlsx"
    path.write_bytes(_zip64_input(conflict=conflict))
    with ZipFile(path) as archive:
        assert len(archive.infolist()) == 20

    def never_construct(*args, **kwargs):
        pytest.fail("Conflicting directory must be refused before ZipFile allocates entries")

    monkeypatch.setattr(xlsx_input, "ZipFile", never_construct)
    with pytest.raises(WorkbookInputError) as error:
        inspect_xlsx(path)
    assert error.value.details["reason"] == "inconsistent_zip64_directory"


@pytest.mark.parametrize("sentinel", [False, True])
@pytest.mark.parametrize("setting,value,reason", [("MEMBERS", "5", "member_count"), ("DIRECTORY_BYTES", "100", "directory_bytes")])
def test_both_zip64_forms_enforce_limits_before_allocation(tmp_path, monkeypatch, sentinel, setting, value, reason):
    path = tmp_path / "zip64.xlsx"
    path.write_bytes(_zip64_input(sentinel=sentinel))
    monkeypatch.setenv(f"SHEETFORGE_XLSX_MAX_{setting}", value)
    monkeypatch.setattr(xlsx_input, "ZipFile", lambda *a, **k: pytest.fail("Unadmitted ZipFile allocation"))
    with pytest.raises(WorkbookInputError) as error:
        inspect_xlsx(path)
    assert error.value.details["reason"] == reason


@pytest.mark.parametrize("sentinel", [False, True])
def test_matching_zip64_directories_without_and_with_sentinels_are_admitted(tmp_path, sentinel):
    path = tmp_path / "zip64.xlsx"
    path.write_bytes(_zip64_input(sentinel=sentinel))
    inspect_xlsx(path)


@pytest.mark.parametrize("variant", ["prepended", "comment_record"])
def test_zipfile_directory_rebasing_cannot_bypass_preinspection(tmp_path, monkeypatch, variant):
    path = make_input(tmp_path)
    if variant == "prepended":
        path.write_bytes(b"prepended" + path.read_bytes())
    else:
        with ZipFile(path, "a") as archive:
            archive.comment = struct.pack("<4s4H2LH", b"PK\x05\x06", 0, 0, 1, 1, 46, 0, 0)
    monkeypatch.setattr(xlsx_input, "ZipFile", lambda *a, **k: pytest.fail("Ambiguous directory allocation"))
    with pytest.raises(WorkbookInputError):
        inspect_xlsx(path)


def test_zip64_directory_is_supported_and_ambiguous_comments_fail_cleanly(tmp_path):
    path = make_input(tmp_path)
    raw = path.read_bytes()
    offset = raw.rfind(b"PK\x05\x06")
    _, _, _, disk_count, count, size, start, _ = struct.unpack_from("<4s4H2LH", raw, offset)
    record = struct.pack("<4sQ2H2L4Q", b"PK\x06\x06", 44, 45, 45, 0, 0, disk_count, count, size, start)
    locator = struct.pack("<4sLQL", b"PK\x06\x07", 0, offset, 1)
    eocd = struct.pack("<4s4H2LH", b"PK\x05\x06", 0, 0, 0xFFFF, 0xFFFF, 0xFFFFFFFF, 0xFFFFFFFF, 0)
    path.write_bytes(raw[:offset] + record + locator + eocd)
    inspect_xlsx(path)
    book = load_workbook_checked(path)
    assert book.active["B2"].value == "safe"
    book.close()
    normal = make_input(tmp_path)
    with ZipFile(normal, "a") as archive:
        archive.comment = b"ordinary comment"
    inspect_xlsx(normal)
    with ZipFile(normal, "a") as archive:
        archive.comment = b"comment PK\x05\x06 ignored"
    with pytest.raises(WorkbookInputError):
        inspect_xlsx(normal)


@pytest.mark.parametrize("extra,reason", [("../outside.xml", "unsafe_member_path"), ("/outside.xml", "unsafe_member_path"), ("xl/worksheets/sheet1.xml", "duplicate_members")])
def test_ambiguous_or_unsafe_zip_members_are_rejected(tmp_path, extra, reason):
    path = make_input(tmp_path)
    if reason == "duplicate_members":
        with pytest.warns(UserWarning, match="Duplicate name"):
            with ZipFile(path, "a") as archive:
                archive.writestr(extra, b"synthetic")
    else:
        with ZipFile(path, "a") as archive:
            archive.writestr(extra, b"synthetic")
    with pytest.raises(WorkbookInputError) as error:
        inspect_xlsx(path)
    assert error.value.details["reason"] == reason


def test_unsupported_compression_is_rejected_before_loading(tmp_path):
    path = make_input(tmp_path)
    with ZipFile(path, "a", ZIP_BZIP2) as archive:
        archive.writestr("custom.bin", b"synthetic")
    with pytest.raises(WorkbookInputError) as error:
        inspect_xlsx(path)
    assert error.value.details["reason"] == "unsupported_compression_or_encryption"


@pytest.mark.parametrize("value", ["0", "-1", "none", "1.5", "9999999999999999999999"])
def test_invalid_operator_configuration_has_no_unbounded_fallback(tmp_path, monkeypatch, value):
    path = make_input(tmp_path)
    monkeypatch.setenv("SHEETFORGE_XLSX_MAX_TOTAL_BYTES", value)
    with pytest.raises(WorkbookInputError) as error:
        inspect_xlsx(path)
    assert error.value.details["reason"] == "invalid_configuration"


@pytest.mark.parametrize("read_only", [False, True])
def test_checked_descriptor_is_owned_and_closed_with_workbook(tmp_path, monkeypatch, read_only):
    path = make_input(tmp_path)
    original = xlsx_input._open_checked
    handles = []

    def capture(filename):
        stream, stamp = original(filename)
        handles.append(stream)
        return stream, stamp

    monkeypatch.setattr(xlsx_input, "_open_checked", capture)
    book = load_workbook_checked(path, read_only=read_only)
    assert list(book.active.values)[1] == (1, "safe")
    if read_only:
        assert not handles[0].closed
    book.close()
    assert handles[0].closed


@pytest.mark.skipif(not hasattr(os, "mkfifo"), reason="requires POSIX FIFO support")
def test_nonregular_input_is_rejected_before_read(tmp_path):
    path = tmp_path / "pipe.xlsx"
    os.mkfifo(path)
    with pytest.raises(WorkbookInputError, match="before expansion"):
        inspect_xlsx(path)


def test_mcp_input_limit_error_is_structured_without_server_root_disclosure(tmp_path, monkeypatch):
    import json

    path = make_input(tmp_path)
    monkeypatch.setenv("SHEETFORGE_XLSX_MAX_MEMBER_BYTES", "100")
    result = json.loads(server.quick_read(str(path), max_rows=1))
    assert not result["ok"]
    assert result["error"]["code"] == "xlsx_input_rejected"
    assert str(tmp_path) not in result["error"]["message"]


def test_failed_load_closes_checked_descriptor_even_if_workbook_cleanup_fails(tmp_path, monkeypatch):
    path = make_input(tmp_path)
    original_open = xlsx_input._open_checked
    original_load = xlsx_input._load_workbook
    streams = []

    def changed_stamp(filename):
        stream, stamp = original_open(filename)
        streams.append(stream)
        return stream, (stamp[0] + 1, *stamp[1:])

    def broken_cleanup(*args, **kwargs):
        book = original_load(*args, **kwargs)

        def close():
            raise OSError("injected cleanup failure")

        book.close = close
        return book

    monkeypatch.setattr(xlsx_input, "_open_checked", changed_stamp)
    monkeypatch.setattr(xlsx_input, "_load_workbook", broken_cleanup)
    with pytest.raises(WorkbookInputError) as error:
        load_workbook_checked(path)
    assert error.value.details["reason"] == "source_changed_during_load"
    assert streams[0].closed
