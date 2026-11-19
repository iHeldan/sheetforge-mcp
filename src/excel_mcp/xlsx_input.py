"""Bound ZIP metadata before openpyxl can allocate or decompress workbook data."""

from dataclasses import dataclass, fields
import os
from pathlib import Path, PurePosixPath
import stat
import struct
from zipfile import BadZipFile, ZIP_DEFLATED, ZIP_STORED, ZipFile

from openpyxl import load_workbook as _load_workbook
from openpyxl.reader.excel import SUPPORTED_FORMATS

from .exceptions import PreconditionFailedError, WorkbookError


class WorkbookInputError(PreconditionFailedError, WorkbookError):
    """A workbook was refused before expansion under the server's input policy."""


@dataclass(frozen=True)
class InputLimits:
    file_bytes: int = 128 * 1024 * 1024
    total_bytes: int = 256 * 1024 * 1024
    member_bytes: int = 64 * 1024 * 1024
    directory_bytes: int = 16 * 1024 * 1024
    members: int = 20_000


def _fail(reason, **details):
    raise WorkbookInputError(
        "XLSX input was refused before expansion; inspect the file or ask the server owner to review its input limits",
        code="xlsx_input_rejected", details={"reason": reason, **details},
    )


def input_limits() -> InputLimits:
    values = {}
    for field in fields(InputLimits):
        name = f"SHEETFORGE_XLSX_MAX_{field.name.upper()}"
        raw = os.environ.get(name)
        if raw is None:
            continue
        raw = raw.strip()
        if not raw or not raw.isascii() or not raw.isdecimal() or len(raw) > 19:
            _fail("invalid_configuration", setting=name)
        value = int(raw)
        if not 0 < value <= 2**63 - 1:
            _fail("invalid_configuration", setting=name)
        values[field.name] = value
    return InputLimits(**values)


def _directory(stream, size, limits):
    stream.seek(max(0, size - 65_557))
    tail = stream.read(65_557)
    offset = len(tail)
    while True:
        offset = tail.rfind(b"PK\x05\x06", 0, offset)
        if offset < 0:
            _fail("invalid_zip_directory")
        if offset + 22 <= len(tail) and offset + 22 + struct.unpack_from("<H", tail, offset + 20)[0] == len(tail):
            break
    _, disk, cd_disk, disk_count, count, cd_size, cd_offset, _ = struct.unpack_from("<4s4H2LH", tail, offset)
    end = size - len(tail) + offset
    if disk or cd_disk or disk_count != count:
        _fail("unsupported_multidisk_zip")
    if count == 0xFFFF or cd_size == 0xFFFFFFFF or cd_offset == 0xFFFFFFFF:
        stream.seek(max(0, end - 20))
        locator = stream.read(20)
        if len(locator) != 20:
            _fail("invalid_zip64_directory")
        signature, disk, position, disks = struct.unpack("<4sLQL", locator)
        if signature != b"PK\x06\x07" or disk or disks != 1 or position + 56 > end - 20:
            _fail("invalid_zip64_directory")
        stream.seek(position)
        header = stream.read(56)
        if len(header) != 56:
            _fail("invalid_zip64_directory")
        signature, record_size, _, _, disk, cd_disk, disk_count, count, cd_size, cd_offset = struct.unpack("<4sQ2H2L4Q", header)
        if signature != b"PK\x06\x06" or record_size < 44 or position + 12 + record_size != end - 20:
            _fail("invalid_zip64_directory")
        if disk or cd_disk or disk_count != count:
            _fail("unsupported_multidisk_zip")
        end = position
    if count > limits.members:
        _fail("member_count", observed=count, limit=limits.members)
    if cd_size > limits.directory_bytes:
        _fail("directory_bytes", observed=cd_size, limit=limits.directory_bytes)
    if cd_offset + cd_size > end:
        _fail("invalid_zip_directory")

    # Count actual records before ZipFile constructs its per-entry Python objects;
    # a forged EOCD entry count must not bypass the allocation bound.
    stream.seek(cd_offset)
    consumed = actual_count = 0
    while consumed < cd_size:
        header = stream.read(46)
        if len(header) != 46 or header[:4] != b"PK\x01\x02":
            _fail("invalid_zip_directory")
        lengths = struct.unpack_from("<3H", header, 28)
        consumed += 46 + sum(lengths)
        actual_count += 1
        if actual_count > limits.members:
            _fail("member_count", observed=actual_count, limit=limits.members)
        if consumed > cd_size:
            _fail("invalid_zip_directory")
        stream.seek(sum(lengths), 1)
    if actual_count != count:
        _fail("invalid_zip_directory")


def _inspect(stream, limits):
    info = os.fstat(stream.fileno())
    if not stat.S_ISREG(info.st_mode):
        _fail("not_regular_file")
    if info.st_size > limits.file_bytes:
        _fail("file_bytes", observed=info.st_size, limit=limits.file_bytes)
    _directory(stream, info.st_size, limits)
    stream.seek(0)
    with ZipFile(stream) as archive:
        entries = archive.infolist()
        if len(entries) > limits.members:
            _fail("member_count", observed=len(entries), limit=limits.members)
        names = set()
        total = 0
        for entry in entries:
            name = PurePosixPath(entry.filename)
            if entry.filename in names:
                _fail("duplicate_members")
            names.add(entry.filename)
            if name.is_absolute() or ".." in name.parts or "\\" in entry.filename or "\0" in entry.orig_filename:
                _fail("unsafe_member_path")
            if entry.flag_bits & 1 or entry.compress_type not in {ZIP_STORED, ZIP_DEFLATED}:
                _fail("unsupported_compression_or_encryption")
            if entry.file_size > limits.member_bytes:
                _fail("member_bytes", observed=entry.file_size, limit=limits.member_bytes)
            total += entry.file_size
            if total > limits.total_bytes:
                _fail("total_bytes", observed=total, limit=limits.total_bytes)
    stream.seek(0)
    return (info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns, info.st_ctime_ns)


def _open_checked(path):
    path = Path(path).resolve(strict=True)
    if not path.is_file():
        _fail("not_regular_file")
    flags = os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0) | getattr(os, "O_NONBLOCK", 0)
    stream = os.fdopen(os.open(path, flags), "rb")
    try:
        stamp = _inspect(stream, input_limits())
        return stream, stamp
    except (BadZipFile, struct.error) as error:
        stream.close()
        raise WorkbookInputError("Invalid XLSX archive", code="xlsx_input_rejected", details={"reason": "invalid_zip"}) from error
    except BaseException:
        stream.close()
        raise


def inspect_xlsx(path):
    stream, _ = _open_checked(path)
    stream.close()


def load_workbook_checked(filename, **kwargs):
    if Path(filename).suffix.lower() not in SUPPORTED_FORMATS:
        _fail("unsupported_extension")
    stream, stamp = _open_checked(filename)
    book = None
    try:
        book = _load_workbook(stream, **kwargs)
        info = os.fstat(stream.fileno())
        if (info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns, info.st_ctime_ns) != stamp:
            _fail("source_changed_during_load")
        if kwargs.get("read_only", False):
            close = book.close

            def close_stream():
                try:
                    close()
                finally:
                    stream.close()

            book.close = close_stream
        else:
            stream.close()
        return book
    except BaseException:
        try:
            if book is not None:
                book.close()
        finally:
            stream.close()
        raise
