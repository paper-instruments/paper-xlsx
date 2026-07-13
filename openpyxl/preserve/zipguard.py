# paper-xlsx: ZIP and OPC custody preflight

import binascii
import io
import struct
import zipfile
import zlib
from collections import Counter

from openpyxl.errors import UnsupportedStructureError


MAX_ENTRIES = 10000
MAX_PART_BYTES = 256 * 1024 * 1024
MAX_TOTAL_BYTES = 512 * 1024 * 1024
RATIO_CAP = 500
RATIO_FLOOR = 64 * 1024 * 1024

_LOCAL_HEADER = struct.Struct("<4s5H3L2H")
_LOCAL_SIGNATURE = b"PK\x03\x04"
_ALLOWED_FLAGS = 0x000E | 0x0800  # deflate options, descriptor, UTF-8 name
_READ_CHUNK = 1024 * 1024


def _refuse(context, detail, *, kind="invalid-zip-package", anchor=None):
    raise UnsupportedStructureError(
        "{0}: {1} Nothing was loaded.".format(context, detail),
        kind=kind,
        anchor=anchor,
    )


def _ascii_name_key(name):
    """OPC part-name comparison is ASCII-case-insensitive, not casefold."""
    return "".join(chr(ord(ch) + 32) if "A" <= ch <= "Z" else ch
                   for ch in name)


def _check_names(infos, context):
    names = [info.filename for info in infos]
    duplicates = sorted(name for name, count in Counter(names).items()
                        if count > 1)
    if duplicates:
        _refuse(
            context,
            "the archive contains duplicate ZIP entry names ({0}); readers may "
            "choose different copies.".format(", ".join(duplicates)),
            kind="duplicate-zip-entry",
        )

    by_key = {}
    collisions = []
    for name in names:
        key = _ascii_name_key(name)
        previous = by_key.setdefault(key, name)
        if previous != name:
            collisions.append((previous, name))
    if collisions:
        rendered = ", ".join("{0!r} / {1!r}".format(a, b)
                             for a, b in collisions)
        _refuse(
            context,
            "the archive contains ASCII-case-colliding OPC part names "
            "({0}); they identify the same part.".format(rendered),
            kind="case-colliding-opc-part",
        )
    return names


def _check_declared_limits(infos, context):
    if len(infos) > MAX_ENTRIES:
        _refuse(context, "the archive declares {0} entries, past the "
                "{1}-entry cap.".format(len(infos), MAX_ENTRIES),
                kind="zip-entry-limit")
    total = sum(info.file_size for info in infos)
    if total > MAX_TOTAL_BYTES:
        _refuse(context, "the archive declares {0} aggregate uncompressed "
                "bytes, past the {1}-byte cap.".format(
                    total, MAX_TOTAL_BYTES), kind="zip-size-limit")
    for info in infos:
        if info.file_size > MAX_PART_BYTES:
            _refuse(context, "part {0!r} declares {1} uncompressed bytes, "
                    "past the {2}-byte part cap.".format(
                        info.filename, info.file_size, MAX_PART_BYTES),
                    kind="zip-size-limit", anchor=info.filename)
        if (info.file_size > RATIO_FLOOR and info.compress_size > 0
                and info.file_size / info.compress_size > RATIO_CAP):
            _refuse(context, "part {0!r} declares an inflation ratio past "
                    "the {1}x cap.".format(info.filename, RATIO_CAP),
                    kind="zip-ratio-limit", anchor=info.filename)


def _read_entry_layout(archive, info, context):
    fp = archive.fp
    fp.seek(info.header_offset)
    header = fp.read(_LOCAL_HEADER.size)
    if len(header) != _LOCAL_HEADER.size:
        _refuse(context, "part {0!r} has a truncated local header."
                .format(info.filename), anchor=info.filename)
    (signature, _version, flags, method, _time, _date, _crc,
     _compressed_size, _file_size, name_len, extra_len) = \
        _LOCAL_HEADER.unpack(header)
    if signature != _LOCAL_SIGNATURE:
        _refuse(context, "part {0!r} has an invalid local-header signature."
                .format(info.filename), anchor=info.filename)
    if flags != info.flag_bits:
        _refuse(context, "part {0!r} disagrees between local and central "
                "general-purpose flags.".format(info.filename),
                anchor=info.filename)
    if flags & ~_ALLOWED_FLAGS:
        if flags & 0x1:
            detail = "part {0!r} is ZIP-encrypted"
        else:
            detail = "part {0!r} uses unsupported ZIP flags 0x{1:04x}"
        _refuse(context, detail.format(info.filename, flags),
                kind="unsupported-zip-flags", anchor=info.filename)
    if method != info.compress_type:
        _refuse(context, "part {0!r} disagrees between local and central "
                "compression methods.".format(info.filename),
                anchor=info.filename)
    if method not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
        _refuse(context, "part {0!r} uses unsupported ZIP compression "
                "method {1}.".format(info.filename, method),
                kind="unsupported-zip-method", anchor=info.filename)

    local_name = fp.read(name_len)
    extra = fp.read(extra_len)
    if len(local_name) != name_len or len(extra) != extra_len:
        _refuse(context, "part {0!r} has a truncated local header."
                .format(info.filename), anchor=info.filename)
    encoding = "utf-8" if flags & 0x0800 else "cp437"
    try:
        decoded_name = local_name.decode(encoding)
    except UnicodeDecodeError:
        _refuse(context, "part {0!r} has an invalid encoded local name."
                .format(info.filename), anchor=info.filename)
    if decoded_name != info.orig_filename:
        _refuse(context, "part {0!r} has different local and central names."
                .format(info.filename), anchor=info.filename)

    # CRC/size disagreement is normalized safely: custody follows the central
    # directory, actual payload inflation is verified below, and raw copy
    # falls back to recompression. This preserves the established ZIP-confusion
    # normalization path without trusting either set of declared checksums.

    data_start = info.header_offset + _LOCAL_HEADER.size + name_len + extra_len
    data_end = data_start + info.compress_size
    if data_end > archive.start_dir:
        _refuse(context, "part {0!r} extends into the central directory."
                .format(info.filename), anchor=info.filename)
    return data_start, data_end, method


def _consume_output(info, output, state, context):
    total, crc = state
    total += len(output)
    if total > info.file_size or total > MAX_PART_BYTES:
        _refuse(context, "part {0!r} inflates past its declared size."
                .format(info.filename), anchor=info.filename)
    return total, binascii.crc32(output, crc)


def _verify_payload(archive, info, data_start, method, context):
    fp = archive.fp
    fp.seek(data_start)
    remaining = info.compress_size
    total = 0
    crc = 0
    decompressor = zlib.decompressobj(-15) \
        if method == zipfile.ZIP_DEFLATED else None
    try:
        while remaining:
            raw = fp.read(min(_READ_CHUNK, remaining))
            if not raw:
                _refuse(context, "part {0!r} has a truncated compressed "
                        "stream.".format(info.filename), anchor=info.filename)
            remaining -= len(raw)
            if decompressor is None:
                total, crc = _consume_output(
                    info, raw, (total, crc), context)
                continue
            pending = raw
            while pending:
                output = decompressor.decompress(pending, _READ_CHUNK)
                pending = decompressor.unconsumed_tail
                total, crc = _consume_output(
                    info, output, (total, crc), context)
        if decompressor is not None:
            while decompressor.unconsumed_tail:
                pending = decompressor.unconsumed_tail
                output = decompressor.decompress(pending, _READ_CHUNK)
                total, crc = _consume_output(
                    info, output, (total, crc), context)
            if not decompressor.eof or decompressor.unused_data:
                _refuse(context, "part {0!r} has an invalid deflate stream."
                        .format(info.filename), anchor=info.filename)
    except zlib.error as exc:
        _refuse(context, "part {0!r} has an invalid deflate stream ({1})."
                .format(info.filename, exc), anchor=info.filename)

    if total != info.file_size:
        _refuse(context, "part {0!r} inflates to {1} bytes, not the declared "
                "{2}.".format(info.filename, total, info.file_size),
                anchor=info.filename)
    if crc & 0xFFFFFFFF != info.CRC:
        _refuse(context, "part {0!r} fails its CRC check."
                .format(info.filename), kind="zip-crc-mismatch",
                anchor=info.filename)
    return total


def validate_archive(archive, *, context="preserve-mode package"):
    """Validate one ZIP before any workbook model is constructed."""
    infos = archive.infolist()
    names = _check_names(infos, context)
    _check_declared_limits(infos, context)

    layouts = []
    original_position = archive.fp.tell()
    try:
        for info in infos:
            data_start, data_end, method = _read_entry_layout(
                archive, info, context)
            layouts.append((info.header_offset, data_end, info, data_start,
                            method))
        previous_end = 0
        for header_start, data_end, info, _data_start, _method in sorted(
                layouts, key=lambda item: item[0]):
            if header_start < previous_end:
                _refuse(context, "part {0!r} overlaps another ZIP entry."
                        .format(info.filename), anchor=info.filename)
            previous_end = data_end

        actual_total = 0
        for _start, _end, info, data_start, method in layouts:
            actual_total += _verify_payload(
                archive, info, data_start, method, context)
            if actual_total > MAX_TOTAL_BYTES:
                _refuse(context, "actual aggregate inflation exceeds the "
                        "{0}-byte cap.".format(MAX_TOTAL_BYTES),
                        kind="zip-size-limit")
    finally:
        archive.fp.seek(original_position)
    return names


def validate_package_bytes(data, *, context="workbook package"):
    try:
        with zipfile.ZipFile(io.BytesIO(data), "r") as archive:
            return validate_archive(archive, context=context)
    except UnsupportedStructureError:
        raise
    except (zipfile.BadZipFile, LargeZipFile, RuntimeError,
            NotImplementedError) as exc:
        _refuse(context, "invalid ZIP structure ({0}).".format(exc))


LargeZipFile = zipfile.LargeZipFile
