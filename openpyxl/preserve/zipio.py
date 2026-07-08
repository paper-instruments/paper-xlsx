# paper-xlsx: deterministic, atomic zip writing + raw compressed-stream copy
# (CONVENTIONS §7; PR-0 D4/D10; evidence: OPEN-QUESTIONS Q1)

"""Zip-layer machinery for the preserve-mode save.

- Deterministic output: every entry gets a fixed timestamp and uniform
  attributes, so part payloads (and, on the raw-copy path, compressed
  streams) are reproducible run-to-run.
- Raw compressed-stream copy: untouched parts are copied without
  recompression (measured 235x faster and byte-identical); guarded by the
  D10 conditions, with transparent fallback to recompression.
- Atomic targets: path targets are written temp-file-then-``os.replace``
  (in-place truncation is the measured corruption hazard); file-like targets
  are built fully in memory and written in one seek/write/truncate pass.
"""

import io
import os
import struct
import tempfile
import zipfile

# Fixed timestamp for deterministic archives (zip epoch).
FIXED_DATE_TIME = (1980, 1, 1, 0, 0, 0)
_EXTERNAL_ATTR = 0o600 << 16


def _probe_private_zipfile_api():
    """The raw-copy fast path uses CPython-private ZipFile internals
    (validated on 3.9-3.13). Probe once; absence disables the fast path."""
    try:
        buf = io.BytesIO()
        zf = zipfile.ZipFile(buf, "w")
        ok = all(hasattr(zf, attr) for attr in
                 ("fp", "filelist", "NameToInfo", "start_dir", "_writecheck"))
        ok = ok and hasattr(zipfile.ZipInfo("x"), "FileHeader")
        zf.close()
        return ok
    except Exception:
        return False


RAW_COPY_AVAILABLE = _probe_private_zipfile_api()

_ZIP64_LIMIT = 0xFFFFFFFF


def raw_copy_supported(info):
    """D10 guards: data-descriptor entries (GP flag bit 3), zip64-sized
    entries, and exotic compression methods take the recompression fallback."""
    if not RAW_COPY_AVAILABLE:
        return False
    if info.flag_bits & 0x8:            # data descriptor: sizes live after payload
        return False
    # entries not read from an archive may lack size/offset attributes
    compress_size = getattr(info, "compress_size", None)
    header_offset = getattr(info, "header_offset", None)
    if compress_size is None or header_offset is None:
        return False
    if compress_size >= _ZIP64_LIMIT or info.file_size >= _ZIP64_LIMIT:
        return False
    if header_offset >= _ZIP64_LIMIT:
        return False
    if info.compress_type not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
        return False
    return True


def _read_raw_stream(zin, info):
    """Read one entry's compressed byte stream straight from the archive."""
    f = zin.fp
    f.seek(info.header_offset)
    header = f.read(30)
    if header[:4] != b"PK\x03\x04":
        raise zipfile.BadZipFile(
            "bad local file header for {0!r}".format(info.filename))
    name_len, extra_len = struct.unpack("<HH", header[26:30])
    f.seek(info.header_offset + 30 + name_len + extra_len)
    return f.read(info.compress_size)


def copy_entry(zin, info, zout):
    """Copy one entry from ``zin`` into ``zout``, raw when possible.

    Payload bytes are identical either way; the raw path also preserves the
    compressed stream. Entry metadata is normalized for determinism.
    """
    if raw_copy_supported(info):
        payload_stream = _read_raw_stream(zin, info)
        new = zipfile.ZipInfo(info.filename, date_time=FIXED_DATE_TIME)
        new.compress_type = info.compress_type
        new.external_attr = _EXTERNAL_ATTR
        new.file_size = info.file_size
        new.CRC = info.CRC
        new.compress_size = info.compress_size
        zout._writecheck(new)
        zout._didModify = True
        new.header_offset = zout.fp.tell()
        zout.fp.write(new.FileHeader())
        zout.fp.write(payload_stream)
        zout.start_dir = zout.fp.tell()
        zout.filelist.append(new)
        zout.NameToInfo[new.filename] = new
    else:
        write_entry(zout, info.filename, zin.read(info.filename),
                    compress_type=info.compress_type
                    if info.compress_type in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED)
                    else zipfile.ZIP_DEFLATED)


def write_entry(zout, name, payload, compress_type=zipfile.ZIP_DEFLATED):
    """Write one entry with deterministic metadata."""
    info = zipfile.ZipInfo(name, date_time=FIXED_DATE_TIME)
    info.compress_type = compress_type
    info.external_attr = _EXTERNAL_ATTR
    zout.writestr(info, payload)


def build_archive_bytes(build):
    """Run ``build(zout)`` against an in-memory archive; return its bytes."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as zout:
        build(zout)
    return buf.getvalue()


def deliver(data, target):
    """Deliver finished archive bytes to a path or a binary file-like.

    Path targets: temp file in the same directory + ``os.replace`` — the
    original survives any mid-write crash (never in-place truncation).
    File-like targets: single seek(0)/write/truncate choreography (the pandas
    handle dance); the in-memory build above is the atomicity mechanism.
    """
    if hasattr(target, "write"):
        if hasattr(target, "seek"):
            target.seek(0)
        target.write(data)
        if hasattr(target, "truncate"):
            target.truncate()
        return

    target = os.fspath(target)
    directory = os.path.dirname(os.path.abspath(target))
    fd, tmp_path = tempfile.mkstemp(prefix=".paper_save_", suffix=".tmp",
                                    dir=directory)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
        os.replace(tmp_path, target)
    except BaseException:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise
