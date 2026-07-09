"""The preserve-mode zip machinery: raw copy, determinism, atomic delivery
."""
from __future__ import annotations

import io
import os
import struct
import zipfile

import pytest

from openpyxl.preserve import zipio


def _sample_zip_bytes(entries=None):
    entries = entries or {"a.xml": b"<a/>" * 100, "b.bin": os.urandom(256)}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, payload in entries.items():
            z.writestr(name, payload)
    return buf.getvalue()


class TestRawCopy:

    def test_raw_copy_payloads_identical(self, fixture_copy, tmp_path):
        src = fixture_copy("gauntlet/gauntlet.xlsx")
        out = io.BytesIO()
        with zipfile.ZipFile(src) as zin, \
                zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                zipio.copy_entry(zin, info, zout)
        with zipfile.ZipFile(src) as zin, zipfile.ZipFile(io.BytesIO(out.getvalue())) as zv:
            assert zv.namelist() == zin.namelist()
            for name in zin.namelist():
                assert zv.read(name) == zin.read(name), name
            assert zv.testzip() is None

    def test_raw_copy_preserves_compressed_stream(self, fixture_copy):
        # stronger than payload identity: the deflate stream itself copies
        src = fixture_copy("minimal/minimal_clean.xlsx")
        out = io.BytesIO()
        with zipfile.ZipFile(src) as zin, \
                zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            infos = zin.infolist()
            for info in infos:
                zipio.copy_entry(zin, info, zout)
        with zipfile.ZipFile(io.BytesIO(out.getvalue())) as zv:
            for info in infos:
                if zipio.raw_copy_supported(info):
                    assert zv.getinfo(info.filename).compress_size == info.compress_size
                    assert zv.getinfo(info.filename).CRC == info.CRC

    def test_data_descriptor_entries_take_the_fallback(self, tmp_path):
        # hand-craft a zip whose entry sets GP flag bit 3 (data descriptor)
        payload = b"<x/>"
        import zlib
        comp = zlib.compressobj(-1, zlib.DEFLATED, -15)
        cdata = comp.compress(payload) + comp.flush()
        crc = zipfile.crc32(payload) & 0xFFFFFFFF
        name = b"dd.xml"
        # local header: sizes zeroed, bit 3 set, descriptor after payload
        local = (b"PK\x03\x04" + struct.pack("<HHHHH", 20, 0x8, 8, 0, 33)
                 + struct.pack("<LLL", 0, 0, 0) + struct.pack("<HH", len(name), 0)
                 + name + cdata + struct.pack("<LLL", crc, len(cdata), len(payload)))
        central = (b"PK\x01\x02" + struct.pack("<HHHHHH", 20, 20, 0x8, 8, 0, 33)
                   + struct.pack("<LLL", crc, len(cdata), len(payload))
                   + struct.pack("<HHHHHLL", len(name), 0, 0, 0, 0, 0, 0) + name)
        eocd = (b"PK\x05\x06" + struct.pack("<HHHHLLH", 0, 0, 1, 1,
                len(central), len(local), 0))
        raw = local + central + eocd
        zin = zipfile.ZipFile(io.BytesIO(raw))
        info = zin.getinfo("dd.xml")
        assert info.flag_bits & 0x8
        assert not zipio.raw_copy_supported(info)   # guard trips
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w") as zout:
            zipio.copy_entry(zin, info, zout)       # fallback recompression
        with zipfile.ZipFile(io.BytesIO(out.getvalue())) as zv:
            assert zv.read("dd.xml") == payload

    def test_exotic_compression_takes_the_fallback(self):
        info = zipfile.ZipInfo("x.bin")
        info.compress_type = 14  # LZMA
        assert not zipio.raw_copy_supported(info)


class TestDeterminism:

    def test_copy_output_is_deterministic(self, fixture_copy):
        src = fixture_copy("minimal/minimal_clean.xlsx")

        def rebuild():
            out = io.BytesIO()
            with zipfile.ZipFile(src) as zin, \
                    zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
                for info in zin.infolist():
                    zipio.copy_entry(zin, info, zout)
            return out.getvalue()

        assert rebuild() == rebuild()

    def test_write_entry_uses_fixed_metadata(self):
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w") as zout:
            zipio.write_entry(zout, "a.xml", b"<a/>")
        info = zipfile.ZipFile(io.BytesIO(out.getvalue())).getinfo("a.xml")
        assert info.date_time == zipio.FIXED_DATE_TIME


class TestDeliver:

    def test_path_delivery_is_atomic_over_existing_file(self, tmp_path):
        target = tmp_path / "out.xlsx"
        target.write_bytes(b"ORIGINAL")
        zipio.deliver(b"NEW CONTENT", str(target))
        assert target.read_bytes() == b"NEW CONTENT"
        assert list(tmp_path.iterdir()) == [target]  # no temp litter

    def test_failed_replace_leaves_original_intact(self, tmp_path, monkeypatch):
        target = tmp_path / "out.xlsx"
        target.write_bytes(b"ORIGINAL")

        def boom(src, dst):
            raise OSError("simulated crash mid-rename")

        monkeypatch.setattr(zipio.os, "replace", boom)
        with pytest.raises(OSError, match="simulated crash"):
            zipio.deliver(b"NEW", str(target))
        assert target.read_bytes() == b"ORIGINAL"          # original survives
        assert list(tmp_path.iterdir()) == [target]        # temp cleaned up

    def test_filelike_delivery_seek_write_truncate(self):
        fh = io.BytesIO(b"OLD MUCH LONGER CONTENT THAN NEW")
        fh.seek(7)  # arbitrary position, as pandas may leave it
        zipio.deliver(b"NEW", fh)
        assert fh.getvalue() == b"NEW"  # truncated, not partially overwritten
