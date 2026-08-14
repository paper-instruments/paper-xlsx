"""Supported delivery helpers and package hardening."""
from __future__ import annotations

import io
import zipfile

import pytest

from openpyxl import load_workbook


def test_copy_format_through_the_splice(fixture_copy, tmp_path):
    from openpyxl.preserve import copy_format
    from openpyxl.styles import Font

    wb = load_workbook(
        fixture_copy("features/schedule.xlsx"), preserve=True)
    ws = wb["Schedule"]
    ws["B2"].font = Font(bold=True, italic=True)
    assert copy_format(ws, "B2", "B3:B4") == 2
    out = tmp_path / "o.xlsx"
    wb.save(out)
    wb2 = load_workbook(out)
    assert wb2["Schedule"]["B3"].font.bold is True
    assert wb2["Schedule"]["B4"].font.italic is True


def test_pivot_refresh_requires_explicit_scope(fixture_copy):
    wb = load_workbook(
        fixture_copy("minimal/minimal_clean.xlsx"), preserve=True)
    with pytest.raises(ValueError, match="exactly one"):
        wb.set_pivot_refresh_on_load()


def test_all_pivot_refresh_patches_only_root_attribute(
        fixture_copy, tmp_path):
    src = fixture_copy("minimal/minimal_clean.xlsx")
    crafted = tmp_path / "pivot.xlsx"
    cache = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             b'<pivotCacheDefinition xmlns="http://schemas.openxml'
             b'formats.org/spreadsheetml/2006/main" r:id="rId1" '
             b'xmlns:r="http://schemas.openxmlformats.org/office'
             b'Document/2006/relationships" recordCount="2">'
             b'<cacheSource type="worksheet"/></pivotCacheDefinition>')
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(crafted, "w") as zout:
        for info in zin.infolist():
            zout.writestr(info, zin.read(info.filename))
        zout.writestr("xl/pivotCache/pivotCacheDefinition1.xml", cache)

    wb = load_workbook(crafted, preserve=True)
    assert wb.set_pivot_refresh_on_load(all=True) == [
        "xl/pivotCache/pivotCacheDefinition1.xml"]
    out = tmp_path / "o.xlsx"
    wb.save(out)
    with zipfile.ZipFile(out) as archive:
        payload = archive.read("xl/pivotCache/pivotCacheDefinition1.xml")
    assert payload == cache.replace(
        b' recordCount="2"', b' recordCount="2" refreshOnLoad="1"')


def test_pivot_refresh_requires_preserve(fixture_copy):
    wb = load_workbook(
        fixture_copy("minimal/minimal_clean.xlsx"), preserve=False)
    with pytest.raises(ValueError, match="preserve"):
        wb.set_pivot_refresh_on_load(all=True)


def test_highly_compressed_valid_extra_part_is_not_an_eligibility_error(
        fixture_copy, tmp_path):
    src = fixture_copy("minimal/minimal_clean.xlsx")
    crafted = tmp_path / "compressed.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(
            crafted, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info.filename, zin.read(info.filename))
        zout.writestr("xl/media/blob.bin", b"\x00" * (2 << 20))
    wb = load_workbook(crafted, preserve=True)
    assert wb.active is not None


def test_zip_confusion_is_normalized(fixture_copy, tmp_path):
    src = fixture_copy("features/schedule.xlsx")
    with open(src, "rb") as handle:
        data = bytearray(handle.read())
    with zipfile.ZipFile(io.BytesIO(bytes(data))) as archive:
        info = archive.getinfo("docProps/core.xml")
    data[info.header_offset + 14] ^= 0xFF
    crafted = tmp_path / "confused.xlsx"
    with open(crafted, "wb") as handle:
        handle.write(data)
    wb = load_workbook(crafted, preserve=True)
    wb["Schedule"]["A2"] = "edit"
    out = tmp_path / "o.xlsx"
    wb.save(out)
    with zipfile.ZipFile(out) as archive:
        assert b"cp:coreProperties" in archive.read("docProps/core.xml")


def test_spooled_save_is_correct_zip(fixture_copy, tmp_path):
    src = fixture_copy("features/schedule.xlsx")
    wb = load_workbook(src, preserve=True)
    wb["Schedule"]["A2"] = "spooled"
    out = tmp_path / "o.xlsx"
    wb.save(out)
    with zipfile.ZipFile(out) as archive:
        assert archive.testzip() is None
    assert load_workbook(out)["Schedule"]["A2"].value == "spooled"
