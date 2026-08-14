"""Supported delivery helpers and package hardening."""
from __future__ import annotations

import io
import zipfile

import pytest

from openpyxl import load_workbook
from openpyxl.errors import AmbiguousTargetError, TargetNotFoundError


def _with_pivot_graph(workbook, pivots, *, refresh_on_load=False):
    """Attach a small relationship-complete pivot graph to source bytes.

    The preservation API indexes package relationships, not openpyxl's pivot
    model. Keeping the fixture at the package layer makes each resolution
    contract explicit without relying on the upstream pivot deserializer.
    ``pivots`` is an iterable of ``(name, cache_id)`` pairs.
    """
    workbook_xml = (
        b'<pivotCaches xmlns:r="http://schemas.openxmlformats.org/office'
        b'Document/2006/relationships">'
        + b"".join(
            b'<pivotCache cacheId="%s" r:id="rIdCache%s"/>'
            % (cache_id.encode("ascii"), cache_id.encode("ascii"))
            for cache_id in sorted({cache_id for _name, cache_id in pivots})
        )
        + b"</pivotCaches>"
    )
    workbook_rels = b"".join(
        b'<Relationship Id="rIdCache%s" Type="http://schemas.openxml'
        b'formats.org/officeDocument/2006/relationships/'
        b'pivotCacheDefinition" Target="pivotCache/'
        b'pivotCacheDefinition%s.xml"/>'
        % (cache_id.encode("ascii"), cache_id.encode("ascii"))
        for cache_id in sorted({cache_id for _name, cache_id in pivots})
    )
    sheet_rels = b"".join(
        b'<Relationship Id="rIdPivot%s" Type="http://schemas.openxml'
        b'formats.org/officeDocument/2006/relationships/pivotTable" '
        b'Target="../pivotTables/pivotTable%s.xml"/>'
        % (str(index).encode("ascii"), str(index).encode("ascii"))
        for index, _pivot in enumerate(pivots, 1)
    )
    cache_ids = sorted({cache_id for _name, cache_id in pivots})
    source = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(workbook._paper_source)) as zin, \
            zipfile.ZipFile(source, "w") as zout:
        for info in zin.infolist():
            payload = zin.read(info.filename)
            if info.filename == "xl/workbook.xml":
                payload = payload.replace(b"</workbook>",
                                          workbook_xml + b"</workbook>")
            elif info.filename == "xl/_rels/workbook.xml.rels":
                payload = payload.replace(b"</Relationships>",
                                          workbook_rels
                                          + b"</Relationships>")
            zout.writestr(info, payload)
        zout.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            b'<Relationships xmlns="http://schemas.openxmlformats.org/'
            b'package/2006/relationships">' + sheet_rels
            + b"</Relationships>")
        for index, (name, cache_id) in enumerate(pivots, 1):
            zout.writestr(
                "xl/pivotTables/pivotTable%s.xml" % index,
                b'<pivotTableDefinition xmlns="http://schemas.openxml'
                b'formats.org/spreadsheetml/2006/main" name="'
                + name.encode("utf-8") + b'" cacheId="'
                + cache_id.encode("ascii") + b'"/>')
        refresh = b' refreshOnLoad="1"' if refresh_on_load else b""
        for cache_id in cache_ids:
            zout.writestr(
                "xl/pivotCache/pivotCacheDefinition%s.xml" % cache_id,
                b'<pivotCacheDefinition xmlns="http://schemas.openxml'
                b'formats.org/spreadsheetml/2006/main" recordCount="0"'
                + refresh + b'><cacheSource type="worksheet"/>'
                b'</pivotCacheDefinition>')
    workbook._paper_source = source.getvalue()


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


def test_targeted_pivots_follow_relationships_and_deduplicate_shared_cache(
        fixture_copy, tmp_path):
    wb = load_workbook(
        fixture_copy("minimal/minimal_clean.xlsx"), preserve=True)
    _with_pivot_graph(wb, [("SalesPivot", "1"), ("MarginPivot", "1")])

    assert wb.set_pivot_refresh_on_load(
        pivots=["Sheet1!SalesPivot", "MarginPivot"]
    ) == ["xl/pivotCache/pivotCacheDefinition1.xml"]
    receipt = wb.save(tmp_path / "targeted-pivots.xlsx", receipt=True)

    assert [effect for effect in receipt.derived_effects
            if effect["kind"] == "pivot_refresh_on_load_enabled"] == [{
                "kind": "pivot_refresh_on_load_enabled",
                "part": "xl/pivotCache/pivotCacheDefinition1.xml",
                "cause": "explicit_request",
            }]


def test_targeted_pivot_reports_ambiguity_and_missing_names(fixture_copy):
    wb = load_workbook(
        fixture_copy("minimal/minimal_clean.xlsx"), preserve=True)
    _with_pivot_graph(wb, [("Duplicate", "1"), ("Duplicate", "2")])

    with pytest.raises(AmbiguousTargetError, match="sheet-qualified"):
        wb.set_pivot_refresh_on_load(pivots=["Duplicate"])
    with pytest.raises(TargetNotFoundError, match="Missing"):
        wb.set_pivot_refresh_on_load(pivots=["Missing"])


def test_targeted_pivot_is_idempotent_when_refresh_is_already_enabled(
        fixture_copy, tmp_path):
    wb = load_workbook(
        fixture_copy("minimal/minimal_clean.xlsx"), preserve=True)
    _with_pivot_graph(wb, [("ReadyPivot", "1")], refresh_on_load=True)
    original = wb._paper_source

    assert wb.set_pivot_refresh_on_load(pivots=["ReadyPivot"]) == [
        "xl/pivotCache/pivotCacheDefinition1.xml"]
    receipt = wb.save(tmp_path / "already-enabled.xlsx", receipt=True)

    with zipfile.ZipFile(io.BytesIO(original)) as before, \
            zipfile.ZipFile(tmp_path / "already-enabled.xlsx") as after:
        part = "xl/pivotCache/pivotCacheDefinition1.xml"
        assert after.read(part) == before.read(part)
    assert not any(effect["kind"] == "pivot_refresh_on_load_enabled"
                   for effect in receipt.derived_effects)


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
