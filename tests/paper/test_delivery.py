"""Supported delivery helpers and package hardening."""
from __future__ import annotations

import io
import warnings
import zipfile

import pytest

from openpyxl import load_workbook
from openpyxl.errors import (
    AmbiguousTargetError,
    ProtectedWriteWarning,
    TargetNotFoundError,
    UnsupportedStructureError,
)


def _with_pivot_graph(workbook, pivots, *, refresh_on_load=False,
                      cache_parts=None, prefixed_cache=False,
                      orphan_parts=()):
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
    cache_ids = sorted({cache_id for _name, cache_id in pivots})
    if cache_parts is None:
        cache_parts = {
            cache_id: "xl/pivotCache/pivotCacheDefinition{0}.xml".format(
                cache_id)
            for cache_id in cache_ids
        }
    workbook_rels = b"".join(
        b'<Relationship Id="rIdCache%s" Type="http://schemas.openxml'
        b'formats.org/officeDocument/2006/relationships/'
        b'pivotCacheDefinition" Target="/%s"/>'
        % (cache_id.encode("ascii"),
           cache_parts[cache_id].encode("utf-8"))
        for cache_id in cache_ids
    )
    sheet_rels = b"".join(
        b'<Relationship Id="rIdPivot%s" Type="http://schemas.openxml'
        b'formats.org/officeDocument/2006/relationships/pivotTable" '
        b'Target="../pivotTables/pivotTable%s.xml"/>'
        % (str(index).encode("ascii"), str(index).encode("ascii"))
        for index, _pivot in enumerate(pivots, 1)
    )
    content_type_overrides = b"".join(
        b'<Override PartName="/xl/pivotTables/pivotTable%s.xml" '
        b'ContentType="application/vnd.openxmlformats-officedocument.'
        b'spreadsheetml.pivotTable+xml"/>' % str(index).encode("ascii")
        for index, _pivot in enumerate(pivots, 1)
    ) + b"".join(
        b'<Override PartName="/%s" ContentType="application/vnd.'
        b'openxmlformats-officedocument.spreadsheetml.'
        b'pivotCacheDefinition+xml"/>' % part.encode("utf-8")
        for part in tuple(cache_parts.values()) + tuple(orphan_parts)
    )
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
            elif info.filename == "[Content_Types].xml":
                payload = payload.replace(
                    b"</Types>", content_type_overrides + b"</Types>")
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
            root = b"x:pivotCacheDefinition" if prefixed_cache \
                else b"pivotCacheDefinition"
            namespace = (b' xmlns:x="http://schemas.openxmlformats.org/'
                         b'spreadsheetml/2006/main"') if prefixed_cache \
                else (b' xmlns="http://schemas.openxmlformats.org/'
                      b'spreadsheetml/2006/main"')
            source_tag = b"x:cacheSource" if prefixed_cache \
                else b"cacheSource"
            zout.writestr(
                cache_parts[cache_id],
                b'<' + root + namespace + b' recordCount="0"'
                + refresh + b'><' + source_tag + b' type="worksheet"/>'
                b'</' + root + b'>')
        for part in orphan_parts:
            zout.writestr(
                part,
                b'<pivotCacheDefinition xmlns="http://schemas.openxml'
                b'formats.org/spreadsheetml/2006/main" recordCount="0"/>'
            )
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


def test_copy_format_rolls_back_model_and_ledger(
        fixture_copy, monkeypatch):
    from copy import copy

    from openpyxl.preserve import copy_format
    from openpyxl.styles import Font
    import openpyxl.preserve.styleverbs as styleverbs

    wb = load_workbook(
        fixture_copy("features/schedule.xlsx"), preserve=True)
    ws = wb["Schedule"]
    ws["B2"].font = Font(bold=True, italic=True)
    before_styles = {
        coordinate: copy(ws[coordinate]._style)
        for coordinate in ("B3", "B4")
    }
    before_dirty = set(wb._paper_ledger.dirty_coordinates(ws))
    before_row = ws._current_row
    calls = []

    def fail_after_first(coordinate):
        calls.append(coordinate)
        raise RuntimeError("injected format-copy failure")

    monkeypatch.setattr(
        styleverbs, "_copy_format_commit_point", fail_after_first)
    with pytest.raises(RuntimeError, match="injected"):
        copy_format(ws, "B2", "B3:B4")
    assert calls == [(3, 2)]
    assert {coordinate: ws[coordinate]._style
            for coordinate in ("B3", "B4")} == before_styles
    assert set(wb._paper_ledger.dirty_coordinates(ws)) == before_dirty
    assert ws._current_row == before_row


def test_copy_format_rolls_back_new_cells_on_interrupt(
        fixture_copy, monkeypatch):
    from openpyxl.preserve import copy_format
    import openpyxl.preserve.styleverbs as styleverbs

    wb = load_workbook(
        fixture_copy("features/schedule.xlsx"), preserve=True)
    ws = wb["Schedule"]
    before_cells = set(ws._cells)
    before_dirty = set(wb._paper_ledger.dirty_coordinates(ws))
    before_row = ws._current_row
    calls = []

    def interrupt_after_second(coordinate):
        calls.append(coordinate)
        if len(calls) == 2:
            raise KeyboardInterrupt("injected format-copy interruption")

    monkeypatch.setattr(
        styleverbs, "_copy_format_commit_point", interrupt_after_second)
    with pytest.raises(KeyboardInterrupt, match="injected"):
        copy_format(ws, "B2", "Z100:Z101")

    assert calls == [(100, 26), (101, 26)]
    assert set(ws._cells) == before_cells
    assert set(wb._paper_ledger.dirty_coordinates(ws)) == before_dirty
    assert ws._current_row == before_row


def test_copy_format_preflights_merges_and_protection(
        fixture_copy, tmp_path):
    from openpyxl.preserve import copy_format

    src = fixture_copy("minimal/minimal_clean.xlsx")
    wb = load_workbook(src, preserve=True)
    ws = wb.active
    ws["A1"].number_format = "$0.00"
    ws.merge_cells("B1:C1")
    before_anchor = ws["B1"]._style
    with pytest.raises(UnsupportedStructureError,
                       match="non-anchor merged cell"):
        copy_format(ws, "A1", "B1:C1")
    assert ws["B1"]._style == before_anchor

    protected = tmp_path / "protected.xlsx"
    wb = load_workbook(src, preserve=True)
    ws = wb.active
    ws["A1"].number_format = "$0.00"
    ws["B1"].number_format = "0.00"
    ws.protection.sheet = True
    wb.save(protected)
    wb = load_workbook(protected, preserve=True)
    ws = wb.active
    wb.strict_protection = True
    before = ws["B1"]._style
    before_cells = set(ws._cells)
    with pytest.raises(UnsupportedStructureError,
                       match="strict_protection"):
        copy_format(ws, "A1", "B1:C1")
    assert ws["B1"]._style == before
    assert set(ws._cells) == before_cells

    wb.strict_protection = False
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        assert copy_format(ws, "A1", "B1:C1") == 2
    assert len([item for item in caught
                if issubclass(item.category, ProtectedWriteWarning)]) == 1


def test_pivot_refresh_requires_explicit_scope(fixture_copy):
    wb = load_workbook(
        fixture_copy("minimal/minimal_clean.xlsx"), preserve=True)
    with pytest.raises(ValueError, match="exactly one"):
        wb.set_pivot_refresh_on_load()


def test_all_pivot_refresh_patches_only_root_attribute(
        fixture_copy, tmp_path):
    wb = load_workbook(
        fixture_copy("minimal/minimal_clean.xlsx"), preserve=True)
    custom_part = "xl/custom/pivot-cache-a.xml"
    orphan_part = "xl/pivotCache/pivotCacheDefinition99.xml"
    _with_pivot_graph(
        wb, [("SalesPivot", "1")], cache_parts={"1": custom_part},
        orphan_parts=(orphan_part,))
    original = wb._paper_source

    assert wb.set_pivot_refresh_on_load(all=True) == [
        custom_part]
    out = tmp_path / "o.xlsx"
    wb.save(out)
    with zipfile.ZipFile(io.BytesIO(original)) as before, \
            zipfile.ZipFile(out) as after:
        assert b'refreshOnLoad="1"' in after.read(custom_part)
        assert after.read(orphan_part) == before.read(orphan_part)


def test_targeted_pivot_uses_current_sheet_name_after_rename(
        fixture_copy, tmp_path):
    wb = load_workbook(
        fixture_copy("minimal/minimal_clean.xlsx"), preserve=True)
    _with_pivot_graph(wb, [("SalesPivot", "1")])
    wb["Sheet1"].title = "Renamed"

    assert wb.set_pivot_refresh_on_load(
        pivots=["Renamed!SalesPivot"]
    ) == ["xl/pivotCache/pivotCacheDefinition1.xml"]
    out = tmp_path / "renamed-pivot.xlsx"
    wb.save(out)
    with zipfile.ZipFile(out) as archive:
        assert b'refreshOnLoad="1"' in archive.read(
            "xl/pivotCache/pivotCacheDefinition1.xml")


def test_prefixed_pivot_cache_is_patched_and_reported(
        fixture_copy, tmp_path):
    wb = load_workbook(
        fixture_copy("minimal/minimal_clean.xlsx"), preserve=True)
    _with_pivot_graph(
        wb, [("SalesPivot", "1")], prefixed_cache=True)

    wb.set_pivot_refresh_on_load(all=True)
    receipt = wb.save(tmp_path / "prefixed-pivot.xlsx", receipt=True)

    assert [effect for effect in receipt.derived_effects
            if effect["kind"] == "pivot_refresh_on_load_enabled"] == [{
                "kind": "pivot_refresh_on_load_enabled",
                "part": "xl/pivotCache/pivotCacheDefinition1.xml",
                "cause": "explicit_request",
            }]


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
