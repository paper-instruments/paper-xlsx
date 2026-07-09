"""Delivery verbs and hardening."""
from __future__ import annotations

import io
import zipfile

import pytest

from openpyxl import Workbook, load_workbook
from openpyxl.errors import (
    AmbiguousTargetError,
    TargetNotFoundError,
    UnsupportedStructureError,
)


class TestSetInput:

    def test_defined_name_then_label(self, fixture_copy):
        from openpyxl.workbook.defined_name import DefinedName

        wb = load_workbook(fixture_copy("features/schedule.xlsx"),
                           preserve=True)
        wb.defined_names["rate"] = DefinedName(
            "rate", attr_text="Schedule!$B$2")
        cell = wb.set_input("rate", 123)     # defined name wins
        assert (cell.parent.title, cell.coordinate) == ("Schedule", "B2")
        cell = wb.set_input("Item 2", 456)   # label fallback via locate
        assert cell.value == 456

    def test_never_overwrites_formulas(self, fixture_copy):
        wb = load_workbook(fixture_copy("features/schedule.xlsx"),
                           preserve=True)
        with pytest.raises(UnsupportedStructureError,
                           match="never overwrites"):
            wb.set_input("Grand total", 5)   # resolves to Summary!B1 (=f)

    def test_unknown_and_ambiguous(self, fixture_copy):
        wb = load_workbook(fixture_copy("features/schedule.xlsx"),
                           preserve=True)
        with pytest.raises(TargetNotFoundError):
            wb.set_input("No Such Thing", 1)
        wb["Summary"]["A9"] = "Item 1"
        wb["Summary"]["B9"] = 0
        with pytest.raises(AmbiguousTargetError) as exc:
            wb.set_input("Item 1", 1)
        assert len(exc.value.options) == 2


class TestProtectAndScrub:

    def test_protect_for_delivery_roundtrip(self, fixture_copy, tmp_path):
        wb = load_workbook(fixture_copy("features/schedule_calc.xlsx"),
                           preserve=True)
        report = wb.protect_for_delivery()
        assert "Schedule" in report["locked_sheets"]
        assert any(a.startswith("Schedule!B") for a
                   in report["unlocked_inputs"])
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        wb2 = load_workbook(out)
        assert wb2["Schedule"].protection.sheet is True
        assert wb2["Schedule"]["B2"].protection.locked is False
        assert wb2["Schedule"]["B12"].protection.locked is True

    def test_scrub_reports_everything(self, fixture_copy, tmp_path):
        from openpyxl.comments import Comment

        wb = load_workbook(fixture_copy("features/schedule.xlsx"),
                           preserve=True)
        wb["Schedule"]["A2"].comment = Comment("secret note", "author")
        hidden = wb.create_sheet("Internal")
        hidden["A1"] = 1
        hidden.sheet_state = "hidden"
        wb.properties.creator = "Jane Analyst"
        report = wb.scrub()
        assert any("comment at Schedule!A2" in r
                   for r in report["removed"])
        assert any("Internal" in r for r in report["removed"])
        assert any("creator" in r for r in report["removed"])
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        wb2 = load_workbook(out)
        assert "Internal" not in wb2.sheetnames
        assert wb2["Schedule"]["A2"].comment is None

    def test_scrub_refused_hidden_sheet_is_reported(self, fixture_copy):
        wb = load_workbook(fixture_copy("features/schedule.xlsx"),
                           preserve=True)
        # Schedule is referenced by Summary!B1: hiding it then scrubbing
        # must REPORT the refusal, never silently drop or remove
        wb["Schedule"].sheet_state = "hidden"
        report = wb.scrub(remove=("hidden-sheets",))
        assert any("Schedule" in skip for skip in report["skipped"])
        assert "Schedule" in wb.sheetnames

    def test_scrub_preserved_comment_machinery_reported(self,
                                                        fixture_copy):
        from openpyxl.comments import Comment

        # a package that already carries comment parts: scrub reports it
        src = fixture_copy("features/schedule.xlsx")
        wb0 = load_workbook(src)
        wb0["Schedule"]["A2"].comment = Comment("existing", "a")
        wb0.save(src)
        wb = load_workbook(src, preserve=True)
        report = wb.scrub(remove=("comments",))
        assert any("machinery" in skip for skip in report["skipped"])

    def test_unknown_scrub_target_refuses(self, fixture_copy):
        wb = load_workbook(fixture_copy("features/schedule.xlsx"),
                           preserve=True)
        with pytest.raises(ValueError, match="unknown scrub"):
            wb.scrub(remove=("vba",))


class TestStyleVerbs:

    def test_copy_format_through_the_splice(self, fixture_copy,
                                            tmp_path):
        from openpyxl.preserve import copy_format
        from openpyxl.styles import Font

        wb = load_workbook(fixture_copy("features/schedule.xlsx"),
                           preserve=True)
        ws = wb["Schedule"]
        ws["B2"].font = Font(bold=True, italic=True)
        assert copy_format(ws, "B2", "B3:B4") == 2
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        wb2 = load_workbook(out)
        assert wb2["Schedule"]["B3"].font.bold is True
        assert wb2["Schedule"]["B4"].font.italic is True

    def test_apply_profile_by_role(self, fixture_copy, tmp_path):
        from openpyxl.preserve import apply_profile

        wb = load_workbook(fixture_copy("features/schedule_calc.xlsx"),
                           preserve=True)
        counts = apply_profile(wb["Schedule"], {
            "inputs": {"number_format": "comma", "fill": "FFF2CC"},
            "calculations": {"bold": True},
        })
        assert counts["inputs"] >= 4
        assert counts["calculations"] >= 1
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        wb2 = load_workbook(out)
        assert wb2["Schedule"]["B2"].number_format == "#,##0.00"
        assert wb2["Schedule"]["B12"].font.bold is True


class TestPivotRefresh:

    def test_patches_pivot_caches(self, fixture_copy, tmp_path):
        src = fixture_copy("minimal/minimal_clean.xlsx")
        crafted = str(tmp_path / "pivot.xlsx")
        cache = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 b'<pivotCacheDefinition xmlns="http://schemas.openxml'
                 b'formats.org/spreadsheetml/2006/main" r:id="rId1" '
                 b'xmlns:r="http://schemas.openxmlformats.org/office'
                 b'Document/2006/relationships" recordCount="2">'
                 b"<cacheSource type=\"worksheet\"/></pivotCacheDefinition>")
        with zipfile.ZipFile(src) as zin, \
                zipfile.ZipFile(crafted, "w") as zout:
            for name in zin.namelist():
                payload = zin.read(name)
                if name == "[Content_Types].xml":
                    payload = payload.replace(
                        b"</Types>",
                        b'<Override PartName="/xl/pivotCache/pivotCache'
                        b'Definition1.xml" ContentType="application/vnd.'
                        b"openxmlformats-officedocument.spreadsheetml."
                        b'pivotCacheDefinition+xml"/></Types>', 1)
                zout.writestr(name, payload)
            zout.writestr("xl/pivotCache/pivotCacheDefinition1.xml",
                          cache)
        wb = load_workbook(crafted, preserve=True)
        patched = wb.set_pivot_refresh_on_load()
        assert patched == ["xl/pivotCache/pivotCacheDefinition1.xml"]
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        with zipfile.ZipFile(out) as z:
            payload = z.read("xl/pivotCache/pivotCacheDefinition1.xml")
        assert b'refreshOnLoad="1"' in payload
        assert b"<cacheSource" in payload            # content intact

    def test_requires_preserve(self, fixture_copy):
        wb = load_workbook(fixture_copy("minimal/minimal_clean.xlsx"))
        with pytest.raises(ValueError, match="preserve"):
            wb.set_pivot_refresh_on_load()


class TestHardening:

    def test_decompression_bomb_refused(self, tmp_path):
        bomb = str(tmp_path / "bomb.xlsx")
        wb = Workbook()
        wb.active["A1"] = 1
        wb.save(bomb)
        crafted = str(tmp_path / "bomb2.xlsx")
        with zipfile.ZipFile(bomb) as zin, \
                zipfile.ZipFile(crafted, "w",
                                zipfile.ZIP_DEFLATED) as zout:
            for name in zin.namelist():
                zout.writestr(name, zin.read(name))
            # 128 MB of zeros compresses > 500x: past the pinned caps
            zout.writestr("xl/media/blob.bin", b"\x00" * (128 << 20))
        with pytest.raises(UnsupportedStructureError, match="inflates"):
            load_workbook(crafted, preserve=True)

    def test_zip_confusion_normalized(self, fixture_copy, tmp_path):
        # a mismatching LOCAL header falls back to recompression from
        # the central directory's view — the save must not raw-copy
        # ambiguous bytes
        import struct

        src = fixture_copy("features/schedule.xlsx")
        with open(src, "rb") as f:
            data = bytearray(f.read())
        with zipfile.ZipFile(io.BytesIO(bytes(data))) as z:
            info = z.getinfo("docProps/core.xml")
        # flip a byte of the LOCAL header's CRC field (offset 14)
        base = info.header_offset + 14
        data[base] = data[base] ^ 0xFF
        crafted = str(tmp_path / "confused.xlsx")
        with open(crafted, "wb") as f:
            f.write(bytes(data))
        wb = load_workbook(crafted, preserve=True)
        wb["Schedule"]["A2"] = "edit"
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        with zipfile.ZipFile(out) as z:
            payload = z.read("docProps/core.xml")   # normalized, valid
        assert b"cp:coreProperties" in payload

    def test_mark_dirty_oversized_range_clamps(self, fixture_copy):
        wb = load_workbook(fixture_copy("features/schedule.xlsx"),
                           preserve=True)
        ws = wb["Schedule"]
        wb.mark_dirty("Schedule!A1:XFD1048576")     # clamps, no explosion
        dirty = wb._paper_ledger.dirty_coordinates(ws)
        assert len(dirty) <= (ws.max_row * ws.max_column)

    def test_spooled_save_is_correct_zip(self, fixture_copy, tmp_path):
        src = fixture_copy("features/schedule.xlsx")
        wb = load_workbook(src, preserve=True)
        wb["Schedule"]["A2"] = "spooled"
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        with zipfile.ZipFile(out) as z:
            assert z.testzip() is None
        assert load_workbook(out)["Schedule"]["A2"].value == "spooled"
