"""One regression test per confirmed finding of the standing per-batch
adversarial gates (PLAN-v0.1 process amendment 2). Sections are tagged by
batch; every test here pins a fix for a live repro the gate produced.
"""
from __future__ import annotations

import io
import zipfile

import pytest

from openpyxl import Workbook, load_workbook
from openpyxl.errors import (
    BoundaryViolationError,
    ProtectedWriteWarning,
    UnsupportedStructureError,
)

from .support.partdiff import part_payloads


class TestBatch1ObjectGuardGaps:

    def test_chart_anchor_mutation_refuses(self, fixture_copy, tmp_path):
        # the chart fingerprint was chart._write() only — the anchor lives
        # in the preserved drawing part, so a chart MOVE vanished silently
        src = fixture_copy("features/chart_image.xlsx")
        wb = load_workbook(src, preserve=True)
        ws = next(w for w in wb.worksheets if w._charts)
        ws._charts[0].anchor = "K25"
        with pytest.raises(UnsupportedStructureError, match="chart"):
            wb.save(str(tmp_path / "o.xlsx"))

    def test_image_data_swap_refuses(self, fixture_copy, tmp_path):
        # same anchor + same path + different pixels: the fingerprint now
        # covers the backing bytes (non-destructively — image._data()
        # CLOSES the ref stream and must never be used for snapshots)
        src = fixture_copy("features/chart_image.xlsx")
        wb = load_workbook(src, preserve=True)
        ws = next(w for w in wb.worksheets if w._images)
        img = ws._images[0]
        original = img.ref.getvalue()
        img.ref = io.BytesIO(original + b"\x00")
        with pytest.raises(UnsupportedStructureError, match="image"):
            wb.save(str(tmp_path / "o.xlsx"))

    def test_chartsheet_chart_mutation_refuses(self, tmp_path):
        # chartsheet-anchored charts were entirely outside the boundary
        from openpyxl.chart import BarChart, Reference

        wb = Workbook()
        ws = wb.active
        for r in range(1, 5):
            ws.append([r, r * 2])
        cs = wb.create_chartsheet("ChartOnly")
        chart = BarChart()
        chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=4))
        cs.add_chart(chart)
        src = str(tmp_path / "cs.xlsx")
        wb.save(src)

        wb2 = load_workbook(src, preserve=True)
        wb2.chartsheets[0]._charts[0].title = "TAMPERED"
        with pytest.raises(UnsupportedStructureError, match="chartsheet"):
            wb2.save(str(tmp_path / "o.xlsx"))


class TestBatch1RecalcGuardGaps:

    def _three_sheet_fixture(self, tmp_path, formula):
        # Excel-producer shape: cached value present, no fullCalcOnLoad
        wb = Workbook()
        s1 = wb.active
        s1.title = "Sheet1"
        wb.create_sheet("Sheet2")
        wb.create_sheet("Sheet3")
        for s in wb.worksheets:
            s["A1"] = 10
        s1["B1"] = formula
        raw = str(tmp_path / "threed_raw.xlsx")
        wb.save(raw)
        out = str(tmp_path / "threed.xlsx")
        with zipfile.ZipFile(raw) as zin, zipfile.ZipFile(out, "w") as zout:
            for name in zin.namelist():
                payload = zin.read(name)
                if name == "xl/workbook.xml":
                    payload = payload.replace(b' fullCalcOnLoad="1"', b"")
                if name == "xl/worksheets/sheet1.xml" and b"<f>" in payload:
                    payload = payload.replace(b"</f>", b"</f><v>999</v>", 1)
                zout.writestr(name, payload)
        return out

    def test_3d_span_reference_forces_recalc_flag(self, tmp_path):
        # 'Sheet1:Sheet3' was recorded as a phantom sheet name nothing
        # could match — the recalc guard and certification taint both
        # silently missed 3-D formulas
        src = self._three_sheet_fixture(tmp_path, "=SUM(Sheet1:Sheet3!A1)")
        wb = load_workbook(src, preserve=True)
        wb["Sheet2"]["A1"] = 42
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        assert b"fullCalcOnLoad" in part_payloads(out)["xl/workbook.xml"]

    def test_3d_span_is_unresolved_in_the_sketch(self):
        from openpyxl.preserve.perception import dependency_sketch

        wb = Workbook()
        wb.create_sheet("Sheet2")
        wb.create_sheet("Sheet3")
        wb.active["B1"] = "=SUM(Sheet1:Sheet3!A1)"
        sketch = dependency_sketch(wb)
        assert any("B1" in a for a in sketch.unresolved)

    def test_case_insensitive_reference_intersection(self):
        from openpyxl.preserve.perception import dependency_sketch

        wb = Workbook()
        ws = wb.active            # 'Sheet'
        wb.create_sheet("Other")
        wb["Other"]["B1"] = "=sheet!A1*2"       # lowercase ref
        sketch = dependency_sketch(wb)
        hits = sketch.cells_referencing("Sheet", (1, 1, 1, 1))
        assert any("B1" in h for h in hits)


class TestBatch1ProtectionGaps:

    def _protected(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "original"
        ws.protection.sheet = True
        src = str(tmp_path / "protected.xlsx")
        wb.save(src)
        return src

    def test_delete_of_locked_cell_is_protection_checked(self, tmp_path):
        # del ws['A1'] evaded what ws['A1']=None refused
        src = self._protected(tmp_path)
        wb = load_workbook(src, preserve=True)
        wb.strict_protection = True
        with pytest.raises(UnsupportedStructureError, match="locked"):
            del wb.active["A1"]
        assert wb.active["A1"].value == "original"

        wb2 = load_workbook(src, preserve=True)
        with pytest.warns(ProtectedWriteWarning):
            del wb2.active["A1"]

    def test_structural_shift_on_protected_sheet_warns_or_refuses(
            self, tmp_path):
        src = self._protected(tmp_path)
        wb = load_workbook(src, preserve=True)
        with pytest.warns(ProtectedWriteWarning, match="insert_rows"):
            wb.active.insert_rows(1)

        wb2 = load_workbook(src, preserve=True)
        wb2.strict_protection = True
        with pytest.raises(UnsupportedStructureError, match="protected"):
            wb2.active.delete_rows(1)


class TestBatch1InputHonestyGaps:

    def test_cfb_sniff_anchors_at_offset_zero(self, fixture_copy, tmp_path):
        # (a) a valid xlsx handed over at a position where CFB bytes sit
        # (embedded OLE payload) must NOT false-refuse
        src = fixture_copy("minimal/minimal_clean.xlsx")
        cfb_payload = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 64
        embedded = str(tmp_path / "embedded.xlsx")
        with zipfile.ZipFile(src) as zin, \
                zipfile.ZipFile(embedded, "w") as zout:
            for name in zin.namelist():
                zout.writestr(name, zin.read(name))
            zout.writestr(zipfile.ZipInfo("xl/embeddings/oleObject1.bin"),
                          cfb_payload)
        with open(embedded, "rb") as f:
            data = f.read()
        offset = data.find(cfb_payload)
        assert offset > 0
        handle = open(embedded, "rb")
        handle.seek(offset)
        wb = load_workbook(handle)             # must not raise
        assert wb.active is not None
        handle.close()

        # (b) a genuine CFB file via a mid-position handle still gets the
        # typed refusal (it evaded to BadZipFile before the fix)
        cfb = str(tmp_path / "enc.xlsx")
        with open(cfb, "wb") as f:
            f.write(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 4096)
        handle = open(cfb, "rb")
        handle.seek(4)
        with pytest.raises(UnsupportedStructureError, match="ENCRYPT"):
            load_workbook(handle)
        assert handle.tell() == 4              # position restored
        handle.close()

    def test_rich_text_mode_suppresses_rich_text_loss_entry(
            self, fixture_copy, tmp_path):
        # under rich_text=True the stock save PRESERVES runs — warning
        # about flattening was loud-but-wrong
        src = fixture_copy("minimal/minimal_clean.xlsx")
        rich = str(tmp_path / "rich.xlsx")
        with zipfile.ZipFile(src) as zin, zipfile.ZipFile(rich, "w") as zout:
            for name in zin.namelist():
                payload = zin.read(name)
                if name.startswith("xl/worksheets/sheet"):
                    payload = payload.replace(
                        b"</sheetData>",
                        b'<row r="9"><c r="A9" t="inlineStr"><is><r><rPr>'
                        b'<b/></rPr><t>bold</t></r></is></c></row>'
                        b"</sheetData>", 1)
                zout.writestr(name, payload)
        import warnings as _w
        with _w.catch_warnings():
            _w.simplefilter("ignore")
            flat = load_workbook(rich)
            modeled = load_workbook(rich, rich_text=True)
        assert "rich-text" in flat._paper_loss_inventory.kinds()
        assert "rich-text" not in modeled._paper_loss_inventory.kinds()


class TestBatch1RemapAndCertifyGaps:

    def test_xlfn_prefixed_functions_are_excluded(self):
        from openpyxl import oracle

        wb = Workbook()
        ws = wb.active
        ws["A1"] = "=_xlfn.LET(x,1,x*2)"
        ws["A2"] = "=_xlfn.RANDARRAY(2)"
        seeds = oracle._exclusion_seeds(wb)
        assert seeds[("Sheet", 1, 1)] == "unsupported:LET"
        assert seeds[("Sheet", 2, 1)] == "volatile"

    def test_external_ref_behind_defined_name_is_excluded(self):
        from openpyxl import oracle
        from openpyxl.workbook.defined_name import DefinedName

        wb = Workbook()
        wb.defined_names["EXTPRICE"] = DefinedName(
            "EXTPRICE", attr_text="'[1]Other'!$A$1")
        wb.active["A1"] = "=EXTPRICE*2"
        seeds = oracle._exclusion_seeds(wb)
        assert seeds[("Sheet", 1, 1)] == "external-link"

    def test_boundary_counts_dimension_only_rows(
            self, fixture_copy, tmp_path):
        # a <row r="1048576" ht="30"/> with no cells evaded ws.max_row
        src = fixture_copy("features/schedule.xlsx")
        floored = str(tmp_path / "floored.xlsx")
        with zipfile.ZipFile(src) as zin, \
                zipfile.ZipFile(floored, "w") as zout:
            for name in zin.namelist():
                payload = zin.read(name)
                if name.startswith("xl/worksheets/sheet") \
                        and b"Schedule" not in payload[:200]:
                    pass
                if name == "xl/worksheets/sheet1.xml":
                    payload = payload.replace(
                        b"</sheetData>",
                        b'<row r="1048576" ht="30" customHeight="1"/>'
                        b"</sheetData>", 1)
                zout.writestr(name, payload)
        wb = load_workbook(floored, preserve=True)
        ws = wb["Schedule"]
        if not ws.row_dimensions.get(1048576):
            pytest.skip("floor row not on the Schedule sheet part")
        with pytest.raises(BoundaryViolationError):
            ws.insert_rows(1)

    def test_insert_beyond_content_does_not_false_refuse(self, fixture_copy):
        # content ends early; inserting at a beyond-content index that
        # arithmetically exceeds the limit must NOT refuse (nothing moves)
        wb = load_workbook(fixture_copy("features/schedule.xlsx"),
                           preserve=True)
        ws = wb["Schedule"]
        remap = ws.insert_rows(1048570, 10)     # no occupied cell shifts
        assert remap is not None
