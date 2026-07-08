"""Phase 2c: the splice writer (CONVENTIONS §3.4; PR-0 D6/D7/D15).

Every assertion follows the reopen rule (save → reopen → assert) and the
part budget is checked literally. The ledger cross-check runs on every
preserve save in this suite (PAPER_LEDGER_CROSSCHECK=1 via conftest).
"""
from __future__ import annotations

import io
import time
import zipfile

import pytest

from openpyxl import load_workbook
from openpyxl.errors import PaperRefusal, UnsupportedStructureError
from openpyxl.package import diff_package

from .support.harness import assert_part_budget
from .support.partdiff import part_payloads

GAUNTLET = "gauntlet/gauntlet.xlsx"


def _model_sheet(path):
    for name, payload in part_payloads(path).items():
        if name.startswith("xl/worksheets/") and b"Quarterly Model" in payload:
            return name, payload
    raise AssertionError("no Model sheet found")


class TestNoOpRoundTrip:
    """The no-op invariant: load(preserve=True) + save == byte-identical
    part payloads, on every fixture class (CONVENTIONS §4)."""

    @pytest.mark.parametrize("fixture", [
        "gauntlet/gauntlet.xlsx",
        "features/lo_authored.xlsx",       # LO producer: sst, declarations
        "features/schedule_calc.xlsx",     # cached values
        "features/shared_formulas.xlsx",   # shared groups + array formula
        "features/macro_stub.xlsm",        # binary vba part
        "minimal/minimal_clean.xlsx",
    ])
    def test_noop_is_byte_identical(self, fixture_copy, tmp_path, fixture):
        src = fixture_copy(fixture)
        wb = load_workbook(src, preserve=True)
        out = str(tmp_path / ("noop" + fixture[-5:]))
        wb.save(out)
        d = diff_package(src, out)
        assert d.clean and not d.equivalent, d

    def test_noop_after_reads_and_materialization(self, fixture_copy, tmp_path):
        src = fixture_copy(GAUNTLET)
        wb = load_workbook(src, preserve=True)
        _ = wb["Model"]["Z99"]
        _ = wb["Model"].row_dimensions[55]
        for _row in wb["Model"].iter_rows():
            pass
        out = str(tmp_path / "noop.xlsx")
        wb.save(out)
        assert diff_package(src, out).clean


class TestSpliceCompletenessTrap:
    """THE signature test (CONVENTIONS §4): a one-cell edit on a sheet
    carrying sparklines, x14 CF, and a drawing reference — everything
    survives, and exactly one part changes."""

    def test_one_cell_edit_preserves_everything(self, fixture_copy, tmp_path):
        src = fixture_copy(GAUNTLET)
        wb = load_workbook(src, preserve=True)
        wb["Model"]["B8"] = 0.15
        out = str(tmp_path / "edit.xlsx")
        wb.save(out)

        sheet_name, _ = _model_sheet(src)
        assert_part_budget(src, out, expect_changed={sheet_name})

        _, sheet_after = _model_sheet(out)
        assert b"sparklineGroups" in sheet_after
        assert b"x14:conditionalFormattings" in sheet_after
        assert b"<x14:id>" in sheet_after            # CF twin pointer
        assert b"<drawing" in sheet_after            # chart attachment
        assert b"<legacyDrawing" in sheet_after

        wb2 = load_workbook(out)
        assert wb2["Model"]["B8"].value == 0.15
        assert len(wb2["Model"].merged_cells.ranges) == 1
        assert len(wb2["Model"].conditional_formatting) == 3
        assert wb2["Model"]["B6"].value == "=B3-B4-B5"

    @pytest.mark.lo_smoke
    def test_spliced_output_loads_in_libreoffice(self, fixture_copy, tmp_path, lo):
        src = fixture_copy(GAUNTLET)
        wb = load_workbook(src, preserve=True)
        wb["Model"]["B8"] = 0.15
        out = str(tmp_path / "edit.xlsx")
        wb.save(out)
        assert lo.lo_loads(out)


class TestCellEdits:

    def test_type_change_number_to_string(self, fixture_copy, tmp_path):
        src = fixture_copy(GAUNTLET)
        wb = load_workbook(src, preserve=True)
        wb["Model"]["B8"] = "TBD"
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        wb2 = load_workbook(out)
        assert wb2["Model"]["B8"].value == "TBD"
        _, sheet = _model_sheet(out)
        assert b't="inlineStr"' in sheet             # PR-0 D1

    def test_new_cell_new_row_and_delete(self, fixture_copy, tmp_path):
        src = fixture_copy(GAUNTLET)
        wb = load_workbook(src, preserve=True)
        wb["Model"]["G3"] = 123                       # new cell, existing row
        wb["Model"]["B40"] = "=SUM(B3:B5)"            # new row
        del wb["Model"]["B9"]                         # delete a cell
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        wb2 = load_workbook(out)
        assert wb2["Model"]["G3"].value == 123
        assert wb2["Model"]["B40"].value == "=SUM(B3:B5)"
        assert wb2["Model"]["B9"].value is None
        # untouched neighbours intact
        assert wb2["Model"]["B3"].value == 100
        assert wb2["Model"]["A9"].value == "Scenario"

    def test_formula_edit_without_calcchain(self, fixture_copy, tmp_path):
        src = fixture_copy("features/schedule.xlsx")
        wb = load_workbook(src, preserve=True)
        wb["Schedule"]["B2"] = 999
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        wb2 = load_workbook(out)
        assert wb2["Schedule"]["B2"].value == 999
        assert wb2["Schedule"]["B12"].value == "=SUM(B2:B11)"

    def test_style_reuse_is_spliceable(self, fixture_copy, tmp_path):
        # assigning a style that already exists in the stylesheet works
        src = fixture_copy(GAUNTLET)
        wb = load_workbook(src, preserve=True)
        wb["Model"]["B9"].style = "paper_input"       # existing named style
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        wb2 = load_workbook(out)
        assert wb2["Model"]["B9"].style == "paper_input"

    def test_new_style_refuses_at_this_stage(self, fixture_copy, tmp_path):
        from openpyxl.styles import Font

        src = fixture_copy(GAUNTLET)
        with open(src, "rb") as f:
            before = f.read()
        wb = load_workbook(src, preserve=True)
        wb["Model"]["A2"].font = Font(name="Menlo", size=7)   # brand-new font
        out = str(tmp_path / "o.xlsx")
        with pytest.raises(UnsupportedStructureError, match="Phase 2d"):
            wb.save(out)
        with open(src, "rb") as f:
            assert f.read() == before


class TestSharedAndArrayFormulas:

    def test_shared_group_dissolves_on_touch(self, fixture_copy, tmp_path):
        src = fixture_copy("features/shared_formulas.xlsx")
        wb = load_workbook(src, preserve=True)
        wb["Calc"]["B3"] = 999                        # member of B2:B6 group
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        parts = part_payloads(out)
        sheet = next(p for n, p in parts.items()
                     if n.startswith("xl/worksheets/") and b"A2*2" in p)
        assert b't="shared"' not in sheet             # group dissolved whole
        wb2 = load_workbook(out)
        assert wb2["Calc"]["B3"].value == 999
        assert wb2["Calc"]["B4"].value == "=A4*2"     # follower kept meaning
        assert wb2["Calc"]["B6"].value == "=A6*2"
        # the array formula's bytes were never touched
        assert b't="array"' in sheet

    def test_untouched_shared_group_passes_through_verbatim(
            self, fixture_copy, tmp_path):
        src = fixture_copy("features/shared_formulas.xlsx")
        wb = load_workbook(src, preserve=True)
        wb["Calc"]["A10"] = "note"                    # outside the group
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        parts = part_payloads(out)
        sheet = next(p for n, p in parts.items()
                     if n.startswith("xl/worksheets/") and b"A2*2" in p)
        assert sheet.count(b'<f t="shared" si="0"/>') == 4   # untouched

    def test_array_formula_edit_refuses_atomically(self, fixture_copy, tmp_path):
        src = fixture_copy("features/shared_formulas.xlsx")
        with open(src, "rb") as f:
            before = f.read()
        wb = load_workbook(src, preserve=True)
        wb["Calc"]["D3"] = 5                          # inside D2:D4 array
        with pytest.raises(UnsupportedStructureError, match="array formula"):
            wb.save(str(tmp_path / "o.xlsx"))
        with open(src, "rb") as f:
            assert f.read() == before


class TestRegionEdits:

    def test_merge_cells_updates_only_the_sheet(self, fixture_copy, tmp_path):
        src = fixture_copy(GAUNTLET)
        wb = load_workbook(src, preserve=True)
        wb["Model"].merge_cells("A14:C15")
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        sheet_name, _ = _model_sheet(src)
        assert_part_budget(src, out, expect_changed={sheet_name})
        wb2 = load_workbook(out)
        assert {"A1:F1", "A14:C15"} == {
            str(r) for r in wb2["Model"].merged_cells.ranges}
        _, sheet = _model_sheet(out)
        assert b"sparklineGroups" in sheet            # traps still intact

    def test_freeze_panes_change(self, fixture_copy, tmp_path):
        src = fixture_copy(GAUNTLET)
        wb = load_workbook(src, preserve=True)
        wb["Model"].freeze_panes = "B4"
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        wb2 = load_workbook(out)
        assert wb2["Model"].freeze_panes == "B4"

    def test_column_width_change(self, fixture_copy, tmp_path):
        src = fixture_copy(GAUNTLET)
        wb = load_workbook(src, preserve=True)
        wb["Model"].column_dimensions["A"].width = 33.5
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        wb2 = load_workbook(out)
        assert wb2["Model"].column_dimensions["A"].width == 33.5

    def test_row_height_change_keeps_cells(self, fixture_copy, tmp_path):
        src = fixture_copy(GAUNTLET)
        wb = load_workbook(src, preserve=True)
        wb["Model"].row_dimensions[3].height = 30
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        wb2 = load_workbook(out)
        assert wb2["Model"].row_dimensions[3].height == 30
        assert wb2["Model"]["B3"].value == 100        # cells untouched

    def test_data_validation_add_on_plain_sheet(self, fixture_copy, tmp_path):
        from openpyxl.worksheet.datavalidation import DataValidation

        src = fixture_copy("features/datavalidation.xlsx")
        wb = load_workbook(src, preserve=True)
        ws = wb.active
        dv = DataValidation(type="list", formula1='"Yes,No"')
        dv.add("C2")
        ws.add_data_validation(dv)
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        wb2 = load_workbook(out)
        assert len(wb2.active.data_validations.dataValidation) == 3

    def test_new_region_inserted_at_schema_position(self, fixture_copy, tmp_path):
        src = fixture_copy("minimal/minimal_clean.xlsx")
        wb = load_workbook(src, preserve=True)
        wb["Sheet1"].merge_cells("F1:G2")             # sheet had no mergeCells
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        wb2 = load_workbook(out)
        assert [str(r) for r in wb2["Sheet1"].merged_cells.ranges] == ["F1:G2"]

    def test_cf_change_refuses_at_this_stage(self, fixture_copy, tmp_path):
        from openpyxl.formatting.rule import CellIsRule
        from openpyxl.styles import PatternFill

        src = fixture_copy(GAUNTLET)
        with open(src, "rb") as f:
            before = f.read()
        wb = load_workbook(src, preserve=True)
        wb["Model"].conditional_formatting.add(
            "B9:B9", CellIsRule(operator="lessThan", formula=["0"],
                                fill=PatternFill("solid", fgColor="FF0000AA")))
        with pytest.raises(UnsupportedStructureError, match="conditionalFormatting"):
            wb.save(str(tmp_path / "o.xlsx"))
        with open(src, "rb") as f:
            assert f.read() == before


class TestStageRefusals:
    """Cross-part operations that land in Phase 2d refuse loudly and
    atomically at this stage (never a silent drop)."""

    def _assert_refuses(self, wb, src, tmp_path, match):
        with open(src, "rb") as f:
            before = f.read()
        out = str(tmp_path / "refused.xlsx")
        with pytest.raises(PaperRefusal, match=match):
            wb.save(out)
        import os
        assert not os.path.exists(out)
        with open(src, "rb") as f:
            assert f.read() == before

    def test_added_sheet_refuses(self, fixture_copy, tmp_path):
        src = fixture_copy(GAUNTLET)
        wb = load_workbook(src, preserve=True)
        wb.create_sheet("New")
        self._assert_refuses(wb, src, tmp_path, "Phase 2d")

    def test_workbook_level_change_refuses(self, fixture_copy, tmp_path):
        from openpyxl.workbook.defined_name import DefinedName

        src = fixture_copy(GAUNTLET)
        wb = load_workbook(src, preserve=True)
        wb.defined_names["Fresh"] = DefinedName("Fresh", attr_text="Model!$A$1")
        self._assert_refuses(wb, src, tmp_path, "workbook-level")

    def test_comment_change_refuses(self, fixture_copy, tmp_path):
        from openpyxl.comments import Comment

        src = fixture_copy(GAUNTLET)
        wb = load_workbook(src, preserve=True)
        wb["Model"]["B9"].comment = Comment("new note", "tester")
        self._assert_refuses(wb, src, tmp_path, "comment")

    def test_data_only_save_refuses(self, fixture_copy, tmp_path):
        src = fixture_copy("features/schedule_calc.xlsx")
        wb = load_workbook(src, preserve=True, data_only=True)
        self._assert_refuses(wb, src, tmp_path, "data_only")

    def test_style_registry_mutation_refuses(self, fixture_copy, tmp_path):
        src = fixture_copy(GAUNTLET)
        wb = load_workbook(src, preserve=True)
        wb["Model"]["B8"].fill.start_color.rgb = "FF12AB34"   # proxy leak
        wb["Model"]["A2"] = "trigger a save plan"
        self._assert_refuses(wb, src, tmp_path, "mutated in place")


class TestProducerGuards:

    def _surgery(self, fixture_copy, tmp_path, mutate_sheet):
        src = fixture_copy("minimal/minimal_clean.xlsx")
        out = str(tmp_path / "surgery.xlsx")
        with zipfile.ZipFile(src) as zin, zipfile.ZipFile(out, "w") as zout:
            for name in zin.namelist():
                payload = zin.read(name)
                if name.startswith("xl/worksheets/sheet"):
                    payload = mutate_sheet(payload)
                zout.writestr(name, payload)
        return out

    def test_ph_attribute_carried_over(self, fixture_copy, tmp_path):
        # PR-0 D6 carry rule: legal extra cell attributes survive the edit
        surgical = self._surgery(
            fixture_copy, tmp_path,
            lambda p: p.replace(b'<c r="B2"', b'<c r="B2" ph="1"', 1))
        wb = load_workbook(surgical, preserve=True)
        wb["Sheet1"]["B2"] = 42
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        parts = part_payloads(out)
        sheet = next(p for n, p in parts.items()
                     if n.startswith("xl/worksheets/"))
        assert b'ph="1"' in sheet
        assert load_workbook(out)["Sheet1"]["B2"].value == 42

    def test_cm_metadata_cell_refuses(self, fixture_copy, tmp_path):
        surgical = self._surgery(
            fixture_copy, tmp_path,
            lambda p: p.replace(b'<c r="B2"', b'<c r="B2" cm="1"', 1))
        wb = load_workbook(surgical, preserve=True)
        wb["Sheet1"]["B2"] = 42
        with pytest.raises(UnsupportedStructureError, match="cell metadata"):
            wb.save(str(tmp_path / "o.xlsx"))

    def test_rless_rows_refuse(self, fixture_copy, tmp_path):
        import re

        surgical = self._surgery(
            fixture_copy, tmp_path,
            lambda p: re.sub(br'<row r="\d+"', b"<row", p))
        wb = load_workbook(surgical, preserve=True)
        wb["Sheet1"]["B2"] = 42
        with pytest.raises(UnsupportedStructureError, match="no r attribute"):
            wb.save(str(tmp_path / "o.xlsx"))

    def test_doctype_refuses(self, fixture_copy, tmp_path):
        surgical = self._surgery(
            fixture_copy, tmp_path,
            lambda p: b"<!DOCTYPE worksheet>" + p)
        wb = load_workbook(surgical, preserve=True)
        wb["Sheet1"]["B2"] = 42
        with pytest.raises(UnsupportedStructureError, match="DOCTYPE"):
            wb.save(str(tmp_path / "o.xlsx"))


class TestSaveTargets:

    def test_save_to_file_like_target(self, fixture_copy):
        src = fixture_copy(GAUNTLET)
        wb = load_workbook(src, preserve=True)
        wb["Model"]["B8"] = 0.2
        buf = io.BytesIO(b"pre-existing garbage that must be truncated away")
        wb.save(buf)
        buf.seek(0)
        wb2 = load_workbook(buf)
        assert wb2["Model"]["B8"].value == 0.2


class TestPerformanceGuardrail:
    """PR-0 D4 (as amended in Phase 2c): preserve save within 2x stock save
    on the large fixture. Measured 1.82x-1.87x; see PR0-API-PROPOSAL.md D4
    for the amendment evidence."""

    def test_splice_save_within_budget(self, fixture_copy, tmp_path, monkeypatch):
        monkeypatch.delenv("PAPER_LEDGER_CROSSCHECK", raising=False)
        src = fixture_copy("large/large150k.xlsx")

        def best_of(n, fn):
            times = []
            for _ in range(n):
                t0 = time.perf_counter()
                fn()
                times.append(time.perf_counter() - t0)
            return min(times)

        wb_stock = load_workbook(src)
        t_stock = best_of(2, lambda: wb_stock.save(str(tmp_path / "stock.xlsx")))

        wb = load_workbook(src, preserve=True)
        wb["Big"]["A2"] = 424242
        t_preserve = best_of(2, lambda: wb.save(str(tmp_path / "preserve.xlsx")))

        assert load_workbook(str(tmp_path / "preserve.xlsx"))["Big"]["A2"].value == 424242
        assert t_preserve <= 2.0 * t_stock, (
            "splice save {0:.3f}s exceeded 2x stock save {1:.3f}s".format(
                t_preserve, t_stock))
