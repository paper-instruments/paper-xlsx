"""Phase 3: honesty organs — the data_only trap, recalc-on-load, format
refusals (PLAN Phase 3; PR-0 §3)."""
from __future__ import annotations

import warnings

import pytest

from openpyxl import load_workbook
from openpyxl.errors import (
    LossySaveWarning,
    PaperRefusal,
    UnsupportedStructureError,
)
from openpyxl.utils.exceptions import InvalidFileException

from .support.partdiff import part_payloads


class TestDataOnlyTrap:

    def test_preserve_data_only_save_refuses_naming_the_override(
            self, fixture_copy, tmp_path):
        src = fixture_copy("features/schedule_calc.xlsx")
        with open(src, "rb") as f:
            before = f.read()
        wb = load_workbook(src, preserve=True, data_only=True)
        with pytest.raises(UnsupportedStructureError,
                           match="allow_formula_loss"):
            wb.save(str(tmp_path / "o.xlsx"))
        with open(src, "rb") as f:
            assert f.read() == before

    def test_override_loses_only_edited_cells(self, fixture_copy, tmp_path):
        # under preserve the trap is defused for untouched cells: only cells
        # the user edited lose formulas — stock destroys ALL of them
        src = fixture_copy("features/schedule_calc.xlsx")
        wb = load_workbook(src, preserve=True, data_only=True)
        wb["Schedule"]["B12"] = 9999                # edit a formula cell
        out = str(tmp_path / "o.xlsx")
        wb.save(out, allow_formula_loss=True)
        wb2 = load_workbook(out)                    # formulas view
        assert wb2["Schedule"]["B12"].value == 9999            # edited: literal
        assert wb2["Schedule"]["B13"].value == "=B12*(1+Growth)"  # untouched: formula
        assert wb2["Summary"]["B1"].value == "=Schedule!B12"

    def test_noop_data_only_save_with_override_is_byte_identical(
            self, fixture_copy, tmp_path):
        from openpyxl.package import diff_package

        src = fixture_copy("features/schedule_calc.xlsx")
        wb = load_workbook(src, preserve=True, data_only=True)
        out = str(tmp_path / "o.xlsx")
        wb.save(out, allow_formula_loss=True)
        assert diff_package(src, out).clean

    def test_stock_data_only_save_warns(self, fixture_copy, tmp_path):
        src = fixture_copy("features/schedule_calc.xlsx")
        wb = load_workbook(src, data_only=True)
        with pytest.warns(LossySaveWarning, match="PERMANENTLY replaces"):
            wb.save(str(tmp_path / "o.xlsx"))

    def test_stock_data_only_warning_silenced_by_override(
            self, fixture_copy, tmp_path):
        src = fixture_copy("features/schedule_calc.xlsx")
        wb = load_workbook(src, data_only=True)
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            wb.save(str(tmp_path / "o.xlsx"), allow_formula_loss=True)
        assert not [w for w in caught
                    if isinstance(w.message, LossySaveWarning)
                    and "PERMANENTLY" in str(w.message)]


class TestRecalcOnLoad:

    def test_formula_edit_sets_full_calc_on_load(self, fixture_copy, tmp_path):
        # schedule_calc is LibreOffice-written: its calcPr lacks the flag
        src = fixture_copy("features/schedule_calc.xlsx")
        assert b"fullCalcOnLoad" not in part_payloads(src)["xl/workbook.xml"]
        wb = load_workbook(src, preserve=True)
        wb["Schedule"]["B20"] = "=SUM(B2:B3)"
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        assert b'fullCalcOnLoad="1"' in part_payloads(out)["xl/workbook.xml"]

    def test_value_only_edit_does_not_touch_workbook_xml(
            self, fixture_copy, tmp_path):
        src = fixture_copy("features/schedule_calc.xlsx")
        wb = load_workbook(src, preserve=True)
        wb["Schedule"]["A1"] = "renamed"            # no formula involved
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        before = part_payloads(src)["xl/workbook.xml"]
        after = part_payloads(out)["xl/workbook.xml"]
        assert before == after

    def test_formula_deletion_also_sets_the_flag(self, fixture_copy, tmp_path):
        src = fixture_copy("features/schedule_calc.xlsx")
        wb = load_workbook(src, preserve=True)
        del wb["Schedule"]["B13"]                   # deletes a formula cell
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        assert b'fullCalcOnLoad="1"' in part_payloads(out)["xl/workbook.xml"]


class TestFormatRefusals:

    def test_preserve_xls_refuses_with_conversion_hint(self, fixture_copy):
        src = fixture_copy("legacy/legacy.xls")
        with pytest.raises(UnsupportedStructureError, match="LibreOffice"):
            load_workbook(src, preserve=True)

    def test_preserve_xlsb_refuses_with_conversion_hint(self, fixture_copy):
        src = fixture_copy("legacy/binary.xlsb")
        with pytest.raises(UnsupportedStructureError, match="convert"):
            load_workbook(src, preserve=True)

    def test_refusals_are_paper_refusals(self, fixture_copy):
        with pytest.raises(PaperRefusal):
            load_workbook(fixture_copy("legacy/legacy.xls"), preserve=True)

    def test_stock_path_keeps_upstream_exception(self, fixture_copy):
        # strict superset: stock behavior unchanged
        with pytest.raises(InvalidFileException):
            load_workbook(fixture_copy("legacy/legacy.xls"))
        with pytest.raises(InvalidFileException):
            load_workbook(fixture_copy("legacy/binary.xlsb"))
