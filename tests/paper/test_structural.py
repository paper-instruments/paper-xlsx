"""Phase 6a: the structural-edit guard — reference-aware refusals under
preserve, loud warning on the stock path (PLAN Phase 6a; PR-0 §8)."""
from __future__ import annotations

import warnings

import pytest

from openpyxl import Workbook, load_workbook
from openpyxl.errors import StructuralShiftWarning, UnsupportedStructureError


class TestPreserveRefusals:
    """Since Phase 6b, fully-modeled sheets REWRITE instead of refusing
    (test_rewrite.py); these tests cover the sheets that must still refuse —
    anything carrying unmodeled range-bearing content."""

    def test_refusal_names_blockers_and_victims(self, fixture_copy):
        src = fixture_copy("gauntlet/gauntlet.xlsx")
        wb = load_workbook(src, preserve=True)
        with pytest.raises(UnsupportedStructureError) as exc:
            wb["Model"].insert_rows(1)
        msg = str(exc.value)
        assert "extLst" in msg or "extension" in msg    # the blocker
        assert "A1:F1" in msg                 # merged banner (victim)
        assert "chart" in msg.lower()         # preserved chart bytes
        assert "xl/charts/chart1.xml" in msg
        assert "Nothing was changed" in msg

    def test_column_shift_analyzed_too(self, fixture_copy):
        src = fixture_copy("gauntlet/gauntlet.xlsx")
        wb = load_workbook(src, preserve=True)
        with pytest.raises(UnsupportedStructureError) as exc:
            wb["Model"].delete_cols(2)
        assert "'Model'!B6" in str(exc.value)  # victim formula named

    def test_refusal_is_atomic(self, fixture_copy):
        src = fixture_copy("gauntlet/gauntlet.xlsx")
        wb = load_workbook(src, preserve=True)
        before_b6 = wb["Model"]["B6"].value
        before_max = wb["Model"].max_row
        with pytest.raises(UnsupportedStructureError):
            wb["Model"].insert_rows(3)
        assert wb["Model"]["B6"].value == before_b6
        assert wb["Model"].max_row == before_max
        assert wb._paper_ledger.cells == {}
        assert wb._paper_ledger.shifts == {}

    def test_added_sheets_stay_exempt(self, fixture_copy):
        wb = load_workbook(fixture_copy("gauntlet/gauntlet.xlsx"),
                           preserve=True)
        ws = wb.create_sheet("Scratch")
        ws.append([1, 2, 3])
        ws.insert_rows(1)                     # no refusal: generated whole
        assert ws.max_row == 2


class TestStockWarning:

    def test_loaded_workbook_warns_on_shift(self, fixture_copy):
        wb = load_workbook(fixture_copy("features/schedule.xlsx"))
        with pytest.warns(StructuralShiftWarning, match="updates NOTHING"):
            wb["Schedule"].insert_rows(5)
        # stock behavior unchanged: the shift still happened
        assert wb["Schedule"]["B13"].value == "=SUM(B2:B11)"

    def test_fresh_workbook_does_not_warn(self):
        wb = Workbook()
        ws = wb.active
        ws.append([1, 2, 3])
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            ws.insert_rows(1)
        assert not [w for w in caught
                    if isinstance(w.message, StructuralShiftWarning)]
