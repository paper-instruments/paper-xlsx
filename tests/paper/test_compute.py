"""The computation layer — explicit scenarios and certified write-back."""
from __future__ import annotations

import zipfile

import pytest

from openpyxl import Workbook, load_workbook, oracle
from openpyxl.errors import UnsupportedStructureError

needs_soffice = pytest.mark.skipif(
    not oracle.available(), reason="LibreOffice not installed")


class TestEvaluate:

    @needs_soffice
    @pytest.mark.lo_smoke
    def test_scenario_run_certified_and_untouched(self, fixture_copy):
        # battery job 12: one evaluate call with explicit certification state
        src = fixture_copy("features/schedule_calc.xlsx")
        with open(src, "rb") as f:
            before = f.read()
        ev = oracle.evaluate(
            src, set={"Schedule!B2": 1000, "Schedule!B3": 0},
            read=["Summary!B1", "Schedule!B12"])
        assert isinstance(ev, oracle.Evaluation)      # pinned return type
        assert ev.status == "ok"
        assert ev.outputs["Summary!B1"] == ev.outputs["Schedule!B12"]
        assert ev.outputs["Summary!B1"] == 6500 - 200 - 300 + 1000
        cert = ev.certification
        # Every formula touched by this scenario is input-dependent and is
        # therefore excluded from independent certification. Zero checked
        # formulas must never be reported as CERTIFIED.
        assert cert.status == "BASELINE_UNVERIFIABLE"
        assert cert.checked == 0
        assert cert.input_excluded
        assert "Summary!B1" in cert.input_excluded    # downstream of input
        payload = ev.to_dict()
        assert payload["schema"] == "evaluation"
        assert payload["version"] == 1
        assert payload["certification"]["status"] == \
            "BASELINE_UNVERIFIABLE"
        with open(src, "rb") as f:
            assert f.read() == before                 # original untouched

    @needs_soffice
    @pytest.mark.lo_smoke
    def test_evaluate_many_pool(self, fixture_copy):
        src = fixture_copy("features/schedule_calc.xlsx")
        cases = [{"Schedule!B2": v} for v in (100, 300)]
        results = oracle.evaluate_many(src, cases, ["Summary!B1"],
                                       pool_size=2)
        assert [e.outputs["Summary!B1"] for e in results] == [6400, 6600]
        assert all(e.certification.status == "BASELINE_UNVERIFIABLE"
                   and e.certification.checked == 0
                   and e.certification.input_excluded
                   for e in results)


class TestWriteBack:

    def test_path_required(self):
        with pytest.raises(ValueError, match="filesystem path"):
            oracle.write_back(b"PK\x03\x04junk")

    @needs_soffice
    @pytest.mark.lo_smoke
    def test_cacheless_write_back_and_clear(self, tmp_path):
        # battery job 24: the cache-less openpyxl file gains real caches
        wb = Workbook()
        ws = wb.active
        ws.title = "M"
        ws["A1"] = 10
        ws["A2"] = 32
        ws["A3"] = "=A1+A2"
        ws["A4"] = "=A3*2"
        p = str(tmp_path / "fresh.xlsx")
        wb.save(p)
        # cache-less = BASELINE_UNVERIFIABLE: gated
        with pytest.raises(UnsupportedStructureError,
                           match="certification-gated"):
            oracle.write_back(p)
        result = oracle.write_back(p, allow_uncertified=True)
        assert isinstance(result, oracle.WriteBackResult)  # pinned type
        assert result.uncertified is True             # the loud stamp
        assert result.cells_written == 2
        # an uncertified write NEVER clears the recalc flag: Excel must
        # not be told to trust caches nobody verified
        assert result.cleared_fullcalc is False
        assert "xl/worksheets/sheet1.xml" in result.package_diff
        payload = result.to_dict()
        assert payload["schema"] == "oracle_write_back"
        wb2 = load_workbook(p, data_only=True)
        assert wb2["M"]["A3"].value == 42
        assert wb2["M"]["A4"].value == 84
        wb3 = load_workbook(p)
        assert wb3["M"]["A3"].value == "=A1+A2"       # formulas intact
        with zipfile.ZipFile(p) as z:
            assert b"fullCalcOnLoad" in z.read("xl/workbook.xml")
        # second run: the caches now verify, so the CERTIFIED pass writes
        # nothing new and MAY clear the flag (full coverage, certified)
        result2 = oracle.write_back(p)
        assert result2.uncertified is False
        assert result2.cells_written == 0
        assert result2.cleared_fullcalc is True
        assert result2.package_diff == ["xl/workbook.xml"]
        with zipfile.ZipFile(p) as z:
            assert b"fullCalcOnLoad" not in z.read("xl/workbook.xml")

    @needs_soffice
    @pytest.mark.lo_smoke
    def test_write_back_is_macro_safe(self, fixture_copy):
        # the splice writes values into the ORIGINAL package: LibreOffice
        # bytes never enter the output, so vbaProject.bin survives
        src = fixture_copy("features/macro_stub.xlsm")
        with zipfile.ZipFile(src) as z:
            vba_before = z.read("xl/vbaProject.bin")
        oracle.write_back(src, allow_uncertified=True)
        with zipfile.ZipFile(src) as z:
            assert z.read("xl/vbaProject.bin") == vba_before


class TestCacheSplice:

    def test_datetime_cache_serializes_as_serial(self, fixture_copy,
                                                 tmp_path):
        import datetime

        src = fixture_copy("features/schedule_calc.xlsx")
        wb = load_workbook(src, preserve=True)
        ws = wb["Schedule"]
        target = next((r, c) for (r, c), cell in sorted(ws._cells.items())
                      if cell.data_type == "f")
        wb._paper_ledger.cache_writes.setdefault(ws, {})[target] = \
            datetime.datetime(2026, 7, 8)
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        wb2 = load_workbook(out)
        cell = wb2["Schedule"].cell(row=target[0], column=target[1])
        assert cell.data_type == "f"                  # formula untouched

    def test_cache_write_on_non_formula_refuses(self, fixture_copy,
                                                tmp_path):
        src = fixture_copy("features/schedule_calc.xlsx")
        wb = load_workbook(src, preserve=True)
        ws = wb["Schedule"]
        wb._paper_ledger.cache_writes.setdefault(ws, {})[(2, 1)] = 1.0
        with pytest.raises(Exception, match="formula"):
            wb.save(str(tmp_path / "o.xlsx"))
