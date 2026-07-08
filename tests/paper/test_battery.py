"""The five-job acceptance battery (PLAN Phase 1).

Two halves:

- ``TestStockCarnageBaseline`` runs each job against STOCK behavior and
  asserts the exact damage measured in Phase 0 (OPEN-QUESTIONS.md Q11). These
  tests PASS today. They are the fork's justification artifact and its
  permanent regression guard: if upstream behavior ever changes, these fail
  and the damage model gets re-derived, not assumed.

- ``TestBatterySafety`` encodes the forever pass-criterion: each job ends
  correct or loudly refused — never silently wrong. Each is a strict xfail
  naming the phase that flips it; when that phase lands, the xfail must be
  removed (strict=True makes an unexpected pass a failure, so flipping is
  forced, not optional).

Fixture note: the corpus is openpyxl-authored (+ zip surgery), so stock
carnage here UNDERSTATES real-file damage — openpyxl round-trips content it
modeled itself (chart survival is asserted below as fixture-specific
fairness, not as a general claim). Real-Excel fixtures land via
FIXTURE-REQUESTS.md and extend this baseline.
"""
from __future__ import annotations

import io
import zipfile

import pytest

from openpyxl import load_workbook

from .support.partdiff import part_payloads


def _sheet_payload(path_or_bytes, marker):
    parts = part_payloads(path_or_bytes)
    for name, payload in parts.items():
        if name.startswith("xl/worksheets/sheet") and marker in payload:
            return name, payload
    raise AssertionError("no sheet part contains {0!r}".format(marker))


class TestStockCarnageBaseline:
    """Damage-model rows reproduced against stock (PLAN's table, corrected by
    Phase-0 evidence). Every assertion here is a measured stock behavior."""

    def test_job1_assumption_flip_kills_extensions_and_shared_groups(
            self, fixture_copy, tmp_path):
        src = fixture_copy("gauntlet/gauntlet.xlsx")
        _, model_before = _sheet_payload(src, b"Quarterly Model")
        _, calc_before = _sheet_payload(src, b"double")
        assert b"sparklineGroups" in model_before
        assert b"x14:conditionalFormattings" in model_before
        assert b"<x14:id>" in model_before
        assert b'<f t="shared"' in calc_before

        # extension drop is warned at LOAD (Q11: not silent, and load-time)
        with pytest.warns(UserWarning, match="extension is not supported"):
            wb = load_workbook(src)
        wb["Model"]["B8"] = 0.15
        out = str(tmp_path / "job1_out.xlsx")
        wb.save(out)

        _, model_after = _sheet_payload(out, b"Quarterly Model")
        _, calc_after = _sheet_payload(out, b"double")
        # the intended edit landed...
        assert b'r="B8"' in model_after and b"0.15" in model_after
        # ...and everything half-understood died:
        assert b"sparklineGroups" not in model_after            # sparklines gone
        assert b"x14:conditionalFormattings" not in model_after  # x14 CF twin gone
        assert b"<x14:id>" not in model_after                    # twin pointer gone
        assert b'<f t="shared"' not in calc_after                # shared group dissolved
        # fixture-specific fairness: openpyxl-authored chart/image survive
        parts_after = part_payloads(out)
        assert any(n.startswith("xl/charts/") for n in parts_after)
        assert any(n.startswith("xl/media/") for n in parts_after)
        # values are stale by construction: formulas carry no cached value
        assert b"<f>SUM(Data!B2:B5)</f>" in model_after

    def test_job2_pandas_append_kills_extensions(self, fixture_copy, tmp_path):
        pd = pytest.importorskip("pandas")
        src = fixture_copy("gauntlet/gauntlet.xlsx")
        _, model_before = _sheet_payload(src, b"Quarterly Model")
        assert b"sparklineGroups" in model_before
        rels_before = part_payloads(src)["xl/_rels/workbook.xml.rels"]

        df = pd.DataFrame({"a": [1, 2], "b": [3, 4]})
        with pd.ExcelWriter(src, engine="openpyxl", mode="a") as writer:
            df.to_excel(writer, sheet_name="Appended", index=False)

        parts = part_payloads(src)
        # the intended append landed...
        wb = load_workbook(src)
        assert "Appended" in wb.sheetnames
        # ...but the sheet nobody touched lost its extensions,
        _, model_after = _sheet_payload(src, b"Quarterly Model")
        assert b"sparklineGroups" not in model_after
        # and the whole package was regenerated (existing rels renumbered)
        assert parts["xl/_rels/workbook.xml.rels"] != rels_before

    def test_job3_insert_rows_corrupts_references_silently(
            self, fixture_copy, tmp_path):
        src = fixture_copy("features/schedule.xlsx")
        wb = load_workbook(src)
        ws = wb["Schedule"]
        assert ws["B12"].value == "=SUM(B2:B11)"
        ws.insert_rows(5)
        out = str(tmp_path / "job3_out.xlsx")
        wb.save(out)

        wb2 = load_workbook(out)
        ws2 = wb2["Schedule"]
        # data now spans B2:B12 (a row was inserted above the range end)...
        # ...but the moved TOTAL still sums the OLD range: silent corruption
        assert ws2["B13"].value == "=SUM(B2:B11)"
        # the defined name still points at the OLD input cell (now empty)
        assert wb2.defined_names["Growth"].value == "Schedule!$B$15"
        assert ws2["B15"].value is None          # input moved to B16
        assert ws2["B16"].value == 0.05
        # the cross-sheet reference still points at the old TOTAL slot,
        # which now holds a data row — a plausible-looking wrong number
        assert wb2["Summary"]["B1"].value == "=Schedule!B12"
        assert ws2["B12"].value == 1100          # was "Item 10"'s amount, shifted

    def test_job4_xlsm_roundtrip_strips_vba_silently(self, fixture_copy, tmp_path):
        src = fixture_copy("features/macro_stub.xlsm")
        vba_before = part_payloads(src)["xl/vbaProject.bin"]

        wb = load_workbook(src)  # no keep_vba
        out = str(tmp_path / "job4_out.xlsm")
        wb.save(out)
        parts = part_payloads(out)
        assert "xl/vbaProject.bin" not in parts               # VBA gone
        assert b"macroEnabled" not in parts["[Content_Types].xml"]

        # contrast: keep_vba=True preserves it byte-identically (the in-tree
        # retention precedent the spine generalizes)
        wb2 = load_workbook(src, keep_vba=True)
        out2 = str(tmp_path / "job4_keepvba.xlsm")
        wb2.save(out2)
        assert part_payloads(out2)["xl/vbaProject.bin"] == vba_before

    def test_job5_data_only_save_destroys_all_formulas(self, fixture_copy, tmp_path):
        src = fixture_copy("features/schedule_calc.xlsx")
        # marker is the formula text: LO-written files keep strings in the
        # shared-strings part, so cell text like "TOTAL" is not in sheet XML
        _, sheet_before = _sheet_payload(src, b"SUM(B2:B11)")
        assert sheet_before.count(b"<f") >= 2  # formulas present in the file

        wb = load_workbook(src, data_only=True)
        assert wb["Schedule"]["B12"].value == 6500  # cached values read fine
        out = str(tmp_path / "job5_out.xlsx")
        wb.save(out)                                # ...and this is the trap

        _, sheet_after = _sheet_payload(out, b"6500")
        assert b"<f" not in sheet_after             # every formula destroyed
        wb2 = load_workbook(out)
        assert wb2["Schedule"]["B12"].value == 6500  # literal, not =SUM(...)

    def test_stale_cached_values_row_pipelines_read_nothing(self, fixture_copy):
        pd = pytest.importorskip("pandas")
        # openpyxl-written formulas carry no cached values -> pandas sees NaN
        stale = fixture_copy("features/schedule.xlsx")
        df = pd.read_excel(stale, sheet_name="Schedule", header=0)
        total_cell = df[df.iloc[:, 0] == "TOTAL"].iloc[0, 1]
        assert pd.isna(total_cell)
        # the LibreOffice-recalculated twin reads real numbers
        calc = fixture_copy("features/schedule_calc.xlsx")
        df2 = pd.read_excel(calc, sheet_name="Schedule", header=0)
        total_cell2 = df2[df2.iloc[:, 0] == "TOTAL"].iloc[0, 1]
        assert total_cell2 == 6500


class TestBatterySafety:
    """The forever criterion: correct or loudly refused — never silently
    wrong. Strict xfails; each names the phase that must flip it."""

    @pytest.mark.xfail(reason="preserve mode splice lands in Phase 2c/2d", strict=True)
    def test_job1_assumption_flip_preserves_everything(self, fixture_copy, tmp_path):
        src = fixture_copy("gauntlet/gauntlet.xlsx")
        wb = load_workbook(src, preserve=True)
        wb["Model"]["B8"] = 0.15
        out = str(tmp_path / "job1_safe.xlsx")
        wb.save(out)
        _, model_after = _sheet_payload(out, b"Quarterly Model")
        assert b"sparklineGroups" in model_after
        assert b"x14:conditionalFormattings" in model_after
        assert b"<x14:id>" in model_after
        wb2 = load_workbook(out)
        assert wb2["Model"]["B8"].value == 0.15

    @pytest.mark.xfail(reason="preserve mode via pandas lands in Phase 2d", strict=True)
    def test_job2_pandas_append_preserves_everything(self, fixture_copy):
        pd = pytest.importorskip("pandas")
        src = fixture_copy("gauntlet/gauntlet.xlsx")
        df = pd.DataFrame({"a": [1, 2], "b": [3, 4]})
        with pd.ExcelWriter(src, engine="openpyxl", mode="a",
                            engine_kwargs={"preserve": True}) as writer:
            df.to_excel(writer, sheet_name="Appended", index=False)
        _, model_after = _sheet_payload(src, b"Quarterly Model")
        assert b"sparklineGroups" in model_after
        wb = load_workbook(src)
        assert "Appended" in wb.sheetnames

    @pytest.mark.xfail(reason="structural-edit guard lands in Phase 6a", strict=True)
    def test_job3_insert_rows_refuses_or_rewrites(self, fixture_copy, tmp_path):
        src = fixture_copy("features/schedule.xlsx")
        with open(src, "rb") as f:
            before = f.read()
        wb = load_workbook(src, preserve=True)
        ws = wb["Schedule"]
        try:
            ws.insert_rows(5)
            wb.save(src)
        except Exception as exc:
            # refusal path: must be typed and atomic
            from openpyxl.errors import PaperRefusal
            assert isinstance(exc, PaperRefusal)
            with open(src, "rb") as f:
                assert f.read() == before
        else:
            # rewrite path (Phase 6b): references must have followed the shift
            wb2 = load_workbook(src)
            assert wb2["Schedule"]["B13"].value == "=SUM(B2:B12)"
            assert wb2.defined_names["Growth"].value == "Schedule!$B$16"
            assert wb2["Summary"]["B1"].value == "=Schedule!B13"

    @pytest.mark.xfail(reason="byte retention + raw copy lands in Phase 2a", strict=True)
    def test_job4_xlsm_roundtrip_preserves_vba(self, fixture_copy, tmp_path):
        src = fixture_copy("features/macro_stub.xlsm")
        vba_before = part_payloads(src)["xl/vbaProject.bin"]
        wb = load_workbook(src, preserve=True)  # note: NO keep_vba flag needed
        out = str(tmp_path / "job4_safe.xlsm")
        wb.save(out)
        parts = part_payloads(out)
        assert parts["xl/vbaProject.bin"] == vba_before
        assert b"macroEnabled" in parts["[Content_Types].xml"]

    @pytest.mark.xfail(reason="data_only save refusal lands in Phase 3", strict=True)
    def test_job5_data_only_save_refuses(self, fixture_copy, tmp_path):
        src = fixture_copy("features/schedule_calc.xlsx")
        with open(src, "rb") as f:
            before = f.read()
        wb = load_workbook(src, preserve=True, data_only=True)
        assert wb["Schedule"]["B12"].value == 6500  # reading stays legal
        out = str(tmp_path / "job5_safe.xlsx")
        from openpyxl.errors import PaperRefusal
        with pytest.raises(PaperRefusal):
            wb.save(out)
        import os
        assert not os.path.exists(out)  # refusal left nothing behind
        with open(src, "rb") as f:
            assert f.read() == before
