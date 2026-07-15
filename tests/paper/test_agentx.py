"""Perception and the agent experience —
locate, search, scan_errors, allowed_values, validate, model map,
receipts, structured refusals, findings, diffs."""
from __future__ import annotations

import io
import zipfile

import pytest

from openpyxl import Workbook, load_workbook
from openpyxl.errors import (
    AmbiguousTargetError,
    PaperRefusal,
    TargetNotFoundError,
    UnsupportedStructureError,
)


class TestLocate:

    def test_locate_right_and_below(self, fixture_copy):
        wb = load_workbook(fixture_copy("features/schedule.xlsx"))
        ws = wb["Summary"]
        assert ws.locate("Grand total").coordinate == "B1"
        ws["D1"] = "Rate"
        ws["D2"] = 0.05
        assert ws.locate("Rate", prefer="below").value == 0.05

    def test_normalized_match(self, fixture_copy):
        wb = load_workbook(fixture_copy("features/schedule.xlsx"))
        ws = wb["Summary"]
        assert ws.locate("  GRAND   TOTAL ").coordinate == "B1"

    def test_exact_beats_normalized(self, fixture_copy):
        wb = load_workbook(fixture_copy("features/schedule.xlsx"))
        ws = wb["Summary"]
        ws["A5"] = "grand total"              # normalized twin
        ws["B5"] = 42
        # exact match wins outright: no ambiguity
        assert ws.locate("Grand total").coordinate == "B1"
        assert ws.locate("grand total").coordinate == "B5"

    def test_refusals_carry_structured_fields(self, fixture_copy):
        wb = load_workbook(fixture_copy("features/schedule.xlsx"))
        ws = wb["Summary"]
        with pytest.raises(TargetNotFoundError) as exc:
            ws.locate("Nothing Here")
        assert exc.value.kind == "label-not-found"
        assert isinstance(exc.value, PaperRefusal)


class TestSearchAndScan:

    def test_search_values_and_formulas(self, fixture_copy):
        wb = load_workbook(fixture_copy("features/schedule.xlsx"))
        hits = wb.search("Grand")
        assert {"address": "Summary!A1", "match": "Grand",
                "kind": "value"} in hits
        formula_hits = wb.search("Schedule!", values=False)
        assert all(h["kind"] == "formula" for h in formula_hits)
        assert formula_hits

    def test_search_regex(self, fixture_copy):
        wb = load_workbook(fixture_copy("features/schedule.xlsx"))
        hits = wb.search(r"Item \d", regex=True, formulas=False)
        assert len(hits) >= 2

    def test_scan_errors_sees_cached_and_formula_refs(self, fixture_copy,
                                                      tmp_path):
        from openpyxl.preserve import scan_errors

        src = fixture_copy("features/schedule.xlsx")
        crafted = str(tmp_path / "err.xlsx")
        with zipfile.ZipFile(src) as zin, \
                zipfile.ZipFile(crafted, "w") as zout:
            for name in zin.namelist():
                payload = zin.read(name)
                if name == "xl/worksheets/sheet1.xml":
                    payload = payload.replace(
                        b"</sheetData>",
                        b'<row r="30"><c r="A30" t="e"><f>1/0</f>'
                        b"<v>#DIV/0!</v></c></row></sheetData>", 1)
                zout.writestr(name, payload)
        wb = load_workbook(crafted, preserve=True)
        wb["Summary"]["C1"] = "=#REF!+1"
        results = scan_errors(wb)
        sources = {r["source"] for r in results}
        assert {"cache", "formula"} <= sources

    def test_allowed_values_literal_and_range(self, fixture_copy):
        from openpyxl.worksheet.datavalidation import DataValidation

        wb = load_workbook(fixture_copy("features/schedule.xlsx"))
        ws = wb["Summary"]
        dv = DataValidation(type="list", formula1='"Yes,No,Maybe"')
        dv.add("D1")
        ws.add_data_validation(dv)
        assert ws.allowed_values("D1") == ["Yes", "No", "Maybe"]
        dv2 = DataValidation(type="list",
                             formula1="=Schedule!$A$2:$A$4")
        dv2.add("D2")
        ws.add_data_validation(dv2)
        assert ws.allowed_values(ws["D2"]) == ["Item 1", "Item 2",
                                               "Item 3"]
        assert ws.allowed_values("E9") is None


class TestValidateAndReceipt:

    def test_validate_raises_what_save_would(self, fixture_copy):
        wb = load_workbook(fixture_copy("features/chart_image.xlsx"),
                           preserve=True)
        wb["Model"]._charts[0].style = 31     # inexpressible mutation
        with pytest.raises(UnsupportedStructureError, match="style"):
            wb.validate()

    def test_validate_clean_session_returns_none(self, fixture_copy):
        wb = load_workbook(fixture_copy("features/schedule.xlsx"),
                           preserve=True)
        wb["Schedule"]["A2"] = "edited"
        assert wb.validate() is None

    def test_save_receipt(self, fixture_copy, tmp_path):
        from openpyxl.preserve.receipts import EditReceipt

        wb = load_workbook(fixture_copy("features/schedule.xlsx"),
                           preserve=True)
        wb["Schedule"]["A2"] = "renamed item"
        out = str(tmp_path / "o.xlsx")
        result = wb.save(out, receipt=True)
        assert isinstance(result, EditReceipt)
        payload = result.to_dict()
        assert payload["schema"] == "edit_receipt"
        assert payload["version"] == 1
        sheet_part = next(iter(result.cells_changed))
        assert result.cells_changed[sheet_part] == {"A2": "changed"}
        assert sheet_part in result.parts_changed

    def test_receipt_requires_preserve(self, fixture_copy, tmp_path):
        wb = load_workbook(fixture_copy("features/schedule.xlsx"))
        with pytest.raises(ValueError, match="preserve"):
            wb.save(str(tmp_path / "o.xlsx"), receipt=True)


class TestModelMap:

    def test_roles_and_schema(self, fixture_copy):
        wb = load_workbook(fixture_copy("features/schedule_calc.xlsx"))
        mm = wb.model_map()
        payload = mm.to_dict()
        assert payload["schema"] == "model_map"
        assert payload["version"] == 1
        schedule = payload["sheets"]["Schedule"]
        # B2:B5 feed B12 -> inputs; B12 feeds B13 and Summary!B1 ->
        # calculation; B13 unreferenced -> output
        for addr in ("B2", "B3", "B4", "B5"):
            assert addr in schedule["inputs"]
        assert "B12" in schedule["calculations"]
        assert "B13" in schedule["outputs"]
        assert "A1" in schedule["constants"]  # header text
        assert mm.inputs("Schedule")

class TestFindings:

    def test_taxonomy_measurements(self, fixture_copy, tmp_path):
        from openpyxl.preserve import findings
        from openpyxl.preserve.hygiene import FINDING_KINDS, Finding

        wb = Workbook()
        ws = wb.active
        ws.title = "M"
        for i in range(1, 6):
            ws.cell(row=i, column=1, value=i * 10)
            ws.cell(row=i, column=2, value="=A{0}*1.17".format(i))
        ws["B3"] = "=A3*2+99.5"               # breaks the run + hardcode
        ws["C1"] = "=NOW()"                   # volatile
        ws["D1"] = "#REF!"
        ws["A9"] = 5
        ws["A10"] = 6
        ws["A11"] = 7
        ws["A12"] = 8
        ws["A13"] = 90000000                  # magnitude outlier
        hidden = wb.create_sheet("Hidden")
        hidden["A1"] = 1
        hidden.sheet_state = "hidden"
        ws.row_dimensions[7].hidden = True
        results = findings(wb)
        kinds = {f.kind for f in results}
        assert all(isinstance(f, Finding) for f in results)
        assert all(f.kind in FINDING_KINDS for f in results)
        assert "inconsistent-row-formula" in kinds
        assert "hardcode-in-formula" in kinds
        assert "volatile" in kinds
        assert "error-cell" in kinds
        assert "hidden-sheet" in kinds
        assert "hidden-rows" in kinds
        assert "magnitude-outlier" in kinds
        for f in results:
            assert f.evidence                 # measurements carry evidence

    def test_orphaned_name(self, fixture_copy):
        from openpyxl.preserve import findings
        from openpyxl.workbook.defined_name import DefinedName

        wb = Workbook()
        wb.active["A1"] = 1
        wb.active["B1"] = "=A1"
        wb.defined_names["dead"] = DefinedName(
            "dead", attr_text="Gone!$A$1")
        kinds = {f.kind: f for f in findings(wb)}
        assert "orphaned-name" in kinds
        assert any("dead" in e for e in kinds["orphaned-name"].evidence)


class TestDiffWorkbooks:

    def test_content_vs_shifted(self, fixture_copy, tmp_path):
        from openpyxl.preserve import diff_workbooks

        src = fixture_copy("features/schedule.xlsx")
        wb = load_workbook(src, preserve=True)
        ws = wb["Schedule"]
        remap = ws.insert_rows(1)             # AddressRemap
        ws["B3"] = 999                        # content change (was B2)
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        report = diff_workbooks(src, out, remaps=[remap])
        payload = report.to_dict()
        assert payload["schema"] == "workbook_diff"
        changed_addrs = {e["address"] for e in report.changed}
        assert "Schedule!B2" in changed_addrs  # the value edit
        # everything else moved intact: classified shifted, not changed
        shifted_from = {e["from"] for e in report.shifted}
        assert "Schedule!B4" in shifted_from or "Schedule!A4" in \
            shifted_from
        assert not any(e["address"].startswith("Schedule!A")
                       and e["address"] != "Schedule!A1"
                       for e in report.changed
                       if e["before"] is not None
                       and e["after"] is not None)

    def test_sheet_membership(self, fixture_copy, tmp_path):
        from openpyxl.preserve import diff_workbooks

        src = fixture_copy("features/schedule.xlsx")
        wb = load_workbook(src, preserve=True)
        ws = wb.create_sheet("New")
        ws["A1"] = 1
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        report = diff_workbooks(src, out)
        assert report.added_sheets == ["New"]
        assert report.removed_sheets == []
