"""Perception helpers for cell diffs and dependency sketches."""
from __future__ import annotations

from openpyxl import Workbook, load_workbook
from openpyxl.package import diff_cells
from openpyxl.preserve.perception import dependency_sketch
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table, TableColumn

class TestDiffCells:

    def test_value_and_formula_changes(self, fixture_copy, tmp_path):
        src = fixture_copy("features/schedule.xlsx")
        wb = load_workbook(src, preserve=True)
        wb["Schedule"]["B2"] = 999
        wb["Schedule"]["B20"] = "=SUM(B2:B4)"
        out = str(tmp_path / "b.xlsx")
        wb.save(out)
        d = diff_cells(src, out)
        by_addr = {c["address"]: c for c in d.changes}
        assert by_addr["'Schedule'!B2"]["old_value"] == 200
        assert by_addr["'Schedule'!B2"]["new_value"] == 999
        assert by_addr["'Schedule'!B20"]["new_formula"] == "=SUM(B2:B4)"
        assert not d.sheets_added and not d.sheets_removed

    def test_identical_files_are_clean(self, fixture_copy):
        a = fixture_copy("features/schedule.xlsx", "a.xlsx")
        b = fixture_copy("features/schedule.xlsx", "b.xlsx")
        assert diff_cells(a, b).clean

    def test_added_sheet_reported(self, fixture_copy, tmp_path):
        src = fixture_copy("minimal/minimal_clean.xlsx")
        wb = load_workbook(src, preserve=True)
        wb.create_sheet("Extra")["A1"] = 1
        out = str(tmp_path / "b.xlsx")
        wb.save(out)
        d = diff_cells(src, out)
        assert d.sheets_added == ["Extra"]

    def test_to_dict_schema(self, fixture_copy):
        a = fixture_copy("minimal/minimal_clean.xlsx")
        doc = diff_cells(a, a).to_dict()
        assert doc["schema"] == "cells_diff" and doc["version"] == 1


class TestDependencySketch:

    def test_known_edges(self, fixture_copy):
        wb = load_workbook(fixture_copy("gauntlet/gauntlet.xlsx"),
                           preserve=True)
        sk = dependency_sketch(wb)
        doc = sk.to_dict()
        assert doc["references"]["'Model'!B6"] == ["B3", "B4", "B5"]
        assert doc["references"]["'Model'!B12"] == ["Data!B2:B5"]

    def test_intersection_query_cross_sheet(self, fixture_copy):
        wb = load_workbook(fixture_copy("gauntlet/gauntlet.xlsx"),
                           preserve=True)
        sk = dependency_sketch(wb)
        # who references Data!B2:B5?
        assert sk.cells_referencing("Data", (2, 2, 2, 5)) == ["'Model'!B12"]
        # nobody references Data column D
        assert sk.cells_referencing("Data", (4, 1, 4, 100)) == []

    def test_defined_name_expansion(self, fixture_copy):
        wb = load_workbook(fixture_copy("features/schedule.xlsx"),
                           preserve=True)
        sk = dependency_sketch(wb)
        # B13 = B12*(1+Growth); Growth -> Schedule!$B$15
        hits = sk.cells_referencing("Schedule", (2, 15, 2, 15))
        assert "'Schedule'!B13" in hits

    def test_common_structured_refs_resolve_to_table_ranges(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["Name", "Amount", "Tax"])
        ws.append(["A", 10, 1])
        ws.append(["B", 20, 2])
        ws.append(["C", 30, 3])
        ws.append(["Total", 60, 6])
        table = Table(displayName="Table1", ref="A1:C5")
        table.totalsRowCount = 1
        ws.add_table(table)
        ws["E1"] = "=SUM(Table1[Amount])"
        ws["E2"] = "=SUM(Table1[[#Data],[Amount]])"
        ws["E3"] = "=SUM(Table1[[Amount]:[Tax]])"
        ws["E4"] = "=SUM(Table1[#Headers])"
        ws["E5"] = "=SUM(Table1[#Totals])"
        ws["E6"] = "=SUM(Table1[#All])"
        ws["E7"] = "=SUM(Table1[[#Headers],[Amount]:[Tax]])"
        ws["E8"] = "=SUM(tAbLe1[amount])"
        path = tmp_path / "structured.xlsx"
        wb.save(path)
        wb = load_workbook(path, preserve=True)

        sk = dependency_sketch(wb)
        refs = {
            address: bounds
            for address, [(_sheet, bounds, _raw)] in sk.references.items()
        }
        assert refs["'Data'!E1"] == (2, 2, 2, 4)
        assert refs["'Data'!E2"] == (2, 2, 2, 4)
        assert refs["'Data'!E3"] == (2, 2, 3, 4)
        assert refs["'Data'!E4"] == (1, 1, 3, 1)
        assert refs["'Data'!E5"] == (1, 5, 3, 5)
        assert refs["'Data'!E6"] == (1, 1, 3, 5)
        assert refs["'Data'!E7"] == (2, 1, 3, 1)
        assert refs["'Data'!E8"] == (2, 2, 2, 4)
        assert sk.unresolved == {}

    def test_current_row_structured_refs_resolve_exact_row(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.append(["Amount", "Calc"])
        ws.append([10, "=[@Amount]"])
        ws.append([20, "=Table1[@Amount]"])
        ws.add_table(Table(displayName="Table1", ref="A1:B3"))
        path = tmp_path / "current-row.xlsx"
        wb.save(path)
        wb = load_workbook(path, preserve=True)

        sk = dependency_sketch(wb)
        refs = {
            address: bounds
            for address, [(_sheet, bounds, _raw)] in sk.references.items()
        }
        assert refs["'Sheet'!B2"] == (1, 2, 1, 2)
        assert refs["'Sheet'!B3"] == (1, 3, 1, 3)
        assert sk.cells_referencing("Sheet", (1, 3, 1, 3)) == ["'Sheet'!B3"]

    def test_external_workbook_reference_stays_unresolved(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "=[Book.xlsx]Sheet!A1"

        sk = dependency_sketch(wb)
        assert sk.to_dict()["unresolved"] == {
            "'Sheet'!A1": ["[Book.xlsx]Sheet!A1"]
        }
        assert "'Sheet'!A1" in sk.cells_referencing("Anywhere", (1, 1, 1, 1))

    def test_current_row_in_multi_cell_formula_stays_unresolved(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["Amount", "Calc"])
        ws.append([10, None])
        ws.append([20, None])
        ws.add_table(Table(displayName="Table1", ref="A1:B3"))
        ws["B2"] = ArrayFormula(ref="B2:B3", text="=[@Amount]")

        sk = dependency_sketch(wb)

        assert sk.to_dict()["unresolved"] == {
            "'Sheet'!B2": ["[@Amount]"]
        }

    def test_unproven_structured_refs_stay_unresolved(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["Amount"])
        ws.append([10])
        ws.add_table(Table(displayName="Only", ref="A1:A2"))
        ws["C1"] = "=Missing[Amount]"
        ws["C2"] = "=Only[Missing]"

        duplicate = Table(displayName="DupCols", ref="A1:B2",
                          tableColumns=[
                              TableColumn(id=1, name="Amount"),
                              TableColumn(id=2, name="amount"),
                          ])
        ws["B1"] = "amount"
        ws._tables.add(duplicate)
        ws["C3"] = "=DupCols[Amount]"
        ws["C5"] = "=Only[[#Headers],[#Data],[Amount]]"

        empty = wb.create_sheet("Empty")
        empty["A1"] = "Amount"
        empty.add_table(Table(displayName="EmptyTable", ref="A1:A1"))
        ws["C6"] = "=EmptyTable[Missing]"

        other = wb.create_sheet("Other")
        other.append(["Amount"])
        other.append([20])
        ws["D1"] = "Amount"
        ws["D2"] = 10
        ws._tables.add(Table(displayName="Table1", ref="D1:D2",
                             tableColumns=[
                                 TableColumn(id=1, name="Amount"),
                             ]))
        other._tables.add(Table(displayName="table1", ref="A1:A2",
                                tableColumns=[
                                    TableColumn(id=1, name="Amount"),
                                ]))
        ws["C4"] = "=Table1[Amount]"

        sk = dependency_sketch(wb)
        assert sk.to_dict()["unresolved"] == {
            "'Sheet'!C1": ["Missing[Amount]"],
            "'Sheet'!C2": ["Only[Missing]"],
            "'Sheet'!C3": ["DupCols[Amount]"],
            "'Sheet'!C4": ["Table1[Amount]"],
            "'Sheet'!C5": ["Only[[#Headers],[#Data],[Amount]]"],
            "'Sheet'!C6": ["EmptyTable[Missing]"],
        }

    def test_missing_table_regions_stay_unresolved(self):
        # a selector whose region does not exist must stay unresolved
        # (always-intersecting), never resolve to an empty range list: an
        # operand recorded in neither map is invisible to the move guards
        # and to certification taint
        wb = Workbook()
        ws = wb.active

        empty = wb.create_sheet("Empty")
        empty["A1"] = "Amount"
        empty.add_table(Table(displayName="EmptyData", ref="A1:A1"))
        ws["C1"] = "=SUM(EmptyData[Amount])"
        ws["C2"] = "=SUM(EmptyData[#Data])"

        headerless = wb.create_sheet("Headerless")
        headerless["A1"] = 10
        headerless.add_table(Table(displayName="NoHeaders", ref="A1:A1",
                                   headerRowCount=0,
                                   tableColumns=[
                                       TableColumn(id=1, name="Amount"),
                                   ]))
        ws["C3"] = "=SUM(NoHeaders[#Headers])"

        plain = wb.create_sheet("Plain")
        plain.append(["Amount"])
        plain.append([10])
        plain.add_table(Table(displayName="NoTotals", ref="A1:A2"))
        ws["C4"] = "=SUM(NoTotals[#Totals])"

        # an ambiguous totals flag collapses the data region: the operand
        # still reads a real cell, so it must not vanish
        ambiguous = wb.create_sheet("Ambiguous")
        ambiguous.append(["Amount"])
        ambiguous.append([10])
        flagged = Table(displayName="Flagged", ref="A1:A2",
                        tableColumns=[TableColumn(id=1, name="Amount")])
        flagged.totalsRowShown = True
        ambiguous._tables.add(flagged)
        ws["C5"] = "=SUM(Flagged[Amount])"

        sk = dependency_sketch(wb)
        assert sk.to_dict() == {
            "schema": "dependency_sketch",
            "version": 1,
            "references": {},
            "unresolved": {
                "'Sheet'!C1": ["EmptyData[Amount]"],
                "'Sheet'!C2": ["EmptyData[#Data]"],
                "'Sheet'!C3": ["NoHeaders[#Headers]"],
                "'Sheet'!C4": ["NoTotals[#Totals]"],
                "'Sheet'!C5": ["Flagged[Amount]"],
            },
        }
        # and they stay conservative: every one is reported for any region
        assert sk.cells_referencing("Empty", (1, 1, 1, 1)) == [
            "'Sheet'!C1", "'Sheet'!C2", "'Sheet'!C3", "'Sheet'!C4",
            "'Sheet'!C5",
        ]


def _full_sketch_feeds(wb):
    """Whether a value edit feeds a formula, from the whole sketch: the
    answer the save gate gave before it learned to stop early."""
    from openpyxl.preserve.perception import _intersects

    sketch = dependency_sketch(wb)
    if sketch.unresolved:
        return True
    for ws, dirty in wb._paper_ledger.value_overwrites.items():
        for refs in sketch.references.values():
            for ref_sheet, bounds, _raw in refs:
                if ref_sheet.casefold() == ws.title.casefold() and any(
                        _intersects(bounds, col, row, col, row)
                        for row, col in dirty):
                    return True
    return False


class TestDirtyFeedsFormulas:

    import pytest

    @pytest.mark.parametrize("formula, edit, feeds", [
        ("=Data!A1*2", "Data!A1", True),
        ("=Data!A1*2", "Data!B9", False),
        ("=data!a1+1", "Data!A1", True),
        ("=SUM(Data!A:A)", "Data!A4", True),
        ("=SUM(Data!A:A)", "Data!C4", False),
        ("=rate*2", "Data!A3", True),
        ("=rate*2", "Data!A2", False),
        ('=INDIRECT("Data!A1")', "Data!B9", True),
        ("=A1+1", "Calc!A1", True),
        ("=A1+1", "Data!A1", False),
    ])
    def test_stops_early_with_the_full_sketch_answer(
            self, tmp_path, formula, edit, feeds):
        from openpyxl.preserve.saver import _dirty_feeds_formulas
        from openpyxl.workbook.defined_name import DefinedName

        from .support.partdiff import part_payloads

        wb = Workbook()
        data = wb.active
        data.title = "Data"
        for row in range(1, 6):
            data.cell(row, 1, row)
        calc = wb.create_sheet("Calc")
        calc["A1"] = 7
        calc["B2"] = "=Data!A5"
        calc["C3"] = formula
        wb.defined_names["rate"] = DefinedName("rate", attr_text="Data!$A$3")
        wb.calculation.fullCalcOnLoad = False
        src = tmp_path / "src.xlsx"
        wb.save(src)

        wb = load_workbook(src, preserve=True)
        sheet, coordinate = edit.split("!")
        wb[sheet][coordinate] = 99
        assert _dirty_feeds_formulas(wb, wb._paper_ledger) is feeds
        assert _full_sketch_feeds(wb) is feeds

        out = tmp_path / "out.xlsx"
        wb.save(out)
        reopened = load_workbook(out)
        assert reopened[sheet][coordinate].value == 99
        assert (b'fullCalcOnLoad="1"' in part_payloads(str(out))[
            "xl/workbook.xml"]) is feeds

    def test_titles_sharing_a_casefold_keep_both_sheets_edits(self, tmp_path):
        from openpyxl.preserve.saver import _dirty_feeds_formulas

        from .support.harness import assert_part_budget
        from .support.partdiff import part_payloads

        wb = Workbook()
        wb.active.title = "ß"
        wb.create_sheet("ss")
        wb.create_sheet("Calc")["C1"] = "='ß'!A1"
        wb.calculation.fullCalcOnLoad = False
        src = tmp_path / "src.xlsx"
        wb.save(src)

        wb = load_workbook(src, preserve=True)
        wb["ß"]["A1"] = 3
        wb["ss"]["B9"] = 4
        assert _dirty_feeds_formulas(wb, wb._paper_ledger) is True
        assert _full_sketch_feeds(wb) is True

        out = tmp_path / "out.xlsx"
        wb.save(out)
        reopened = load_workbook(out)
        assert (reopened["ß"]["A1"].value, reopened["ss"]["B9"].value) == (
            3, 4)
        assert b'fullCalcOnLoad="1"' in part_payloads(str(out))[
            "xl/workbook.xml"]
        assert_part_budget(src, out, expect_changed={
            "xl/workbook.xml", "xl/worksheets/sheet1.xml",
            "xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml"})

    import pytest

    def test_a_resolver_bug_is_not_read_as_an_unresolved_reference(
            self, fixture_copy, monkeypatch):
        # Only a malformed table's TypeError or ValueError means
        # "unresolved"; any other exception is a bug and must surface.
        import pytest

        from openpyxl.preserve import perception

        def broken(*_args, **_kwargs):
            raise AttributeError("resolver bug")

        wb = load_workbook(fixture_copy("features/tables.xlsx"),
                           preserve=True)
        wb["Data"]["C2"] = "=SUM(RegionTable[Amount])"
        monkeypatch.setattr(
            perception, "_structured_reference_ranges", broken)
        with pytest.raises(AttributeError, match="resolver bug"):
            dependency_sketch(wb)

    @pytest.mark.parametrize("fixture", [
        "features/table_open_ended_ref.xlsx",
        "features/table_zero_row_ref.xlsx",
    ])
    def test_a_malformed_table_is_unresolved_not_a_crash(
            self, fixture_copy, tmp_path, fixture):
        # A damaged table ref made resolving RegionTable[Amount] raise, which
        # failed every save; with the early exit it failed only when no
        # earlier formula answered, so the answer depended on cell order.
        from openpyxl.preserve.saver import _dirty_feeds_formulas

        from .support.harness import assert_part_budget
        from .support.partdiff import part_payloads

        source = fixture_copy(fixture)
        wb = load_workbook(source, preserve=True)
        sketch = dependency_sketch(wb)
        assert sketch.unresolved == {"'Data'!C3": ["RegionTable[Amount]"]}
        wb["Data"]["B2"] = 21
        assert _dirty_feeds_formulas(wb, wb._paper_ledger) is True
        assert _full_sketch_feeds(wb) is True

        out = tmp_path / "out.xlsx"
        wb.save(out)
        reopened = load_workbook(out)
        assert reopened["Data"]["B2"].value == 21
        assert reopened["Data"]["C3"].value == "=SUM(RegionTable[Amount])"
        cached = load_workbook(out, data_only=True)["Data"]
        assert (cached["C2"].value, cached["C3"].value) == (None, None)
        assert b'fullCalcOnLoad="1"' in part_payloads(str(out))[
            "xl/workbook.xml"]
        assert_part_budget(source, out, expect_changed={
            "xl/workbook.xml", "xl/worksheets/sheet1.xml"})
