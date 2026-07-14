"""Perception helpers for cell diffs and dependency sketches."""
from __future__ import annotations

from openpyxl import Workbook, load_workbook
from openpyxl.package import diff_cells
from openpyxl.preserve.perception import dependency_sketch


def test_workbook_has_no_manifest_api():
    assert not hasattr(Workbook(), "manifest")


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

    def test_structured_refs_are_conservative(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "=SUM(Table1[Amount])"
        sk = dependency_sketch(wb)
        assert "'Sheet'!A1" in sk.to_dict()["unresolved"]
        # unresolved references hit EVERY intersection query
        assert "'Sheet'!A1" in sk.cells_referencing("Anywhere", (1, 1, 1, 1))
