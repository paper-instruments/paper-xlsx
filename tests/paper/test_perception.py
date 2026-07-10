"""Perception — manifest, diff_cells, dependency sketch
."""
from __future__ import annotations

import json
import os

import pytest

from openpyxl import Workbook, load_workbook
from openpyxl.package import diff_cells
from openpyxl.preserve.perception import dependency_sketch

GOLDENS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "goldens")


class TestManifest:

    def test_gauntlet_manifest_matches_golden(self, fixture_copy):
        # golden files update only via explicit command with human-reviewed
        # diffs
        wb = load_workbook(fixture_copy("gauntlet/gauntlet.xlsx"),
                           preserve=True)
        doc = wb.manifest().to_dict()
        with open(os.path.join(GOLDENS, "gauntlet_manifest.json")) as f:
            golden = json.load(f)
        assert json.loads(json.dumps(doc, sort_keys=True)) == golden

    def test_confession_comes_from_the_package_not_the_model(
            self, fixture_copy):
        wb = load_workbook(fixture_copy("features/macro_stub.xlsm"),
                           preserve=True)
        conf = wb.manifest().to_dict()["confession"]
        assert conf["vba_present"] is True     # model never parses VBA

    def test_preservation_block_stock_vs_preserve(self, fixture_copy):
        src = fixture_copy("gauntlet/gauntlet.xlsx")
        preserve = load_workbook(src, preserve=True).manifest().to_dict()
        assert preserve["preservation"]["mode"] == "preserve"
        stock = load_workbook(src).manifest().to_dict()
        assert stock["preservation"]["mode"] == "stock"
        assert "worksheet-extension" in stock["preservation"]["at_risk"]

    def test_volatile_functions_detected_per_pinned_table(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "=NOW()"
        ws["A2"] = "=RANDBETWEEN(1,10)"
        ws["A3"] = "=OFFSET(A1,1,0)"
        ws["A4"] = "=SUM(B1:B3)"
        vol = wb.manifest().to_dict()["volatile_functions"]
        assert set(vol["nondeterministic"]) == {"NOW", "RANDBETWEEN"}
        assert set(vol["deterministic"]) == {"OFFSET"}

    def test_fresh_workbook_manifest_works(self):
        doc = Workbook().manifest().to_dict()
        assert doc["schema"] == "workbook_manifest"
        assert doc["confession"]["vba_present"] is False


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
