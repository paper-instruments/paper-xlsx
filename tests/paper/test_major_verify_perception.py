from __future__ import annotations

import time
import zipfile

from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula


def test_local_defined_name_wins_and_sheet_lookup_is_case_insensitive():
    wb = Workbook()
    ws = wb.active
    ws.title = "Inputs"
    ws["A1"] = 10
    ws["B1"] = 20
    ws["C1"] = "=rate+'inputs'!A1"
    wb.defined_names["Rate"] = DefinedName("Rate", attr_text="Inputs!$A$1")
    ws.defined_names["RATE"] = DefinedName("RATE", attr_text="Inputs!$B$1")

    model = wb.model_map()
    assert "B1" in model.sheets["Inputs"]["inputs"]
    assert "A1" in model.sheets["Inputs"]["inputs"]


def test_huge_sparse_range_intersects_only_populated_cells():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = 1
    ws["XFD1048576"] = 2
    ws["B1"] = "=SUM(A1:XFD1048576)"
    started = time.monotonic()
    model = wb.model_map()
    assert time.monotonic() - started < 2
    assert set(model.sheets["Sheet"]["inputs"]) == {"A1", "XFD1048576"}


def test_array_formula_text_participates_in_dependencies():
    wb = Workbook()
    ws = wb.active
    ws["B1"] = 5
    ws["A1"] = ArrayFormula(ref="A1:A2", text="=B1*{1;2}")

    from openpyxl.preserve.perception import dependency_sketch

    sketch = dependency_sketch(wb)
    assert sketch.references["'Sheet'!A1"][0][1] == (2, 1, 2, 1)


def test_preserve_save_retains_macro_package_content_type(fixture_copy,
                                                          tmp_path):
    source = fixture_copy("features/macro_stub.xlsm")
    wb = load_workbook(source, preserve=True)
    from openpyxl.xml.constants import XLSM

    assert wb.mime_type == XLSM
    wb.active["A1"] = "edited"
    output = tmp_path / "retained.xlsm"
    wb.save(output)
    with zipfile.ZipFile(output) as archive:
        content_types = archive.read("[Content_Types].xml")
        assert b"macroEnabled.main+xml" in content_types
        assert "xl/vbaProject.bin" in archive.namelist()
