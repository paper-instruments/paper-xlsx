from __future__ import annotations

import io
import re
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


def test_array_formula_text_participates_in_manifest_and_dependencies():
    wb = Workbook()
    ws = wb.active
    ws["B1"] = 5
    ws["A1"] = ArrayFormula(ref="A1:A2", text="=B1*{1;2}")

    manifest = wb.manifest().to_dict()
    assert manifest["sheets"][0]["formula_count"] == 1
    from openpyxl.preserve.perception import dependency_sketch

    sketch = dependency_sketch(wb)
    assert sketch.references["'Sheet'!A1"][0][1] == (2, 1, 2, 1)


def test_manifest_cache_detection_allows_xml_whitespace():
    wb = Workbook()
    wb.active["A1"] = "=1"
    raw = io.BytesIO()
    wb.save(raw)
    crafted = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(raw.getvalue())) as zin, \
            zipfile.ZipFile(crafted, "w") as zout:
        for info in zin.infolist():
            payload = zin.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                payload, replacements = re.subn(
                    br"</f>\s*<v(?:\s*/>|></v>)",
                    b"</f>\n  <v>1</v>", payload, count=1)
                assert replacements == 1
            zout.writestr(info, payload)

    loaded = load_workbook(io.BytesIO(crafted.getvalue()), preserve=True)
    assert loaded.manifest().to_dict()["computation"]["certifiable"] is True


def test_preserve_manifest_does_not_claim_retained_content_is_dropped(
        fixture_copy):
    wb = load_workbook(fixture_copy("features/macro_stub.xlsm"),
                       preserve=True)
    doc = wb.manifest().to_dict()
    assert doc["confession"]["vba_present"] is True
    assert doc["confession"]["at_risk_content"] == []


def test_chart_auxiliary_parts_are_not_counted_as_charts(fixture_copy):
    wb = load_workbook(fixture_copy("features/chart_image.xlsx"),
                       preserve=True)
    source = wb._paper_source
    import io
    import zipfile

    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(source)) as zin, \
            zipfile.ZipFile(out, "w") as zout:
        for info in zin.infolist():
            zout.writestr(info, zin.read(info))
        zout.writestr("xl/charts/style1.xml", b"<style/>")
        zout.writestr("xl/charts/colors1.xml", b"<colors/>")
    wb2 = load_workbook(io.BytesIO(out.getvalue()), preserve=True)
    assert wb2.manifest().to_dict()["confession"]["chart_parts"] == 1


def test_package_media_count_survives_missing_model_images(fixture_copy):
    wb = load_workbook(fixture_copy("features/chart_image.xlsx"),
                       preserve=True)
    wb["Model"]._images = []
    sheet = next(item for item in wb.manifest().to_dict()["sheets"]
                 if item["title"] == "Model")
    assert sheet["images"] == 1


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
