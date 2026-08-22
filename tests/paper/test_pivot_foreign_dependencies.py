"""PR 9: selected-pivot dependency and consumer constraints."""
from __future__ import annotations

import io
import zipfile

from openpyxl import load_workbook

from .test_pivot_adoption_qualification import _codes, _load_payload, _rewrite_bytes
from .test_pivot_graph import _basic_package


def test_getpivotdata_is_an_operation_constraint(tmp_path):
    payload = _basic_package()
    payload = _rewrite_bytes(
        payload, "xl/worksheets/sheet1.xml",
        lambda body: body.replace(
            b"</sheetData>",
            b'<row r="20"><c r="A20" t="str"><f>'
            b'GETPIVOTDATA("Amount",$B$3)</f><v></v></c></row></sheetData>',
            1,
        ),
    )
    wb = _load_payload(tmp_path, payload, name="gpd.xlsx")
    wb["Summary"]["A20"] = '=GETPIVOTDATA("Amount",$B$3)'
    result = wb["Summary"].pivots["SalesByRegion"].qualify_adoption()
    constraint_codes = [item.code for item in result.operation_constraints]
    assert "pivot-dependent-reference" in constraint_codes
    assert result.eligible is False


def test_pivotchart_lexical_reference_is_detected(tmp_path):
    payload = _basic_package()
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(payload)) as before, \
            zipfile.ZipFile(output, "w") as after:
        for info in before.infolist():
            after.writestr(info, before.read(info.filename))
        after.writestr(
            "xl/charts/chart1.xml",
            '<?xml version="1.0"?>'
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/'
            'drawingml/2006/chart">'
            '<c:pivotSource><c:name>SalesByRegion</c:name></c:pivotSource>'
            '</c:chartSpace>',
        )
    wb = _load_payload(tmp_path, output.getvalue(), name="chart.xlsx")
    result = wb["Summary"].pivots["SalesByRegion"].qualify_adoption()
    assert "foreign-dependent-object" in _codes(result)


def test_slicer_cache_reference_is_detected(tmp_path):
    payload = _basic_package()
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(payload)) as before, \
            zipfile.ZipFile(output, "w") as after:
        for info in before.infolist():
            after.writestr(info, before.read(info.filename))
        after.writestr(
            "xl/slicerCaches/slicerCache1.xml",
            '<?xml version="1.0"?>'
            '<slicerCacheDefinition name="Slicer_Region" '
            'sourceName="SalesByRegion">'
            '<slicerCache pivotCacheId="1"/>'
            '</slicerCacheDefinition>',
        )
    wb = _load_payload(tmp_path, output.getvalue(), name="slicer.xlsx")
    result = wb["Summary"].pivots["SalesByRegion"].qualify_adoption()
    assert "foreign-dependent-object" in _codes(result)


def test_duplicate_internal_external_rid_cannot_qualify(tmp_path):
    payload = _basic_package()

    def add_duplicate(body):
        closing = b"</Relationships>"
        duplicate = (
            b'<Relationship Id="rIdPivot1" Type="urn:external" '
            b'Target="https://example.invalid/pivot" '
            b'TargetMode="External"/>'
        )
        return body.replace(closing, duplicate + closing, 1)

    payload = _rewrite_bytes(
        payload, "xl/worksheets/_rels/sheet1.xml.rels", add_duplicate)
    wb = _load_payload(tmp_path, payload, name="dup.xlsx")
    result = wb["Summary"].pivots["SalesByRegion"].qualify_adoption()
    assert result.eligible is False
    assert "duplicate-relationship-id" in _codes(result) \
        or "invalid-pivot-graph" in _codes(result)
