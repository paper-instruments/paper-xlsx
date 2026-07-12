import io
import datetime
import re
import zipfile

import pytest

from openpyxl import Workbook
from openpyxl import oracle
from openpyxl.errors import UnsupportedStructureError
from openpyxl.workbook.defined_name import DefinedName


def _package_with_caches(formulas, caches):
    workbook = Workbook()
    sheet = workbook.active
    for coordinate, formula in formulas.items():
        sheet[coordinate] = formula
    target = io.BytesIO()
    workbook.save(target)

    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(target.getvalue())) as source, \
            zipfile.ZipFile(output, "w") as destination:
        for info in source.infolist():
            payload = source.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                for coordinate, (data_type, value) in caches.items():
                    pattern = re.compile(
                        rb'(<c r="' + coordinate.encode("ascii")
                        + rb'"[^>]*>\s*<f[^>]*>.*?</f>)\s*'
                        + rb'(?:<v\s*/>|<v></v>)', re.S)
                    type_attribute = (b' t="' + data_type.encode("ascii")
                                      + b'"') if data_type else b""
                    replacement = (
                        b'<c r="' + coordinate.encode("ascii")
                        + b'"' + type_attribute + b'><f>'
                        + formulas[coordinate][1:].encode("utf-8")
                        + b'</f><v>' + str(value).encode("utf-8") + b'</v>')
                    payload, count = pattern.subn(replacement, payload)
                    assert count == 1
            destination.writestr(info, payload)
    return output.getvalue()


def test_formula_result_comparison_requires_matching_excel_types():
    assert not oracle._formula_results_match("#N/A", "e", "#N/A", "s")
    assert not oracle._formula_results_match(True, "b", 1, "n")
    assert not oracle._formula_results_match("42", "s", 42, "n")


def test_numeric_comparison_accepts_rounding_noise_not_material_difference():
    adjacent = 1.0 + 2 ** -52
    assert oracle._values_match(1.0, adjacent)
    assert not oracle._values_match(1_000_000_000_000.0,
                                    1_000_000_000_100.0)


def test_error_scan_distinguishes_text_from_excel_error_type():
    workbook = Workbook()
    workbook.active["A1"] = "#N/A"
    workbook.active["A1"].data_type = "s"
    workbook.active["A2"] = "#N/A"
    workbook.active["A2"].data_type = "e"
    target = io.BytesIO()
    workbook.save(target)

    assert oracle._scan_errors(target.getvalue()) == [
        {"sheet": "Sheet", "cell": "A2", "value": "#N/A"}]


def test_matching_formula_error_never_certifies():
    package = _package_with_caches(
        {"A1": "=1/0"}, {"A1": ("e", "#DIV/0!")})

    result, _recalculated = oracle._certify_impl(
        package, 1, recalculated=package)

    assert result.status == "DIVERGED"
    assert result.divergences == [{
        "address": "Sheet!A1",
        "cached": "#DIV/0!",
        "computed": "#DIV/0!",
        "reason": "formula-error",
    }]


def test_excluded_formula_prevents_complete_certification():
    package = _package_with_caches(
        {"A1": "=RAND()", "B1": "=1+1"},
        {"A1": (None, 0.5), "B1": (None, 2)})

    result, _recalculated = oracle._certify_impl(
        package, 1, recalculated=package)

    assert result.status == "BASELINE_UNVERIFIABLE"
    assert result.checked == 1
    assert result.volatile_excluded == ["Sheet!A1"]


def test_cache_write_type_gate_rejects_mismatched_serializer_type():
    assert oracle._cache_write_preserves_type(42, "n", None)
    assert not oracle._cache_write_preserves_type("42", "n", None)
    assert not oracle._cache_write_preserves_type("#N/A", "s", None)
    assert not oracle._cache_write_preserves_type(
        datetime.date(2025, 1, 1), "d", None)


def test_defined_name_formula_exclusions_prevent_certification():
    workbook = Workbook()
    workbook.defined_names.add(DefinedName("Clock", attr_text="TODAY()"))
    workbook.active["A1"] = "=Clock"
    target = io.BytesIO()
    workbook.save(target)
    package = target.getvalue()

    result, _recalculated = oracle._certify_impl(
        package, 1, recalculated=package)

    assert result.status == "BASELINE_UNVERIFIABLE"
    assert result.volatile_excluded == ["Sheet!A1"]


def test_defined_name_formula_propagates_input_taint():
    workbook = Workbook()
    workbook.active["A1"] = 1
    workbook.defined_names.add(DefinedName(
        "Calc", attr_text="SUM(Sheet!$A$1)"))
    workbook.active["B1"] = "=Calc"
    target = io.BytesIO()
    workbook.save(target)
    package = target.getvalue()

    result, _ = oracle._certify_impl(
        package, 1, recalculated=package,
        input_seeds=[("Sheet", 1, 1)])

    assert result.status == "BASELINE_UNVERIFIABLE"
    assert result.input_excluded == ["Sheet!B1"]


def test_in_place_recalc_refuses_source_changed_during_calculation(
        tmp_path, monkeypatch):
    source = tmp_path / "source.xlsx"
    package = _package_with_caches(
        {"A1": "=1+1"}, {"A1": (None, 2)})
    source.write_bytes(package)

    def replace_source(data, _timeout):
        source.write_bytes(b"concurrent replacement")
        return data

    monkeypatch.setattr(oracle, "_recalculate_bytes", replace_source)
    with pytest.raises(UnsupportedStructureError) as refusal:
        oracle.recalc(source, in_place=True)

    assert refusal.value.kind == "destination-identity-changed"
    assert source.read_bytes() == b"concurrent replacement"
