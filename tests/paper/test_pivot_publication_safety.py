"""Adversarial publication and in-session composition regressions."""
from __future__ import annotations

import os
import zipfile
from copy import copy
from datetime import date

import pytest

from openpyxl import load_workbook
from openpyxl.errors import BoundaryViolationError
from openpyxl.errors import UnsupportedStructureError
from openpyxl.pivot.api_types import (
    PivotAxisField,
    PivotItemFilter,
    PivotMeasure,
)
from openpyxl.worksheet.table import Table

from .support.harness import save_and_reopen
from .support.partdiff import part_payloads
from .test_pivot_create_package import _create_by_region
from .test_pivot_refresh import _preserved_matrix


_TABLE = "features/tables.xlsx"
_CACHE = "xl/pivotCache/pivotCacheDefinition1.xml"
_RECORDS = "xl/pivotCache/pivotCacheRecords1.xml"
_PIVOT = "xl/pivotTables/pivotTable1.xml"


@pytest.mark.parametrize("mutation", ["value", "style"])
def test_direct_staged_output_edit_refuses_save(fixture_copy, tmp_path, mutation):
    source = fixture_copy(_TABLE)
    workbook = load_workbook(source, preserve=True)
    _create_by_region(workbook["Data"])
    output = workbook["Data"]["F4"]
    if mutation == "value":
        output.value = 999
    else:
        output.number_format = "$#,##0"

    destination = str(tmp_path / (mutation + ".xlsx"))
    with pytest.raises(BoundaryViolationError) as exc:
        workbook.save(destination)
    assert exc.value.kind == "pivot-output-collision"
    assert not os.path.exists(destination)


def test_direct_reopened_output_edit_refuses_save(fixture_copy, tmp_path):
    source = fixture_copy(_TABLE)
    workbook = load_workbook(source, preserve=True)
    _create_by_region(workbook["Data"])
    created = str(tmp_path / "created.xlsx")
    workbook = save_and_reopen(workbook, created, preserve=True)
    workbook["Data"]["F4"] = True

    destination = str(tmp_path / "edited.xlsx")
    with pytest.raises(BoundaryViolationError) as exc:
        workbook.save(destination)
    assert exc.value.kind == "pivot-output-collision"
    assert not os.path.exists(destination)
    assert workbook["Data"].pivots["ByRegion"].capabilities.can_delete is False


def test_unsaved_create_delete_refuses_to_erase_later_blank_slot(tmp_path):
    _source, workbook = _preserved_matrix(
        tmp_path,
        ["Region", "Amount"],
        [["East", 10], ["West", 7]],
        table="Sales",
    )
    worksheet = workbook["Summary"]
    pivot = worksheet.pivots.create(
        name="ByRegion",
        source="Sales",
        destination="B4",
        rows=[],
        columns=["Region"],
        values=["Amount"],
    )
    assert worksheet["B4"].value is None
    worksheet["B4"] = "keep me"

    with pytest.raises(BoundaryViolationError) as exc:
        pivot.delete()
    assert exc.value.kind == "pivot-output-collision"
    assert worksheet["B4"].value == "keep me"


def test_create_then_delete_restores_custom_format_registry(tmp_path):
    source, workbook = _preserved_matrix(
        tmp_path,
        ["Region", "Amount"],
        [["East", 10], ["West", 7]],
        table="Sales",
    )
    pivot = workbook["Summary"].pivots.create(
        name="ByRegion", source="Sales", destination="A1",
        rows=["Region"],
        values=[PivotMeasure("Amount", number_format="USD #,##0")],
    )
    assert "USD #,##0" in workbook._number_formats
    pivot.delete()
    assert "USD #,##0" not in workbook._number_formats

    destination = str(tmp_path / "cancelled-format.xlsx")
    workbook.save(destination)
    assert part_payloads(destination) == part_payloads(source)


def test_refresh_then_rename_keeps_cache_and_output_in_sync(fixture_copy, tmp_path):
    source = fixture_copy(_TABLE)
    workbook = load_workbook(source, preserve=True)
    _create_by_region(workbook["Data"])
    created = str(tmp_path / "created.xlsx")
    workbook = save_and_reopen(workbook, created, preserve=True)
    before = part_payloads(created)

    workbook["Data"]["B2"] = 99
    renamed = workbook["Data"].pivots["ByRegion"].refresh().rename(
        "RegionalSales")
    receipt = workbook.save(str(tmp_path / "renamed.xlsx"), receipt=True)
    after = part_payloads(str(tmp_path / "renamed.xlsx"))

    assert renamed.name == "RegionalSales"
    assert after[_CACHE] != before[_CACHE]
    assert b' name="RegionalSales"' in after[_PIVOT]
    effects = [item for item in receipt.derived_effects
               if item["kind"] == "pivot_renamed"]
    assert effects == [{
        "kind": "pivot_renamed",
        "name": "RegionalSales",
        "sheet": "Data",
        "cache_rebuilt": True,
    }]
    reopened = load_workbook(str(tmp_path / "renamed.xlsx"), preserve=True)
    assert reopened["Data"]["F4"].value == 99
    assert reopened["Data"].pivots["RegionalSales"].capabilities.can_refresh_on_open


def test_repoint_move_rename_composition_keeps_one_coherent_graph(tmp_path):
    _source, workbook = _preserved_matrix(
        tmp_path,
        ["Region", "Amount"],
        [["East", 10], ["West", 7]],
        table="Sales",
    )
    data = workbook["Data"]
    data["D1"] = "Region"
    data["E1"] = "Amount"
    data["D2"] = "North"
    data["E2"] = 41
    data.add_table(Table(displayName="Other", ref="D1:E2"))
    pivot = workbook["Summary"].pivots.create(
        name="ByRegion", source="Sales", destination="A1",
        rows=["Region"], values=["Amount"])
    pivot = pivot.repoint_source("Other")
    pivot = pivot.move("D4")
    pivot = pivot.rename("NorthOnly")

    destination = str(tmp_path / "composed.xlsx")
    workbook = save_and_reopen(workbook, destination, preserve=True)
    reopened = workbook["Summary"].pivots["NorthOnly"]
    assert reopened.source.name == "Other"
    assert reopened.destination == "D4"
    assert workbook["Summary"]["D5"].value == "North"
    assert workbook["Summary"]["E5"].value == 41
    assert reopened.capabilities.can_headless_refresh is True


def test_rename_is_a_lexical_definition_patch(fixture_copy, tmp_path):
    source = fixture_copy(_TABLE)
    workbook = load_workbook(source, preserve=True)
    _create_by_region(workbook["Data"])
    created = str(tmp_path / "created.xlsx")
    workbook.save(created)
    before = part_payloads(created)

    workbook = load_workbook(created, preserve=True)
    workbook["Data"].pivots["ByRegion"].rename("RegionalSales")
    renamed = str(tmp_path / "renamed.xlsx")
    workbook.save(renamed)
    after = part_payloads(renamed)

    expected = before[_PIVOT].replace(
        b'name="ByRegion"', b'name="RegionalSales"', 1)
    assert after[_PIVOT] == expected
    assert after[_RECORDS] == before[_RECORDS]


def test_refresh_rebuilds_cache_when_aggregate_is_unchanged(tmp_path):
    _source, workbook = _preserved_matrix(
        tmp_path,
        ["Region", "Amount"],
        [["East", 20], ["East", 30]],
        table="Sales",
    )
    workbook["Summary"].pivots.create(
        name="ByRegion", source="Sales", destination="A1",
        rows=["Region"], values=[PivotMeasure("Amount", "sum")])
    created = str(tmp_path / "created.xlsx")
    workbook = save_and_reopen(workbook, created, preserve=True)
    before = part_payloads(created)[_CACHE]

    workbook["Data"]["B2"] = 21
    workbook["Data"]["B3"] = 29
    workbook["Summary"].pivots["ByRegion"].refresh()
    refreshed = str(tmp_path / "refreshed.xlsx")
    workbook.save(refreshed)

    assert part_payloads(refreshed)[_CACHE] != before
    assert load_workbook(refreshed, preserve=True)["Summary"]["B2"].value == 50


def test_custom_measure_format_survives_reopen_and_refresh(tmp_path):
    _source, workbook = _preserved_matrix(
        tmp_path,
        ["Region", "Amount"],
        [["East", 20], ["West", 30]],
        table="Sales",
    )
    workbook["Summary"].pivots.create(
        name="ByRegion", source="Sales", destination="A1",
        rows=["Region"],
        values=[PivotMeasure("Amount", "sum", number_format="$#,##0")],
    )
    created = str(tmp_path / "formatted.xlsx")
    workbook = save_and_reopen(workbook, created, preserve=True)
    pivot = workbook["Summary"].pivots["ByRegion"]
    assert pivot.spec.values[0].number_format == "$#,##0"

    workbook["Data"]["B2"] = 25
    pivot.refresh()
    refreshed = str(tmp_path / "formatted-refreshed.xlsx")
    workbook = save_and_reopen(workbook, refreshed, preserve=True)
    assert workbook["Summary"].pivots["ByRegion"].spec.values[0].number_format \
        == "$#,##0"
    assert b'numFmtId="164"' in part_payloads(refreshed)[_PIVOT]


def test_unmodeled_standard_semantics_disable_reserializing_mutators(
        fixture_copy, tmp_path):
    source = fixture_copy(_TABLE)
    workbook = load_workbook(source, preserve=True)
    _create_by_region(workbook["Data"])
    created = str(tmp_path / "created.xlsx")
    workbook.save(created)
    tampered = str(tmp_path / "tampered.xlsx")
    _rewrite_part(
        created,
        tampered,
        _PIVOT,
        lambda payload: payload.replace(
            b'showRowStripes="0"', b'showRowStripes="1"', 1),
    )

    workbook = load_workbook(tampered, preserve=True)
    pivot = workbook["Data"].pivots["ByRegion"]
    assert pivot.capabilities.can_edit_layout is False
    assert pivot.capabilities.can_headless_refresh is False
    assert pivot.capabilities.can_rename is True
    with pytest.raises(UnsupportedStructureError):
        pivot.update(layout="outline")
    renamed = pivot.rename("RegionalSales")
    destination = str(tmp_path / "renamed-tampered.xlsx")
    workbook.save(destination)
    assert renamed.name == "RegionalSales"
    assert b'showRowStripes="1"' in part_payloads(destination)[_PIVOT]


def test_show_data_as_is_not_projected_as_an_ordinary_sum(fixture_copy, tmp_path):
    source = fixture_copy(_TABLE)
    workbook = load_workbook(source, preserve=True)
    _create_by_region(workbook["Data"])
    created = str(tmp_path / "created.xlsx")
    workbook.save(created)
    tampered = str(tmp_path / "show-data-as.xlsx")
    _rewrite_part(
        created,
        tampered,
        _PIVOT,
        lambda payload: payload.replace(
            b'showDataAs="normal"', b'showDataAs="percentOfTotal"', 1),
    )

    workbook = load_workbook(tampered, preserve=True)
    pivot = workbook["Data"].pivots["ByRegion"]
    assert pivot.spec is None
    assert pivot.capabilities.can_edit_layout is False
    assert pivot.capabilities.can_headless_refresh is False


def test_date_and_manual_item_order_reopen_with_lifecycle_capabilities(tmp_path):
    _source, workbook = _preserved_matrix(
        tmp_path,
        ["AsOf", "Region", "Amount"],
        [
            [date(2026, 1, 1), "East", 20],
            [date(2026, 2, 1), "West", 30],
        ],
        table="Sales",
    )
    workbook["Summary"].pivots.create(
        name="ByDate", source="Sales", destination="A1",
        rows=["AsOf"], values=["Amount"])
    workbook["Summary"].pivots.create(
        name="ByRegion", source="Sales", destination="E1",
        rows=[PivotAxisField("Region", items=["West", "East"])],
        values=["Amount"])
    created = str(tmp_path / "ordered.xlsx")
    workbook = save_and_reopen(workbook, created, preserve=True)

    by_date = workbook["Summary"].pivots["ByDate"]
    by_region = workbook["Summary"].pivots["ByRegion"]
    assert by_date.capabilities.can_headless_refresh is True
    assert by_date.capabilities.can_delete is True
    assert by_region.capabilities.can_headless_refresh is True
    assert by_region.capabilities.can_delete is True
    assert workbook["Summary"]["E2"].value == "West"
    assert workbook["Summary"]["E3"].value == "East"


def test_direct_edit_to_reopened_report_filter_cell_refuses_save(tmp_path):
    _source, workbook = _preserved_matrix(
        tmp_path,
        ["Region", "Status", "Amount"],
        [["East", "Closed", 20], ["West", "Open", 30]],
        table="Sales",
    )
    workbook["Summary"].pivots.create(
        name="ClosedOnly", source="Sales", destination="A3",
        rows=["Region"],
        filters=[PivotItemFilter("Status", include=["Closed"])],
        values=["Amount"],
    )
    created = str(tmp_path / "filtered.xlsx")
    workbook = save_and_reopen(workbook, created, preserve=True)
    workbook["Summary"]["B1"] = "tampered"

    destination = str(tmp_path / "filtered-tampered.xlsx")
    with pytest.raises(BoundaryViolationError) as exc:
        workbook.save(destination)
    assert exc.value.kind == "pivot-output-collision"
    assert not os.path.exists(destination)


def _rewrite_part(source, destination, part, transform):
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(destination, "w") as after:
        for info in before.infolist():
            payload = before.read(info.filename)
            if info.filename == part:
                payload = transform(payload)
            after.writestr(copy(info), payload)
