"""PR 8: scale and determinism for Paper-managed pivots."""
from __future__ import annotations

import os

import pytest

from openpyxl import Workbook, load_workbook
from openpyxl.errors import BoundaryViolationError
from openpyxl.pivot.source import DEFAULT_LIMITS
from openpyxl.worksheet.table import Table

from .support.harness import save_and_reopen


def _rows(count):
    regions = ("East", "West", "North", "South")
    products = ("A", "B", "C")
    rows = []
    for index in range(count):
        rows.append((
            regions[index % 4],
            products[index % 3],
            float(index % 50),
        ))
    return rows


def test_moderate_source_is_deterministic(tmp_path):
    count = 2000 if os.environ.get("PAPER_PIVOT_SCALE") != "1" else 100000
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(("Region", "Product", "Amount"))
    for row in _rows(count):
        ws.append(row)
    ws.add_table(Table(displayName="Sales", ref="A1:C%s" % (count + 1)))
    summary = wb.create_sheet("Summary")
    first = str(tmp_path / "a.xlsx")
    wb.save(first)
    wb = load_workbook(first, preserve=True)
    wb["Summary"].pivots.create(
        name="ByRegion", source="Sales", destination="A1",
        rows=["Region"], columns=["Product"], values=["Amount"])
    wb = save_and_reopen(wb, str(tmp_path / "b.xlsx"), preserve=True)
    again = load_workbook(first, preserve=True)
    again["Summary"].pivots.create(
        name="ByRegion", source="Sales", destination="A1",
        rows=["Region"], columns=["Product"], values=["Amount"])
    again = save_and_reopen(again, str(tmp_path / "c.xlsx"), preserve=True)
    from .support.partdiff import part_payloads
    left = part_payloads(str(tmp_path / "b.xlsx"))
    right = part_payloads(str(tmp_path / "c.xlsx"))
    for name in (
        "xl/pivotCache/pivotCacheRecords1.xml",
        "xl/pivotTables/pivotTable1.xml",
    ):
        assert left[name] == right[name]


def test_output_cardinality_refuses_before_cells(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(("Region", "Amount"))
    ws.append(("East", 1))
    ws.append(("West", 2))
    ws.add_table(Table(displayName="Sales", ref="A1:B3"))
    summary = wb.create_sheet("Summary")
    path = str(tmp_path / "src.xlsx")
    wb.save(path)
    wb = load_workbook(path, preserve=True)
    from openpyxl.pivot.plan import plan_pivot
    from openpyxl.pivot.source import snapshot_from_workbook, PivotLimits
    from openpyxl.pivot.api_types import PivotMeasure, PivotSource, PivotSpec, PivotAxisField
    snapshot = snapshot_from_workbook(wb, "Sales")
    spec = PivotSpec(
        name="ByRegion",
        source=PivotSource.table("Sales"),
        destination="A1",
        rows=(PivotAxisField("Region"),),
        values=(PivotMeasure("Amount"),),
    )
    with pytest.raises(BoundaryViolationError) as exc:
        plan_pivot(spec, snapshot, limits=PivotLimits(output_cells=1))
    assert exc.value.kind == "pivot-output-too-large"
    assert wb["Summary"]["A1"].value is None
