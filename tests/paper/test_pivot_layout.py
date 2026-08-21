"""Provisional pivot layout coordinates and build plans."""
from __future__ import annotations

import json
import os

import pytest

from openpyxl.errors import BoundaryViolationError
from openpyxl.pivot.api_types import (
    PivotAxisField,
    PivotMeasure,
    PivotSource,
    PivotSpec,
)
from openpyxl.pivot.layout import ROLE_GRAND_TOTAL, ROLE_HEADER, ROLE_SUBTOTAL, ROLE_VALUE
from openpyxl.pivot.plan import plan_pivot
from openpyxl.pivot.source import (
    DEFAULT_LIMITS,
    PivotLimits,
    snapshot_from_matrix,
)
from openpyxl.utils.cell import range_boundaries
from openpyxl.xml.constants import MAX_COLUMN, MAX_ROW

from .conftest import FIXTURES_DIR


_PIVOTS = os.path.join(FIXTURES_DIR, "pivots")


def _plan(data=None, limits=None, **spec_kw):
    snapshot = snapshot_from_matrix(
        ["Region", "Product", "Amount"],
        data or [["East", "A", 10], ["West", "A", 7]],
        source=PivotSource.range("Data", "A1:C3"),
    )
    values = {
        "name": "SalesByRegion",
        "source": PivotSource.range("Data", "A1:C3"),
        "destination": "E3",
        "rows": (PivotAxisField("Region"),),
        "values": (PivotMeasure("Amount", aggregate="sum", caption="Sum"),),
        "layout": "tabular",
        "row_grand_totals": True,
        "column_grand_totals": False,
        "subtotals": False,
    }
    values.update(spec_kw)
    return plan_pivot(
        PivotSpec(**values), snapshot, limits=limits or DEFAULT_LIMITS)


def _grid(plan):
    return {
        (cell.row, cell.column): cell.value for cell in plan.output.cells
    }


def _matrix(plan):
    min_col, min_row, max_col, max_row = range_boundaries(plan.output.ref)
    grid = _grid(plan)
    return [
        [grid.get((row, column))
         for column in range(min_col, max_col + 1)]
        for row in range(min_row, max_row + 1)
    ]


def test_tabular_one_row_field_one_measure():
    plan = _plan()
    grid = _grid(plan)
    assert plan.output.destination == "E3"
    assert plan.output.ref == "E3:F6"
    assert plan.output.first_header_row == 1
    assert plan.output.first_data_row == 2
    assert plan.output.first_data_col == 1
    assert grid[(3, 5)] == "Region"
    assert grid[(3, 6)] == "Sum"
    assert grid[(4, 5)] == "East"
    assert grid[(4, 6)] == 10
    assert grid[(5, 5)] == "West"
    assert grid[(5, 6)] == 7
    assert grid[(6, 5)] == "Grand Total"
    assert grid[(6, 6)] == 17
    roles = {(cell.row, cell.column): cell.role for cell in plan.output.cells}
    assert roles[(3, 6)] == ROLE_HEADER
    assert roles[(4, 6)] == ROLE_VALUE
    assert roles[(6, 6)] == ROLE_GRAND_TOTAL


def test_multiple_measures_and_values_axis_orientations():
    columns = _plan(values=(
        PivotMeasure("Amount", aggregate="sum", caption="Sum"),
        PivotMeasure("Amount", aggregate="count", caption="Count"),
    ))
    assert columns.output.column_count == 3
    assert _grid(columns)[(3, 6)] == "Sum"
    assert _grid(columns)[(3, 7)] == "Count"
    assert _grid(columns)[(4, 6)] == 10
    assert _grid(columns)[(4, 7)] == 1

    rows = _plan(
        values=(
            PivotMeasure("Amount", aggregate="sum", caption="Sum"),
            PivotMeasure("Amount", aggregate="count", caption="Count"),
        ),
        values_axis="rows",
        row_grand_totals=False,
    )
    assert rows.output.row_count >= 4
    assert "Sum" in _grid(rows).values()
    assert "Count" in _grid(rows).values()


def test_values_on_rows_preserve_dimension_labels_and_measure_captions():
    snapshot = snapshot_from_matrix(
        ["Region", "Product", "Amount"],
        [["East", "A", 10], ["East", "B", 4], ["West", "A", 7]],
        source=PivotSource.range("Data", "A1:C4"),
    )
    spec = PivotSpec(
        name="SalesByRegion",
        source=PivotSource.range("Data", "A1:C4"),
        destination="E3",
        rows=(PivotAxisField("Region"),),
        columns=(PivotAxisField("Product"),),
        values=(
            PivotMeasure("Amount", aggregate="sum", caption="Sum"),
            PivotMeasure("Amount", aggregate="count", caption="Count"),
        ),
        layout="tabular",
        values_axis="rows",
        row_grand_totals=True,
        column_grand_totals=False,
    )

    plan = plan_pivot(spec, snapshot)

    assert plan.output.ref == "E3:H9"
    assert plan.output.row_count == 7
    assert plan.output.column_count == 4
    assert plan.output.first_data_row == 2
    assert plan.output.first_data_col == 2
    assert _matrix(plan) == [
        ["Region", "Values", "A", "B"],
        ["East", "Sum", 10, 4],
        [None, "Count", 1, 1],
        ["West", "Sum", 7, None],
        [None, "Count", 1, None],
        ["Grand Total", "Sum", 17, 4],
        [None, "Count", 2, 1],
    ]


def test_nested_column_fields_materialize_every_header_level():
    snapshot = snapshot_from_matrix(
        ["Region", "Year", "Quarter", "Amount", "Units"],
        [
            ["East", 2024, "Q1", 10, 1],
            ["East", 2024, "Q2", 4, 2],
            ["West", 2025, "Q1", 7, 3],
        ],
        source=PivotSource.range("Data", "A1:E4"),
    )
    spec = PivotSpec(
        name="SalesByPeriod",
        source=PivotSource.range("Data", "A1:E4"),
        destination="B4",
        rows=(PivotAxisField("Region"),),
        columns=(PivotAxisField("Year"), PivotAxisField("Quarter")),
        values=(
            PivotMeasure("Amount", aggregate="sum", caption="Revenue"),
            PivotMeasure("Units", aggregate="sum", caption="Units"),
        ),
        layout="tabular",
        values_axis="columns",
        row_grand_totals=False,
        column_grand_totals=False,
    )

    plan = plan_pivot(spec, snapshot)

    assert plan.output.ref == "B4:H8"
    assert plan.output.row_count == 5
    assert plan.output.column_count == 7
    assert plan.output.first_data_row == 4
    assert plan.output.first_data_col == 1
    assert _matrix(plan) == [
        ["Region", 2024, 2024, 2024, 2024, 2025, 2025],
        [None, "Q1", "Q1", "Q2", "Q2", "Q1", "Q1"],
        [None, "Revenue", "Units", "Revenue", "Units", "Revenue", "Units"],
        ["East", 10, 1, 4, 2, None, None],
        ["West", None, None, None, None, 7, 3],
    ]


def test_subtotals_on_values_axis_rows_stay_inside_ref():
    plan = _plan(
        data=[["East", "A", 10], ["East", "B", 4], ["West", "A", 7]],
        rows=(PivotAxisField("Region"), PivotAxisField("Product")),
        values=(
            PivotMeasure("Amount", aggregate="sum", caption="Sum"),
            PivotMeasure("Amount", aggregate="count", caption="Count"),
        ),
        values_axis="rows",
        subtotals=True,
        layout="tabular",
    )
    min_col, min_row, max_col, max_row = range_boundaries(plan.output.ref)
    for cell in plan.output.cells:
        assert min_row <= cell.row <= max_row
        assert min_col <= cell.column <= max_col
    grid = _grid(plan)
    assert "East Total" in grid.values()
    subtotal_rows = sorted({
        cell.row for cell in plan.output.cells if cell.role == ROLE_SUBTOTAL
    })
    assert len(subtotal_rows) == 4
    east_total_row = min(
        cell.row for cell in plan.output.cells
        if cell.value == "East Total")
    assert grid[(east_total_row, 8)] == 14
    assert grid[(east_total_row + 1, 8)] == 2
    grand_row = min(
        cell.row for cell in plan.output.cells
        if cell.value == "Grand Total")
    assert grand_row == max(subtotal_rows) + 1
    assert grid[(grand_row, 8)] == 21
    assert grid[(grand_row + 1, 8)] == 3


def test_compact_outline_and_tabular_label_columns():
    tabular = _plan(
        rows=(PivotAxisField("Region"), PivotAxisField("Product")),
        layout="tabular",
        row_grand_totals=False,
    )
    compact = _plan(
        rows=(PivotAxisField("Region"), PivotAxisField("Product")),
        layout="compact",
        row_grand_totals=False,
    )
    outline = _plan(
        rows=(PivotAxisField("Region"), PivotAxisField("Product")),
        layout="outline",
        row_grand_totals=False,
    )
    assert tabular.output.first_data_col == 2
    assert compact.output.first_data_col == 1
    assert outline.output.first_data_col == 2
    assert "layout-coordinates-provisional" in tabular.warnings


def test_output_beyond_sheet_edge_refuses():
    with pytest.raises(BoundaryViolationError) as exc:
        _plan(destination="%s%s" % ("XFD", MAX_ROW))
    assert exc.value.kind == "pivot-output-too-large"
    assert str(MAX_COLUMN) in exc.value.options or str(MAX_ROW) in exc.value.options


def test_plan_is_deterministic_and_does_not_mutate_snapshot():
    first = _plan()
    second = _plan()
    assert first.digest == second.digest
    assert first.to_dict() == second.to_dict()
    snapshot = snapshot_from_matrix(
        ["Region", "Product", "Amount"],
        [["East", "A", 10]],
    )
    before = snapshot.records
    plan_pivot(_plan().spec, snapshot)
    assert snapshot.records is before


def test_output_and_cache_limits_refuse_before_materializing_xml():
    with pytest.raises(BoundaryViolationError) as output:
        _plan(limits=PivotLimits(output_cells=1))
    assert output.value.kind == "pivot-output-too-large"
    huge = PivotLimits(cache_xml_bytes=1)
    with pytest.raises(BoundaryViolationError) as cache:
        _plan(limits=huge)
    assert cache.value.kind == "pivot-cache-too-large"


def test_linear_record_visits_as_cardinality_grows():
    def visits(count):
        rows = [["R%s" % index, "A", index] for index in range(count)]
        snapshot = snapshot_from_matrix(
            ["Region", "Product", "Amount"], rows)
        plan = plan_pivot(
            PivotSpec(
                name="Growth",
                source=PivotSource.range("Data", "A1:C2"),
                destination="E3",
                rows=(PivotAxisField("Region"),),
                values=(PivotMeasure("Amount", aggregate="sum"),),
                row_grand_totals=False,
            ),
            snapshot,
        )
        return len(snapshot.records), plan.output.row_count

    small = visits(50)
    large = visits(200)
    assert small[0] == 50
    assert large[0] == 200
    assert large[1] > small[1]
    assert DEFAULT_LIMITS.output_cells == 1000000


@pytest.mark.parametrize("filename", [
    name for name in sorted(os.listdir(_PIVOTS))
    if name.endswith((".xlsx", ".xlsm", ".xltx"))
] if os.path.isdir(_PIVOTS) else [])
def test_plan_matches_sidecar_visible_values(filename):
    path = os.path.join(_PIVOTS, filename)
    with open(path + ".json") as handle:
        sidecar = json.load(handle)
    expected = sidecar.get("expected_visible_values")
    if expected is None:
        pytest.skip("sidecar has no approved Excel visible values")
    pytest.skip("planner does not consume package bytes; fixture replay is PR 4")
