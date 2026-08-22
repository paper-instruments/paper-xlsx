"""Deterministic pivot aggregation."""
from __future__ import annotations

from datetime import date

import pytest

from openpyxl.errors import BoundaryViolationError
from openpyxl.pivot.aggregate import aggregate_snapshot, display_item
from openpyxl.pivot.api_types import (
    PivotAxisField,
    PivotItemFilter,
    PivotMeasure,
    PivotSource,
    PivotSpec,
)
from openpyxl.pivot.source import PivotLimits, snapshot_from_matrix, typed_value


def _spec(**overrides):
    values = {
        "name": "SalesByRegion",
        "source": PivotSource.range("Data", "A1:C4"),
        "destination": "E3",
        "rows": (PivotAxisField("Region"),),
        "values": (PivotMeasure("Amount", aggregate="sum"),),
    }
    values.update(overrides)
    return PivotSpec(**values)


def _snapshot(rows=None):
    return snapshot_from_matrix(
        ["Region", "Product", "Amount"],
        rows or [
            ["East", "A", 10],
            ["East", "B", 4],
            ["West", "A", 7],
            ["West", "A", -1],
        ],
    )


def test_sum_count_average_min_max_and_count_numbers():
    snapshot = snapshot_from_matrix(
        ["Region", "Amount", "Label"],
        [
            ["East", 10, "x"],
            ["East", None, "y"],
            ["East", 5, "z"],
        ],
    )
    result = aggregate_snapshot(snapshot, _spec(
        source=PivotSource.range("Data", "A1:C4"),
        values=(
            PivotMeasure("Amount", aggregate="sum", caption="Sum"),
            PivotMeasure("Amount", aggregate="count", caption="Count"),
            PivotMeasure("Amount", aggregate="count_numbers", caption="N"),
            PivotMeasure("Amount", aggregate="average", caption="Avg"),
            PivotMeasure("Amount", aggregate="min", caption="Min"),
            PivotMeasure("Amount", aggregate="max", caption="Max"),
        ),
    ))
    east = result.cells[(result.row_keys[0], ())]
    assert [item.value for item in east] == [15, 2, 2, 7.5, 5, 10]
    assert result.included_row_count == 3


def test_count_versus_count_numbers_over_text_bool_and_blank():
    snapshot = snapshot_from_matrix(
        ["Region", "Flag"],
        [["East", "x"], ["East", True], ["East", 3], ["East", None]],
    )
    result = aggregate_snapshot(snapshot, _spec(
        source=PivotSource.range("Data", "A1:B5"),
        values=(
            PivotMeasure("Flag", aggregate="count", caption="Count"),
            PivotMeasure("Flag", aggregate="count_numbers", caption="Nums"),
        ),
    ))
    values = result.cells[(result.row_keys[0], ())]
    assert values[0].value == 3
    assert values[1].value == 1


def test_sum_refuses_text_or_boolean_measures():
    snapshot = snapshot_from_matrix(
        ["Region", "Flag"],
        [["East", True]],
    )
    with pytest.raises(BoundaryViolationError) as exc:
        aggregate_snapshot(snapshot, _spec(
            source=PivotSource.range("Data", "A1:B2"),
            values=(PivotMeasure("Flag", aggregate="sum"),),
        ))
    assert exc.value.kind == "invalid-pivot-source"


@pytest.mark.parametrize("aggregate", ["min", "max"])
@pytest.mark.parametrize("values", [
    [date(2024, 1, 1), 1],
    [1, date(2024, 1, 1)],
])
def test_min_max_refuse_mixed_dates_and_numbers(aggregate, values):
    snapshot = snapshot_from_matrix(
        ["Region", "Amount"],
        [["East", value] for value in values],
    )
    with pytest.raises(BoundaryViolationError) as exc:
        aggregate_snapshot(snapshot, _spec(
            source=PivotSource.range("Data", "A1:B3"),
            values=(PivotMeasure("Amount", aggregate=aggregate),),
        ))
    assert exc.value.kind == "invalid-pivot-source"


def test_filters_and_repeated_keys_preserve_first_seen_order():
    result = aggregate_snapshot(_snapshot(), _spec(
        filters=(PivotItemFilter("Product", include=["A"]),),
    ))
    assert [key[0].value for key in result.row_keys] == ["East", "West"]
    east = result.cells[(result.row_keys[0], ())][0].value
    west = result.cells[(result.row_keys[1], ())][0].value
    assert east == 10
    assert west == 6


def test_empty_post_filter_still_returns_metadata():
    result = aggregate_snapshot(_snapshot(), _spec(
        filters=(PivotItemFilter("Product", include=["North"]),),
    ))
    assert result.included_row_count == 0
    assert result.row_keys == ()
    assert result.grand_total[0].value is None


def test_two_measures_over_the_same_field_and_grand_totals():
    result = aggregate_snapshot(_snapshot(), _spec(
        values=(
            PivotMeasure("Amount", aggregate="sum", caption="Sum"),
            PivotMeasure("Amount", aggregate="count", caption="Count"),
        ),
        row_grand_totals=True,
    ))
    assert result.grand_total[0].value == 20
    assert result.grand_total[1].value == 4


def test_subtotals_for_nested_rows():
    result = aggregate_snapshot(_snapshot(), _spec(
        rows=(PivotAxisField("Region"), PivotAxisField("Product")),
        subtotals=True,
    ))
    east = (result.row_keys[0][0],)
    assert east in result.row_subtotals
    assert result.row_subtotals[east][0].value == 14


def test_cardinality_limit_refuses_before_output_allocation():
    snapshot = _snapshot()
    with pytest.raises(BoundaryViolationError) as exc:
        aggregate_snapshot(
            snapshot, _spec(), limits=PivotLimits(aggregate_states=0))
    assert exc.value.kind == "pivot-cardinality-too-large"


def test_dimension_captions_match_excel_for_boolean_and_blank_items():
    assert display_item(typed_value(True)) == "TRUE"
    assert display_item(typed_value(False)) == "FALSE"
    assert display_item(typed_value(None)) == "(blank)"
