# paper-xlsx: bounded deterministic pivot aggregation

"""Aggregate a typed source snapshot against a resolved PivotSpec.

The engine processes records in source order, uses first-seen item order as
the provisional default, and never imports pandas. Groups with no accepted
measure values materialize as blank rather than zero. Pending Excel
transcripts may replace blank/empty-string/count-coercion rules later.
"""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl.errors import BoundaryViolationError
from openpyxl.pivot.api_types import SUPPORTED_AGGREGATES
from openpyxl.pivot.source import (
    KIND_BLANK,
    KIND_BOOLEAN,
    KIND_DATE,
    KIND_DATETIME,
    KIND_NUMBER,
    KIND_TEXT,
    TypedValue,
    typed_value,
)


@dataclass(frozen=True)
class MeasureValue:
    field: str
    aggregate: str
    caption: str
    number_format: str | None
    value: object
    input_count: int


@dataclass(frozen=True)
class AggregateResult:
    row_keys: tuple
    column_keys: tuple
    cells: object
    row_subtotals: object
    column_totals: object
    row_totals: object
    grand_total: tuple
    source_row_count: int
    included_row_count: int
    measure_input_counts: tuple
    warnings: tuple = ()


def aggregate_snapshot(snapshot, spec, limits=None):
    """Filter, group, and aggregate ``snapshot`` according to ``spec``."""
    from openpyxl.pivot.source import DEFAULT_LIMITS

    limits = limits or DEFAULT_LIMITS
    field_index = snapshot.field_index
    _validate_spec_fields(spec, field_index)
    included = _filter_records(snapshot.records, spec.filters, field_index)
    row_fields = [item.field for item in spec.rows]
    column_fields = [item.field for item in spec.columns]
    measures = spec.values

    row_keys, row_order = _ordered_keys(
        included, row_fields, field_index, spec.rows)
    column_keys, column_order = _ordered_keys(
        included, column_fields, field_index, spec.columns)
    if not row_fields:
        row_keys = ((),)
        row_order = [()]
    if not column_fields:
        column_keys = ((),)
        column_order = [()]

    states = len(row_order) * len(column_order) * len(measures)
    if states > limits.aggregate_states:
        raise BoundaryViolationError(
            "pivot would materialize %s aggregate states; the limit is %s"
            % (states, limits.aggregate_states),
            kind="pivot-cardinality-too-large",
            options=[str(states), str(limits.aggregate_states)],
        )

    buckets = {}
    for record in included:
        row_key = _key(record, row_fields, field_index)
        col_key = _key(record, column_fields, field_index) if column_fields \
            else ()
        for measure in measures:
            index = field_index[measure.field]
            accepted = _accept_measure(record.values[index], measure.aggregate)
            bucket = buckets.setdefault(
                (row_key, col_key, measure.field, measure.aggregate), [])
            if accepted is not None:
                bucket.append(accepted)

    cells = {}
    input_counts = []
    for row_key in row_order:
        for col_key in column_order:
            values = []
            for measure in measures:
                raw = buckets.get(
                    (row_key, col_key, measure.field, measure.aggregate), [])
                value = _reduce(raw, measure.aggregate)
                values.append(MeasureValue(
                    measure.field, measure.aggregate,
                    measure.caption or _default_caption(measure),
                    measure.number_format, value, len(raw),
                ))
                input_counts.append(len(raw))
            cells[(row_key, col_key)] = tuple(values)

    row_totals = {}
    if spec.column_grand_totals and column_fields:
        for row_key in row_order:
            row_totals[row_key] = _totals_for(
                included, measures, field_index,
                lambda record, row_key=row_key: _key(
                    record, row_fields, field_index) == row_key)

    column_totals = {}
    if spec.row_grand_totals and column_fields:
        for col_key in column_order:
            column_totals[col_key] = _totals_for(
                included, measures, field_index,
                lambda record, col_key=col_key: (
                    not column_fields
                    or _key(record, column_fields, field_index) == col_key))

    grand = _totals_for(included, measures, field_index, lambda record: True)

    subtotals = {}
    if spec.subtotals and len(row_fields) > 1:
        prefixes = []
        seen = set()
        for row_key in row_order:
            for depth in range(1, len(row_fields)):
                prefix = row_key[:depth]
                if prefix in seen:
                    continue
                seen.add(prefix)
                prefixes.append(prefix)
                if column_fields:
                    by_column = {}
                    for col_key in column_order:
                        by_column[col_key] = _totals_for(
                            included, measures, field_index,
                            lambda record, prefix=prefix, col_key=col_key: (
                                _key(record, row_fields, field_index)[:len(prefix)]
                                == prefix
                                and _key(record, column_fields, field_index)
                                == col_key))
                    by_column[None] = _totals_for(
                        included, measures, field_index,
                        lambda record, prefix=prefix: _key(
                            record, row_fields, field_index)[:len(prefix)] == prefix)
                    subtotals[prefix] = by_column
                else:
                    subtotals[prefix] = _totals_for(
                        included, measures, field_index,
                        lambda record, prefix=prefix: _key(
                            record, row_fields, field_index)[:len(prefix)] == prefix)

    return AggregateResult(
        row_keys=tuple(row_order),
        column_keys=tuple(column_order),
        cells=cells,
        row_subtotals=subtotals,
        column_totals=column_totals,
        row_totals=row_totals,
        grand_total=grand,
        source_row_count=len(snapshot.records),
        included_row_count=len(included),
        measure_input_counts=tuple(input_counts),
    )


def _validate_spec_fields(spec, field_index):
    used = []
    for item in spec.rows + spec.columns + spec.filters:
        if item.field not in field_index:
            raise BoundaryViolationError(
                "pivot field %r is not in the source"
                % item.field,
                kind="invalid-pivot-source",
                options=sorted(field_index),
            )
        if item.field in used:
            raise BoundaryViolationError(
                "field %r cannot appear on more than one axis"
                % item.field,
                kind="invalid-pivot-source",
                options=[item.field],
            )
        used.append(item.field)
    for measure in spec.values:
        if measure.field not in field_index:
            raise BoundaryViolationError(
                "measure field %r is not in the source"
                % measure.field,
                kind="invalid-pivot-source",
                options=sorted(field_index),
            )
        if measure.aggregate not in SUPPORTED_AGGREGATES:
            raise BoundaryViolationError(
                "unsupported aggregate %r" % measure.aggregate,
                kind="unsupported-pivot-feature",
            )


def _filter_records(records, filters, field_index):
    included = []
    compiled = []
    for item in filters:
        allowed = None
        if item.include is not None:
            allowed = {typed_value(value) for value in item.include}
        elif item.exclude is not None:
            allowed = None
            excluded = {typed_value(value) for value in item.exclude}
        else:
            excluded = None
            allowed = None
        compiled.append((field_index[item.field], allowed,
                         None if item.exclude is None
                         else {typed_value(value) for value in item.exclude}))
    for record in records:
        keep = True
        for index, allowed, excluded in compiled:
            value = record.values[index]
            if allowed is not None and value not in allowed:
                keep = False
                break
            if excluded is not None and value in excluded:
                keep = False
                break
        if keep:
            included.append(record)
    return included


def _ordered_keys(records, fields, field_index, axis_fields):
    if not fields:
        return (), []
    explicit = None
    if len(axis_fields) == 1 and axis_fields[0].items is not None:
        explicit = tuple(typed_value(item) for item in axis_fields[0].items)
    order = []
    seen = set()
    for record in records:
        key = _key(record, fields, field_index)
        if key not in seen:
            seen.add(key)
            order.append(key)
    if explicit is not None:
        wanted = set(explicit)
        found = {key[0] for key in order}
        missing = [item for item in explicit if item not in found]
        extra = found - wanted
        if missing or extra:
            raise BoundaryViolationError(
                "explicit item order must list each typed item exactly once",
                kind="invalid-pivot-source",
                options=[repr(item.value) for item in missing],
            )
        order = [(item,) for item in explicit]
    return tuple(order), order


def _key(record, fields, field_index):
    return tuple(record.values[field_index[name]] for name in fields)


def _accept_measure(value, aggregate):
    if value.kind == KIND_BLANK:
        return None
    if aggregate == "count":
        return 1
    if aggregate == "count_numbers":
        return value.value if value.kind == KIND_NUMBER else None
    if value.kind == KIND_NUMBER:
        return value.value
    if aggregate in ("min", "max") and value.kind in (KIND_DATE, KIND_DATETIME):
        return value
    if aggregate in ("sum", "average", "min", "max"):
        if value.kind in (KIND_TEXT, KIND_BOOLEAN, KIND_DATE, KIND_DATETIME):
            raise BoundaryViolationError(
                "%s cannot aggregate %s values"
                % (aggregate, value.kind),
                kind="invalid-pivot-source",
                options=[value.kind],
            )
    return None


def _reduce(values, aggregate):
    if not values:
        return None
    if aggregate == "count":
        return len(values)
    if aggregate == "count_numbers":
        return len(values)
    if aggregate == "sum":
        return _sum(values)
    if aggregate == "average":
        return _sum(values) / len(values)
    if aggregate == "min":
        return _extreme(values, min)
    if aggregate == "max":
        return _extreme(values, max)
    return None


def _sum(values):
    total = 0.0
    for item in values:
        total += float(item)
    if total.is_integer():
        return int(total)
    return total


def _extreme(values, chooser):
    typed = [isinstance(item, TypedValue) for item in values]
    if any(typed) and not all(typed):
        raise BoundaryViolationError(
            "min/max cannot mix numbers and dates",
            kind="invalid-pivot-source",
        )
    if values and all(typed):
        kinds = {item.kind for item in values}
        if kinds != {values[0].kind}:
            raise BoundaryViolationError(
                "min/max cannot mix date and datetime values",
                kind="invalid-pivot-source",
            )
        chosen = chooser(values, key=lambda item: item.value)
        return chosen.value
    numbers = [float(item) for item in values]
    chosen = chooser(numbers)
    if chosen.is_integer():
        return int(chosen)
    return chosen


def _totals_for(records, measures, field_index, predicate):
    values = []
    for measure in measures:
        accepted = []
        for record in records:
            if not predicate(record):
                continue
            item = _accept_measure(
                record.values[field_index[measure.field]], measure.aggregate)
            if item is not None:
                accepted.append(item)
        values.append(MeasureValue(
            measure.field, measure.aggregate,
            measure.caption or _default_caption(measure),
            measure.number_format, _reduce(accepted, measure.aggregate),
            len(accepted),
        ))
    return tuple(values)


def _default_caption(measure):
    labels = {
        "sum": "Sum of %s",
        "count": "Count of %s",
        "count_numbers": "Count of %s",
        "average": "Average of %s",
        "min": "Min of %s",
        "max": "Max of %s",
    }
    return labels[measure.aggregate] % measure.field


def display_item(value):
    """Provisional visible caption. Blank is None, not English '(blank)'."""
    if value is None or (isinstance(value, TypedValue)
                         and value.kind == KIND_BLANK):
        return None
    if isinstance(value, TypedValue):
        return value.value
    return value
