# paper-xlsx: visible pivot matrix and location planning

"""Lay out an aggregate result as a rectangular output matrix.

Coordinates are provisional hypotheses until Excel transcripts approve them.
The engine assigns cell roles and number formats only; it never writes
worksheet cells. Compact, outline, and tabular differ in how row-field labels
occupy columns, not in the underlying aggregate.
"""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl.errors import BoundaryViolationError
from openpyxl.pivot.aggregate import display_item
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.xml.constants import MAX_COLUMN, MAX_ROW


ROLE_HEADER = "header"
ROLE_ROW_LABEL = "row_label"
ROLE_COLUMN_LABEL = "column_label"
ROLE_VALUE = "value"
ROLE_SUBTOTAL = "subtotal"
ROLE_GRAND_TOTAL = "grand_total"
ROLE_BLANK = "blank"
ROLE_FILTER = "filter"


@dataclass(frozen=True)
class OutputCell:
    row: int
    column: int
    value: object
    role: str
    number_format: str | None = None
    field: str | None = None


@dataclass(frozen=True)
class PivotOutput:
    cells: tuple
    ref: str
    first_header_row: int
    first_data_row: int
    first_data_col: int
    row_count: int
    column_count: int
    destination: str
    roles: object

    def value_at(self, row, column):
        return self.roles.get((row, column))


@dataclass(frozen=True)
class ValuesRowEvent:
    kind: str
    key: tuple
    measure_index: int | None = None


def layout_result(spec, result, destination, limits=None):
    """Map ``result`` onto a rectangle whose origin is ``destination``."""
    from openpyxl.pivot.source import DEFAULT_LIMITS

    limits = limits or DEFAULT_LIMITS
    origin_col, origin_row = _origin(destination)
    row_depth = len(spec.rows)
    col_depth = len(spec.columns)
    measure_count = len(spec.values)
    values_on_columns = spec.values_axis == "columns"
    captions = [
        measure.caption or _caption(measure) for measure in spec.values]

    body_row_keys = list(result.row_keys)
    body_col_keys = list(result.column_keys)
    row_events = None
    if not values_on_columns:
        row_events = _values_row_events(spec, result)
    elif spec.subtotals and row_depth > 1:
        body_row_keys = _with_subtotal_rows(
            body_row_keys, result.row_subtotals,
            totals_first=spec.layout in ("compact", "outline"))

    header_rows = _header_row_count(
        spec, col_depth, measure_count, values_on_columns)
    label_cols = _label_column_count(spec.layout, row_depth)
    if not values_on_columns and row_depth and spec.layout != "compact":
        label_cols += 1
    if values_on_columns:
        value_cols = max(1, len(body_col_keys)) * measure_count
        body_rows = len(body_row_keys)
    else:
        value_cols = max(1, len(body_col_keys))
        body_rows = len(row_events)

    total_rows = header_rows + body_rows
    if spec.row_grand_totals:
        total_rows += 1 if values_on_columns else measure_count
    total_cols = label_cols + value_cols
    if spec.column_grand_totals and spec.columns:
        total_cols += measure_count if values_on_columns else 1

    filter_cells = len(spec.filters) * 2
    output_cells = total_rows * total_cols + filter_cells
    if output_cells > limits.output_cells:
        raise BoundaryViolationError(
            "pivot would materialize %s output cells; the limit is %s"
            % (output_cells, limits.output_cells),
            kind="pivot-output-too-large",
            options=[str(output_cells), str(limits.output_cells)],
        )
    last_row = origin_row + total_rows - 1
    last_col = origin_col + total_cols - 1
    filter_start_row = origin_row - len(spec.filters) - 1
    if spec.filters and filter_start_row < 1:
        raise BoundaryViolationError(
            "pivot destination %s does not leave room for %s report "
            "filter row(s) and the required spacer"
            % (destination, len(spec.filters)),
            kind="pivot-output-too-large",
            options=[destination, str(len(spec.filters))],
        )
    if last_row > MAX_ROW or last_col > MAX_COLUMN:
        raise BoundaryViolationError(
            "pivot output would exceed the worksheet grid at %s%s"
            % (get_column_letter(min(last_col, MAX_COLUMN)),
               min(last_row, MAX_ROW)),
            kind="pivot-output-too-large",
            options=[destination, str(MAX_ROW), str(MAX_COLUMN)],
        )

    cells = {}
    _write_filters(cells, spec, filter_start_row, origin_col)
    _write_headers(
        cells, spec, origin_row, origin_col, header_rows, label_cols,
        body_col_keys, captions, values_on_columns)
    _write_body(
        cells, spec, result, origin_row + header_rows, origin_col,
        label_cols, body_row_keys, body_col_keys, captions, values_on_columns,
        row_events=row_events)
    if spec.row_grand_totals:
        _write_row_grand_total(
            cells, spec, result, origin_row + header_rows + body_rows,
            origin_col, label_cols, body_col_keys, captions, values_on_columns)
    if spec.column_grand_totals and spec.columns:
        _write_column_grand_total(
            cells, spec, result, origin_row, origin_col, header_rows,
            label_cols, value_cols, body_row_keys, captions, values_on_columns,
            row_events=row_events)

    ref = "%s:%s" % (
        _coord(origin_col, origin_row),
        _coord(last_col, last_row),
    )
    ordered = tuple(
        cells[key] for key in sorted(cells, key=lambda item: (item[0], item[1]))
    )
    return PivotOutput(
        cells=ordered,
        ref=ref,
        first_header_row=_first_header_row(spec, values_on_columns),
        first_data_row=header_rows,
        first_data_col=label_cols,
        row_count=total_rows,
        column_count=total_cols,
        destination=_coord(origin_col, origin_row),
        roles={(cell.row, cell.column): cell for cell in ordered},
    )


def _write_filters(cells, spec, start_row, origin_col):
    for offset, item in enumerate(spec.filters):
        row = start_row + offset
        selected = item.include or ()
        if item.include is None and not item.exclude:
            display = "(All)"
        elif len(selected) == 1:
            display = display_item(selected[0])
        else:
            display = "(Multiple Items)"
        _put(cells, row, origin_col, item.field, ROLE_FILTER, field=item.field)
        _put(cells, row, origin_col + 1, display, ROLE_FILTER,
             field=item.field)


def _origin(destination):
    column, row = coordinate_from_string(destination)
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(column), row


def _coord(column, row):
    return "%s%s" % (get_column_letter(column), row)


def _header_row_count(spec, col_depth, measure_count, values_on_columns):
    if values_on_columns and spec.layout == "tabular":
        return max(2, col_depth + (2 if measure_count > 1 else 1))
    if not values_on_columns:
        return max(1, col_depth + 1 if col_depth else 1)
    extra = 1 if measure_count > 1 or col_depth == 0 else 0
    return max(1, col_depth + extra)


def _first_header_row(spec, values_on_columns):
    if values_on_columns and spec.layout == "tabular" and not spec.columns:
        return 2
    return 1


def _label_column_count(layout, row_depth):
    if row_depth == 0:
        return 1
    if layout == "compact":
        return 1
    return row_depth


def _with_subtotal_rows(row_keys, subtotals, totals_first=False):
    if not subtotals:
        return list(row_keys)
    ordered = []
    if totals_first:
        previous = None
        for key in row_keys:
            for depth in range(1, len(key)):
                prefix = key[:depth]
                if previous is None or previous[:depth] != prefix:
                    ordered.append(("__subtotal__", prefix))
            ordered.append(key)
            previous = key
        return ordered
    for index, key in enumerate(row_keys):
        ordered.append(key)
        for depth in range(len(key) - 1, 0, -1):
            prefix = key[:depth]
            nxt = row_keys[index + 1] if index + 1 < len(row_keys) else None
            if nxt is None or nxt[:depth] != prefix:
                ordered.append(("__subtotal__", prefix))
    return ordered


def _values_row_events(spec, result):
    """Return Excel's visible row hierarchy when Values is a row field."""
    events = []
    measure_count = len(spec.values)
    if spec.layout == "tabular":
        keys = list(result.row_keys)
        if spec.subtotals and len(spec.rows) > 1:
            keys = _with_subtotal_rows(
                keys, result.row_subtotals, totals_first=False)
        for key in keys:
            is_subtotal = key and key[0] == "__subtotal__"
            actual = key[1] if is_subtotal else key
            kind = "subtotal" if is_subtotal else "data"
            for measure_index in range(measure_count):
                events.append(ValuesRowEvent(
                    kind, actual, measure_index))
        return tuple(events)

    previous = None
    for key in result.row_keys:
        common = _common_prefix_length(previous, key)
        if previous is not None and spec.subtotals:
            for depth in range(len(previous) - 1, common, -1):
                prefix = previous[:depth]
                for measure_index in range(measure_count):
                    events.append(ValuesRowEvent(
                        "subtotal", prefix, measure_index))
        for depth in range(common + 1, len(key) + 1):
            events.append(ValuesRowEvent("dimension", key[:depth]))
        for measure_index in range(measure_count):
            events.append(ValuesRowEvent("data", key, measure_index))
        previous = key
    if previous is not None and spec.subtotals:
        for depth in range(len(previous) - 1, 0, -1):
            prefix = previous[:depth]
            for measure_index in range(measure_count):
                events.append(ValuesRowEvent(
                    "subtotal", prefix, measure_index))
    return tuple(events)


def _common_prefix_length(left, right):
    if left is None:
        return 0
    return next(
        (index for index, pair in enumerate(zip(left, right))
         if pair[0] != pair[1]),
        min(len(left), len(right)),
    )


def _write_headers(cells, spec, origin_row, origin_col, header_rows,
                   label_cols, column_keys, captions, values_on_columns):
    if values_on_columns and spec.layout == "tabular":
        _write_expanded_headers(
            cells, spec, origin_row, origin_col, header_rows, label_cols,
            column_keys, captions)
        return
    if not values_on_columns and spec.columns:
        if spec.layout == "compact":
            _write_compact_values_on_rows_headers(
                cells, spec, origin_row, origin_col, header_rows,
                label_cols, column_keys)
            return
        _write_values_on_rows_headers(
            cells, spec, origin_row, origin_col, header_rows, label_cols,
            column_keys)
        return
    if not values_on_columns and spec.layout == "compact":
        _put(cells, origin_row, origin_col, "Row Labels", ROLE_HEADER,
             field=spec.rows[0].field if spec.rows else None)
        _put(cells, origin_row, origin_col + label_cols, None, ROLE_BLANK)
        return
    if spec.rows and spec.layout != "compact":
        for index, field in enumerate(spec.rows):
            _put(cells, origin_row, origin_col + index, field.field,
                 ROLE_HEADER, field=field.field)
    elif spec.rows:
        _put(cells, origin_row, origin_col, "Row Labels", ROLE_HEADER,
             field=spec.rows[0].field)
    else:
        _put(cells, origin_row, origin_col, None, ROLE_BLANK)

    measure_count = len(captions) if values_on_columns else 1
    for col_index, col_key in enumerate(column_keys):
        for depth, item in enumerate(col_key):
            field = spec.columns[depth].field
            for measure_index in range(measure_count):
                column = (
                    origin_col + label_cols + col_index * measure_count
                    + measure_index
                )
                _put(cells, origin_row + depth, column, display_item(item),
                     ROLE_COLUMN_LABEL, field=field)

    if values_on_columns:
        if header_rows > len(spec.columns):
            header_row = origin_row + header_rows - 1
            for col_index, _col_key in enumerate(column_keys):
                for measure_index, caption in enumerate(captions):
                    column = (
                        origin_col + label_cols + col_index * measure_count
                        + measure_index
                    )
                    _put(cells, header_row, column, caption, ROLE_HEADER)
    else:
        _put(cells, origin_row, origin_col + label_cols - 1,
             "Values", ROLE_HEADER)
        if not spec.columns and spec.layout == "tabular":
            _put(cells, origin_row, origin_col + label_cols,
                 "Total", ROLE_HEADER)


def _write_expanded_headers(cells, spec, origin_row, origin_col,
                            header_rows, label_cols, column_keys, captions):
    bottom_row = origin_row + header_rows - 1
    if spec.rows:
        for index, field in enumerate(spec.rows):
            _put(cells, bottom_row, origin_col + index, field.field,
                 ROLE_HEADER, field=field.field)
    else:
        _put(cells, bottom_row, origin_col, None, ROLE_BLANK)

    if len(captions) == 1:
        _put(cells, origin_row, origin_col, captions[0], ROLE_HEADER)
    elif not spec.columns:
        _put(cells, origin_row, origin_col, "Values", ROLE_HEADER)

    value_start = origin_col + label_cols
    if spec.columns:
        for depth, field in enumerate(spec.columns):
            _put(cells, origin_row, value_start + depth, field.field,
                 ROLE_COLUMN_LABEL, field=field.field)
        if len(captions) > 1:
            _put(cells, origin_row, value_start + len(spec.columns),
                 "Values", ROLE_HEADER)
        measure_count = len(captions)
        previous_key = None
        for col_index, col_key in enumerate(column_keys):
            for depth, item in enumerate(col_key):
                field = spec.columns[depth].field
                repeated = previous_key is not None \
                    and previous_key[:depth + 1] == col_key[:depth + 1]
                for measure_index in range(measure_count):
                    column = (
                        value_start + col_index * measure_count
                        + measure_index
                    )
                    value = display_item(item) \
                        if measure_index == 0 and not repeated else None
                    _put(cells, origin_row + depth + 1, column, value,
                         ROLE_COLUMN_LABEL if value is not None else ROLE_BLANK,
                         field=field)
            previous_key = col_key
        if len(captions) > 1:
            for col_index, _col_key in enumerate(column_keys):
                for measure_index, caption in enumerate(captions):
                    column = (
                        value_start + col_index * measure_count
                        + measure_index
                    )
                    _put(cells, bottom_row, column, caption, ROLE_HEADER)
        return

    if len(captions) == 1:
        _put(cells, bottom_row, value_start, "Total", ROLE_HEADER)
    else:
        for measure_index, caption in enumerate(captions):
            _put(cells, bottom_row, value_start + measure_index,
                 caption, ROLE_HEADER)


def _write_values_on_rows_headers(cells, spec, origin_row, origin_col,
                                   header_rows, label_cols, column_keys):
    bottom_row = origin_row + header_rows - 1
    if spec.rows:
        for index, field in enumerate(spec.rows):
            _put(cells, bottom_row, origin_col + index, field.field,
                 ROLE_HEADER, field=field.field)
    else:
        _put(cells, bottom_row, origin_col, None, ROLE_BLANK)
    _put(cells, bottom_row, origin_col + label_cols - 1,
         "Values", ROLE_HEADER)

    value_start = origin_col + label_cols
    for depth, field in enumerate(spec.columns):
        _put(cells, origin_row, value_start + depth, field.field,
             ROLE_COLUMN_LABEL, field=field.field)

    previous_key = None
    for col_index, col_key in enumerate(column_keys):
        for depth, item in enumerate(col_key):
            repeated = previous_key is not None \
                and previous_key[:depth + 1] == col_key[:depth + 1]
            value = None if repeated else display_item(item)
            _put(cells, origin_row + depth + 1,
                 value_start + col_index, value,
                 ROLE_COLUMN_LABEL if value is not None else ROLE_BLANK,
                 field=spec.columns[depth].field)
        previous_key = col_key
    if spec.column_grand_totals:
        _put(cells, bottom_row, value_start + len(column_keys),
             "Grand Total", ROLE_HEADER)


def _write_compact_values_on_rows_headers(
        cells, spec, origin_row, origin_col, header_rows, label_cols,
        column_keys):
    bottom_row = origin_row + header_rows - 1
    value_start = origin_col + label_cols
    _put(cells, origin_row, origin_col, None, ROLE_BLANK)
    _put(cells, origin_row, value_start, "Column Labels",
         ROLE_COLUMN_LABEL)
    _put(cells, bottom_row, origin_col, "Row Labels", ROLE_HEADER,
         field=spec.rows[0].field if spec.rows else None)
    previous_key = None
    for col_index, col_key in enumerate(column_keys):
        for depth, item in enumerate(col_key):
            repeated = previous_key is not None \
                and previous_key[:depth + 1] == col_key[:depth + 1]
            value = None if repeated else display_item(item)
            _put(cells, origin_row + depth + 1,
                 value_start + col_index, value,
                 ROLE_COLUMN_LABEL if value is not None else ROLE_BLANK,
                 field=spec.columns[depth].field)
        previous_key = col_key
    if spec.column_grand_totals:
        _put(cells, bottom_row, value_start + len(column_keys),
             "Grand Total", ROLE_HEADER)


def _write_body(cells, spec, result, start_row, origin_col, label_cols,
                row_keys, column_keys, captions, values_on_columns,
                row_events=None):
    if not values_on_columns:
        _write_values_row_body(
            cells, spec, result, start_row, origin_col, label_cols,
            column_keys, captions, row_events)
        return
    cursor = start_row
    previous_key = None
    for row_key in row_keys:
        is_sub = isinstance(row_key, tuple) and row_key and row_key[0] == "__subtotal__"
        if is_sub:
            prefix = row_key[1]
            _write_label_row(
                cells, spec, cursor, origin_col, prefix, ROLE_SUBTOTAL,
                suffix="" if spec.layout in ("compact", "outline")
                else " Total")
            totals = result.row_subtotals.get(prefix, ())
            if isinstance(totals, dict):
                measure_values = [
                    totals.get(col_key, ()) for col_key in column_keys]
            else:
                measure_values = [totals]
            role = ROLE_SUBTOTAL
            if spec.layout == "outline":
                previous_key = prefix
        else:
            _write_label_row(
                cells, spec, cursor, origin_col, row_key, ROLE_ROW_LABEL,
                previous=previous_key)
            measure_values = [
                result.cells.get((row_key, col_key), ()) for col_key in column_keys]
            role = ROLE_VALUE
            previous_key = row_key
        _write_value_row(
            cells, cursor, origin_col + label_cols, column_keys,
            _flatten_measures(measure_values), captions, True, role)
        cursor += 1


def _write_values_row_body(cells, spec, result, start_row, origin_col,
                           label_cols, column_keys, captions, events):
    previous_leaf = None
    previous_dimension = None
    for offset, event in enumerate(events):
        row = start_row + offset
        if event.kind == "dimension":
            _write_label_row(
                cells, spec, row, origin_col, event.key, ROLE_ROW_LABEL,
                previous=previous_dimension)
            previous_dimension = event.key
            continue

        measure_index = event.measure_index
        if event.kind == "data":
            if spec.layout == "tabular" and measure_index == 0:
                _write_label_row(
                    cells, spec, row, origin_col, event.key, ROLE_ROW_LABEL,
                    previous=previous_leaf)
                previous_leaf = event.key
            caption_col = origin_col if spec.layout == "compact" \
                else origin_col + label_cols - 1
            _put(cells, row, caption_col, captions[measure_index], ROLE_HEADER)
            groups = [
                result.cells.get((event.key, col_key), ())
                for col_key in column_keys
            ]
            role = ROLE_VALUE
        else:
            label = "%s %s" % (
                display_item(event.key[-1]), captions[measure_index])
            _put(cells, row, origin_col, label, ROLE_SUBTOTAL)
            totals = result.row_subtotals.get(event.key, ())
            groups = [totals.get(col_key, ()) for col_key in column_keys] \
                if isinstance(totals, dict) else [totals]
            role = ROLE_SUBTOTAL

        values = [
            group[measure_index] if measure_index < len(group) else None
            for group in groups
        ]
        _write_value_row(
            cells, row, origin_col + label_cols, column_keys,
            values, captions, True, role)


def _write_row_grand_total(cells, spec, result, start_row, origin_col,
                           label_cols, column_keys, captions, values_on_columns):
    _put(cells, start_row, origin_col, "Grand Total", ROLE_GRAND_TOTAL)
    if values_on_columns:
        if spec.columns:
            values = []
            for col_key in column_keys:
                values.extend(result.column_totals.get(col_key, result.grand_total))
            _write_value_row(
                cells, start_row, origin_col + label_cols, column_keys,
                values, captions, True, ROLE_GRAND_TOTAL)
        else:
            _write_value_row(
                cells, start_row, origin_col + label_cols, column_keys,
                result.grand_total, captions, True, ROLE_GRAND_TOTAL)
    else:
        start_col = origin_col + label_cols
        for measure_index, (caption, measure) in enumerate(
                zip(captions, result.grand_total)):
            row = start_row + measure_index
            _put(cells, row, origin_col, "Total %s" % caption,
                 ROLE_GRAND_TOTAL)
            if spec.rows and spec.layout != "compact":
                _put(cells, row, origin_col + label_cols - 1, None,
                     ROLE_BLANK)
            if spec.columns:
                for col_index, col_key in enumerate(column_keys):
                    group = result.column_totals.get(col_key, result.grand_total)
                    item = group[measure_index] if measure_index < len(group) else measure
                    _put(cells, row, start_col + col_index, item.value,
                         ROLE_GRAND_TOTAL, number_format=item.number_format,
                         field=item.field)
            else:
                _put(cells, row, start_col, measure.value, ROLE_GRAND_TOTAL,
                     number_format=measure.number_format, field=measure.field)


def _write_column_grand_total(cells, spec, result, origin_row, origin_col,
                              header_rows, label_cols, value_cols, row_keys,
                              captions, values_on_columns, row_events=None):
    start_col = origin_col + label_cols + value_cols
    header_row = origin_row + header_rows - 1
    measure_count = len(captions) if values_on_columns else 1
    if values_on_columns and measure_count > 1:
        for measure_index, caption in enumerate(captions):
            caption_row = origin_row + 1 \
                if spec.layout == "tabular" else header_row - 1
            _put(cells, caption_row, start_col + measure_index,
                 "Total %s" % caption, ROLE_HEADER)
            _put(cells, header_row, start_col + measure_index, None,
                 ROLE_BLANK)
    else:
        for measure_index, caption in enumerate(
                captions if values_on_columns else ("Grand Total",)):
            _put(cells, header_row, start_col + measure_index,
                 "Grand Total" if measure_index == 0
                 or not values_on_columns else caption, ROLE_HEADER)
    cursor = origin_row + header_rows
    if not values_on_columns:
        for event in row_events:
            if event.kind == "dimension":
                cursor += 1
                continue
            if event.kind == "data":
                totals = result.row_totals.get(
                    event.key, result.grand_total)
            else:
                by_col = result.row_subtotals.get(event.key, ())
                totals = by_col.get(None, result.grand_total) \
                    if isinstance(by_col, dict) else by_col
            measure = totals[event.measure_index] \
                if event.measure_index < len(totals) else None
            if measure is not None:
                _put(cells, cursor, start_col, measure.value,
                     ROLE_SUBTOTAL if event.kind == "subtotal"
                     else ROLE_GRAND_TOTAL,
                     number_format=measure.number_format,
                     field=measure.field)
            cursor += 1
        if spec.row_grand_totals:
            for measure in result.grand_total:
                _put(cells, cursor, start_col, measure.value,
                     ROLE_GRAND_TOTAL, number_format=measure.number_format,
                     field=measure.field)
                cursor += 1
        return
    body_keys = [
        key for key in row_keys
        if not (isinstance(key, tuple) and key and key[0] == "__subtotal__")
    ]
    if spec.subtotals and len(spec.rows) > 1:
        body_keys = list(row_keys)
    for row_key in body_keys:
        is_sub = isinstance(row_key, tuple) and row_key and row_key[0] == "__subtotal__"
        lookup_key = row_key[1] if is_sub else row_key
        if is_sub:
            by_col = result.row_subtotals.get(lookup_key)
            totals = by_col.get(None, result.grand_total) if isinstance(by_col, dict) \
                else (by_col or result.grand_total)
        else:
            totals = result.row_totals.get(lookup_key, result.grand_total)
        if values_on_columns:
            _write_value_row(
                cells, cursor, start_col, (), totals, captions, True,
                ROLE_SUBTOTAL if is_sub else ROLE_GRAND_TOTAL)
            cursor += 1
        else:
            for measure_index, measure in enumerate(totals):
                _put(cells, cursor + measure_index, start_col, measure.value,
                     ROLE_GRAND_TOTAL, number_format=measure.number_format,
                     field=measure.field)
            cursor += len(totals)
    if spec.row_grand_totals:
        if values_on_columns:
            _write_value_row(
                cells, cursor, start_col, (), result.grand_total, captions,
                True, ROLE_GRAND_TOTAL)
        else:
            for measure_index, measure in enumerate(result.grand_total):
                _put(cells, cursor + measure_index, start_col, measure.value,
                     ROLE_GRAND_TOTAL, number_format=measure.number_format,
                     field=measure.field)


def _write_label_row(cells, spec, row, origin_col, key, role, suffix="",
                     previous=None):
    if not spec.rows:
        _put(cells, row, origin_col, None, ROLE_BLANK)
        return
    values = key
    if spec.layout == "compact":
        label = display_item(values[-1]) if values else None
        if suffix:
            label = "%s%s" % (display_item(values[-1]), suffix)
        _put(cells, row, origin_col, label, role)
        return
    for index, field in enumerate(spec.rows):
        if index < len(values):
            label = display_item(values[index])
            if previous is not None and index < len(previous) \
                    and previous[:index + 1] == values[:index + 1]:
                label = None
            if suffix and index == len(values) - 1:
                label = "%s%s" % (label, suffix) if label is not None \
                    else suffix.strip()
            _put(cells, row, origin_col + index, label, role, field=field.field)
        else:
            _put(cells, row, origin_col + index, None, ROLE_BLANK)


def _write_value_row(cells, row, start_col, column_keys, measures, captions,
                     values_on_columns, role):
    if not measures:
        return
    for offset, measure in enumerate(measures):
        value = measure.value if hasattr(measure, "value") else None
        number_format = getattr(measure, "number_format", None)
        field = getattr(measure, "field", None)
        _put(cells, row, start_col + offset, value, role,
             number_format=number_format, field=field)


def _flatten_measures(groups):
    flat = []
    for group in groups:
        flat.extend(group)
    return flat


def _put(cells, row, column, value, role, number_format=None, field=None):
    cells[(row, column)] = OutputCell(
        row, column, value, role, number_format, field)


def _caption(measure):
    from openpyxl.pivot.aggregate import _default_caption
    return _default_caption(measure)
