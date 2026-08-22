# paper-xlsx: owned PivotTable package payloads

"""Convert a ``PivotBuildPlan`` into Paper-owned cache, records, and table bytes.

The builder uses inherited serializers only on objects it constructs. It never
reads or reserializes foreign parts. Cache IDs and part names are supplied by
the caller; class-level ``_id`` counters are not consulted.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime

from openpyxl.errors import BoundaryViolationError
from openpyxl.pivot.cache import (
    CacheDefinition,
    CacheField,
    CacheSource,
    SharedItems,
    WorksheetSource,
)
from openpyxl.pivot.fields import (
    Boolean,
    DateTimeField,
    Index,
    Missing,
    Number,
    Text,
)
from openpyxl.pivot.qualify import PAPER_TAG
from openpyxl.pivot.record import Record, RecordList
from openpyxl.pivot.source import (
    KIND_BLANK,
    KIND_BOOLEAN,
    KIND_DATE,
    KIND_DATETIME,
    KIND_ERROR,
    KIND_NUMBER,
    KIND_TEXT,
)
from openpyxl.pivot.table import (
    DataField,
    FieldItem,
    Location,
    PageField,
    PivotField,
    PivotTableStyle,
    RowColField,
    RowColItem,
    TableDefinition,
)
from openpyxl.xml.constants import REL_NS, SHEET_MAIN_NS
from openpyxl.xml.functions import tostring


_RECORDS_REL_ID = "rId1"
_CREATED_VERSION = 4
_MIN_REFRESHABLE_VERSION = 3
_VALUES_FIELD = -2
_DEFAULT_STYLE = "PivotStyleMedium2"
_AGGREGATE_XML = {
    "sum": "sum",
    "count": "count",
    "count_numbers": "countNums",
    "average": "average",
    "min": "min",
    "max": "max",
}
BUILT_IN_PIVOT_STYLES = frozenset(
    "PivotStyle%s%d" % (tone, number)
    for tone in ("Light", "Medium", "Dark")
    for number in range(1, 29)
)


@dataclass(frozen=True)
class PivotPayloads:
    cache_definition: bytes
    cache_records: bytes
    pivot_table: bytes

    def to_dict(self):
        return {
            "cache_definition_bytes": len(self.cache_definition),
            "cache_records_bytes": len(self.cache_records),
            "pivot_table_bytes": len(self.pivot_table),
        }


def build_pivot_payloads(
        plan, cache_id, workbook=None, *,
        records_relationship_id=_RECORDS_REL_ID,
        pivot_cache_relationship_id=_RECORDS_REL_ID):
    """Serialize a Paper-owned pivot with its retained relationship IDs."""
    cache_id = int(cache_id)
    indexed_fields = _indexed_field_names(plan.spec)
    fields = _cache_fields(plan, indexed_fields)
    records = _record_list(plan, indexed_fields)
    cache = CacheDefinition(
        saveData=True,
        enableRefresh=True,
        refreshOnLoad=False,
        backgroundQuery=False,
        createdVersion=_CREATED_VERSION,
        refreshedVersion=_CREATED_VERSION,
        minRefreshableVersion=_MIN_REFRESHABLE_VERSION,
        recordCount=len(records.r),
        cacheSource=_cache_source(plan.spec.source),
        cacheFields=fields,
        id=records_relationship_id,
    )
    table = _table_definition(
        plan, cache_id, workbook,
        pivot_cache_relationship_id=pivot_cache_relationship_id,
    )
    _assert_consistent(plan, cache, records, table, cache_id)
    return PivotPayloads(
        cache_definition=_serialize_root(cache),
        cache_records=_serialize_root(records),
        pivot_table=_serialize_root(table),
    )


def _cache_source(source):
    if source.kind == "table":
        worksheet = WorksheetSource(name=source.name)
    elif source.kind == "range":
        worksheet = WorksheetSource(ref=source.ref, sheet=source.sheet)
    else:
        raise BoundaryViolationError(
            "pivot builder only emits worksheet table or range sources",
            kind="unsupported-pivot-source",
            options=[source.kind],
        )
    return CacheSource(type="worksheet", worksheetSource=worksheet)


def _indexed_field_names(spec):
    return frozenset(
        item.field
        for collection in (spec.rows, spec.columns, spec.filters)
        for item in collection
    )


def _cache_fields(plan, indexed_fields):
    fields = []
    for name, items in zip(plan.fields, plan.shared_items):
        shared = SharedItems(
            _fields=[_typed_item(item) for item in items]
            if name in indexed_fields else (),
            **_shared_flags(items)
        )
        fields.append(CacheField(
            name=name,
            sharedItems=shared,
            uniqueList=None,
            sqlType=None,
            hierarchy=None,
            level=None,
            databaseField=None,
        ))
    return fields


def _typed_item(item):
    if item.kind == KIND_BLANK:
        return Missing()
    if item.kind == KIND_TEXT:
        return Text(v=item.value)
    if item.kind == KIND_NUMBER:
        return Number(v=float(item.value))
    if item.kind == KIND_BOOLEAN:
        return Boolean(v=item.value)
    if item.kind in (KIND_DATE, KIND_DATETIME):
        value = item.value
        if isinstance(value, date) and not isinstance(value, datetime):
            value = datetime(value.year, value.month, value.day)
        return DateTimeField(v=value)
    if item.kind == KIND_ERROR:
        from openpyxl.pivot.fields import Error
        return Error(v=item.value)
    raise BoundaryViolationError(
        "cannot serialize pivot shared item kind %r" % item.kind,
        kind="invalid-pivot-source",
        options=[item.kind],
    )


def _shared_flags(items):
    kinds = {item.kind for item in items}
    numbers = [item.value for item in items if item.kind == KIND_NUMBER]
    dates = [item.value for item in items
             if item.kind in (KIND_DATE, KIND_DATETIME)]
    has_string = bool(kinds & {KIND_TEXT, KIND_ERROR})
    has_number = KIND_NUMBER in kinds
    has_date = bool(kinds & {KIND_DATE, KIND_DATETIME})
    has_blank = KIND_BLANK in kinds
    has_bool = KIND_BOOLEAN in kinds
    mixed = len(kinds - {KIND_BLANK}) > 1
    flags = {
        "containsBlank": True if has_blank else None,
        "containsString": (
            False if (has_number or has_date) and not (has_string or has_bool)
            else True if (has_string or has_bool) and mixed else None
        ),
        "containsNumber": True if has_number else None,
        "containsDate": True if has_date else None,
        "containsNonDate": True if has_date and (
            has_string or has_number or has_bool) else None,
        "containsInteger": True if numbers and all(
            float(value).is_integer() for value in numbers) else None,
        "containsMixedTypes": True if mixed else None,
        "containsSemiMixedTypes": (
            True if has_blank and (
                has_string or has_number or has_bool or has_date)
            else False if has_number or has_date else None
        ),
    }
    if numbers:
        flags["minValue"] = float(min(numbers))
        flags["maxValue"] = float(max(numbers))
    if dates:
        normalized = []
        for value in dates:
            if isinstance(value, datetime):
                normalized.append(value)
            elif isinstance(value, date):
                normalized.append(datetime(value.year, value.month, value.day))
            else:
                normalized.append(value)
        flags["minDate"] = min(normalized)
        flags["maxDate"] = max(normalized)
    return flags


def _record_list(plan, indexed_fields):
    catalogs = [ {item: index for index, item in enumerate(items)}
                 for items in plan.shared_items]
    records = []
    for record in plan.records:
        fields = []
        for name, value, catalog in zip(
                plan.fields, record.values, catalogs):
            if name in indexed_fields:
                fields.append(Index(v=catalog[value]))
            else:
                fields.append(_typed_item(value))
        records.append(Record(_fields=fields))
    return RecordList(r=records)


def _table_definition(plan, cache_id, workbook=None,
                     pivot_cache_relationship_id=_RECORDS_REL_ID):
    spec = plan.spec
    lookups = [
        {item: index for index, item in enumerate(shared)}
        for shared in plan.shared_items
    ]
    flags = _layout_flags(spec.layout)
    row_names = [item.field for item in spec.rows]
    col_names = [item.field for item in spec.columns]
    filter_names = [item.field for item in spec.filters]
    measure_fields = {item.field for item in spec.values}
    values_on_rows = spec.values_axis == "rows"
    values_field = _needs_values_field(spec)

    pivot_fields = []
    item_lookups = {}
    for name in plan.fields:
        index = plan.field_indexes[name]
        shared = plan.shared_items[index]
        if name in row_names:
            axis = spec.rows[row_names.index(name)]
            field, item_lookups[name] = _axis_pivot_field(
                name, "axisRow", shared, axis, spec, flags)
            pivot_fields.append(field)
        elif name in col_names:
            axis = spec.columns[col_names.index(name)]
            field, item_lookups[name] = _axis_pivot_field(
                name, "axisCol", shared, axis, spec, flags)
            pivot_fields.append(field)
        elif name in filter_names:
            item = spec.filters[filter_names.index(name)]
            pivot_fields.append(_filter_pivot_field(
                name, shared, item, flags))
        else:
            pivot_fields.append(PivotField(
                name=name,
                dataField=True if name in measure_fields else None,
                compact=flags["field_compact"],
                outline=flags["field_outline"],
                defaultSubtotal=False,
                showAll=False,
            ))

    row_fields = [RowColField(x=plan.field_indexes[name]) for name in row_names]
    col_fields = [RowColField(x=plan.field_indexes[name]) for name in col_names]
    if values_field and values_on_rows:
        row_fields.append(RowColField(x=_VALUES_FIELD))
    if values_field and not values_on_rows:
        col_fields.append(RowColField(x=_VALUES_FIELD))

    if values_on_rows:
        row_items = _values_axis_row_items(plan, row_names, item_lookups)
    else:
        row_items = _item_tuples(
            _visible_row_keys(plan), row_names, item_lookups, None)
        if spec.row_grand_totals:
            row_items.append(RowColItem(t="grand"))

    col_keys = list(plan.aggregate.column_keys)
    if spec.columns:
        col_items = _item_tuples(
            col_keys, col_names, item_lookups,
            spec.values if (not values_on_rows and values_field) else None)
        if spec.column_grand_totals:
            if not values_on_rows and values_field:
                for measure_index, _measure in enumerate(spec.values):
                    col_items.append(RowColItem(
                        t="grand", i=measure_index,
                        x=(Index(v=measure_index),)))
            else:
                col_items.append(RowColItem(t="grand"))
    elif values_field and not values_on_rows:
        col_items = []
        previous = None
        for index, _measure in enumerate(spec.values):
            item, previous = _data_row_col_item(
                (index,), previous, measure_index=index)
            col_items.append(item)
    else:
        col_items = [RowColItem()]

    page_fields = []
    for item in spec.filters:
        index = plan.field_indexes[item.field]
        lookup = lookups[index]
        selected = _selected_items(item, plan.shared_items[index])
        if item.include is not None and len(selected) == 1:
            page_item = lookup[selected[0]]
        else:
            page_item = None
        page_fields.append(PageField(fld=index, item=page_item, hier=-1))

    from openpyxl.pivot.aggregate import _default_caption
    data_fields = []
    for measure in spec.values:
        caption = measure.caption or _default_caption(measure)
        data_fields.append(DataField(
            name=caption,
            fld=plan.field_indexes[measure.field],
            subtotal=_AGGREGATE_XML[measure.aggregate],
            baseField=0,
            baseItem=0,
            numFmtId=_number_format_id(workbook, measure.number_format),
        ))

    style_name = spec.style or _DEFAULT_STYLE
    return TableDefinition(
        name=spec.name,
        cacheId=cache_id,
        dataOnRows=values_on_rows,
        dataCaption="Values",
        id=pivot_cache_relationship_id,
        tag=PAPER_TAG,
        createdVersion=_CREATED_VERSION,
        updatedVersion=_CREATED_VERSION,
        minRefreshableVersion=_MIN_REFRESHABLE_VERSION,
        compact=flags["compact"],
        outline=flags["outline"],
        compactData=flags["compactData"],
        outlineData=flags["outlineData"],
        gridDropZones=flags["gridDropZones"],
        useAutoFormatting=False,
        itemPrintTitles=False,
        rowGrandTotals=spec.row_grand_totals,
        colGrandTotals=spec.column_grand_totals,
        showHeaders=True,
        location=Location(
            ref=plan.output.ref,
            firstHeaderRow=plan.output.first_header_row,
            firstDataRow=plan.output.first_data_row,
            firstDataCol=plan.output.first_data_col,
            rowPageCount=len(spec.filters) if spec.filters else None,
            colPageCount=1 if spec.filters else None,
        ),
        pivotFields=pivot_fields,
        rowFields=tuple(row_fields),
        rowItems=tuple(row_items),
        colFields=tuple(col_fields),
        colItems=tuple(col_items),
        pageFields=tuple(page_fields),
        dataFields=tuple(data_fields),
        pivotTableStyleInfo=PivotTableStyle(
            name=style_name,
            showRowHeaders=True,
            showColHeaders=True,
            showRowStripes=False,
            showColStripes=False,
            showLastColumn=True,
        ),
    )


def _layout_flags(layout):
    if layout == "compact":
        return {
            "compact": True,
            "outline": True,
            "compactData": True,
            "outlineData": True,
            "gridDropZones": False,
            "field_compact": True,
            "field_outline": True,
        }
    if layout == "outline":
        return {
            "compact": False,
            "outline": True,
            "compactData": False,
            "outlineData": True,
            "gridDropZones": False,
            "field_compact": False,
            "field_outline": True,
        }
    return {
        "compact": False,
        "outline": False,
        "compactData": False,
        "outlineData": False,
        "gridDropZones": True,
        "field_compact": False,
        "field_outline": False,
    }


def _needs_values_field(spec):
    return len(spec.values) > 1 or spec.values_axis == "rows"


def _axis_pivot_field(name, axis, shared, axis_field, spec, flags):
    items, item_lookup = _field_items(shared, axis_field.items)
    default_subtotal = None
    if axis == "axisRow" and len(spec.rows) > 1 \
            and spec.rows[-1].field != name:
        if not spec.subtotals:
            default_subtotal = False
            items = items[:-1]
    elif axis == "axisCol" and len(spec.columns) > 1 \
            and spec.columns[-1].field != name:
        # v1 subtotals apply only to row fields. Excel otherwise supplies
        # automatic subtotals for outer column fields during refresh.
        default_subtotal = False
        items = items[:-1]
    field = PivotField(
        name=name,
        axis=axis,
        compact=flags["field_compact"],
        outline=flags["field_outline"],
        subtotalTop=spec.layout != "tabular",
        defaultSubtotal=default_subtotal,
        showAll=False,
        items=items,
    )
    return field, item_lookup


def _filter_pivot_field(name, shared, item, flags):
    selected = set(_selected_items(item, shared))
    items = []
    for position, value in enumerate(shared):
        hidden = value not in selected
        items.append(FieldItem(
            t=None, sd=None, x=position, h=True if hidden else None))
    items.append(FieldItem(t="default", sd=None))
    return PivotField(
        name=name,
        axis="axisPage",
        compact=flags["field_compact"],
        outline=flags["field_outline"],
        defaultSubtotal=None,
        showAll=False,
        multipleItemSelectionAllowed=(
            True if item.include is None or len(selected) != 1 else None),
        items=items,
    )


def _field_items(shared, explicit):
    if explicit is None:
        order = list(range(len(shared)))
    else:
        from openpyxl.pivot.source import typed_value
        lookup = {item: index for index, item in enumerate(shared)}
        order = [lookup[typed_value(value)] for value in explicit]
    items = [FieldItem(t=None, sd=None, x=position) for position in order]
    items.append(FieldItem(t="default", sd=None))
    return items, {shared[position]: index for index, position in enumerate(order)}


def _selected_items(item, shared):
    from openpyxl.pivot.source import typed_value
    if item.include is not None:
        return tuple(typed_value(value) for value in item.include)
    excluded = {typed_value(value) for value in (item.exclude or ())}
    return tuple(value for value in shared if value not in excluded)


def _visible_row_keys(plan):
    keys = list(plan.aggregate.row_keys)
    spec = plan.spec
    if spec.subtotals and len(spec.rows) > 1:
        from openpyxl.pivot.layout import _with_subtotal_rows
        keys = _with_subtotal_rows(
            keys, plan.aggregate.row_subtotals,
            totals_first=spec.layout in ("compact", "outline"))
    return keys


def _item_tuples(keys, field_names, item_lookups, measures):
    items = []
    previous = None
    for key in keys:
        if isinstance(key, tuple) and key and key[0] == "__subtotal__":
            previous = None
            prefix = key[1]
            indexes = _key_indexes(prefix, field_names, item_lookups)
            if measures is None:
                items.append(RowColItem(t="default", x=indexes))
            else:
                for measure_index, _measure in enumerate(measures):
                    items.append(RowColItem(
                        t="default", i=measure_index,
                        x=indexes + (Index(v=measure_index),)))
            continue
        indexes = _key_indexes(key, field_names, item_lookups)
        if measures is None:
            item, previous = _data_row_col_item(
                tuple(index.v for index in indexes), previous)
            items.append(item)
        else:
            for measure_index, _measure in enumerate(measures):
                values = tuple(index.v for index in indexes) + (measure_index,)
                item, previous = _data_row_col_item(
                    values, previous, measure_index=measure_index)
                items.append(item)
    return items


def _values_axis_row_items(plan, field_names, item_lookups):
    from openpyxl.pivot.layout import _values_row_events

    items = []
    previous = None
    for event in _values_row_events(plan.spec, plan.aggregate):
        indexes = tuple(
            index.v for index in _key_indexes(
                event.key, field_names, item_lookups)
        )
        if event.kind == "dimension":
            item, previous = _data_row_col_item(indexes, previous)
        elif event.kind == "data":
            values = indexes + (event.measure_index,)
            item, previous = _data_row_col_item(
                values, previous, measure_index=event.measure_index)
        else:
            item = RowColItem(
                t="default",
                i=event.measure_index or None,
                x=tuple(Index(v=None if value == 0 else value)
                        for value in indexes),
            )
            previous = None
        items.append(item)
    if plan.spec.row_grand_totals:
        for measure_index, _measure in enumerate(plan.spec.values):
            items.append(RowColItem(
                t="grand",
                i=measure_index or None,
                x=(Index(v=None),),
            ))
    return items


def _data_row_col_item(values, previous, measure_index=None):
    repeated = 0
    if previous is not None:
        repeated = next(
            (index for index, pair in enumerate(zip(previous, values))
             if pair[0] != pair[1]),
            min(len(previous), len(values)),
        )
    indexes = tuple(
        Index(v=None if value == 0 else value)
        for value in values[repeated:]
    )
    item = RowColItem(
        t=None,
        r=repeated or None,
        i=measure_index or None,
        x=indexes,
    )
    return item, values


def _key_indexes(key, field_names, item_lookups):
    if not field_names:
        return ()
    indexes = []
    for offset, name in enumerate(field_names):
        if offset >= len(key):
            break
        indexes.append(Index(v=item_lookups[name][key[offset]]))
    return tuple(indexes)


def _number_format_id(workbook, number_format):
    if not number_format or workbook is None:
        return None
    from openpyxl.styles.numbers import (
        BUILTIN_FORMATS_MAX_SIZE,
        builtin_format_id,
    )

    builtin = builtin_format_id(number_format)
    if builtin is not None:
        return builtin
    formats = workbook._number_formats
    if number_format not in formats:
        formats.add(number_format)
    return formats.index(number_format) + BUILTIN_FORMATS_MAX_SIZE


def _assert_consistent(plan, cache, records, table, cache_id):
    if cache.recordCount != len(records.r):
        raise BoundaryViolationError(
            "pivot cache recordCount %s does not match %s records"
            % (cache.recordCount, len(records.r)),
            kind="invalid-pivot-source",
        )
    if cache.recordCount != len(plan.records):
        raise BoundaryViolationError(
            "pivot cache records do not match the source snapshot",
            kind="invalid-pivot-source",
        )
    if len(cache.cacheFields) != len(plan.fields):
        raise BoundaryViolationError(
            "pivot cache field count does not match the source snapshot",
            kind="invalid-pivot-source",
        )
    if table.cacheId != cache_id:
        raise BoundaryViolationError(
            "pivot table cacheId does not match the allocated cache",
            kind="invalid-pivot-graph",
        )
    if len(table.pivotFields) != len(plan.fields):
        raise BoundaryViolationError(
            "pivot field catalog does not match the source snapshot",
            kind="invalid-pivot-source",
        )
    if len(table.dataFields) != len(plan.spec.values):
        raise BoundaryViolationError(
            "pivot data field count does not match the spec",
            kind="invalid-pivot-source",
        )
    if table.location.ref != plan.output.ref:
        raise BoundaryViolationError(
            "pivot location does not match the layout range",
            kind="invalid-pivot-source",
        )


def _serialize_root(obj):
    tree = obj.to_tree()
    tree.set("xmlns", SHEET_MAIN_NS)
    if getattr(obj, "id", None):
        tree.set("{%s}id" % REL_NS, obj.id)
    return tostring(tree)
