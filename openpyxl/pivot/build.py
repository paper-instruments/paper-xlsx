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
    PivotField,
    PivotTableStyle,
    RowColField,
    RowColItem,
    TableDefinition,
)
from openpyxl.xml.constants import REL_NS, SHEET_MAIN_NS
from openpyxl.xml.functions import tostring


_RECORDS_REL_ID = "rId1"
_CACHE_REL_ID = "rId1"
_CREATED_VERSION = 4
_MIN_REFRESHABLE_VERSION = 3


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


def build_pivot_payloads(plan, cache_id):
    """Serialize one Paper-owned pivot and its dedicated cache."""
    cache_id = int(cache_id)
    fields = _cache_fields(plan)
    records = _record_list(plan)
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
        id=_RECORDS_REL_ID,
    )
    table = _table_definition(plan, cache_id)
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


def _cache_fields(plan):
    fields = []
    for name, items in zip(plan.fields, plan.shared_items):
        shared = SharedItems(
            _fields=[_shared_item(item) for item in items],
            **_shared_flags(items)
        )
        fields.append(CacheField(name=name, sharedItems=shared, uniqueList=True))
    return fields


def _shared_item(item):
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
        "containsString": True if has_string or has_bool else None,
        "containsNumber": True if has_number else None,
        "containsDate": True if has_date else None,
        "containsNonDate": True if (has_string or has_number or has_bool) else None,
        "containsInteger": True if numbers and all(
            float(value).is_integer() for value in numbers) else None,
        "containsMixedTypes": True if mixed else None,
        "containsSemiMixedTypes": True if has_blank and (
            has_string or has_number or has_bool or has_date) else None,
    }
    if numbers:
        flags["minValue"] = float(min(numbers))
        flags["maxValue"] = float(max(numbers))
    if dates:
        flags["minDate"] = min(dates)
        flags["maxDate"] = max(dates)
    return flags


def _record_list(plan):
    catalogs = [ {item: index for index, item in enumerate(items)}
                 for items in plan.shared_items]
    records = []
    for record in plan.records:
        fields = []
        for value, catalog in zip(record.values, catalogs):
            fields.append(Index(v=catalog[value]))
        records.append(Record(_fields=fields))
    return RecordList(r=records)


def _table_definition(plan, cache_id):
    spec = plan.spec
    field_indexes = plan.field_indexes
    row_field = spec.rows[0].field
    measure = spec.values[0]
    row_index = field_indexes[row_field]
    measure_index = field_indexes[measure.field]
    shared_row = plan.shared_items[row_index]
    shared_lookup = {item: index for index, item in enumerate(shared_row)}

    pivot_fields = []
    for name in plan.fields:
        if name == row_field:
            items = [FieldItem(x=position) for position in range(len(shared_row))]
            items.append(FieldItem(t="default"))
            pivot_fields.append(PivotField(
                name=name,
                axis="axisRow",
                compact=False,
                outline=False,
                subtotalTop=False,
                defaultSubtotal=False,
                showAll=False,
                items=items,
            ))
        elif name == measure.field:
            pivot_fields.append(PivotField(
                name=name,
                dataField=True,
                compact=False,
                outline=False,
                defaultSubtotal=False,
                showAll=False,
            ))
        else:
            pivot_fields.append(PivotField(
                name=name,
                compact=False,
                outline=False,
                defaultSubtotal=False,
                showAll=False,
            ))

    row_items = []
    for key in plan.aggregate.row_keys:
        item = key[0]
        row_items.append(RowColItem(x=(Index(v=shared_lookup[item]),)))
    if spec.row_grand_totals:
        row_items.append(RowColItem(t="grand"))

    caption = measure.caption
    if caption is None:
        from openpyxl.pivot.aggregate import _default_caption
        caption = _default_caption(measure)

    return TableDefinition(
        name=spec.name,
        cacheId=cache_id,
        dataOnRows=False,
        dataCaption="Values",
        tag=PAPER_TAG,
        createdVersion=_CREATED_VERSION,
        updatedVersion=_CREATED_VERSION,
        minRefreshableVersion=_MIN_REFRESHABLE_VERSION,
        compact=False,
        outline=False,
        compactData=False,
        outlineData=False,
        gridDropZones=True,
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
        ),
        pivotFields=pivot_fields,
        rowFields=(RowColField(x=row_index),),
        rowItems=row_items,
        colItems=(RowColItem(t="default"),),
        dataFields=(DataField(
            name=caption,
            fld=measure_index,
            subtotal="sum",
        ),),
        pivotTableStyleInfo=PivotTableStyle(
            name="PivotStyleMedium2",
            showRowHeaders=True,
            showColHeaders=True,
            showRowStripes=False,
            showColStripes=False,
            showLastColumn=True,
        ),
        id=_CACHE_REL_ID,
    )


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
    expected_rows = len(plan.aggregate.row_keys)
    if plan.spec.row_grand_totals:
        expected_rows += 1
    if len(table.rowItems) != expected_rows:
        raise BoundaryViolationError(
            "pivot rowItems count does not match the layout",
            kind="invalid-pivot-source",
        )
    if len(table.dataFields) != 1:
        raise BoundaryViolationError(
            "PR 4 pivot tables carry exactly one data field",
            kind="unsupported-pivot-feature",
        )


def _serialize_root(obj):
    tree = obj.to_tree()
    tree.set("xmlns", SHEET_MAIN_NS)
    if getattr(obj, "id", None):
        tree.set("{%s}id" % REL_NS, obj.id)
    return tostring(tree)
