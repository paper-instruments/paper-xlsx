# paper-xlsx: semantic projection of a loaded pivot graph

"""Project a relationship-resolved pivot graph into public value types.

Projection is read-only. A field index without a matching cache field, an
item index outside its shared-item set, or an unknown aggregate prevents a
complete ``PivotSpec`` and disables every capability that depends on that
projection. Source records and output cells are never enumerated.
"""

from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass

from openpyxl.pivot.api_types import (
    SUPPORTED_AGGREGATES,
    PivotAxisField,
    PivotItemFilter,
    PivotMeasure,
    PivotSource,
    PivotSpec,
)
from openpyxl.pivot.graph import (
    _SHARED_ITEM_TAGS,
    _attr,
    _children,
    _first,
    _local,
    _parse_xml,
)
from openpyxl.utils import get_column_letter, range_boundaries


_AGGREGATE_SET = frozenset(SUPPORTED_AGGREGATES)
_TRUE = frozenset(("1", "true", "True"))
_FALSE = frozenset(("0", "false", "False"))


@dataclass(frozen=True)
class ProjectionReason:
    code: str
    context: tuple = ()

    def to_dict(self):
        return {
            "code": self.code,
            "context": {key: value for key, value in self.context},
        }


@dataclass(frozen=True)
class PivotProjection:
    complete: bool
    spec: PivotSpec | None
    source: PivotSource | None
    destination: str | None
    output_range: str | None
    rows: tuple | None
    columns: tuple | None
    filters: tuple | None
    values: tuple | None
    reasons: tuple = ()

    def to_dict_fields(self):
        payload = {}
        if self.source is not None:
            payload["source"] = self.source.to_dict()
        if self.destination is not None:
            payload["destination"] = self.destination
        if self.output_range is not None:
            payload["output_range"] = self.output_range
        if self.rows is not None:
            payload["rows"] = [item.field for item in self.rows]
        if self.columns is not None:
            payload["columns"] = [item.field for item in self.columns]
        if self.filters is not None:
            payload["filters"] = [item.to_dict() for item in self.filters]
        if self.values is not None:
            payload["values"] = [
                {"field": item.field, "aggregate": item.aggregate}
                for item in self.values
            ]
        return payload


def _reason(code, **context):
    items = tuple(sorted(
        (key, value) for key, value in context.items() if value is not None))
    return ProjectionReason(code, items)


def project_pivot(node, cache, source=None, workbook=None):
    """Project one graph node into a semantic specification when possible."""
    reasons = []
    details = _load_details(node, cache, source)
    source_obj = _project_source(node, cache, workbook, reasons)
    output_range = node.output_range
    destination = _destination_from_ref(output_range)
    if output_range and destination is None:
        reasons.append(_reason(
            "invalid-output-location",
            part=node.identity.pivot_part,
            ref=output_range,
        ))
    field_names = () if cache is None else cache.field_names
    shared_counts = () if cache is None else tuple(
        len(kinds) for kinds in cache.shared_item_kinds)
    shared_values = details.get("shared_items") or ()

    rows = _project_axis(
        node.row_fields, field_names, shared_counts, details.get("row_items_by_field"),
        shared_values, node.identity.pivot_part, "rows", reasons)
    columns = _project_axis(
        node.column_fields, field_names, shared_counts,
        details.get("column_items_by_field"), shared_values,
        node.identity.pivot_part, "columns", reasons)
    filters = _project_filters(
        details.get("page_fields") or _page_fields_from_node(node),
        field_names, shared_counts, shared_values,
        details.get("page_includes") or {},
        node.identity.pivot_part, reasons)
    values = _project_values(
        node.data_fields, field_names, node.identity.pivot_part, reasons)

    complete = not reasons
    spec = None
    if complete and source_obj is not None and destination and values:
        try:
            spec = PivotSpec(
                name=node.identity.name,
                source=source_obj,
                destination=destination,
                rows=rows or (),
                columns=columns or (),
                filters=filters or (),
                values=values,
                layout=details.get("layout") or "compact",
                values_axis=details.get("values_axis") or "columns",
                row_grand_totals=details.get("row_grand_totals", True),
                column_grand_totals=details.get("column_grand_totals", True),
                subtotals=details.get("subtotals", False),
                style=details.get("style"),
            )
        except (TypeError, ValueError) as exc:
            reasons.append(_reason(
                "incomplete-semantic-projection",
                part=node.identity.pivot_part,
                detail=str(exc),
            ))
            complete = False
            spec = None
    elif complete:
        complete = False
        if source_obj is None:
            reasons.append(_reason(
                "unsupported-source",
                part=(None if cache is None else cache.definition_part),
            ))
        if not destination:
            reasons.append(_reason(
                "missing-output-location",
                part=node.identity.pivot_part,
            ))
        if not values:
            reasons.append(_reason(
                "missing-measure",
                part=node.identity.pivot_part,
            ))

    return PivotProjection(
        complete=complete and spec is not None,
        spec=spec,
        source=source_obj,
        destination=destination,
        output_range=output_range,
        rows=rows,
        columns=columns,
        filters=filters,
        values=values,
        reasons=tuple(reasons),
    )


def _page_fields_from_node(node):
    return tuple((index, None) for index in node.page_fields)


def _project_source(node, cache, workbook, reasons):
    descriptor = None
    if node.source_descriptor is not None:
        descriptor = node.source_descriptor
    elif cache is not None:
        descriptor = cache.source_descriptor
    if descriptor is None:
        return None
    if descriptor.kind == "table" and descriptor.name:
        return PivotSource.table(descriptor.name)
    if descriptor.kind == "range" and descriptor.sheet and descriptor.ref:
        try:
            return PivotSource.range(descriptor.sheet, descriptor.ref)
        except ValueError:
            reasons.append(_reason(
                "unsupported-source",
                sheet=descriptor.sheet,
                ref=descriptor.ref,
            ))
            return None
    if descriptor.kind == "defined-name" and descriptor.name:
        resolved = _resolve_defined_name(workbook, descriptor.name)
        if resolved is not None:
            return resolved
        return PivotSource(kind="defined-name", name=descriptor.name,
                           sheet=descriptor.sheet)
    reasons.append(_reason(
        "unsupported-source",
        kind=descriptor.kind,
        name=descriptor.name,
        sheet=descriptor.sheet,
        ref=descriptor.ref,
    ))
    return None


def _resolve_defined_name(workbook, name):
    if workbook is None:
        return None
    folded = name.casefold()
    holders = [getattr(workbook, "defined_names", {})]
    holders.extend(
        ws.defined_names for ws in getattr(workbook, "worksheets", ()))
    for holder in holders:
        for key, defn in (holder or {}).items():
            declared = getattr(defn, "name", None)
            if not (
                (isinstance(key, str) and key.casefold() == folded)
                or (isinstance(declared, str) and declared.casefold() == folded)
            ):
                continue
            attr = getattr(defn, "attr_text", None) or getattr(defn, "value", None)
            if not isinstance(attr, str):
                continue
            try:
                return PivotSource.parse(attr)
            except (TypeError, ValueError):
                return None
    return None


def _destination_from_ref(ref):
    if not ref:
        return None
    try:
        min_col, min_row, _max_col, _max_row = range_boundaries(ref)
    except (TypeError, ValueError):
        return None
    if None in (min_col, min_row):
        return None
    return "%s%s" % (get_column_letter(min_col), min_row)


def _project_axis(indexes, field_names, shared_counts, item_indexes,
                  shared_values, part, axis, reasons):
    if not indexes:
        return ()
    fields = []
    for index in indexes:
        if index == -2:
            continue
        name = _field_name(index, field_names, part, axis, reasons)
        if name is None:
            return None
        items = None
        chosen = None if item_indexes is None else item_indexes.get(index)
        catalog = shared_values[index] if index < len(shared_values) else ()
        if chosen:
            if list(chosen) == list(range(len(catalog))):
                items = None
            else:
                values = []
                count = shared_counts[index] if index < len(shared_counts) else 0
                for item_index in chosen:
                    if item_index < 0 or (count and item_index >= count):
                        reasons.append(_reason(
                            "invalid-item-index",
                            part=part,
                            axis=axis,
                            field=name,
                            index=str(item_index),
                        ))
                        return None
                    if item_index < len(catalog):
                        values.append(catalog[item_index])
                    else:
                        values.append(item_index)
                items = tuple(values) if values else None
        fields.append(PivotAxisField(name, items=items))
    return tuple(fields)


def _project_filters(page_fields, field_names, shared_counts, shared_values,
                     page_includes, part, reasons):
    if not page_fields:
        return ()
    filters = []
    for field_index, item_index in page_fields:
        name = _field_name(
            field_index, field_names, part, "filters", reasons)
        if name is None:
            return None
        include = None
        selected = page_includes.get(field_index)
        if item_index is not None and item_index != -1:
            count = shared_counts[field_index] \
                if field_index < len(shared_counts) else 0
            if count and (item_index < 0 or item_index >= count):
                reasons.append(_reason(
                    "invalid-item-index",
                    part=part,
                    axis="filters",
                    field=name,
                    index=str(item_index),
                ))
                return None
            value = None
            if field_index < len(shared_values) \
                    and item_index < len(shared_values[field_index]):
                value = shared_values[field_index][item_index]
            if value is None and count and item_index < count:
                include = (item_index,)
            elif value is not None:
                include = (value,)
        elif selected:
            values = []
            shared = shared_values[field_index] if field_index < len(shared_values) else ()
            for shared_index in selected:
                if shared_index < len(shared):
                    values.append(shared[shared_index])
                else:
                    values.append(shared_index)
            include = tuple(values) if values else None
        filters.append(PivotItemFilter(name, include=include))
    return tuple(filters)


def _project_values(data_fields, field_names, part, reasons):
    if not data_fields:
        return None
    values = []
    for item in data_fields:
        index = item.get("field")
        aggregate = item.get("aggregate")
        name = _field_name(index, field_names, part, "values", reasons)
        if name is None:
            return None
        if aggregate not in _AGGREGATE_SET:
            reasons.append(_reason(
                "unknown-aggregate",
                part=part,
                field=name,
                aggregate=str(aggregate),
            ))
            return None
        caption = item.get("name")
        values.append(PivotMeasure(
            name,
            aggregate=aggregate,
            caption=caption if caption else None,
        ))
    return tuple(values)


def _field_name(index, field_names, part, axis, reasons):
    if not isinstance(index, int) or index < 0 or index >= len(field_names):
        reasons.append(_reason(
            "missing-field",
            part=part,
            axis=axis,
            index=None if index is None else str(index),
        ))
        return None
    name = field_names[index]
    if not name:
        reasons.append(_reason(
            "missing-field",
            part=part,
            axis=axis,
            index=str(index),
        ))
        return None
    return name


def _load_details(node, cache, source):
    details = {
        "page_fields": _page_fields_from_node(node),
        "layout": None,
        "values_axis": "columns",
        "row_grand_totals": True,
        "column_grand_totals": True,
        "subtotals": False,
        "style": None,
        "shared_items": (),
        "row_items": None,
        "column_items": None,
        "row_items_by_field": {},
        "column_items_by_field": {},
        "page_includes": {},
    }
    if not source or not node.identity.pivot_part:
        return details
    try:
        with zipfile.ZipFile(io.BytesIO(source)) as zin:
            names = set(zin.namelist())
            if node.identity.pivot_part in names:
                _enrich_from_pivot(
                    _parse_xml(zin.read(node.identity.pivot_part)), details)
            cache_part = None if cache is None else cache.definition_part
            if cache_part and cache_part in names:
                details["shared_items"] = _shared_items(
                    _parse_xml(zin.read(cache_part)))
    except (OSError, zipfile.BadZipFile, ValueError):
        return details
    return details


def _enrich_from_pivot(root, details):
    if root is None:
        return
    compact = _bool_attr(root, "compact", default=True)
    outline = _bool_attr(root, "outline", default=True)
    if compact:
        details["layout"] = "compact"
    elif outline:
        details["layout"] = "outline"
    else:
        details["layout"] = "tabular"
    if _bool_attr(root, "dataOnRows", default=False):
        details["values_axis"] = "rows"
    details["row_grand_totals"] = _bool_attr(
        root, "rowGrandTotals", default=True)
    details["column_grand_totals"] = _bool_attr(
        root, "colGrandTotals", default=True)
    style = _first(root, "pivotTableStyleInfo")
    if style is not None:
        details["style"] = _attr(style, "name")
    page = _first(root, "pageFields")
    if page is not None:
        fields = []
        for child in _children(page, "pageField"):
            raw = _attr(child, "fld")
            item = _attr(child, "item")
            try:
                field_index = int(raw)
            except (TypeError, ValueError):
                continue
            item_index = None
            if item is not None:
                try:
                    item_index = int(item)
                except (TypeError, ValueError):
                    item_index = None
            fields.append((field_index, item_index))
        details["page_fields"] = tuple(fields)
    container = _first(root, "pivotFields")
    fields = _children(container, "pivotField") if container is not None else []
    row_indexes = []
    column_indexes = []
    row_items_by_field = {}
    column_items_by_field = {}
    page_includes = {}
    subtotals = False
    for field_index, field in enumerate(fields):
        axis = _attr(field, "axis")
        items = tuple(_item_indexes(field))
        if axis == "axisRow":
            row_indexes.append(items)
            row_items_by_field[field_index] = items
            if _bool_attr(field, "defaultSubtotal", default=True):
                subtotals = True
        elif axis == "axisCol":
            column_indexes.append(items)
            column_items_by_field[field_index] = items
        elif axis == "axisPage":
            page_includes[field_index] = tuple(
                _visible_item_indexes(field))
    if row_indexes:
        details["row_items"] = tuple(row_indexes)
    if column_indexes:
        details["column_items"] = tuple(column_indexes)
    details["row_items_by_field"] = row_items_by_field
    details["column_items_by_field"] = column_items_by_field
    details["page_includes"] = page_includes
    details["subtotals"] = subtotals


def _item_indexes(field):
    items = _first(field, "items")
    if items is None:
        return ()
    values = []
    for child in _children(items, "item"):
        raw = _attr(child, "x")
        if raw is None:
            continue
        try:
            values.append(int(raw))
        except (TypeError, ValueError):
            continue
    return values


def _visible_item_indexes(field):
    items = _first(field, "items")
    if items is None:
        return ()
    values = []
    for child in _children(items, "item"):
        if _bool_attr(child, "h", default=False):
            continue
        raw = _attr(child, "x")
        if raw is None:
            continue
        try:
            values.append(int(raw))
        except (TypeError, ValueError):
            continue
    return values


def _shared_items(root):
    if root is None:
        return ()
    container = _first(root, "cacheFields")
    if container is None:
        return ()
    values = []
    for field in _children(container, "cacheField"):
        shared = _first(field, "sharedItems")
        items = []
        if shared is not None:
            for child in list(shared):
                if _local(child.tag) in _SHARED_ITEM_TAGS:
                    items.append(_shared_item_value(child))
        values.append(tuple(items))
    return tuple(values)


def _shared_item_value(element):
    tag = _local(element.tag)
    if tag == "m":
        return None
    raw = _attr(element, "v")
    if tag == "n":
        try:
            number = float(raw)
        except (TypeError, ValueError):
            return raw
        if number.is_integer():
            return int(number)
        return number
    if tag == "b":
        return raw in _TRUE
    return raw


def _bool_attr(element, name, default=False):
    raw = _attr(element, name)
    if raw is None:
        return default
    if raw in _TRUE:
        return True
    if raw in _FALSE:
        return False
    return default
