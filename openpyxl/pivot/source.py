# paper-xlsx: typed pivot source snapshots

"""Resolve table/range sources into typed records without evaluating formulas.

The snapshot layer may read an already-loaded workbook. It must not create
cells, dirty the ledger, open ZIP members, or invoke LibreOffice. Fully blank
rows and formula empty strings follow the provisional hypothesis that they
remain records; Excel transcripts in later PRs may replace that rule.
"""

from __future__ import annotations

import hashlib
import json
import math
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal

from openpyxl.errors import (
    AmbiguousTargetError,
    BoundaryViolationError,
    TargetNotFoundError,
)
from openpyxl.pivot.api_types import PivotSource
from openpyxl.utils import range_boundaries


KIND_BLANK = "blank"
KIND_TEXT = "text"
KIND_NUMBER = "number"
KIND_BOOLEAN = "boolean"
KIND_DATE = "date"
KIND_DATETIME = "datetime"
KIND_ERROR = "error"


@dataclass(frozen=True)
class PivotLimits:
    source_rows: int = 500000
    source_fields: int = 256
    source_cells: int = 10000000
    distinct_items_field: int = 100000
    distinct_items_total: int = 500000
    aggregate_states: int = 1000000
    output_cells: int = 1000000
    cache_xml_bytes: int = 256 * 1024 * 1024
    peak_memory: int = 512 * 1024 * 1024


DEFAULT_LIMITS = PivotLimits()


@dataclass(frozen=True)
class TypedValue:
    kind: str
    value: object = None

    def to_json(self):
        if self.kind == KIND_BLANK:
            return {"kind": self.kind}
        if self.kind in (KIND_DATE, KIND_DATETIME):
            return {"kind": self.kind, "value": self.value.isoformat()}
        return {"kind": self.kind, "value": self.value}


@dataclass(frozen=True)
class TypedRecord:
    values: tuple
    source_row: int


@dataclass(frozen=True)
class SourceSnapshot:
    source: PivotSource
    fields: tuple
    records: tuple
    shared_items: tuple
    formula_coordinates: tuple
    identity: str
    bounds: tuple
    warnings: tuple = ()
    calculation_provenance: object = None

    @property
    def field_index(self):
        return {name: index for index, name in enumerate(self.fields)}


def typed_value(raw):
    """Convert a Python/cell value to a typed pivot item.

    Boolean is checked before numbers. ``None`` and ``""`` are blank under the
    provisional empty-string hypothesis. NaN, infinity, and Excel errors refuse
    at snapshot time so they cannot be silently dropped.
    """
    if raw is None or raw == "":
        return TypedValue(KIND_BLANK)
    if isinstance(raw, bool):
        return TypedValue(KIND_BOOLEAN, raw)
    if isinstance(raw, datetime):
        return TypedValue(KIND_DATETIME, raw)
    if isinstance(raw, date):
        return TypedValue(KIND_DATE, raw)
    if isinstance(raw, Decimal):
        raw = float(raw)
    if isinstance(raw, (int, float)):
        number = float(raw)
        if not math.isfinite(number):
            raise BoundaryViolationError(
                "pivot sources cannot include NaN or infinity",
                kind="invalid-pivot-source",
                options=[str(raw)],
            )
        if isinstance(raw, int) or number.is_integer():
            return TypedValue(KIND_NUMBER, int(number))
        return TypedValue(KIND_NUMBER, number)
    if isinstance(raw, str):
        if raw.startswith("#") and raw.isupper():
            raise BoundaryViolationError(
                "pivot sources cannot include Excel error values",
                kind="invalid-pivot-source",
                options=[raw],
            )
        return TypedValue(KIND_TEXT, raw)
    raise BoundaryViolationError(
        "unsupported pivot source value type %s" % type(raw).__name__,
        kind="invalid-pivot-source",
        options=[repr(raw)],
    )


def snapshot_from_matrix(headers, rows, source=None, formula_coordinates=(),
                         limits=DEFAULT_LIMITS, generation=None):
    """Build a snapshot from already-extracted headers and row values."""
    fields = _validate_headers(headers, limits)
    if limits.source_rows and len(rows) > limits.source_rows:
        raise BoundaryViolationError(
            "pivot source has %s data rows; the limit is %s"
            % (len(rows), limits.source_rows),
            kind="pivot-source-too-large",
            options=[str(len(rows)), str(limits.source_rows)],
        )
    cells = len(rows) * len(fields)
    if cells > limits.source_cells:
        raise BoundaryViolationError(
            "pivot source has %s cells; the limit is %s"
            % (cells, limits.source_cells),
            kind="pivot-source-too-large",
            options=[str(cells), str(limits.source_cells)],
        )
    if not rows:
        raise BoundaryViolationError(
            "pivot source must contain at least one data row",
            kind="invalid-pivot-source",
        )
    records = []
    catalogs = [[] for _ in fields]
    catalog_sets = [set() for _ in fields]
    total_items = 0
    for offset, row in enumerate(rows):
        if len(row) != len(fields):
            raise BoundaryViolationError(
                "source row %s has %s values; expected %s"
                % (offset + 1, len(row), len(fields)),
                kind="invalid-pivot-source",
            )
        values = []
        for index, raw in enumerate(row):
            item = typed_value(raw)
            values.append(item)
            if item not in catalog_sets[index]:
                if len(catalogs[index]) >= limits.distinct_items_field:
                    raise BoundaryViolationError(
                        "field %r has more than %s distinct items"
                        % (fields[index], limits.distinct_items_field),
                        kind="pivot-cardinality-too-large",
                        options=[fields[index],
                                 str(limits.distinct_items_field)],
                    )
                if total_items >= limits.distinct_items_total:
                    raise BoundaryViolationError(
                        "pivot source exceeds %s distinct items"
                        % limits.distinct_items_total,
                        kind="pivot-cardinality-too-large",
                        options=[str(limits.distinct_items_total)],
                    )
                catalogs[index].append(item)
                catalog_sets[index].add(item)
                total_items += 1
        records.append(TypedRecord(tuple(values), offset + 1))
    source = source or PivotSource.range("Source", "A1:%s%s" % (
        _column_letter(len(fields)), len(rows) + 1))
    shared = tuple(tuple(items) for items in catalogs)
    identity = _source_identity(
        source, fields, records, formula_coordinates, generation)
    return SourceSnapshot(
        source=source,
        fields=tuple(fields),
        records=tuple(records),
        shared_items=shared,
        formula_coordinates=tuple(formula_coordinates),
        identity=identity,
        bounds=(source.sheet, 1, 1, len(fields), len(rows) + 1)
        if source.kind == "range" else (source.name, None, None, None, None),
    )


def snapshot_from_workbook(workbook, source, limits=DEFAULT_LIMITS):
    """Read a table or range through already-materialized cells only."""
    source = PivotSource.parse(source)
    table = None
    if source.kind == "table":
        worksheet, table = _find_table(workbook, source.name)
        ref = table.ref
    else:
        worksheet, ref = resolve_source_ref(workbook, source)
    min_col, min_row, max_col, max_row = _closed_range_boundaries(
        worksheet, ref)
    _refuse_merged_source(worksheet, min_col, min_row, max_col, max_row)
    data_max_row = max_row
    if table is not None:
        data_max_row -= _validate_table_source(
            worksheet, table, min_col, min_row, max_col, max_row)
    field_count = max_col - min_col + 1
    data_rows = data_max_row - min_row
    if field_count > limits.source_fields:
        raise BoundaryViolationError(
            "pivot source has %s fields; the limit is %s"
            % (field_count, limits.source_fields),
            kind="pivot-source-too-large",
            options=[str(field_count), str(limits.source_fields)],
        )
    if data_rows < 1:
        raise BoundaryViolationError(
            "pivot source must contain at least one data row",
            kind="invalid-pivot-source",
            anchor="%s!%s" % (worksheet.title, ref),
        )
    headers = []
    for column in range(min_col, max_col + 1):
        headers.append(_cell_value(worksheet, min_row, column))
    rows = []
    formulas = []
    cells = getattr(worksheet, "_cells", {})
    for row in range(min_row + 1, data_max_row + 1):
        values = []
        for column in range(min_col, max_col + 1):
            cell = cells.get((row, column))
            values.append(None if cell is None else cell.value)
            if cell is not None and _is_formula(cell):
                formulas.append("%s!%s" % (worksheet.title, cell.coordinate))
        rows.append(values)
    generation = _edit_generation(workbook)
    snapshot = snapshot_from_matrix(
        headers, rows, source=source if source.kind != "table"
        else PivotSource.table(source.name),
        formula_coordinates=formulas, limits=limits, generation=generation)
    return SourceSnapshot(
        source=snapshot.source,
        fields=snapshot.fields,
        records=snapshot.records,
        shared_items=snapshot.shared_items,
        formula_coordinates=snapshot.formula_coordinates,
        identity=snapshot.identity,
        bounds=(worksheet.title, min_col, min_row, max_col, max_row),
        warnings=snapshot.warnings,
    )


def resolve_source_ref(workbook, source):
    source = PivotSource.parse(source)
    if source.kind == "range":
        worksheet = _worksheet(workbook, source.sheet)
        return worksheet, source.ref
    if source.kind == "table":
        worksheet, table = _find_table(workbook, source.name)
        return worksheet, table.ref
    if source.kind == "defined-name":
        resolved = _defined_name_ref(workbook, source.name)
        if resolved is None:
            raise TargetNotFoundError(
                "defined name %r is not a sheet-qualified range"
                % source.name,
                kind="pivot-not-found",
                options=[source.name],
            )
        return resolve_source_ref(workbook, resolved)
    raise BoundaryViolationError(
        "unsupported pivot source kind %r" % source.kind,
        kind="unsupported-pivot-source",
    )


def _closed_range_boundaries(worksheet, ref):
    try:
        bounds = range_boundaries(ref)
    except (TypeError, ValueError) as exc:
        raise BoundaryViolationError(
            "pivot source %r is not a rectangular A1 range" % ref,
            kind="invalid-pivot-source",
            anchor="%s!%s" % (worksheet.title, ref),
        ) from exc
    if None in bounds:
        raise BoundaryViolationError(
            "pivot sources must have finite row and column bounds",
            kind="unsupported-pivot-source",
            anchor="%s!%s" % (worksheet.title, ref),
            options=[ref],
        )
    return bounds


def _refuse_merged_source(worksheet, min_col, min_row, max_col, max_row):
    for merged in worksheet.merged_cells.ranges:
        if not (
            merged.max_col < min_col or max_col < merged.min_col
            or merged.max_row < min_row or max_row < merged.min_row
        ):
            raise BoundaryViolationError(
                "pivot source intersects merged cells at %s" % merged.coord,
                kind="unsupported-pivot-source",
                anchor="%s!%s" % (worksheet.title, merged.coord),
                options=[merged.coord],
            )


def _validate_table_source(worksheet, table, min_col, min_row,
                           max_col, max_row):
    name = table.displayName or table.name
    if table.tableType not in (None, "worksheet") \
            or table.connectionId is not None:
        raise BoundaryViolationError(
            "table %r is query-backed or externally connected" % name,
            kind="unsupported-pivot-source",
            options=[name],
        )
    header_rows = table.headerRowCount
    if header_rows is None:
        header_rows = 1
    if isinstance(header_rows, bool) or header_rows != 1:
        raise BoundaryViolationError(
            "table %r must have exactly one header row" % name,
            kind="unsupported-pivot-source",
            anchor="%s!%s" % (worksheet.title, table.ref),
            options=[str(header_rows)],
        )
    totals_count = table.totalsRowCount
    if isinstance(totals_count, bool) or totals_count not in (None, 0, 1):
        raise BoundaryViolationError(
            "table %r has unsupported totals-row metadata" % name,
            kind="unsupported-pivot-source",
            anchor="%s!%s" % (worksheet.title, table.ref),
            options=[str(totals_count)],
        )
    totals_rows = int(totals_count == 1 or bool(table.totalsRowShown))
    if max_row - min_row + 1 < header_rows + totals_rows + 1:
        raise BoundaryViolationError(
            "table %r must contain at least one data row" % name,
            kind="invalid-pivot-source",
            anchor="%s!%s" % (worksheet.title, table.ref),
        )
    columns = list(table.tableColumns)
    width = max_col - min_col + 1
    if columns and len(columns) != width:
        raise BoundaryViolationError(
            "table %r column metadata does not match its range" % name,
            kind="unsupported-pivot-source",
            anchor="%s!%s" % (worksheet.title, table.ref),
            options=[str(len(columns)), str(width)],
        )
    if columns:
        headers = [
            _cell_value(worksheet, min_row, column)
            for column in range(min_col, max_col + 1)
        ]
        declared = [column.name for column in columns]
        if declared != headers:
            raise BoundaryViolationError(
                "table %r column metadata does not match its headers" % name,
                kind="unsupported-pivot-source",
                anchor="%s!%s" % (worksheet.title, table.ref),
            )
    return totals_rows


def _validate_headers(headers, limits):
    if not headers:
        raise BoundaryViolationError(
            "pivot source must contain at least one field",
            kind="invalid-pivot-source",
        )
    if len(headers) > limits.source_fields:
        raise BoundaryViolationError(
            "pivot source has %s fields; the limit is %s"
            % (len(headers), limits.source_fields),
            kind="pivot-source-too-large",
            options=[str(len(headers)), str(limits.source_fields)],
        )
    names = []
    folded = []
    for header in headers:
        if not isinstance(header, str) or not header:
            raise BoundaryViolationError(
                "pivot source headers must be nonempty strings",
                kind="invalid-pivot-source",
                options=[repr(header)],
            )
        key = header.casefold()
        if key in folded:
            raise BoundaryViolationError(
                "pivot source headers must be unique under case-insensitive "
                "comparison",
                kind="invalid-pivot-source",
                options=[header],
            )
        names.append(header)
        folded.append(key)
    return names


def _cell_value(worksheet, row, column):
    cell = getattr(worksheet, "_cells", {}).get((row, column))
    if cell is None:
        return None
    return cell.value


def _is_formula(cell):
    if getattr(cell, "data_type", None) == "f":
        return True
    value = getattr(cell, "value", None)
    return isinstance(value, str) and value.startswith("=")


def _worksheet(workbook, title):
    for worksheet in workbook.worksheets:
        if worksheet.title == title:
            return worksheet
    raise TargetNotFoundError(
        "worksheet %r was not found" % title,
        kind="pivot-not-found",
        options=[ws.title for ws in workbook.worksheets],
    )


def _find_table(workbook, name):
    folded = name.casefold()
    hits = []
    for worksheet in workbook.worksheets:
        tables = getattr(worksheet, "tables", None) or {}
        # TableList.items() returns (name, ref); the Table objects live
        # on the dict subclass itself.
        for key, table in dict.items(tables):
            declared = (
                key,
                getattr(table, "name", None),
                getattr(table, "displayName", None),
            )
            if any(isinstance(item, str) and item.casefold() == folded
                   for item in declared):
                hits.append((worksheet, table))
    if not hits:
        raise TargetNotFoundError(
            "table %r was not found" % name,
            kind="pivot-not-found",
            options=[name],
        )
    unique = {(id(table), table.ref) for _ws, table in hits}
    if len(unique) > 1:
        raise AmbiguousTargetError(
            "table name %r is not unique" % name,
            kind="ambiguous-pivot",
            options=[name],
        )
    return hits[0]


def _defined_name_ref(workbook, name):
    folded = name.casefold()
    holders = [getattr(workbook, "defined_names", {})]
    holders.extend(ws.defined_names for ws in workbook.worksheets)
    for holder in holders:
        for key, defn in (holder or {}).items():
            declared = getattr(defn, "name", None)
            if not (
                (isinstance(key, str) and key.casefold() == folded)
                or (isinstance(declared, str) and declared.casefold() == folded)
            ):
                continue
            text = getattr(defn, "attr_text", None) or getattr(defn, "value", None)
            if not isinstance(text, str):
                return None
            try:
                return PivotSource.parse(text)
            except (TypeError, ValueError):
                return None
    return None


def _edit_generation(workbook):
    ledger = getattr(workbook, "_paper_ledger", None)
    if ledger is None:
        return None
    dirty = []
    for worksheet in workbook.worksheets:
        dirty.extend(
            (worksheet.title, row, column)
            for row, column in ledger.dirty_coordinates(worksheet)
        )
    dirty.sort()
    return tuple(dirty)


def _source_identity(source, fields, records, formulas, generation):
    payload = {
        "source": source.to_dict(),
        "fields": list(fields),
        "records": [
            [item.to_json() for item in record.values]
            for record in records
        ],
        "formulas": list(formulas),
        "generation": list(generation) if generation else None,
    }
    encoded = json.dumps(payload, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def _column_letter(index):
    from openpyxl.utils import get_column_letter
    return get_column_letter(index)
