# paper-xlsx: foreign pivot raw-cell ownership proof

"""Prove a selected foreign output footprint from persisted cache records.

``location.ref`` is never the complete footprint. Report-filter cells sit
above the body with a spacer. Managed sparse ownership is not this proof.
"""

from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass

from openpyxl.pivot.graph import _attr, _local, _parse_xml
from openpyxl.pivot.layout import ROLE_FILTER
from openpyxl.pivot.qualify import (
    QualificationReason,
    _snapshot_from_cache_package,
)
from openpyxl.utils.cell import coordinate_from_string, range_boundaries
from openpyxl.utils import column_index_from_string


@dataclass(frozen=True)
class ForeignCellPayload:
    row: int
    column: int
    present: bool
    value: object = None
    value_type: str | None = None
    formula: str | None = None
    style_id: str | None = None
    empty_string: bool = False

    def to_dict(self):
        return {
            "row": self.row,
            "column": self.column,
            "present": self.present,
            "value": self.value,
            "value_type": self.value_type,
            "formula": self.formula,
            "style_id": self.style_id,
            "empty_string": self.empty_string,
        }


@dataclass(frozen=True)
class ForeignOutputOwnership:
    location_ref: str
    body_coordinates: tuple
    filter_coordinates: tuple
    payloads: object

    def to_dict(self):
        return {
            "location_ref": self.location_ref,
            "body_coordinates": [
                [row, column] for row, column in self.body_coordinates
            ],
            "filter_coordinates": [
                [row, column] for row, column in self.filter_coordinates
            ],
            "payload_count": len(self.payloads),
        }


def prove_foreign_output(workbook, node, projection):
    """Return ``(ownership, reasons)``. Ownership is None when unproved."""
    reasons = []
    if projection.spec is None or not node.output_range:
        return None, (_reason(
            "foreign-output-unproved",
            part=node.identity.pivot_part,
            output_range=node.output_range,
        ),)
    if not node.cache_definition_part or not node.cache_records_part:
        return None, (_reason(
            "foreign-cache-records-unavailable",
            part=node.identity.pivot_part,
        ),)
    package = getattr(workbook, "_paper_source", None)
    if not package:
        return None, (_reason(
            "foreign-output-unproved",
            part=node.identity.pivot_part,
        ),)
    try:
        from openpyxl.pivot.plan import plan_pivot

        snapshot = _snapshot_from_cache_package(package, node, projection)
        plan = plan_pivot(projection.spec, snapshot)
    except Exception:
        return None, (_reason(
            "foreign-output-unproved",
            part=node.identity.pivot_part,
            detail="cache-reconstruction-failed",
        ),)
    if plan.output.ref != node.output_range:
        return None, (_reason(
            "foreign-output-unproved",
            part=node.identity.pivot_part,
            expected=plan.output.ref,
            actual=node.output_range,
        ),)
    worksheet = _current_worksheet(workbook, node)
    if worksheet is None:
        return None, (_reason(
            "foreign-output-unproved",
            part=node.identity.pivot_part,
            sheet=node.sheet_title,
        ),)
    raw_cells = _raw_sheet_cells(package, node.identity.worksheet_part)
    if raw_cells is None:
        return None, (_reason(
            "foreign-output-unproved",
            part=node.identity.worksheet_part,
        ),)
    filter_coords = tuple(
        (cell.row, cell.column)
        for cell in plan.output.cells
        if cell.role == ROLE_FILTER
    )
    body_coords = tuple(
        (cell.row, cell.column)
        for cell in plan.output.cells
        if cell.role != ROLE_FILTER
    )
    footprint = _footprint(node.output_range, plan.output.cells)
    live = getattr(worksheet, "_cells", {})
    dirty = set()
    ledger = getattr(workbook, "_paper_ledger", None)
    if ledger is not None:
        dirty = set(ledger.dirty_coordinates(worksheet))
    payloads = {}
    explained = {
        (cell.row, cell.column): cell
        for cell in plan.output.cells
    }
    for row, column in footprint:
        if (row, column) in dirty:
            reasons.append(_reason(
                "foreign-output-unproved",
                part=node.identity.pivot_part,
                coordinate=_coord(column, row),
                detail="dirty-output-cell",
            ))
            return None, tuple(reasons)
        planned = explained.get((row, column))
        raw = raw_cells.get((row, column))
        existing = live.get((row, column))
        if planned is not None and planned.value is not None:
            mismatch = _semantic_mismatch(planned, existing, raw)
            if mismatch is not None:
                reasons.append(_reason(
                    "foreign-output-unproved",
                    part=node.identity.pivot_part,
                    coordinate=_coord(column, row),
                    detail=mismatch,
                ))
                return None, tuple(reasons)
        elif raw is not None:
            reasons.append(_reason(
                "foreign-output-unproved",
                part=node.identity.pivot_part,
                coordinate=_coord(column, row),
                detail="unexplained-cell-node",
            ))
            return None, tuple(reasons)
        if existing is not None:
            if getattr(existing, "_comment", None) is not None \
                    or getattr(existing, "comment", None) is not None:
                reasons.append(_reason(
                    "foreign-output-unproved",
                    part=node.identity.pivot_part,
                    coordinate=_coord(column, row),
                    detail="comment",
                ))
                return None, tuple(reasons)
            if getattr(existing, "_hyperlink", None) is not None:
                reasons.append(_reason(
                    "foreign-output-unproved",
                    part=node.identity.pivot_part,
                    coordinate=_coord(column, row),
                    detail="hyperlink",
                ))
                return None, tuple(reasons)
        payloads[(row, column)] = ForeignCellPayload(
            row=row,
            column=column,
            present=raw is not None,
            value=None if raw is None else raw.get("value"),
            value_type=None if raw is None else raw.get("type"),
            formula=None if raw is None else raw.get("formula"),
            style_id=None if raw is None else raw.get("style"),
            empty_string=bool(raw and raw.get("empty_string")),
        )
    if _footprint_intersects_merge(worksheet, footprint):
        return None, (_reason(
            "foreign-output-unproved",
            part=node.identity.pivot_part,
            detail="merged-output-cell",
        ),)
    return ForeignOutputOwnership(
        location_ref=node.output_range,
        body_coordinates=body_coords,
        filter_coordinates=filter_coords,
        payloads=payloads,
    ), ()


def _semantic_mismatch(planned, existing, raw):
    actual = None if existing is None else existing.value
    if existing is None or existing.data_type == "f":
        return "formula-or-missing-cell"
    if type(actual) is not type(planned.value) or actual != planned.value:
        return "value-mismatch"
    if planned.number_format is not None \
            and existing.number_format != planned.number_format:
        return "number-format-mismatch"
    if raw is None:
        return "missing-raw-cell"
    if raw.get("formula"):
        return "formula-cell"
    return None


def _footprint(location_ref, cells):
    coords = {(cell.row, cell.column) for cell in cells}
    try:
        min_col, min_row, max_col, max_row = range_boundaries(location_ref)
    except (TypeError, ValueError):
        return tuple(sorted(coords))
    for row in range(min_row, max_row + 1):
        for column in range(min_col, max_col + 1):
            coords.add((row, column))
    return tuple(sorted(coords))


def _raw_sheet_cells(package, worksheet_part):
    try:
        with zipfile.ZipFile(io.BytesIO(package)) as archive:
            if worksheet_part not in set(archive.namelist()):
                return None
            root = _parse_xml(archive.read(worksheet_part))
    except (KeyError, OSError, ValueError, zipfile.BadZipFile):
        return None
    if root is None:
        return None
    cells = {}
    for element in root.iter():
        if _local(element.tag) != "c":
            continue
        ref = _attr(element, "r")
        if not ref:
            continue
        try:
            column_letter, row = coordinate_from_string(ref)
            column = column_index_from_string(column_letter)
        except (TypeError, ValueError):
            continue
        value_type = _attr(element, "t")
        style_id = _attr(element, "s")
        formula = None
        value = None
        empty_string = False
        for child in list(element):
            tag = _local(child.tag)
            if tag == "f":
                formula = (child.text or "")
            elif tag == "v":
                value = child.text
            elif tag == "is":
                texts = [
                    (item.text or "")
                    for item in child.iter()
                    if _local(item.tag) == "t"
                ]
                value = "".join(texts)
                empty_string = value == ""
        if value_type == "inlineStr" and value == "":
            empty_string = True
        cells[(row, column)] = {
            "type": value_type,
            "style": style_id,
            "formula": formula,
            "value": value,
            "empty_string": empty_string,
        }
    return cells


def _current_worksheet(workbook, node):
    sheet_title = node.sheet_title
    ledger = getattr(workbook, "_paper_ledger", None)
    if ledger is not None:
        for item, original in getattr(ledger, "renames", {}).items():
            if original == sheet_title:
                sheet_title = item.title
                break
    for worksheet in workbook.worksheets:
        if worksheet.title == sheet_title:
            return worksheet
    return None


def _footprint_intersects_merge(worksheet, footprint):
    coords = set(footprint)
    for merged in worksheet.merged_cells.ranges:
        for row in range(merged.min_row, merged.max_row + 1):
            for column in range(merged.min_col, merged.max_col + 1):
                if (row, column) in coords:
                    return True
    return False


def _coord(column, row):
    from openpyxl.utils import get_column_letter
    return "%s%s" % (get_column_letter(column), row)


def _reason(code, **context):
    items = tuple(sorted(
        (key, value) for key, value in context.items() if value is not None))
    return QualificationReason(None, code, items)
