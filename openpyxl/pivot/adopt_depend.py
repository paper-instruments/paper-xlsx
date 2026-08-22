# paper-xlsx: targeted foreign-pivot dependency scan

"""Find objects that depend on one selected pivot or its cache.

Ordinary ``Worksheet.pivots`` inspection stays metadata-bounded. This scan
runs only during explicit ``qualify_adoption()``.
"""

from __future__ import annotations

import io
import zipfile
import re

from openpyxl.pivot.qualify import QualificationReason
from openpyxl.utils.cell import range_boundaries
from openpyxl.xml.constants import REL_NS


PIVOT_TABLE_REL = REL_NS + "/pivotTable"
PIVOT_CACHE_REL = REL_NS + "/pivotCacheDefinition"
PIVOT_RECORDS_REL = REL_NS + "/pivotCacheRecords"
_ALLOWED_SELECTED_RELS = frozenset((
    PIVOT_TABLE_REL, PIVOT_CACHE_REL, PIVOT_RECORDS_REL,
))
_SLICER_MARKERS = (
    "slicerCache", "slicerCacheDefinition", "timelineCache",
    "timelineCacheDefinition",
)
_GETPIVOTDATA = re.compile(r"GETPIVOTDATA\s*\(", re.IGNORECASE)
_DYNAMIC_FUNCS = re.compile(
    r"\b(INDIRECT|OFFSET|RTD|CUBEVALUE|CUBEMEMBER|WEBSERVICE)\s*\(",
    re.IGNORECASE,
)
_EXCLUDED_SOURCE_FUNCS = re.compile(
    r"\b(INDIRECT|OFFSET|RTD|CUBEVALUE|CUBEMEMBER|WEBSERVICE|CALL|REGISTER)\s*\(",
    re.IGNORECASE,
)
_CELL_REF = re.compile(r'(?<![A-Za-z])(\$?[A-Za-z]{1,3}\$?\d+)\b')


def scan_selected_dependencies(workbook, node, cache, graph, projection,
                               ownership):
    """Return ``(blocker_reasons, operation_constraints, dependents)``."""
    blockers = []
    constraints = []
    dependents = []
    blockers.extend(_relationship_blockers(node, cache, graph))
    package = getattr(workbook, "_paper_source", None)
    if package:
        chart_hits = _scan_chart_parts(package, node)
        slicer_hits = _scan_slicer_parts(package, node, cache)
        dependents.extend(chart_hits)
        dependents.extend(slicer_hits)
        if chart_hits:
            blockers.append(_reason(
                "foreign-dependent-object",
                part=node.identity.pivot_part,
                kind="pivot-chart",
                objects=",".join(chart_hits),
            ))
        if slicer_hits:
            blockers.append(_reason(
                "foreign-dependent-object",
                part=node.identity.pivot_part,
                kind="slicer-or-timeline",
                objects=",".join(slicer_hits),
            ))
    if projection.spec is not None:
        constraints.extend(_workbook_consumer_constraints(
            workbook, node, projection, ownership))
    return tuple(blockers), tuple(constraints), tuple(dependents)


def source_requires_calculation(workbook, projection):
    """Detect formula-backed sources without invoking LibreOffice."""
    source = None if projection is None else projection.source
    if source is None or source.kind not in ("table", "range"):
        return False, None, ()
    try:
        from openpyxl.pivot.source import snapshot_from_workbook

        snapshot = snapshot_from_workbook(workbook, source)
    except Exception:
        return False, None, ()
    if not snapshot.formula_coordinates:
        return False, None, ()
    excluded = []
    cells = _source_formula_texts(workbook, snapshot)
    for address, formula in cells:
        if formula and _EXCLUDED_SOURCE_FUNCS.search(formula):
            excluded.append(_reason(
                "unsupported-pivot-source",
                coordinate=address,
                detail="statically-excluded-formula",
            ))
    return True, "libreoffice", tuple(excluded)


def _relationship_blockers(node, cache, graph):
    reasons = []
    incoming = graph.incoming_relationships
    selected = {node.identity.pivot_part}
    if cache is not None:
        if cache.definition_part:
            selected.add(cache.definition_part)
        if cache.records_part:
            selected.add(cache.records_part)
    for part in selected:
        for rel in incoming.get(part, ()):
            if rel.rel_type not in _ALLOWED_SELECTED_RELS:
                reasons.append(_reason(
                    "foreign-dependent-object",
                    part=part,
                    relationship_type=rel.rel_type,
                    owner=rel.owner_part,
                ))
    return reasons


def _scan_chart_parts(package, node):
    hits = []
    name = node.identity.name or ""
    try:
        with zipfile.ZipFile(io.BytesIO(package)) as archive:
            for part in archive.namelist():
                if not part.startswith("xl/charts/") or not part.endswith(".xml"):
                    continue
                try:
                    payload = archive.read(part)
                except KeyError:
                    continue
                text = payload.decode("utf-8", "replace")
                if "pivotSource" in text or "pivotCache" in text:
                    if name and name in text:
                        hits.append(part)
                        continue
                    hits.append(part)
                    continue
                if name and name in text and b"c:tx" in payload:
                    # lexical pivot-name reference without a package edge
                    if ("'%s'" % name) in text or '"%s"' % name in text:
                        hits.append(part)
    except (OSError, ValueError, zipfile.BadZipFile):
        return hits
    return hits


def _scan_slicer_parts(package, node, cache):
    hits = []
    markers = (
        node.identity.pivot_part,
        node.identity.name or "",
        None if cache is None else cache.definition_part,
        None if cache is None else cache.cache_id,
    )
    try:
        with zipfile.ZipFile(io.BytesIO(package)) as archive:
            for part in archive.namelist():
                lowered = part.lower()
                if "slicer" not in lowered and "timeline" not in lowered:
                    continue
                try:
                    payload = archive.read(part)
                except KeyError:
                    continue
                text = payload.decode("utf-8", "replace")
                if any(marker and str(marker) in text for marker in markers):
                    hits.append(part)
                    continue
                if any(token in text for token in _SLICER_MARKERS):
                    hits.append(part)
    except (OSError, ValueError, zipfile.BadZipFile):
        return hits
    return hits


def scan_consumer_constraints(workbook, name, sheet, footprint):
    """Find GETPIVOTDATA/formula/name consumers. Not used during enumeration."""
    constraints = []
    footprint = set(footprint or ())
    name = name or ""
    for worksheet in workbook.worksheets:
        for cell in getattr(worksheet, "_cells", {}).values():
            formula = _formula_text(cell)
            if not formula:
                continue
            address = "%s!%s" % (worksheet.title, cell.coordinate)
            if _GETPIVOTDATA.search(formula) and name and name in formula:
                constraints.append(_capability_reason(
                    "can_rename", "pivot-dependent-reference",
                    kind="getpivotdata", coordinate=address, name=name,
                ))
                constraints.append(_capability_reason(
                    "can_delete", "pivot-dependent-reference",
                    kind="getpivotdata", coordinate=address, name=name,
                ))
            if _DYNAMIC_FUNCS.search(formula) and _formula_mentions_sheet(
                    formula, sheet):
                constraints.append(_capability_reason(
                    "can_move", "pivot-dependent-reference",
                    kind="unresolved-dynamic-reference",
                    coordinate=address,
                ))
            if worksheet.title == sheet and (
                    cell.row, cell.column) not in footprint:
                if _formula_hits_footprint(formula, sheet, footprint):
                    constraints.append(_capability_reason(
                        "can_move", "pivot-dependent-reference",
                        kind="formula", coordinate=address,
                    ))
                    constraints.append(_capability_reason(
                        "can_delete", "pivot-dependent-reference",
                        kind="formula", coordinate=address,
                    ))
    defined = getattr(workbook, "defined_names", None)
    values = []
    if defined is not None:
        if hasattr(defined, "values"):
            values = list(defined.values())
        elif hasattr(defined, "definedName"):
            values = list(defined.definedName)
    for item in values:
        attr = getattr(item, "attr_text", None) or getattr(item, "value", None)
        item_name = getattr(item, "name", None)
        if not attr:
            continue
        if name and name in str(attr):
            constraints.append(_capability_reason(
                "can_rename", "pivot-dependent-reference",
                kind="defined-name", name=item_name,
            ))
        if _formula_hits_footprint(str(attr), sheet, footprint):
            constraints.append(_capability_reason(
                "can_move", "pivot-dependent-reference",
                kind="defined-name", name=item_name,
            ))
    return constraints


def _workbook_consumer_constraints(workbook, node, projection, ownership):
    name = node.identity.name or ""
    sheet = node.sheet_title
    footprint = set()
    if ownership is not None:
        footprint.update(ownership.body_coordinates)
        footprint.update(ownership.filter_coordinates)
    elif node.output_range:
        try:
            min_col, min_row, max_col, max_row = range_boundaries(
                node.output_range)
        except (TypeError, ValueError):
            min_col = min_row = max_col = max_row = None
        if min_col is not None:
            for row in range(min_row, max_row + 1):
                for column in range(min_col, max_col + 1):
                    footprint.add((row, column))
    return scan_consumer_constraints(workbook, name, sheet, footprint)


def _formula_text(cell):
    if getattr(cell, "data_type", None) == "f":
        value = cell.value
        return value if isinstance(value, str) else str(value or "")
    value = getattr(cell, "value", None)
    if isinstance(value, str) and value.startswith("="):
        return value
    return None


def _formula_mentions_sheet(formula, sheet):
    if not sheet or not formula:
        return False
    return sheet in formula


def _formula_hits_footprint(formula, sheet, footprint):
    if not formula or not footprint:
        return False
    from openpyxl.utils import column_index_from_string
    from openpyxl.utils.cell import coordinate_from_string

    for match in _CELL_REF.finditer(formula):
        try:
            column_letter, row = coordinate_from_string(match.group(1))
            column = column_index_from_string(column_letter)
        except (TypeError, ValueError):
            continue
        if (row, column) in footprint:
            return True
    return False


def _source_formula_texts(workbook, snapshot):
    texts = []
    for address in snapshot.formula_coordinates:
        if "!" not in address:
            continue
        sheet, coord = address.split("!", 1)
        worksheet = None
        for item in workbook.worksheets:
            if item.title == sheet:
                worksheet = item
                break
        if worksheet is None:
            continue
        cell = worksheet[coord]
        texts.append((address, _formula_text(cell)))
    return texts


def _capability_reason(capability, code, **context):
    items = tuple(sorted(
        (key, value) for key, value in context.items() if value is not None))
    return QualificationReason(capability, code, items)


def _reason(code, **context):
    items = tuple(sorted(
        (key, value) for key, value in context.items() if value is not None))
    return QualificationReason(None, code, items)
