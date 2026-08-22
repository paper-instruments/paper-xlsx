# paper-xlsx: targeted pivot cache refresh

"""Resolve pivot names and patch only cache refresh metadata."""

import io
import warnings
import zipfile
from collections import deque
from xml.etree.ElementTree import ParseError, fromstring

from openpyxl.errors import (
    AmbiguousTargetError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries
from openpyxl.xml.constants import MAX_COLUMN, MAX_ROW, SHEET_MAIN_NS

from . import crosspart


def _cache_root(payload):
    """Validate a pivot-cache root, including legal prefixed spelling."""
    try:
        element = fromstring(payload)
    except (ParseError, ValueError, TypeError) as exc:
        raise UnsupportedStructureError(
            "pivot cache definition XML is malformed",
            kind="unsupported-pivot-cache") from exc
    expected = "{{{0}}}pivotCacheDefinition".format(SHEET_MAIN_NS)
    if element.tag != expected:
        raise UnsupportedStructureError(
            "pivot cache definition has an unexpected root element",
            kind="unsupported-pivot-cache")
    return crosspart.scan_small(
        payload, "pivotCacheDefinition", max_depth=1,
        allow_prefixed_root=True)


def _index(wb):
    """Return ``(qualified-name map, cache parts)`` from source bytes."""
    from openpyxl.pivot.graph import load_workbook_pivot_graph

    graph = load_workbook_pivot_graph(wb)
    graph.raise_for_refresh_index()
    ledger = getattr(wb, "_paper_ledger", None)
    current_by_original = {
        original: ws.title
        for ws, original in getattr(ledger, "renames", {}).items()
    }
    qualified = graph.qualified_name_map(current_by_original)
    cache_parts = set(graph.registered_cache_parts)
    operations = getattr(ledger, "pivot_operations", {})
    for operation in operations.values():
        if not getattr(operation, "replace_existing", False):
            continue
        allocation = operation.allocation
        for node in graph.pivots:
            if node.identity.pivot_part != allocation.pivot_part:
                continue
            current_title = current_by_original.get(
                node.sheet_title, node.sheet_title)
            existing = qualified.get(node.identity.name, [])
            qualified[node.identity.name] = [
                item for item in existing
                if item != (current_title, node.cache_definition_part)
            ]
            if not qualified[node.identity.name]:
                qualified.pop(node.identity.name, None)
            break
        if operation.kind == "delete" or getattr(operation, "noop", False):
            cache_parts.discard(allocation.cache_part)
            continue
        worksheet = getattr(operation, "worksheet", None)
        sheet_title = worksheet.title if worksheet is not None \
            else operation.sheet
        qualified.setdefault(operation.name, []).append(
            (sheet_title, allocation.cache_part))
        cache_parts.add(allocation.cache_part)
    qualified = {
        name: list(dict.fromkeys(items))
        for name, items in qualified.items()
    }
    referenced = {
        part for items in qualified.values() for _sheet, part in items
    }
    for operation in operations.values():
        original = getattr(operation, "original_cache_part", None)
        if original and original not in referenced:
            cache_parts.discard(original)
    return qualified, sorted(cache_parts)


def _worksheet(wb, title):
    matches = [ws for ws in wb.worksheets
               if ws.title.casefold() == title.casefold()]
    return matches[0] if len(matches) == 1 else None


def _bounds(ref):
    try:
        bounds = range_boundaries(ref)
    except (TypeError, ValueError):
        return None
    min_col, min_row, max_col, max_row = bounds
    if min_col is not None and not 1 <= min_col <= MAX_COLUMN:
        return None
    if max_col is not None and not 1 <= max_col <= MAX_COLUMN:
        return None
    if min_row is not None and not 1 <= min_row <= MAX_ROW:
        return None
    if max_row is not None and not 1 <= max_row <= MAX_ROW:
        return None
    if min_col is not None and max_col is not None and min_col > max_col:
        return None
    if min_row is not None and max_row is not None and min_row > max_row:
        return None
    return bounds


def _named_source(wb, name):
    """Return one exact named source and a fingerprint of every candidate."""
    folded = name.casefold()
    candidates = []
    for ws in wb.worksheets:
        for key in ws.tables:
            table = ws.tables[key]
            declared = (getattr(table, "name", None),
                        getattr(table, "displayName", None))
            key_matches = isinstance(key, str) and key.casefold() == folded
            declared_matches = any(
                isinstance(item, str) and item.casefold() == folded
                for item in declared)
            if not (key_matches or declared_matches):
                continue
            state = (
                "table", ws.title, key, declared[0], declared[1], table.ref)
            resolved = None
            bounds = _bounds(table.ref)
            if declared_matches and bounds is not None:
                resolved = (ws, bounds)
            candidates.append((state, resolved))

    holders = [(None, wb.defined_names)] + [
        (ws, ws.defined_names) for ws in wb.worksheets]
    for owner, definitions in holders:
        for key, defn in definitions.items():
            declared = getattr(defn, "name", None)
            key_matches = isinstance(key, str) and key.casefold() == folded
            declared_matches = isinstance(declared, str) \
                and declared.casefold() == folded
            if not (key_matches or declared_matches):
                continue
            value = getattr(defn, "attr_text", None)
            state = (
                "defined-name", owner.title if owner is not None else "",
                key, declared, value)
            resolved = None
            if declared_matches and isinstance(value, str):
                if owner is not None and "!" not in value:
                    bounds = _bounds(value)
                    if bounds is not None:
                        resolved = (owner, bounds)
                else:
                    try:
                        destinations = list(defn.destinations)
                    except (AttributeError, TypeError, ValueError):
                        destinations = []
                    if len(destinations) == 1:
                        title, ref = destinations[0]
                        ws = _worksheet(wb, title)
                        bounds = _bounds(ref)
                        if ws is not None and bounds is not None:
                            resolved = (ws, bounds)
            candidates.append((state, resolved))

    state = tuple(sorted(
        (candidate for candidate, _resolved in candidates), key=repr))
    if len(candidates) == 1 and candidates[0][1] is not None:
        resolved = candidates[0][1]
        # Compare the effective binding, not harmless spelling differences
        # in a static name. The candidate kind prevents a defined name and a
        # table that happen to cover the same cells from being conflated.
        return resolved, ("resolved", candidates[0][0][0],
                          resolved[0], resolved[1])
    return None, ("unresolved", state)


def _cache_source(wb, payload):
    """Return ``(worksheet, bounds, label, binding)`` for a local source.

    ``None`` means the cache is external and cannot be affected by local cell
    edits. A tuple of ``(None, None, label)`` is an unresolved local source and
    therefore intersects conservatively.
    """
    try:
        root = fromstring(payload)
    except (ParseError, TypeError, ValueError):
        return None, None, "unresolved pivot source", ("malformed",)
    namespace = "{{{0}}}".format(SHEET_MAIN_NS)
    sources = root.findall(namespace + "cacheSource")
    if len(sources) != 1:
        return None, None, "unresolved pivot source", ("source-count",)
    source = sources[0]
    source_type = source.attrib.get("type")
    if source_type == "external":
        return None
    if source_type != "worksheet":
        label = "unsupported {0!r} pivot source".format(source_type)
        return None, None, label, ("unsupported", source_type)
    worksheet_sources = source.findall(namespace + "worksheetSource")
    if len(worksheet_sources) != 1:
        return (None, None, "unresolved worksheet pivot source",
                ("worksheet-source-count", len(worksheet_sources)))
    worksheet_source = worksheet_sources[0]
    ref = worksheet_source.attrib.get("ref")
    title = worksheet_source.attrib.get("sheet")
    name = worksheet_source.attrib.get("name")
    if ref is not None and title is not None and name is None:
        ws = _worksheet(wb, title)
        bounds = _bounds(ref)
        if ws is not None and bounds is not None:
            binding = ("direct", ws, bounds)
            return ws, bounds, "{0}!{1}".format(ws.title, ref), binding
        return (None, None, "unresolved worksheet pivot source",
                ("direct", title, ref, None, bounds))
    elif name is not None and ref is None:
        resolved, candidates = _named_source(wb, name)
        binding = ("named", name, candidates)
        if resolved is not None:
            ws, bounds = resolved
            return ws, bounds, name, binding
        return None, None, "unresolved worksheet pivot source", binding
    return (None, None, "unresolved worksheet pivot source",
            ("worksheet-source", title, ref, name))


def snapshot_sources(wb):
    """Capture semantic local-source bindings when preservation arms."""
    _index_map, cache_parts = _index(wb)
    snapshots = {}
    with zipfile.ZipFile(io.BytesIO(wb._paper_source)) as zin:
        for part in cache_parts:
            source = _cache_source(wb, zin.read(part))
            snapshots[part] = None if source is None else source[3]
    return snapshots


def parts_referencing_sheet(wb, sheet_title):
    """Return local pivot-cache parts that may depend on one worksheet."""
    target = _worksheet(wb, sheet_title)
    if target is None:
        return []
    _index_map, cache_parts = _index(wb)
    hits = []
    with zipfile.ZipFile(io.BytesIO(wb._paper_source)) as zin:
        for part in cache_parts:
            source = _cache_source(wb, zin.read(part))
            if source is None:  # external source
                continue
            ws, _bounds_value, _label, binding = source
            if ws is target:
                hits.append(part)
                continue
            if ws is not None:
                continue
            # A direct source still names its sheet even when its range is
            # invalid. Other unresolved local sources cannot be proven
            # independent of any particular sheet, so a structural edit must
            # refuse rather than strand an unmodeled reference.
            if binding[:1] == ("direct",):
                title = binding[1]
                if isinstance(title, str) \
                        and title.casefold() == sheet_title.casefold():
                    hits.append(part)
            else:
                hits.append(part)
    return sorted(set(hits))


def _address_key(address):
    title, separator, coordinate = address.rpartition("!")
    if not separator:
        return None
    if title.startswith("'") and title.endswith("'"):
        title = title[1:-1].replace("''", "'")
    try:
        row, column = coordinate_to_tuple(coordinate)
    except (TypeError, ValueError):
        return None
    return title, row, column


def _bounds_hit(sheet, bounds, cells):
    min_col, min_row, max_col, max_row = _closed_bounds(bounds)
    folded = sheet.casefold()
    return any(
        title.casefold() == folded
        and min_row <= row <= max_row
        and min_col <= column <= max_col
        for title, row, column in cells)


def _closed_bounds(bounds):
    min_col, min_row, max_col, max_row = bounds
    return (
        1 if min_col is None else min_col,
        1 if min_row is None else min_row,
        MAX_COLUMN if max_col is None else max_col,
        MAX_ROW if max_row is None else max_row,
    )


def _ranges_hit(sheet, bounds, ranges):
    """Whether one sheet range intersects any other range on that sheet."""
    min_col, min_row, max_col, max_row = _closed_bounds(bounds)
    folded = sheet.casefold()
    for title, other in ranges:
        if title.casefold() != folded:
            continue
        o_min_col, o_min_row, o_max_col, o_max_row = \
            _closed_bounds(other)
        if not (o_max_col < min_col or o_min_col > max_col
                or o_max_row < min_row or o_min_row > max_row):
            return True
    return False


def _source_formula_model(wb):
    """Load the original formula view retained by preservation."""
    from openpyxl.reader.excel import load_workbook

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        return load_workbook(
            io.BytesIO(wb._paper_source), data_only=False, preserve=False)


def _use_current_defined_names(model, wb):
    """Overlay current name bindings on a retained formula model."""
    model.defined_names = wb.defined_names.copy()
    current = {ws.title.casefold(): ws for ws in wb.worksheets}
    for ws in model.worksheets:
        live = current.get(ws.title.casefold())
        if live is not None:
            ws.defined_names = live.defined_names.copy()


def _dependency_model(wb, retained=None):
    """Return formulas with the workbook's current defined-name bindings."""
    if not wb.data_only:
        return wb
    formulas = retained if retained is not None else _source_formula_model(wb)
    _use_current_defined_names(formulas, wb)
    return formulas


def _formula_outputs(wb):
    """Map multi-cell formula anchors to their declared result ranges."""
    from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

    outputs = {}
    for ws in wb.worksheets:
        for (row, column), cell in ws._cells.items():
            formula = cell._value
            if not isinstance(formula, (ArrayFormula, DataTableFormula)):
                continue
            bounds = _bounds(formula.ref)
            if bounds is not None:
                outputs[(ws.title, row, column)] = (ws.title, bounds)
    return outputs


def _table_semantics(table):
    """Return the table properties that affect structured references."""
    def formula(value):
        if value is None:
            return None
        return value.array, value.attr_text

    columns = tuple(
        (column.id, column.uniqueName, column.name,
         column.totalsRowFunction, column.totalsRowLabel,
         column.queryTableFieldId,
         formula(column.calculatedColumnFormula),
         formula(column.totalsRowFormula))
        for column in table.tableColumns
    )
    return (
        table.name, table.displayName, table.ref, table.tableType,
        table.headerRowCount, table.insertRow, table.insertRowShift,
        table.totalsRowCount, table.totalsRowShown, columns,
    )


def _table_binding(wb, raw):
    """Fingerprint the exact table named by one structured operand."""
    if not isinstance(raw, str) or "[" not in raw:
        return None
    folded = raw.casefold()
    matches = []
    for ws in wb.worksheets:
        for key in ws.tables:
            table = ws.tables[key]
            names = (key, table.name, table.displayName)
            if not any(isinstance(name, str)
                       and folded.startswith(name.casefold() + "[")
                       for name in names):
                continue
            matches.append((ws.title, key, _table_semantics(table)))
    if not matches:
        return None
    return tuple(sorted(matches, key=repr))


def _operand_binding(wb, address, raw):
    """Fingerprint a name or table behind an unresolved formula operand."""
    if not isinstance(raw, str) or raw.startswith("="):
        return None
    table = _table_binding(wb, raw)
    if table is not None:
        return "table", table
    key = _address_key(address)
    ws = _worksheet(wb, key[0]) if key is not None else None
    if ws is None:
        return None
    from .perception import _defined_name

    definition = _defined_name(wb, ws, raw)
    if definition is None:
        return None
    return "defined-name", definition.name, definition.attr_text


def _dependency_signature(sketch, wb, address):
    references = tuple(sorted([
        (sheet.casefold(), _closed_bounds(bounds), raw)
        for sheet, bounds, raw in sketch.references.get(address, ())
    ], key=repr))
    unresolved = tuple(sorted([
        (raw, _operand_binding(wb, address, raw))
        for raw in sketch.unresolved.get(address, ())
    ], key=repr))
    return references, unresolved


def _tables_changed(wb, ledger):
    """Whether any retained table model differs from its armed snapshot."""
    from openpyxl.xml.functions import tostring
    from .ledger import _settled

    for ws in wb.worksheets:
        before = getattr(ledger, "object_snapshots", {}).get(
            ws, {}).get("table", {})
        current = {}
        for key in ws.tables:
            table = ws.tables[key]
            current[key] = _settled(
                lambda value=table: tostring(value.to_tree()))[0]
        if current != before:
            return True
    return False


def _formula_binding_changes(wb, ledger):
    """Formula outputs whose effective name/table dependencies changed."""
    snapshot = getattr(ledger, "workbook_snapshot", None) or {}
    current = crosspart.render_workbook_elements(wb)
    names_changed = current.get("definedNames") != \
        snapshot.get("definedNames")
    if not names_changed and not _tables_changed(wb, ledger):
        return set(), set()

    from .perception import dependency_sketch

    original = _source_formula_model(wb)
    old_sketch = dependency_sketch(original)
    old_outputs = _formula_outputs(original)
    old_addresses = set(old_sketch.references) | set(old_sketch.unresolved)
    old_signatures = {
        address: _dependency_signature(old_sketch, original, address)
        for address in old_addresses
    }
    if wb.data_only:
        current_model = original
        _use_current_defined_names(current_model, wb)
    else:
        current_model = wb
    new_sketch = dependency_sketch(current_model)
    new_outputs = _formula_outputs(current_model)
    addresses = old_addresses | set(new_sketch.references) \
        | set(new_sketch.unresolved)
    cells = set()
    ranges = set()
    for address in addresses:
        if old_signatures.get(address, ((), ())) == \
                _dependency_signature(new_sketch, wb, address):
            continue
        key = _address_key(address)
        if key is not None:
            cells.add(key)
            for outputs in (old_outputs, new_outputs):
                output = outputs.get(key)
                if output is not None:
                    ranges.add(output)
    return cells, ranges


def _style_dependency_changes(wb, ledger, direct_changes):
    """Return style-edited cells and an optional retained source model."""
    candidates = {
        ws: set(coordinates) - set(direct_changes.get(ws, ()))
        for ws, coordinates in ledger.cells.items()
        if set(coordinates) - set(direct_changes.get(ws, ()))
    }
    if not candidates:
        return set(), None
    retained = _source_formula_model(wb)
    changed = set()
    for ws, coordinates in candidates.items():
        original_title = ledger.renames.get(ws, ws.title)
        original = _worksheet(retained, original_title)
        if original is None:
            continue
        for row, column in coordinates:
            current_cell = ws._cells.get((row, column))
            original_cell = original._cells.get((row, column))
            current_style = getattr(current_cell, "_style", None)
            original_style = getattr(original_cell, "_style", None)
            if current_style != original_style:
                changed.add((ws.title, row, column))
    return changed, retained


def _column_nodes(payload):
    """Return rendered column spans with their complete semantic attributes."""
    if payload is None:
        return set()
    try:
        root = fromstring(payload)
        nodes = set()
        for child in root:
            if child.tag.rsplit("}", 1)[-1] != "col":
                continue
            min_col = int(child.attrib["min"])
            max_col = int(child.attrib["max"])
            if not 1 <= min_col <= max_col <= MAX_COLUMN:
                raise ValueError("column span is outside worksheet bounds")
            nodes.add((min_col, max_col,
                       tuple(sorted(child.attrib.items()))))
        return nodes
    except (KeyError, ParseError, TypeError, ValueError):
        return None


def _changed_column_ranges(before, after):
    """Return exact column bands whose rendered dimension state changed."""
    old = _column_nodes(before)
    new = _column_nodes(after)
    if old is None or new is None:
        return {(1, 1, MAX_COLUMN, MAX_ROW)}
    return {
        (min_col, 1, max_col, MAX_ROW)
        for min_col, max_col, _attributes in old.symmetric_difference(new)
    }


def _changed_filter_ranges(before, after):
    """Return row bands covered by old/new AutoFilter ranges."""
    ranges = set()
    for payload in (before, after):
        if payload is None:
            continue
        try:
            ref = fromstring(payload).attrib.get("ref")
        except (ParseError, TypeError, ValueError):
            return {(1, 1, MAX_COLUMN, MAX_ROW)}
        bounds = _bounds(ref)
        if bounds is None:
            return {(1, 1, MAX_COLUMN, MAX_ROW)}
        _min_col, min_row, _max_col, max_row = bounds
        if min_row is None or max_row is None:
            return {(1, 1, MAX_COLUMN, MAX_ROW)}
        # Filtering hides complete worksheet rows, so formulas in columns
        # outside the filter rectangle can still observe the visibility
        # change through SUBTOTAL or AGGREGATE.
        ranges.add((1, min_row, MAX_COLUMN, max_row))
    return ranges or {(1, 1, MAX_COLUMN, MAX_ROW)}


def _calculation_context_changes(wb, ledger, direct_changes):
    """Formula-precedent cells/ranges changed without a value overwrite."""
    from .regions import diff_regions, diff_row_attrs

    cells, retained = _style_dependency_changes(
        wb, ledger, direct_changes)
    visibility_ranges = set()
    format_ranges = set()
    for ws in wb.worksheets:
        if ws in ledger.added_sheets:
            continue
        armed_rows = ledger.row_attr_snapshots.get(ws, {})
        rows = diff_row_attrs(
            ws, armed_rows)
        for row, current in rows.items():
            before = dict(armed_rows.get(row, ()))
            target = visibility_ranges if before.get("hidden") != \
                current.get("hidden") else format_ranges
            target.add((ws.title, (1, row, MAX_COLUMN, row)))
        armed = ledger.region_snapshots.get(ws, {})
        regions = diff_regions(ws, armed)
        if "cols" in regions:
            format_ranges.update(
                (ws.title, bounds) for bounds in _changed_column_ranges(
                    armed.get("cols"), regions["cols"]))
        if "autoFilter" in regions:
            visibility_ranges.update(
                (ws.title, bounds) for bounds in _changed_filter_ranges(
                    armed.get("autoFilter"), regions["autoFilter"]))
        if "sheetFormatPr" in regions:
            format_ranges.add((ws.title, (1, 1, MAX_COLUMN, MAX_ROW)))
    return cells, visibility_ranges, format_ranges, retained


class _DependencyIndex:
    """Consumable reverse index from referenced cells and ranges.

    A reference group is removed the first time a tainted cell or output
    range intersects it: every formula in that group has then been tainted,
    so later intersections cannot add information.
    """

    def __init__(self, sketch):
        point_groups = {}
        range_groups = {}
        for address, references in sketch.references.items():
            for sheet, bounds, _raw in references:
                closed = _closed_bounds(bounds)
                min_col, min_row, max_col, max_row = closed
                folded = sheet.casefold()
                if min_col == max_col and min_row == max_row:
                    point_groups.setdefault(
                        folded, {}).setdefault(
                            (min_row, min_col), set()).add(address)
                else:
                    range_groups.setdefault(
                        folded, {}).setdefault(
                            closed, set()).add(address)

        self._points = point_groups
        self._ranges = range_groups

    def _pop_point(self, sheet, row, column):
        points = self._points.get(sheet)
        if points is None:
            return ()
        formulas = points.pop((row, column), ())
        if not points:
            self._points.pop(sheet)
        return formulas

    def _pop_points_in(self, sheet, bounds):
        points = self._points.get(sheet)
        if points is None:
            return set()
        min_col, min_row, max_col, max_row = bounds
        formulas = set()
        matches = [
            point for point in points
            if min_row <= point[0] <= max_row
            and min_col <= point[1] <= max_col
        ]
        for point in matches:
            formulas.update(points.pop(point))
        if not points:
            self._points.pop(sheet)
        return formulas

    def _pop_ranges_in(self, sheet, bounds):
        ranges = self._ranges.get(sheet)
        if ranges is None:
            return set()
        min_col, min_row, max_col, max_row = bounds
        formulas = set()
        matches = []
        for other, addresses in ranges.items():
            other_min_col, other_min_row, other_max_col, other_max_row = other
            if other_max_col < min_col or other_min_col > max_col \
                    or other_max_row < min_row or other_min_row > max_row:
                continue
            matches.append(other)
            formulas.update(addresses)
        for other in matches:
            ranges.pop(other)
        if not ranges:
            self._ranges.pop(sheet)
        return formulas

    def pop_cell(self, cell):
        """Return formulas first reached by one newly tainted cell."""
        sheet, row, column = cell
        folded = sheet.casefold()
        formulas = set(self._pop_point(folded, row, column))
        formulas.update(self._pop_ranges_in(
            folded, (column, row, column, row)))
        return formulas

    def pop_range(self, item):
        """Return formulas first reached by one newly tainted range."""
        sheet, bounds = item
        folded = sheet.casefold()
        closed = _closed_bounds(bounds)
        formulas = self._pop_points_in(folded, closed)
        formulas.update(self._pop_ranges_in(folded, closed))
        return formulas


def _dirty_closure(wb, dirty, *, dependency_cells=(),
                   dependency_ranges=(), activation_ranges=(), retained=None,
                   force_recalculation=False):
    """Dirty cells plus every formula result they may affect transitively."""
    tainted = {
        (ws.title, row, column)
        for ws, coordinates in dirty.items()
        for row, column in coordinates
        if row is not None and column is not None
    }
    context_cells = set(dependency_cells)
    context_ranges = set(dependency_ranges)
    other_context_ranges = set(activation_ranges)
    if not tainted and not context_cells and not context_ranges \
            and not other_context_ranges \
            and not force_recalculation:
        return tainted, set()
    from .perception import dependency_sketch

    formulas = _dependency_model(wb, retained=retained)
    sketch = dependency_sketch(formulas)
    outputs = _formula_outputs(formulas)
    addresses = set(sketch.references) | set(sketch.unresolved) \
        | set(sketch.volatile) | set(sketch.contextual)
    keys = {address: _address_key(address) for address in addresses}
    dependency_index = _DependencyIndex(sketch)
    tainted_ranges = set()
    pending = deque()
    enqueued_cells = set()

    def enqueue_cell(key):
        if key is None:
            return
        tainted.add(key)
        if key not in enqueued_cells:
            enqueued_cells.add(key)
            pending.append(("cell", key))

    def enqueue_range(item):
        if item not in tainted_ranges:
            tainted_ranges.add(item)
            pending.append(("range", item))

    def taint_formula(address):
        enqueue_cell(keys.get(address))

    for key in tuple(tainted):
        enqueue_cell(key)
    for address in sketch.volatile:
        taint_formula(address)
    # An unresolved formula may read any edited cell. Taint it rather than
    # guessing that its dynamic/structured/external reference is unrelated.
    for address in sketch.unresolved:
        taint_formula(address)
    for address in sketch.contextual:
        references = sketch.references.get(address, ())
        if any(_ranges_hit(sheet, bounds, context_ranges)
               for sheet, bounds, _raw in references):
            taint_formula(address)

    while pending:
        kind, item = pending.popleft()
        if kind == "cell":
            output = outputs.get(item)
            if output is not None:
                enqueue_range(output)
            dependents = dependency_index.pop_cell(item)
        else:
            dependents = dependency_index.pop_range(item)
        for address in dependents:
            taint_formula(address)
    return tainted, tainted_ranges


def source_impacts(wb, ledger, *, force_recalculation=False):
    """Return local pivot caches that workbook edits can make stale."""
    index, _cache_parts = _index(wb)
    if not index:
        return []
    changes = {ws: set(coords)
               for ws, coords in ledger.value_overwrites.items() if coords}
    for ws, writes in ledger.cache_writes.items():
        if writes:
            changes.setdefault(ws, set()).update(writes)
    (dependency_cells, dependency_ranges, activation_ranges, retained) = \
        _calculation_context_changes(wb, ledger, changes)
    tainted, tainted_ranges = _dirty_closure(
        wb, changes, dependency_cells=dependency_cells,
        dependency_ranges=dependency_ranges,
        activation_ranges=activation_ranges, retained=retained,
        force_recalculation=force_recalculation)
    binding_cells, binding_ranges = _formula_binding_changes(wb, ledger)
    tainted.update(binding_cells)
    tainted_ranges.update(binding_ranges)
    pivots_by_cache = {}
    for pivot_name, entries in index.items():
        for title, cache_part in entries:
            pivots_by_cache.setdefault(cache_part, set()).add(
                "{0}!{1}".format(title, pivot_name))
    impacts = []
    snapshots = getattr(ledger, "pivot_source_snapshots", {})
    with zipfile.ZipFile(io.BytesIO(wb._paper_source)) as zin:
        names = set(zin.namelist())
        for cache_part, pivots in sorted(pivots_by_cache.items()):
            if cache_part not in names:
                # Newly allocated isolated caches are rebuilt from the
                # current source; they are not yet in the preserved package.
                continue
            source = _cache_source(wb, zin.read(cache_part))
            if source is None:
                continue
            ws, bounds, label, binding = source
            source_changed = cache_part in snapshots \
                and snapshots[cache_part] != binding
            if ws is None:
                binding_intersects = bool(
                    binding_cells or binding_ranges)
                recalculation_intersects = force_recalculation and bool(
                    tainted or tainted_ranges)
                intersects = bool(changes or dependency_cells
                                  or dependency_ranges
                                  or activation_ranges) \
                    or binding_intersects \
                    or recalculation_intersects
            else:
                binding_intersects = (
                    _bounds_hit(ws.title, bounds, binding_cells)
                    or _ranges_hit(ws.title, bounds, binding_ranges))
                intersects = (
                    _bounds_hit(ws.title, bounds, tainted)
                    or _ranges_hit(ws.title, bounds, tainted_ranges))
            intersects = intersects or source_changed
            if intersects:
                impacts.append({
                    "part": cache_part,
                    "pivots": sorted(pivots),
                    "source": label,
                    "cause": "source_changed" if source_changed
                    else "input_changed",
                    "formula_binding_changed": binding_intersects,
                })
    return impacts


def _request_impacted_refreshes(wb, ledger, *, force_recalculation=False):
    """Select every cache affected by one explicit recalculation policy."""
    impacts = source_impacts(
        wb, ledger, force_recalculation=force_recalculation)
    ledger.pivot_refresh_requests.update(
        impact["part"] for impact in impacts)
    return impacts


def validate_source_freshness(wb, ledger):
    """Refuse silent stale pivot caches without explicit refresh consent."""
    impacts = source_impacts(wb, ledger)
    rebuilt = set()
    for operation in getattr(ledger, "pivot_operations", {}).values():
        if getattr(operation, "noop", False):
            continue
        if operation.kind == "delete" \
                or getattr(operation, "cache_rebuild", False):
            rebuilt.add(operation.allocation.cache_part)
    unsafe = [impact for impact in impacts
              if impact["part"] not in ledger.pivot_refresh_requests
              and impact["part"] not in rebuilt]
    if not unsafe:
        return impacts
    details = "; ".join(
        "{0} from {1}".format(", ".join(impact["pivots"]),
                              impact["source"])
        for impact in unsafe)
    raise UnsupportedStructureError(
        "workbook edits affect or may affect preserved pivot cache(s): {0}. "
        "Saving would leave plausible but stale pivot results. Explicitly call "
        "wb.set_pivot_refresh_on_load(pivots=[...]) for these pivots to "
        "accept that Excel must refresh them on open; headless readers may "
        "observe the old cached results until then. Nothing was written."
        .format(details),
        kind="stale-pivot-cache",
        anchor=unsafe[0]["pivots"][0])


def resolve_requests(wb, pivots=None, *, all=False):
    """Resolve requested pivots to their cache-definition parts.

    :param wb: Preserve-mode workbook containing the pivots.
    :type wb: openpyxl.workbook.workbook.Workbook
    :param pivots: Pivot names, optionally qualified by worksheet name.
    :type pivots: iterable of str or None
    :param all: Select every loaded pivot cache instead of named pivots.
    :type all: bool
    :return: Sorted package part names for the selected pivot caches.
    :rtype: list of str
    """
    if (pivots is None) == (not all):
        raise ValueError("pass exactly one of pivots=[...] or all=True")
    index, cache_parts = _index(wb)
    if all:
        return cache_parts
    if isinstance(pivots, str):
        raise TypeError("pivots must be an iterable of pivot names")
    requested = list(pivots)
    if not requested:
        raise ValueError("pivots must contain at least one pivot name")
    resolved = set()
    for name in requested:
        if not isinstance(name, str) or not name:
            raise TypeError("pivot names must be non-empty strings")
        if "!" in name:
            title, pivot_name = name.rsplit("!", 1)
            if title.startswith("'") and title.endswith("'"):
                title = title[1:-1].replace("''", "'")
            matches = [item for item in index.get(pivot_name, [])
                       if item[0].casefold() == title.casefold()]
        else:
            pivot_name = name
            matches = index.get(pivot_name, [])
        if not matches:
            raise TargetNotFoundError(
                "no loaded pivot named {0!r} was found".format(name),
                kind="pivot-not-found", anchor=name)
        if len(matches) > 1:
            options = ["{0}!{1}".format(title, pivot_name)
                       for title, _part in matches]
            raise AmbiguousTargetError(
                "pivot name {0!r} is ambiguous; use a sheet-qualified "
                "name".format(name), kind="ambiguous-pivot", anchor=name,
                options=options)
        resolved.add(matches[0][1])
    return sorted(resolved)


def plan_refresh(zin, parts, plan):
    """Add pivot refresh requests to a package byte plan.

    :param zin: Open workbook package.
    :type zin: zipfile.ZipFile
    :param parts: Pivot cache-definition part names to update.
    :type parts: iterable of str
    :param plan: Planned replacement bytes keyed by part name.
    :type plan: dict
    :return: Part names whose refresh metadata changed.
    :rtype: list of str
    """
    patched = []
    for part in sorted(parts):
        payload = plan.get(part, zin.read(part))
        root = _cache_root(payload)
        changed = False
        enable_refresh = root.attrs.get("enableRefresh")
        if enable_refresh is not None \
                and enable_refresh.casefold() not in ("1", "true"):
            start, end, head = crosspart._patch_attr(
                payload, root, "enableRefresh", "1")
            payload = payload[:start] + head + payload[end:]
            root = _cache_root(payload)
            changed = True
        if root.attrs.get("refreshOnLoad", "false").casefold() \
                not in ("1", "true"):
            start, end, head = crosspart._patch_attr(
                payload, root, "refreshOnLoad", "1")
            payload = payload[:start] + head + payload[end:]
            changed = True
        if changed:
            plan[part] = payload
            patched.append(part)
    return patched
