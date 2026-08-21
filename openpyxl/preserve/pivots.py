# paper-xlsx: targeted pivot cache refresh

"""Resolve pivot names and patch only cache refresh metadata."""

import io
import warnings
import zipfile
from xml.etree.ElementTree import ParseError, fromstring

from openpyxl.errors import (
    AmbiguousTargetError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries
from openpyxl.xml.constants import MAX_COLUMN, MAX_ROW, SHEET_MAIN_NS

from . import crosspart
from .tables import _rels_path, _resolve_target

_PIVOT_TABLE_REL = "/pivotTable"
_PIVOT_CACHE_REL = "/pivotCacheDefinition"


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


def _pivot_root(payload):
    """Validate a pivot-table root, including legal prefixed spelling."""
    try:
        element = fromstring(payload)
    except (ParseError, ValueError, TypeError) as exc:
        raise UnsupportedStructureError(
            "pivot table definition XML is malformed",
            kind="unsupported-pivot-table") from exc
    expected = "{{{0}}}pivotTableDefinition".format(SHEET_MAIN_NS)
    if element.tag != expected:
        raise UnsupportedStructureError(
            "pivot table definition has an unexpected root element",
            kind="unsupported-pivot-table")
    return crosspart.scan_small(
        payload, "pivotTableDefinition", max_depth=1,
        allow_prefixed_root=True)


def _relationship_targets(zin, owner_part, suffix):
    rels_part = _rels_path(owner_part)
    if rels_part not in zin.namelist():
        return {}
    root = crosspart.scan_small(zin.read(rels_part), "Relationships",
                                max_depth=1)
    out = {}
    for child in root.children:
        if child.local() != "Relationship" \
                or not child.attrs.get("Type", "").endswith(suffix):
            continue
        out[child.attrs.get("Id")] = _resolve_target(
            owner_part, child.attrs.get("Target", ""))
    return out


def _index(wb):
    """Return ``(qualified-name map, cache parts)`` from source bytes."""
    from .saver import _package_info

    with zipfile.ZipFile(io.BytesIO(wb._paper_source)) as zin:
        names = set(zin.namelist())
        wb_part, sheets = _package_info(zin)
        cache_targets = _relationship_targets(
            zin, wb_part, _PIVOT_CACHE_REL)
        cache_by_id = {}
        workbook_root = crosspart.scan_small(
            zin.read(wb_part), "workbook", max_depth=3)
        pivot_cache_groups = [
            child for child in workbook_root.children
            if child.local() == "pivotCaches"
        ]
        for group in pivot_cache_groups:
            for child in group.children:
                if child.local() != "pivotCache":
                    continue
                cache_id = child.attrs.get("cacheId")
                rid = child.attrs.get("id") or child.attrs.get("r:id")
                target = cache_targets.get(rid)
                if cache_id is not None and target in names:
                    cache_by_id[cache_id] = target

        qualified = {}
        ledger = getattr(wb, "_paper_ledger", None)
        current_by_original = {
            original: ws.title
            for ws, original in getattr(ledger, "renames", {}).items()
        }
        for title, sheet_part in sheets.items():
            current_title = current_by_original.get(title, title)
            pivot_targets = _relationship_targets(
                zin, sheet_part, _PIVOT_TABLE_REL)
            for pivot_part in pivot_targets.values():
                if pivot_part not in names:
                    continue
                root = _pivot_root(zin.read(pivot_part))
                pivot_name = root.attrs.get("name")
                cache_part = cache_by_id.get(root.attrs.get("cacheId"))
                if pivot_name and cache_part:
                    qualified.setdefault(pivot_name, []).append(
                        (current_title, cache_part))
        cache_parts = sorted(set(cache_by_id.values()))
        return qualified, cache_parts


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


def _dependency_model(wb):
    """Return formulas with the workbook's current defined-name bindings."""
    if not wb.data_only:
        return wb
    formulas = _source_formula_model(wb)
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


def _dirty_closure(wb, dirty):
    """Dirty cells plus every formula result they may affect transitively."""
    tainted = {
        (ws.title, row, column)
        for ws, coordinates in dirty.items()
        for row, column in coordinates
        if row is not None and column is not None
    }
    if not tainted:
        return tainted, set()
    from .perception import dependency_sketch

    formulas = _dependency_model(wb)
    sketch = dependency_sketch(formulas)
    outputs = _formula_outputs(formulas)
    tainted_ranges = set()

    def expand_outputs():
        before = len(tainted_ranges)
        for key in tainted:
            output = outputs.get(key)
            if output is not None:
                tainted_ranges.add(output)
        return len(tainted_ranges) != before

    # An unresolved formula may read any edited cell. Taint it rather than
    # guessing that its dynamic/structured/external reference is unrelated.
    for address in sketch.unresolved:
        key = _address_key(address)
        if key is not None:
            tainted.add(key)
    changed = True
    while changed:
        changed = expand_outputs()
        for address, references in sketch.references.items():
            key = _address_key(address)
            if key is None or key in tainted:
                continue
            if any(_bounds_hit(sheet, bounds, tainted)
                   or _ranges_hit(sheet, bounds, tainted_ranges)
                   for sheet, bounds, _raw in references):
                tainted.add(key)
                changed = True
    return tainted, tainted_ranges


def source_impacts(wb, ledger):
    """Return local pivot caches that workbook edits can make stale."""
    index, _cache_parts = _index(wb)
    if not index:
        return []
    changes = {ws: set(coords)
               for ws, coords in ledger.value_overwrites.items() if coords}
    for ws, writes in ledger.cache_writes.items():
        if writes:
            changes.setdefault(ws, set()).update(writes)
    tainted, tainted_ranges = _dirty_closure(wb, changes)
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
        for cache_part, pivots in sorted(pivots_by_cache.items()):
            source = _cache_source(wb, zin.read(cache_part))
            if source is None:
                continue
            ws, bounds, label, binding = source
            source_changed = cache_part in snapshots \
                and snapshots[cache_part] != binding
            if ws is None:
                binding_intersects = bool(
                    binding_cells or binding_ranges)
                intersects = bool(changes) or binding_intersects
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


def validate_source_freshness(wb, ledger):
    """Refuse silent stale pivot caches without explicit refresh consent."""
    impacts = source_impacts(wb, ledger)
    unsafe = [impact for impact in impacts
              if impact["part"] not in ledger.pivot_refresh_requests]
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
