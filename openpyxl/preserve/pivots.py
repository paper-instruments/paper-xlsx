# paper-xlsx: targeted pivot cache refresh

"""Resolve pivot names and patch only cache refresh metadata."""

import io
import zipfile
from xml.etree.ElementTree import ParseError, fromstring

from openpyxl.errors import (
    AmbiguousTargetError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries
from openpyxl.xml.constants import SHEET_MAIN_NS

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
        def descendants(node):
            for child in node.children:
                yield child
                yield from descendants(child)

        for child in descendants(workbook_root):
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
                root = crosspart.scan_small(
                    zin.read(pivot_part), "pivotTableDefinition", max_depth=1)
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
    if any(value is None for value in bounds):
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
    min_col, min_row, max_col, max_row = bounds
    if min_col is None:
        min_col, max_col = 1, 1 << 20
    if min_row is None:
        min_row, max_row = 1, 1 << 22
    folded = sheet.casefold()
    return any(
        title.casefold() == folded
        and min_row <= row <= max_row
        and min_col <= column <= max_col
        for title, row, column in cells)


def _dirty_closure(wb, dirty):
    """Dirty cells plus every formula result they may affect transitively."""
    tainted = {
        (ws.title, row, column)
        for ws, coordinates in dirty.items()
        for row, column in coordinates
        if row is not None and column is not None
    }
    if not tainted:
        return tainted
    from .perception import dependency_sketch

    sketch = dependency_sketch(wb)
    # An unresolved formula may read any edited cell. Taint it rather than
    # guessing that its dynamic/structured/external reference is unrelated.
    for address in sketch.unresolved:
        key = _address_key(address)
        if key is not None:
            tainted.add(key)
    changed = True
    while changed:
        changed = False
        for address, references in sketch.references.items():
            key = _address_key(address)
            if key is None or key in tainted:
                continue
            if any(_bounds_hit(sheet, bounds, tainted)
                   for sheet, bounds, _raw in references):
                tainted.add(key)
                changed = True
    return tainted


def source_impacts(wb, ledger):
    """Return local pivot caches that workbook edits can make stale."""
    dirty = {ws: set(coords) for ws, coords in ledger.value_overwrites.items()
             if coords}
    tainted = _dirty_closure(wb, dirty)
    index, _cache_parts = _index(wb)
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
            intersects = bool(dirty) if ws is None else _bounds_hit(
                ws.title, bounds, tainted)
            intersects = intersects or source_changed
            if intersects:
                impacts.append({
                    "part": cache_part,
                    "pivots": sorted(pivots),
                    "source": label,
                    "cause": "source_changed" if source_changed
                    else "input_changed",
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
            title = title.strip("'").replace("''", "'")
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
        if root.attrs.get("refreshOnLoad") in ("1", "true", "True"):
            continue
        start, end, head = crosspart._patch_attr(
            payload, root, "refreshOnLoad", "1")
        plan[part] = payload[:start] + head + payload[end:]
        patched.append(part)
    return patched
