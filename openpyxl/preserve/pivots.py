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
from openpyxl.utils.cell import range_boundaries
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


def _defined_name_source(wb, name):
    """Resolve one static workbook- or worksheet-scoped defined name."""
    matches = []
    for key, defn in wb.defined_names.items():
        if key.casefold() == name.casefold():
            matches.append((None, defn))
    for ws in wb.worksheets:
        for key, defn in ws.defined_names.items():
            if key.casefold() == name.casefold():
                matches.append((ws, defn))
    resolved = []
    for owner, defn in matches:
        value = getattr(defn, "attr_text", None)
        if not isinstance(value, str):
            continue
        if owner is not None and "!" not in value:
            bounds = _bounds(value)
            if bounds is not None:
                resolved.append((owner, bounds))
            continue
        try:
            destinations = list(defn.destinations)
        except (AttributeError, TypeError, ValueError):
            continue
        if len(destinations) != 1:
            continue
        title, ref = destinations[0]
        ws = _worksheet(wb, title)
        bounds = _bounds(ref)
        if ws is not None and bounds is not None:
            resolved.append((ws, bounds))
    return resolved[0] if len(resolved) == 1 else None


def _named_source(wb, name):
    """Resolve one named worksheet table or one static defined name."""
    tables = []
    for ws in wb.worksheets:
        for key in ws.tables:
            table = ws.tables[key]
            names = (key, getattr(table, "name", None),
                     getattr(table, "displayName", None))
            if any(isinstance(item, str)
                   and item.casefold() == name.casefold()
                   for item in names):
                bounds = _bounds(table.ref)
                if bounds is not None:
                    tables.append((ws, bounds))
    defined = _defined_name_source(wb, name)
    matches = tables + ([defined] if defined is not None else [])
    return matches[0] if len(matches) == 1 else None


def _cache_source(wb, payload):
    """Return ``(worksheet, bounds, label)`` for one exact local source.

    ``None`` means the cache is external and cannot be affected by local cell
    edits. A tuple of ``(None, None, label)`` is an unresolved local source and
    therefore intersects conservatively.
    """
    try:
        root = fromstring(payload)
    except (ParseError, TypeError, ValueError):
        return None, None, "unresolved pivot source"
    namespace = "{{{0}}}".format(SHEET_MAIN_NS)
    sources = root.findall(namespace + "cacheSource")
    if len(sources) != 1:
        return None, None, "unresolved pivot source"
    source = sources[0]
    source_type = source.attrib.get("type")
    if source_type == "external":
        return None
    if source_type != "worksheet":
        return None, None, "unsupported {0!r} pivot source".format(
            source_type)
    worksheet_sources = source.findall(namespace + "worksheetSource")
    if len(worksheet_sources) != 1:
        return None, None, "unresolved worksheet pivot source"
    worksheet_source = worksheet_sources[0]
    ref = worksheet_source.attrib.get("ref")
    title = worksheet_source.attrib.get("sheet")
    name = worksheet_source.attrib.get("name")
    if ref is not None and title is not None and name is None:
        ws = _worksheet(wb, title)
        bounds = _bounds(ref)
        if ws is not None and bounds is not None:
            return ws, bounds, "{0}!{1}".format(ws.title, ref)
    elif name is not None and ref is None:
        resolved = _named_source(wb, name)
        if resolved is not None:
            ws, bounds = resolved
            return ws, bounds, name
    return None, None, "unresolved worksheet pivot source"


def source_impacts(wb, ledger):
    """Return pivot caches made stale by recorded local value writes."""
    dirty = {ws: set(coords) for ws, coords in ledger.value_overwrites.items()
             if coords}
    if not dirty:
        return []
    index, _cache_parts = _index(wb)
    pivots_by_cache = {}
    for pivot_name, entries in index.items():
        for title, cache_part in entries:
            pivots_by_cache.setdefault(cache_part, set()).add(
                "{0}!{1}".format(title, pivot_name))
    impacts = []
    with zipfile.ZipFile(io.BytesIO(wb._paper_source)) as zin:
        for cache_part, pivots in sorted(pivots_by_cache.items()):
            source = _cache_source(wb, zin.read(cache_part))
            if source is None:
                continue
            ws, bounds, label = source
            intersects = ws is None
            if ws is not None:
                min_col, min_row, max_col, max_row = bounds
                intersects = any(
                    min_row <= row <= max_row
                    and min_col <= column <= max_col
                    for row, column in dirty.get(ws, ()))
            if intersects:
                impacts.append({
                    "part": cache_part,
                    "pivots": sorted(pivots),
                    "source": label,
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
        "edited cells affect or may affect preserved pivot cache(s): {0}. "
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
