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
