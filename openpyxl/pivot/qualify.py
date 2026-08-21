# paper-xlsx: pivot validity and per-operation capability qualification

"""Finite, testable qualification rules for loaded PivotTables.

A pivot is graph-valid only when required relationships, parts, counts, and
identities are internally consistent. Capabilities are positive and
operation-specific. Foreign pivots remain foreign even when their semantic
projection is complete. V1 grants them at most ``can_refresh_on_open``.
"""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl.pivot.api_types import (
    PivotCapabilities,
    PivotSource,
)
from openpyxl.utils.cell import range_boundaries


PAPER_TAG = "paper-xlsx:pivot-v1"
SUPPORTED_SOURCE_KINDS = frozenset(("table", "range", "defined-name"))
UNSUPPORTED_SOURCE_KINDS = frozenset((
    "external", "consolidation", "scenario", "unknown", "named",
))
# Empty until a fixture-backed extension is proven orthogonal to an edit.
EXTENSION_ALLOWLIST = frozenset()
MUTATION_CAPABILITIES = (
    "can_headless_refresh",
    "can_rebuild_cache",
    "can_edit_layout",
    "can_repoint_source",
    "can_move",
    "can_rename",
    "can_delete",
)
ALL_CAPABILITIES = ("can_refresh_on_open",) + MUTATION_CAPABILITIES
# Shared caches refuse cache rebuild, source/catalog changes, movement,
# layout/update, and delete. Layout-only shared-cache edits are deferred in
# v1 because ``update()`` rebuilds cache records. Rename is excluded: it
# replaces only the selected pivot definition.
CACHE_ISOLATION_CAPABILITIES = (
    "can_headless_refresh",
    "can_rebuild_cache",
    "can_edit_layout",
    "can_repoint_source",
    "can_move",
    "can_delete",
)
OWNERSHIP_CAPABILITIES = (
    "can_headless_refresh",
    "can_rebuild_cache",
    "can_repoint_source",
    "can_move",
    "can_delete",
)


@dataclass(frozen=True)
class QualificationReason:
    capability: str | None
    code: str
    context: tuple = ()

    def to_dict(self):
        return {
            "capability": self.capability,
            "code": self.code,
            "context": {key: value for key, value in self.context},
        }


@dataclass(frozen=True)
class PivotQualification:
    valid: bool
    origin: str
    capabilities: PivotCapabilities
    refresh_on_open_scope: tuple
    reasons: tuple
    source_supported: bool
    cache_shared: bool
    extensions: tuple


def qualify_pivot(node, cache, projection, graph, workbook=None,
                  ownership_proved=False):
    """Qualify one loaded pivot against the v1 capability contract."""
    reasons = []
    origin = "paper" if node.tag == PAPER_TAG else "foreign"
    cache_shared = bool(
        cache is not None and len(cache.referenced_by) > 1)
    extensions = tuple(node.extension_fingerprints)
    if cache is not None:
        extensions = extensions + tuple(cache.extension_fingerprints)

    graph_reasons = _graph_reasons(node, cache, graph)
    reasons.extend(graph_reasons)
    valid = node.valid and (cache is None or cache.valid) and not any(
        reason.code in _INVALIDATING_CODES for reason in graph_reasons)

    source_supported, source_reasons = _source_supported(
        node, cache, projection, workbook)
    reasons.extend(source_reasons)

    flags = {name: False for name in ALL_CAPABILITIES}
    scope = _refresh_scope(node, cache)

    if valid and _unambiguous_cache(node, cache, graph):
        flags["can_refresh_on_open"] = True
    else:
        reasons.append(_reason(
            "can_refresh_on_open",
            "invalid-pivot-graph" if not valid else "dangling-pivot-cache",
            part=node.identity.pivot_part,
            cache_id=node.cache_id,
        ))
        scope = ()

    if origin == "foreign":
        for name in MUTATION_CAPABILITIES:
            reasons.append(_reason(
                name, "foreign-operation-deferred",
                part=node.identity.pivot_part,
            ))
        return _result(
            valid, origin, flags, scope, reasons, source_supported,
            cache_shared, extensions)

    if not valid:
        for name in MUTATION_CAPABILITIES:
            reasons.append(_reason(
                name, "graph-invalid",
                part=node.identity.pivot_part,
            ))
        return _result(
            False, origin, flags, scope, reasons, source_supported,
            cache_shared, extensions)

    projection_codes = {item.code for item in projection.reasons}
    for item in projection.reasons:
        mapped = item.code
        if mapped in ("missing-field", "invalid-item-index",
                      "unknown-aggregate", "incomplete-semantic-projection",
                      "missing-measure", "missing-output-location",
                      "invalid-output-location"):
            for name in (
                "can_headless_refresh", "can_rebuild_cache",
                "can_edit_layout", "can_repoint_source", "can_rename",
            ):
                reasons.append(_reason(
                    name, mapped, **dict(item.context)))

    if cache is not None and cache.has_grouping:
        _disable(flags, reasons, (
            "can_headless_refresh", "can_rebuild_cache", "can_edit_layout",
            "can_repoint_source", "can_rename",
        ), "unsupported-grouping", part=cache.definition_part)
    if cache is not None and cache.has_calculated:
        _disable(flags, reasons, (
            "can_headless_refresh", "can_rebuild_cache", "can_edit_layout",
            "can_repoint_source", "can_rename",
        ), "unsupported-calculated", part=cache.definition_part)

    disallowed = _disallowed_extensions(extensions)
    if disallowed:
        _disable(flags, reasons, (
            "can_headless_refresh", "can_rebuild_cache", "can_edit_layout",
            "can_repoint_source", "can_rename", "can_delete",
        ), "unsupported-extension",
            part=node.identity.pivot_part,
            uri=disallowed[0].uri)

    if not source_supported:
        _disable(flags, reasons, (
            "can_headless_refresh", "can_rebuild_cache", "can_edit_layout",
            "can_repoint_source", "can_rename",
        ), "unsupported-pivot-source",
            part=None if cache is None else cache.definition_part)

    semantic_ok = (
        projection.complete
        and source_supported
        and not projection_codes
        and not (cache is not None and (cache.has_grouping or cache.has_calculated))
        and not disallowed
    )
    if semantic_ok:
        flags["can_edit_layout"] = True
        _drop_reasons(reasons, "can_edit_layout")

    name_unique = _name_is_unique(node, graph)
    if semantic_ok and name_unique and node.identity.name:
        flags["can_rename"] = True
        _drop_reasons(reasons, "can_rename")
    elif not name_unique or not node.identity.name:
        reasons.append(_reason(
            "can_rename",
            "ambiguous-pivot" if node.identity.name else "missing-name",
            name=node.identity.name or None,
            sheet=node.sheet_title,
        ))

    if cache_shared:
        _disable(flags, reasons, CACHE_ISOLATION_CAPABILITIES,
                 "pivot-cache-shared",
                 part=None if cache is None else cache.definition_part,
                 referenced_by=",".join(
                     "%s!%s" % item for item in cache.referenced_by))

    if origin == "paper" and not cache_shared and projection.complete \
            and workbook is not None and _output_owned(workbook, node, projection):
        ownership_proved = True

    if not ownership_proved:
        _disable(flags, reasons, OWNERSHIP_CAPABILITIES,
                 "output-ownership-unproved",
                 part=node.identity.pivot_part,
                 output_range=node.output_range)

    if semantic_ok and not cache_shared and ownership_proved:
        flags["can_headless_refresh"] = True
        flags["can_rebuild_cache"] = True
        flags["can_repoint_source"] = True
        flags["can_move"] = True
        flags["can_delete"] = True
        for name in OWNERSHIP_CAPABILITIES:
            _drop_reasons(reasons, name)

    return _result(
        valid, origin, flags, scope, reasons, source_supported,
        cache_shared, extensions)


_INVALIDATING_CODES = frozenset((
    "duplicate-cache-id",
    "duplicate-incoming",
    "dangling-workbook-cache",
    "missing-part",
))


def _graph_reasons(node, cache, graph):
    reasons = []
    for item in node.reasons:
        reasons.append(QualificationReason(None, item.code, item.context))
    if cache is not None:
        for item in cache.reasons:
            reasons.append(QualificationReason(None, item.code, item.context))
    pivot_part = node.identity.pivot_part
    cache_part = node.cache_definition_part
    for item in graph.reasons:
        context = dict(item.context)
        part = context.get("part")
        parts = context.get("parts") or ""
        if part in (pivot_part, cache_part) or (
                cache_part and cache_part in parts.split(",")) or (
                node.cache_id and context.get("cache_id") == node.cache_id):
            reasons.append(QualificationReason(None, item.code, item.context))
    return reasons


def _unambiguous_cache(node, cache, graph):
    if cache is None or not cache.definition_part:
        return False
    if node.cache_id and node.cache_id not in graph.caches_by_id:
        return False
    if cache.definition_part not in graph.caches_by_part:
        return False
    return cache.definition_part == node.cache_definition_part


def _refresh_scope(node, cache):
    if cache is None:
        if node.identity.name:
            return ("%s!%s" % (node.sheet_title, node.identity.name),)
        return ()
    return tuple(
        "%s!%s" % (sheet, name)
        for sheet, name in cache.referenced_by
    )


def _source_supported(node, cache, projection, workbook):
    reasons = []
    source = projection.source
    descriptor = node.source_descriptor or (
        None if cache is None else cache.source_descriptor)
    kind = None if descriptor is None else descriptor.kind
    if source is None and kind in UNSUPPORTED_SOURCE_KINDS:
        code = "unsupported-data-model" if kind in (
            "external", "consolidation", "scenario") else "unsupported-source"
        reasons.append(_reason(None, code, kind=kind))
        return False, reasons
    if source is None:
        reasons.append(_reason(None, "unsupported-source"))
        return False, reasons
    if source.kind not in SUPPORTED_SOURCE_KINDS:
        reasons.append(_reason(None, "unsupported-source", kind=source.kind))
        return False, reasons
    header_reason = _header_reason(source, workbook)
    if header_reason is not None:
        reasons.append(header_reason)
        return False, reasons
    return True, reasons


def _header_reason(source, workbook):
    if workbook is None or source.kind != "range":
        return None
    worksheet = None
    for ws in workbook.worksheets:
        if ws.title == source.sheet:
            worksheet = ws
            break
    if worksheet is None:
        return _reason(None, "missing-source-sheet", sheet=source.sheet)
    try:
        min_col, min_row, max_col, _max_row = range_boundaries(source.ref)
    except (TypeError, ValueError):
        return _reason(None, "invalid-pivot-source", ref=source.ref)
    headers = []
    folded = []
    for column in range(min_col, max_col + 1):
        cell = getattr(worksheet, "_cells", {}).get((min_row, column))
        value = None if cell is None else cell.value
        headers.append(value)
        if isinstance(value, str):
            folded.append(value.casefold())
        elif value is None:
            folded.append(None)
        else:
            folded.append(str(value).casefold())
    if any(not isinstance(value, str) or not value for value in headers):
        return _reason(
            None, "invalid-pivot-source",
            sheet=source.sheet, ref=source.ref, detail="blank-or-nonstring-header")
    seen = {}
    for name in folded:
        if name in seen:
            return _reason(
                None, "duplicate-header",
                sheet=source.sheet, header=headers[folded.index(name)])
        seen[name] = True
    return None


def _disallowed_extensions(extensions):
    return tuple(
        item for item in extensions
        if item.uri not in EXTENSION_ALLOWLIST
    )


def _name_is_unique(node, graph):
    if not node.identity.name:
        return False
    folded = node.identity.name.casefold()
    matches = [
        item for item in graph.pivots
        if item.identity.name
        and item.identity.name.casefold() == folded
    ]
    return len(matches) == 1


def _disable(flags, reasons, names, code, **context):
    for name in names:
        flags[name] = False
        reasons.append(_reason(name, code, **context))


def _drop_reasons(reasons, capability):
    kept = [
        item for item in reasons
        if item.capability != capability
    ]
    reasons[:] = kept


def _reason(capability, code, **context):
    items = tuple(sorted(
        (key, value) for key, value in context.items() if value is not None))
    return QualificationReason(capability, code, items)


def _result(valid, origin, flags, scope, reasons, source_supported,
            cache_shared, extensions):
    reasons = _dedupe(reasons)
    _ensure_false_capabilities_explained(flags, reasons)
    return PivotQualification(
        valid=valid,
        origin=origin,
        capabilities=PivotCapabilities(**flags),
        refresh_on_open_scope=tuple(scope),
        reasons=tuple(reasons),
        source_supported=source_supported,
        cache_shared=cache_shared,
        extensions=tuple(extensions),
    )


def _dedupe(reasons):
    seen = set()
    unique = []
    for item in reasons:
        key = (item.capability, item.code, item.context)
        if key in seen:
            continue
        seen.add(key)
        unique.append(item)
    unique.sort(key=lambda item: (
        item.capability or "",
        item.code,
        item.context,
    ))
    return unique


def _ensure_false_capabilities_explained(flags, reasons):
    explained = {item.capability for item in reasons if item.capability}
    graph_invalid = any(item.capability is None for item in reasons)
    for name, enabled in flags.items():
        if enabled or name in explained:
            continue
        code = "graph-invalid" if graph_invalid else "capability-disabled"
        reasons.append(_reason(name, code))


def _output_owned(workbook, node, projection):
    """Prove Paper-managed output against cache records, not live source."""
    return _reconstruct_owned_output(workbook, node, projection) is not None


def _reconstruct_owned_output(workbook, node, projection):
    """Return materialized owned output cells, or None if unproved.

    Blank layout slots are not owned. Extra values there must not be
    treated as pivot output merely because they sit inside ``location.ref``.
    """
    if projection.spec is None or not node.output_range:
        return None
    package = getattr(workbook, "_paper_source", None)
    if not package or not node.cache_definition_part or not node.cache_records_part:
        return None
    try:
        from openpyxl.pivot.plan import plan_pivot

        snapshot = _snapshot_from_cache_package(package, node, projection)
        spec = _spec_without_explicit_items(projection.spec)
        plan = plan_pivot(spec, snapshot)
    except Exception:
        return None
    if plan.output.ref != node.output_range:
        return None
    worksheet = None
    for item in workbook.worksheets:
        if item.title == node.sheet_title:
            worksheet = item
            break
    if worksheet is None:
        return None
    cells = getattr(worksheet, "_cells", {})
    owned = []
    for cell in plan.output.cells:
        if cell.value is None:
            continue
        existing = cells.get((cell.row, cell.column))
        actual = None if existing is None else existing.value
        if actual != cell.value:
            return None
        owned.append((cell.row, cell.column, cell.value, cell.role))
    return tuple(owned)


def _snapshot_from_cache_package(package, node, projection):
    import io
    import zipfile

    from openpyxl.pivot.inspect import _shared_item_value, _shared_items
    from openpyxl.pivot.graph import _children, _local, _parse_xml
    from openpyxl.pivot.source import snapshot_from_matrix

    with zipfile.ZipFile(io.BytesIO(package)) as zin:
        names = set(zin.namelist())
        if node.cache_definition_part not in names or node.cache_records_part not in names:
            raise ValueError("missing cache parts")
        definition = _parse_xml(zin.read(node.cache_definition_part))
        shared = _shared_items(definition)
        headers = []
        fields = _children(_first_local(definition, "cacheFields"), "cacheField") \
            if _first_local(definition, "cacheFields") is not None else []
        for field in fields:
            from openpyxl.pivot.graph import _attr
            headers.append(_attr(field, "name") or "")
        records_root = _parse_xml(zin.read(node.cache_records_part))
        rows = []
        for record in _children(records_root, "r"):
            values = []
            field_index = 0
            for child in list(record):
                tag = _local(child.tag)
                if tag == "x":
                    from openpyxl.pivot.graph import _attr
                    index = int(_attr(child, "v") or 0)
                    catalog = shared[field_index] if field_index < len(shared) else ()
                    values.append(catalog[index] if index < len(catalog) else None)
                else:
                    values.append(_shared_item_value(child))
                field_index += 1
            rows.append(values)
    return snapshot_from_matrix(headers, rows, source=projection.source)


def _first_local(root, tag):
    from openpyxl.pivot.graph import _first
    return _first(root, tag)


def _spec_without_explicit_items(spec):
    from dataclasses import replace
    from openpyxl.pivot.api_types import PivotAxisField

    def _clear(fields):
        return tuple(PivotAxisField(field.field) for field in fields)

    return replace(spec, rows=_clear(spec.rows), columns=_clear(spec.columns))
