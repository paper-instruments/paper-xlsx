# paper-xlsx: pivot package-graph mutations

"""Allocate and register Paper-owned pivot parts on a ``PartPlan``.

Part names and cache IDs come from the complete package plus staged
operations, never from inherited module counters. ``PartPlan.add_part()``
collision refusal stays hard; replacements use the private ``replace_part``.
"""

from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass

from openpyxl.errors import (
    RelationshipPolicyError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from openpyxl.pivot.cache import CacheDefinition
from openpyxl.pivot.record import RecordList
from openpyxl.pivot.table import TableDefinition
from openpyxl.preserve import crosspart
from openpyxl.preserve.lifecycle import _rels_path
from openpyxl.preserve.saver import _package_info


_PIVOT_TABLE_DIR = "xl/pivotTables/pivotTable"
_CACHE_DEF_DIR = "xl/pivotCache/pivotCacheDefinition"
_CACHE_REC_DIR = "xl/pivotCache/pivotCacheRecords"


@dataclass(frozen=True)
class IsolatedCacheAllocation:
    cache_part: str
    records_part: str
    cache_id: int
    records_relationship_id: str = "rId1"


@dataclass(frozen=True)
class PivotAllocation:
    pivot_part: str
    cache_part: str
    records_part: str
    cache_id: int
    worksheet_part: str
    workbook_part: str
    pivot_cache_relationship_id: str = "rId1"
    records_relationship_id: str = "rId1"

    @property
    def pivot_rels_part(self):
        return _rels_path(self.pivot_part)

    @property
    def cache_rels_part(self):
        return _rels_path(self.cache_part)

    def owned_parts(self):
        return (
            self.pivot_part,
            self.pivot_rels_part,
            self.cache_part,
            self.cache_rels_part,
            self.records_part,
        )


def allocate_create(workbook, worksheet, graph, ledger):
    """Choose unused conventional part names and a free cache ID."""
    source = getattr(workbook, "_paper_source", None)
    if source is None:
        raise UnsupportedStructureError(
            "pivot creation requires a preserve-mode package graph",
            kind="invalid-pivot-graph",
        )
    with zipfile.ZipFile(io.BytesIO(source)) as archive:
        names = set(archive.namelist())
        workbook_part, sheet_parts = _package_info(archive)
    original_title = ledger.renames.get(worksheet, worksheet.title)
    worksheet_part = sheet_parts.get(original_title)
    if not worksheet_part:
        raise RelationshipPolicyError(
            "cannot locate the package part for sheet %r via workbook "
            "relationships" % worksheet.title,
            kind="invalid-pivot-graph",
            anchor=worksheet.title,
        )
    reserved = set(names)
    used_ids = set()
    for op in getattr(ledger, "pivot_operations", {}).values():
        allocation = op.allocation
        reserved.update(allocation.owned_parts())
        used_ids.add(int(allocation.cache_id))
    for cache_id in getattr(graph, "caches_by_id", {}):
        try:
            used_ids.add(int(cache_id))
        except (TypeError, ValueError):
            continue
    return PivotAllocation(
        pivot_part=_next_numbered(_PIVOT_TABLE_DIR, reserved),
        cache_part=_next_numbered(_CACHE_DEF_DIR, reserved),
        records_part=_next_numbered(_CACHE_REC_DIR, reserved),
        cache_id=_next_cache_id(used_ids),
        worksheet_part=worksheet_part,
        workbook_part=workbook_part,
    )


def allocate_isolated_cache(workbook, graph, ledger):
    """Choose a free dedicated cache closure for shared-cache isolation."""
    source = getattr(workbook, "_paper_source", None)
    if source is None:
        raise UnsupportedStructureError(
            "shared-cache isolation requires a preserve-mode package graph",
            kind="invalid-pivot-graph",
        )
    with zipfile.ZipFile(io.BytesIO(source)) as archive:
        names = set(archive.namelist())
    reserved = set(names)
    used_ids = set()
    for operation in getattr(ledger, "pivot_operations", {}).values():
        allocation = operation.allocation
        reserved.update(allocation.owned_parts())
        used_ids.add(int(allocation.cache_id))
        original = getattr(operation, "original_cache_part", None)
        if original:
            reserved.add(original)
            reserved.add(_rels_path(original))
        original_records = getattr(operation, "original_records_part", None)
        if original_records:
            reserved.add(original_records)
    for cache_id in getattr(graph, "caches_by_id", {}):
        try:
            used_ids.add(int(cache_id))
        except (TypeError, ValueError):
            continue
    return IsolatedCacheAllocation(
        cache_part=_next_numbered(_CACHE_DEF_DIR, reserved),
        records_part=_next_numbered(_CACHE_REC_DIR, reserved),
        cache_id=_next_cache_id(used_ids),
        records_relationship_id="rId1",
    )


def plan_creates(context):
    """Add staged pivot-operation payloads to the save ``PartPlan``."""
    ledger = context.ledger
    operations = getattr(ledger, "pivot_operations", None)
    if not operations:
        return
    registry = []
    removals = []
    for operation in operations.values():
        if getattr(operation, "noop", False):
            continue
        strategy = getattr(operation, "publication_strategy", None)
        if operation.kind == "create" and not getattr(
                operation, "replace_existing", False):
            registry.extend(_plan_add(context, operation))
        elif operation.kind == "delete" or strategy == "remove":
            _plan_remove(context, operation)
            removals.append(int(operation.allocation.cache_id))
        elif strategy == "shared-isolation":
            added, dropped = _plan_isolate(context, operation)
            registry.extend(added)
            removals.extend(dropped)
        elif operation.kind in (
                "create", "refresh", "repoint", "move", "update", "rename",
                "adopt") or strategy in (
                "dedicated-replacement", "managed-replacement"):
            _plan_replace(context, operation)
        else:
            raise UnsupportedStructureError(
                "unknown pivot publication kind %r. Nothing was written."
                % operation.kind,
                kind="unsupported-pivot-operation",
                anchor=operation.allocation.pivot_part,
            )
    context.part_plan.pivot_cache_registry = tuple(registry)
    context.part_plan.pivot_cache_removals = tuple(removals)


def _plan_add(context, operation):
    part_plan = context.part_plan
    archive = context.archive
    names = context.names
    allocation = operation.allocation
    payloads = operation.payloads
    _validate_allocation_free(allocation, names, part_plan)
    cache_rid = part_plan.reserve_rid(
        context.workbook_rels_part,
        archive.read(context.workbook_rels_part)
        if context.workbook_rels_part in names else None)
    sheet_rels = _rels_path(allocation.worksheet_part)
    pivot_rid = part_plan.reserve_rid(
        sheet_rels,
        archive.read(sheet_rels) if sheet_rels in names else None)
    part_plan.add_part(
        allocation.records_part,
        payloads.cache_records,
        content_type=RecordList.mime_type,
        relate_from=allocation.cache_part,
        rel_type=RecordList.rel_type,
        rel_id=allocation.records_relationship_id,
    )
    part_plan.add_part(
        allocation.cache_part,
        payloads.cache_definition,
        content_type=CacheDefinition.mime_type,
        relate_from=context.workbook_part,
        rel_type=CacheDefinition.rel_type,
        rel_id=cache_rid,
    )
    part_plan.add_part(
        allocation.pivot_part,
        payloads.pivot_table,
        content_type=TableDefinition.mime_type,
        relate_from=allocation.worksheet_part,
        rel_type=TableDefinition.rel_type,
        rel_id=pivot_rid,
    )
    part_plan.rel_appends.setdefault(
        allocation.pivot_rels_part, []).append((
            allocation.pivot_cache_relationship_id,
            CacheDefinition.rel_type,
            _relative_target(allocation.pivot_part, allocation.cache_part),
            None,
        ))
    return ((allocation.cache_id, cache_rid),)


def _plan_isolate(context, operation):
    """Add a dedicated cache and retarget only the selected pivot-to-cache edge."""
    allocation = operation.allocation
    payloads = operation.payloads
    part_plan = context.part_plan
    archive = context.archive
    names = context.names
    for name in (allocation.cache_part, allocation.records_part,
                 allocation.cache_rels_part):
        if name in names or name in part_plan.added or name in part_plan.replaced:
            raise RelationshipPolicyError(
                "planned isolated cache part %r already exists in the package"
                % name,
                kind="invalid-pivot-graph",
                anchor=name,
            )
    context.part_plan.replace_part(allocation.pivot_part, payloads.pivot_table)
    original_cache = getattr(operation, "original_cache_part", None)
    part_plan.retarget_rel(
        allocation.pivot_rels_part,
        allocation.pivot_cache_relationship_id,
        _relative_target(allocation.pivot_part, allocation.cache_part),
        expected_type=CacheDefinition.rel_type,
        expected_target=original_cache,
    )
    cache_rid = part_plan.reserve_rid(
        context.workbook_rels_part,
        archive.read(context.workbook_rels_part)
        if context.workbook_rels_part in names else None)
    part_plan.add_part(
        allocation.records_part,
        payloads.cache_records,
        content_type=RecordList.mime_type,
        relate_from=allocation.cache_part,
        rel_type=RecordList.rel_type,
        rel_id=allocation.records_relationship_id,
    )
    part_plan.add_part(
        allocation.cache_part,
        payloads.cache_definition,
        content_type=CacheDefinition.mime_type,
        relate_from=context.workbook_part,
        rel_type=CacheDefinition.rel_type,
        rel_id=cache_rid,
    )
    removals = _drop_original_cache_if_unused(context, operation)
    return ((allocation.cache_id, cache_rid),), removals


def _plan_replace(context, operation):
    allocation = operation.allocation
    payloads = operation.payloads
    context.part_plan.replace_part(allocation.pivot_part, payloads.pivot_table)
    if getattr(operation, "cache_rebuild", False):
        cache_payload = payloads.cache_definition
        if allocation.cache_part in context.ledger.pivot_refresh_requests:
            from openpyxl.preserve.pivots import plan_refresh

            replacements = {allocation.cache_part: cache_payload}
            plan_refresh(
                context.archive, (allocation.cache_part,), replacements)
            cache_payload = replacements[allocation.cache_part]
        context.part_plan.replace_part(
            allocation.cache_part, cache_payload)
        context.part_plan.replace_part(
            allocation.records_part, payloads.cache_records)


def _plan_remove(context, operation):
    allocation = operation.allocation
    sheet_rels = _rels_path(allocation.worksheet_part)
    workbook_rels = context.workbook_rels_part
    context.part_plan.remove_part(
        allocation.pivot_part,
        referencing_rels=((sheet_rels, allocation.pivot_part),),
    )
    strategy = getattr(operation, "publication_strategy", None)
    cache_exists = (
        allocation.cache_part in context.names
        or allocation.cache_part in context.part_plan.added
    )
    if strategy == "shared-isolation" and not cache_exists:
        _drop_original_cache_if_unused(context, operation)
        return
    if cache_exists:
        context.part_plan.remove_part(
            allocation.cache_part,
            referencing_rels=((workbook_rels, allocation.cache_part),),
        )
    if allocation.records_part and (
            allocation.records_part in context.names
            or allocation.records_part in context.part_plan.added):
        context.part_plan.remove_part(allocation.records_part)
    if strategy == "shared-isolation":
        _drop_original_cache_if_unused(context, operation)


def _departed_original_cache_parts(ledger, original_cache):
    departed = set()
    for operation in getattr(ledger, "pivot_operations", {}).values():
        if getattr(operation, "noop", False):
            continue
        if getattr(operation, "original_cache_part", None) == original_cache:
            departed.add(operation.allocation.pivot_part)
    return departed


def _original_cache_survivors(context, original_cache):
    """Pivots or unexpected consumers that still need the original cache."""
    from openpyxl.pivot.adopt_depend import _ALLOWED_SELECTED_RELS
    from openpyxl.pivot.graph import load_workbook_pivot_graph

    graph = load_workbook_pivot_graph(context.workbook)
    departed = _departed_original_cache_parts(
        context.ledger, original_cache)
    survivors = [
        node for node in graph.pivots
        if node.cache_definition_part == original_cache
        and node.identity.pivot_part not in departed
    ]
    if survivors:
        return survivors
    for rel in graph.incoming_relationships.get(original_cache, ()):
        if rel.owner_part == graph.workbook_part:
            continue
        if rel.owner_part in departed:
            continue
        if rel.rel_type not in _ALLOWED_SELECTED_RELS:
            return [rel]
        if rel.rel_type.endswith("/pivotCacheDefinition"):
            return [rel]
    return []


def _drop_original_cache_if_unused(context, operation):
    """Remove the original shared cache only when no survivor remains."""
    original = getattr(operation, "original_cache_part", None)
    if not original or original in context.part_plan.dropped:
        return []
    survivors = _original_cache_survivors(context, original)
    if survivors:
        if getattr(operation, "remove_original_cache", False):
            raise RelationshipPolicyError(
                "cannot remove original shared cache %r; remaining "
                "references still exist. Nothing was written." % original,
                kind="invalid-pivot-graph",
                anchor=original,
            )
        return []
    records = getattr(operation, "original_records_part", None)
    context.part_plan.remove_part(
        original,
        referencing_rels=((context.workbook_rels_part, original),),
    )
    if records and records not in context.part_plan.dropped \
            and records in context.names:
        context.part_plan.remove_part(records)
    if getattr(operation, "original_cache_id", None) is not None:
        return [int(operation.original_cache_id)]
    return []


def splice_workbook_caches(payload, entries):
    """Insert or append ``<pivotCaches>`` entries in workbook schema order."""
    if not entries:
        return payload
    root = crosspart.scan_small(payload, "workbook", max_depth=2)
    by_local = {}
    for child in root.children:
        by_local.setdefault(child.local(), []).append(child)
    rel_ns = (
        b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/'
        b'2006/relationships"'
    )
    children = b"".join(
        b'<pivotCache%s cacheId="%s" r:id="%s"/>' % (
            rel_ns, str(cache_id).encode("ascii"), rid.encode("ascii"))
        for cache_id, rid in entries
    )
    existing = by_local.get("pivotCaches")
    if not existing:
        container = (
            b'<pivotCaches%s count="%d">%s</pivotCaches>'
            % (rel_ns, len(entries), children)
        )
        return crosspart.apply_edits(
            payload, [crosspart._wb_insert_edit(
                root, by_local, "pivotCaches", container)])
    edits = crosspart._append_with_count(
        payload, existing[0], children, len(entries))
    return crosspart.apply_edits(payload, edits)


def apply_workbook_cache_registry(context, workbook_xml):
    """Splice staged cache registry entries into workbook.xml bytes."""
    payload = workbook_xml
    if payload is None:
        payload = context.archive.read(context.workbook_part)
    removals = getattr(context.part_plan, "pivot_cache_removals", ())
    if removals:
        payload = drop_workbook_caches(payload, removals)
    entries = getattr(context.part_plan, "pivot_cache_registry", ())
    if not entries:
        return payload
    return splice_workbook_caches(payload, entries)


def drop_workbook_caches(payload, cache_ids):
    """Remove ``pivotCache`` registry rows for deleted Paper-owned caches."""
    wanted = {str(cache_id) for cache_id in cache_ids}
    root = crosspart.scan_small(payload, "workbook", max_depth=3)
    container = None
    for child in root.children:
        if child.local() == "pivotCaches":
            container = child
            break
    if container is None:
        return payload
    edits = []
    kept = 0
    for child in container.children:
        if child.local() != "pivotCache":
            continue
        cache_id = child.attrs.get("cacheId")
        if cache_id in wanted:
            edits.append((child.start, child.end, b""))
        else:
            kept += 1
    if not edits:
        return payload
    if kept == 0:
        return crosspart.apply_edits(
            payload, [(container.start, container.end, b"")])
    if "count" in container.attrs:
        edits.append(crosspart._patch_attr(
            payload, container, "count", str(kept)))
    return crosspart.apply_edits(payload, edits)


def _next_numbered(prefix, reserved):
    number = 1
    while True:
        name = "%s%d.xml" % (prefix, number)
        rels = _rels_path(name)
        if name not in reserved and rels not in reserved:
            return name
        number += 1


def _next_cache_id(used):
    # Excel allocates positive cache IDs. Although the schema type is an
    # unsigned integer, Excel repairs a newly authored cache with ID zero.
    cache_id = 1
    while cache_id in used:
        cache_id += 1
    return cache_id


def _validate_allocation_free(allocation, names, part_plan):
    for name in allocation.owned_parts():
        if name in names or name in part_plan.added or name in getattr(
                part_plan, "replaced", {}):
            raise RelationshipPolicyError(
                "planned pivot part %r already exists in the package"
                % name,
                kind="invalid-pivot-graph",
                anchor=name,
            )


def _relative_target(from_part, to_part):
    from openpyxl.preserve.lifecycle import _relative_target as resolve
    return resolve(from_part, to_part)
