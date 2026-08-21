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
class PivotAllocation:
    pivot_part: str
    cache_part: str
    records_part: str
    cache_id: int
    worksheet_part: str
    workbook_part: str

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


def plan_creates(context):
    """Add staged create payloads to the save ``PartPlan``."""
    ledger = context.ledger
    operations = getattr(ledger, "pivot_operations", None)
    if not operations:
        return
    part_plan = context.part_plan
    archive = context.archive
    names = context.names
    registry = []
    for operation in operations.values():
        if operation.kind != "create":
            continue
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
            rel_id="rId1",
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
                "rId1",
                CacheDefinition.rel_type,
                _relative_target(allocation.pivot_part, allocation.cache_part),
                None,
            ))
        registry.append((allocation.cache_id, cache_rid))
    part_plan.pivot_cache_registry = tuple(registry)


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
    entries = getattr(context.part_plan, "pivot_cache_registry", ())
    if not entries:
        return workbook_xml
    payload = workbook_xml
    if payload is None:
        payload = context.archive.read(context.workbook_part)
    return splice_workbook_caches(payload, entries)


def _next_numbered(prefix, reserved):
    number = 1
    while True:
        name = "%s%d.xml" % (prefix, number)
        rels = _rels_path(name)
        if name not in reserved and rels not in reserved:
            return name
        number += 1


def _next_cache_id(used):
    cache_id = 0
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
