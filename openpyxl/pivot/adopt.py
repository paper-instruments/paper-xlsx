# paper-xlsx: dedicated-cache foreign PivotTable adoption

"""Convert one qualified foreign pivot into the Paper-managed lifecycle.

``adopt()`` re-runs qualification against current state, builds the complete
replacement plan, then mutates cells and the ledger. Dedicated caches are
replaced in place. Shared caches allocate a new dedicated closure and
retarget only the selected pivot-to-cache edge.
"""

from __future__ import annotations

import hashlib
import io
import uuid
import zipfile
from dataclasses import dataclass

from openpyxl.errors import (
    BoundaryViolationError,
    RelationshipPolicyError,
    UnsupportedStructureError,
)
from openpyxl.pivot.adopt_qualify import analyze_adoption
from openpyxl.pivot.build import build_pivot_payloads
from openpyxl.pivot.calculate import snapshot_for_pivot
from openpyxl.pivot.create import (
    PivotCreateOperation,
    _assert_output_legal,
    _capture_cell_payloads,
    _cell_snapshots,
    _checkpoint,
    _content_identity,
    _restore_cells,
    _restore_ledger_cells,
    _restore_styles,
    _write_output_cells,
    staged_qualification,
)
from openpyxl.pivot.graph import PivotIdentity, load_workbook_pivot_graph
from openpyxl.pivot.plan import plan_pivot
from openpyxl.pivot.qualify import (
    QualificationReason,
    _snapshot_from_cache_package,
)
from openpyxl.preserve.pivotgraph import (
    PivotAllocation,
    allocate_isolated_cache,
)


@dataclass(frozen=True)
class ForeignPivotAdoptionPlan:
    """Immutable operation-scoped adoption plan. Not a public escape hatch."""

    original: PivotIdentity
    strategy: str
    spec: object
    source_identity: str
    persisted_cache_identity: str
    source_changed_since_cached_snapshot: bool
    old_output_cells: tuple
    allocation: PivotAllocation
    payloads: object
    current_plan: object
    payload_hashes: tuple
    original_cache_id: object


_REASON_KINDS = {
    "foreign-output-unproved": ("BoundaryViolationError", "pivot-output-collision"),
    "foreign-cache-isolation-unproved": (
        "RelationshipPolicyError", "pivot-cache-shared"),
    "invalid-pivot-graph": ("RelationshipPolicyError", "invalid-pivot-graph"),
    "foreign-graph-changed": ("UnsupportedStructureError", "stale-pivot"),
    "foreign-relationship-changed": (
        "RelationshipPolicyError", "invalid-pivot-graph"),
    "foreign-core-semantics-unclassified": (
        "UnsupportedStructureError", "unsupported-pivot-feature"),
    "unsupported-extension": (
        "UnsupportedStructureError", "unsupported-pivot-feature"),
    "foreign-dependent-object": (
        "UnsupportedStructureError", "unsupported-pivot-feature"),
    "unsupported-grouping": (
        "UnsupportedStructureError", "unsupported-pivot-feature"),
    "unsupported-calculated": (
        "UnsupportedStructureError", "unsupported-pivot-feature"),
}


def adopt_pivot(handle):
    """Stage dedicated replacement or shared-cache isolation."""
    from openpyxl.pivot.api import (
        PivotTable,
        invalidate_pivot_overlay,
        require_pivot_inspection,
        _session_for,
    )
    from openpyxl.pivot.mutate import (
        _Cell,
        _clear_coordinates,
        _clear_obsolete,
        _owned_from_plan_cells,
    )

    worksheet = handle._worksheet
    require_pivot_inspection(worksheet, api="PivotTable.adopt")
    state = handle._state()
    if state.qualification.origin == "paper":
        return handle

    analysis = analyze_adoption(handle)
    public = analysis.public
    if not public.eligible:
        raise _refusal_for(public.reasons)
    if public.strategy not in ("dedicated-replacement", "shared-isolation"):
        raise RelationshipPolicyError(
            "this pivot cannot be adopted with a known publication "
            "strategy. Nothing was changed.",
            kind="pivot-cache-shared",
            options=list(analysis.cache_dependents),
            anchor=handle._identity.pivot_part,
        )

    workbook = worksheet.parent
    ledger = workbook._paper_ledger
    graph = load_workbook_pivot_graph(workbook)
    node = graph.pivots_by_identity.get(handle._identity)
    if node is None:
        raise UnsupportedStructureError(
            "the selected pivot graph changed before adoption. "
            "Nothing was changed.",
            kind="stale-pivot",
            anchor=handle._identity.pivot_part,
        )
    cache = graph.caches_by_part.get(node.cache_definition_part)
    spec = state.projection.spec
    if spec is None:
        raise UnsupportedStructureError(
            "this foreign pivot cannot be adopted "
            "(foreign-semantic-incomplete). Nothing was changed.",
            kind="unsupported-pivot-operation",
            anchor=handle._identity.pivot_part,
        )

    _checkpoint("start", workbook)
    snapshot = snapshot_for_pivot(workbook, spec.source)
    plan = plan_pivot(spec, snapshot)
    persisted = _snapshot_from_cache_package(
        getattr(workbook, "_paper_source", None), node, state.projection)
    persisted_identity = _content_identity(persisted)
    current_identity = _content_identity(snapshot)
    old_cells = _owned_from_plan_cells(plan_pivot(spec, persisted).output.cells)
    _assert_output_legal(
        worksheet, plan, snapshot, graph, ledger,
        ignore_coordinates=tuple(
            (row, column) for row, column, _value, _role in old_cells),
        ignore_name=state.name,
    )
    _checkpoint("validated", workbook)

    allocation, isolation = _allocation_for_strategy(
        public.strategy, workbook, ledger, graph, node, cache)
    payloads = build_pivot_payloads(
        plan,
        allocation.cache_id,
        workbook,
        records_relationship_id=allocation.records_relationship_id,
        pivot_cache_relationship_id=allocation.pivot_cache_relationship_id,
    )
    adoption_plan = ForeignPivotAdoptionPlan(
        original=handle._identity,
        strategy=public.strategy,
        spec=spec,
        source_identity=current_identity,
        persisted_cache_identity=persisted_identity,
        source_changed_since_cached_snapshot=(
            persisted_identity != current_identity),
        old_output_cells=old_cells,
        allocation=allocation,
        payloads=payloads,
        current_plan=plan,
        payload_hashes=analysis.payload_hashes,
        original_cache_id=int(node.cache_id),
    )
    _checkpoint("planned", workbook)

    dirty_before = set(ledger.dirty_coordinates(worksheet))
    overwrites_before = set(ledger.value_overwrites.get(worksheet, ()))
    ops_before = dict(ledger.pivot_operations)
    cell_snapshots = _cell_snapshots(worksheet, plan.output.cells)
    old_snapshots = _cell_snapshots(
        worksheet,
        [_Cell(row, column, value, "value")
         for row, column, value, _role in old_cells],
    )
    formats_before = tuple(workbook._number_formats)
    styles_before = tuple(workbook._cell_styles)
    try:
        _write_output_cells(worksheet, plan.output.cells)
        _clear_obsolete(worksheet, old_cells, plan)
        operation = PivotCreateOperation(
            kind="adopt",
            session_id=uuid.uuid4().hex,
            name=spec.name,
            sheet=worksheet.title,
            spec=spec,
            source_identity=current_identity,
            allocation=allocation,
            payloads=payloads,
            output_range=plan.output.ref,
            output_cells=tuple(
                (cell.row, cell.column, cell.value, cell.role)
                for cell in plan.output.cells
            ),
            owned_coordinates=tuple(
                (cell.row, cell.column) for cell in plan.output.cells
                if cell.value is not None
            ),
            clear_coordinates=_clear_coordinates(old_cells, plan),
            replace_existing=True,
            relationship_id=node.identity.relationship_id,
            cache_rebuild=True,
            validate_source_identity=True,
            source_snapshot=snapshot,
            published_cell_payloads=_capture_cell_payloads(
                worksheet,
                tuple((cell.row, cell.column) for cell in plan.output.cells
                      if cell.value is not None),
            ),
            rollback_cells=cell_snapshots + old_snapshots,
            rollback_dirty_coordinates=tuple(sorted(dirty_before)),
            rollback_value_overwrites=tuple(sorted(overwrites_before)),
            rollback_number_formats=formats_before,
            rollback_cell_styles=styles_before,
            qualification=staged_qualification(),
            worksheet=worksheet,
            semantic_effects=("adopt",),
            baseline_spec=spec,
            origin_before="foreign",
            publication_strategy=public.strategy,
            final_action="adopt",
            original_payload_hashes=adoption_plan.payload_hashes,
            original_cache_id=int(node.cache_id),
            persisted_cache_identity=persisted_identity,
            **isolation,
        )
        staged = dict(ledger.pivot_operations)
        staged[operation.session_id] = operation
        ledger.pivot_operations = staged
        _checkpoint("ledger", workbook)
    except Exception:
        _restore_cells(worksheet, cell_snapshots + old_snapshots)
        _restore_ledger_cells(ledger, worksheet, dirty_before, overwrites_before)
        ledger.pivot_operations = ops_before
        _restore_styles(workbook, formats_before, styles_before)
        raise

    invalidate_pivot_overlay(workbook)
    identity = PivotIdentity(
        worksheet_part=allocation.worksheet_part,
        pivot_part=allocation.pivot_part,
        relationship_id=node.identity.relationship_id,
        name=spec.name,
    )
    return PivotTable(worksheet, identity, _session_for(workbook).generation)


def validate_adoption_graph(workbook, ledger):
    """Refuse when the original foreign graph drifted after planning."""
    package = getattr(workbook, "_paper_source", None)
    if not package:
        return
    for operation in getattr(ledger, "pivot_operations", {}).values():
        if getattr(operation, "noop", False):
            continue
        if getattr(operation, "origin_before", "paper") != "foreign":
            continue
        hashes = getattr(operation, "original_payload_hashes", ()) or ()
        if not hashes:
            continue
        allocation = operation.allocation
        original_cache = getattr(operation, "original_cache_part", None)
        mapping = {
            "pivot": allocation.pivot_part,
            "cache": original_cache or allocation.cache_part,
        }
        try:
            with zipfile.ZipFile(io.BytesIO(package)) as archive:
                for kind, expected in hashes:
                    part = mapping.get(kind)
                    if not part:
                        continue
                    actual = hashlib.sha256(archive.read(part)).hexdigest()
                    if actual != expected:
                        raise UnsupportedStructureError(
                            "the original pivot graph changed after adoption "
                            "was staged. Nothing was written.",
                            kind="stale-pivot",
                            anchor=part,
                        )
                for part, expected in getattr(
                        operation, "sibling_payload_hashes", ()) or ():
                    actual = hashlib.sha256(archive.read(part)).hexdigest()
                    if actual != expected:
                        raise UnsupportedStructureError(
                            "a shared-cache sibling changed after adoption "
                            "was staged. Nothing was written.",
                            kind="stale-pivot",
                            anchor=part,
                        )
        except KeyError as exc:
            raise UnsupportedStructureError(
                "the original pivot graph changed after adoption "
                "was staged. Nothing was written.",
                kind="stale-pivot",
                anchor=allocation.pivot_part,
            ) from exc
        if getattr(operation, "publication_strategy", None) != "shared-isolation":
            continue
        graph = load_workbook_pivot_graph(workbook)
        planned = set(getattr(operation, "sibling_parts", ()) or ())
        current = {
            node.identity.pivot_part
            for node in graph.pivots
            if node.cache_definition_part == (
                original_cache or allocation.cache_part)
            and node.identity.pivot_part != allocation.pivot_part
        }
        if planned != current:
            raise UnsupportedStructureError(
                "the shared-cache sibling set changed after adoption "
                "was staged. Nothing was written.",
                kind="stale-pivot",
                anchor=original_cache or allocation.pivot_part,
            )
        from openpyxl.pivot.adopt_depend import _ALLOWED_SELECTED_RELS

        watched = {allocation.pivot_part}
        if original_cache:
            watched.add(original_cache)
        records = getattr(operation, "original_records_part", None)
        if records:
            watched.add(records)
        for part in watched:
            for rel in graph.incoming_relationships.get(part, ()):
                if rel.rel_type not in _ALLOWED_SELECTED_RELS:
                    raise RelationshipPolicyError(
                        "an unexpected relationship now references the "
                        "selected pivot graph. Nothing was written.",
                        kind="invalid-pivot-graph",
                        anchor=part,
                    )


def _allocation_for_strategy(strategy, workbook, ledger, graph, node, cache):
    """Bind either the existing dedicated closure or a new isolated cache."""
    if cache is None or not node.cache_definition_part:
        raise RelationshipPolicyError(
            "cannot adopt a pivot without a resolved cache. "
            "Nothing was changed.",
            kind="invalid-pivot-graph",
            anchor=node.identity.pivot_part,
        )
    if strategy == "dedicated-replacement":
        if (not cache.records_part or not node.cache_relationship_id
                or not cache.records_relationship_id):
            raise RelationshipPolicyError(
                "cannot resolve the pivot's internal cache relationships. "
                "Nothing was changed.",
                kind="invalid-pivot-graph",
                anchor=node.identity.pivot_part,
            )
        allocation = PivotAllocation(
            pivot_part=node.identity.pivot_part,
            cache_part=node.cache_definition_part,
            records_part=cache.records_part,
            cache_id=int(node.cache_id),
            worksheet_part=node.identity.worksheet_part,
            workbook_part=graph.workbook_part,
            pivot_cache_relationship_id=node.cache_relationship_id,
            records_relationship_id=cache.records_relationship_id,
        )
        return allocation, {}
    if strategy == "shared-isolation":
        isolated = allocate_isolated_cache(workbook, graph, ledger)
        allocation = PivotAllocation(
            pivot_part=node.identity.pivot_part,
            cache_part=isolated.cache_part,
            records_part=isolated.records_part,
            cache_id=isolated.cache_id,
            worksheet_part=node.identity.worksheet_part,
            workbook_part=graph.workbook_part,
            pivot_cache_relationship_id=node.cache_relationship_id,
            records_relationship_id=isolated.records_relationship_id,
        )
        return allocation, _isolation_fields(
            workbook, ledger, graph, node, cache)
    raise RelationshipPolicyError(
        "this pivot cannot be adopted with a known publication "
        "strategy. Nothing was changed.",
        kind="pivot-cache-shared",
        anchor=node.identity.pivot_part,
    )


def _isolation_fields(workbook, ledger, graph, node, cache):
    """Capture sibling identity/hash bindings for shared-cache isolation."""
    original_cache = node.cache_definition_part
    siblings = [
        other for other in graph.pivots
        if other.cache_definition_part == original_cache
        and other.identity.pivot_part != node.identity.pivot_part
    ]
    staged = {node.identity.pivot_part}
    for operation in getattr(ledger, "pivot_operations", {}).values():
        if getattr(operation, "publication_strategy", None) != "shared-isolation":
            continue
        if getattr(operation, "original_cache_part", None) == original_cache:
            staged.add(operation.allocation.pivot_part)
    remaining = [
        sibling for sibling in siblings
        if sibling.identity.pivot_part not in staged
    ]
    hashes = []
    source = getattr(workbook, "_paper_source", None)
    if source:
        with zipfile.ZipFile(io.BytesIO(source)) as archive:
            for sibling in siblings:
                part = sibling.identity.pivot_part
                digest = sibling.payload_sha256 or hashlib.sha256(
                    archive.read(part)).hexdigest()
                hashes.append((part, digest))
    else:
        hashes = [
            (sibling.identity.pivot_part, sibling.payload_sha256)
            for sibling in siblings
            if sibling.payload_sha256
        ]
    return {
        "original_cache_part": original_cache,
        "original_records_part": None if cache is None else cache.records_part,
        "sibling_parts": tuple(
            sibling.identity.pivot_part for sibling in siblings),
        "sibling_identities": tuple(
            (sibling.identity.name, sibling.sheet_title)
            for sibling in siblings),
        "sibling_payload_hashes": tuple(hashes),
        "remove_original_cache": not remaining,
    }


def _refusal_for(reasons):
    reason = reasons[0] if reasons else QualificationReason(
        None, "unsupported-pivot-operation")
    family, kind = _REASON_KINDS.get(
        reason.code,
        ("UnsupportedStructureError", "unsupported-pivot-operation"),
    )
    message = (
        "this foreign pivot cannot be adopted (%s). Nothing was changed."
        % reason.code
    )
    context = dict(reason.context)
    if family == "BoundaryViolationError":
        return BoundaryViolationError(
            message, kind=kind, anchor=context.get("part"))
    if family == "RelationshipPolicyError":
        return RelationshipPolicyError(
            message, kind=kind, anchor=context.get("part"))
    return UnsupportedStructureError(
        message, kind=kind, anchor=context.get("part"),
        options=[reason.code],
    )
