# paper-xlsx: refresh, repoint, move, and later lifecycle rebuilds

"""Rebuild a Paper-managed dedicated-cache pivot from a complete spec.

Shared-cache and foreign pivots refuse before any cell or ledger mutation.
In-session creates are updated in place so save still adds one graph.
"""

from __future__ import annotations

import io
import uuid
import zipfile
from dataclasses import replace

from openpyxl.errors import (
    BoundaryViolationError,
    RelationshipPolicyError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from openpyxl.pivot.api_types import PivotSource, PivotSpec
from openpyxl.pivot.build import PivotPayloads, build_pivot_payloads
from openpyxl.pivot.calculate import snapshot_for_pivot
from openpyxl.pivot.create import (
    PivotCreateOperation,
    _assert_name_available,
    _assert_output_legal,
    _capture_cell_payloads,
    _cell_snapshots,
    _checkpoint,
    _content_identity,
    _restore_cells,
    _restore_ledger_cells,
    _restore_styles,
    _validate_published_cells,
    _write_output_cells,
)
from openpyxl.pivot.graph import PivotIdentity, load_workbook_pivot_graph
from openpyxl.pivot.plan import plan_pivot
from openpyxl.preserve.pivotgraph import PivotAllocation


def refresh_pivot(handle):
    """Rebuild cache records and visible output from the current source."""
    return _rebuild(handle, kind="refresh", allow_self_overlap=True)


def repoint_pivot(handle, source, spec=None):
    """Point a dedicated-cache pivot at a new table or range source."""
    state = handle._state()
    _prepare_mutation(state, handle, "repoint", "can_repoint_source")
    new_source = PivotSource.parse(source)
    current = state.projection.spec
    if spec is None:
        spec = replace(current, source=new_source)
    elif not isinstance(spec, PivotSpec):
        raise TypeError("spec must be a PivotSpec")
    else:
        spec = replace(spec, source=new_source)
    return _rebuild(handle, kind="repoint", spec=spec, allow_self_overlap=True)


def move_pivot(handle, destination, destination_sheet=None):
    """Move a managed pivot's output on the same worksheet."""
    if destination_sheet is not None:
        raise UnsupportedStructureError(
            "v1 supports same-sheet pivot moves only. Nothing was changed.",
            kind="unsupported-pivot-operation",
            options=["destination_sheet"],
        )
    state = handle._state()
    _prepare_mutation(state, handle, "move", "can_move")
    spec = replace(state.projection.spec, destination=destination)
    return _rebuild(handle, kind="move", spec=spec, allow_self_overlap=True)


def update_pivot(handle, **changes):
    """Replace the complete spec with only the supplied fields changed."""
    state = handle._state()
    _prepare_mutation(state, handle, "update", "can_edit_layout")
    if not changes:
        return handle
    if set(changes) == {"name"}:
        return rename_pivot(handle, changes["name"])
    spec = _spec_with_changes(state.projection.spec, changes)
    return _rebuild(handle, kind="update", spec=spec, allow_self_overlap=True)


def rename_pivot(handle, name):
    """Patch only the PivotTable name, retaining every other XML byte."""
    from openpyxl.pivot.api import (
        PivotTable,
        invalidate_pivot_overlay,
        require_pivot_inspection,
        _session_for,
    )

    worksheet = handle._worksheet
    require_pivot_inspection(worksheet, api="PivotTable.rename")
    state = handle._state()
    _prepare_mutation(state, handle, "rename", "can_rename", refuse_shared=False)
    if name == state.name:
        return handle
    workbook = worksheet.parent
    _assert_consumers_allow(workbook, handle, "rename")
    ledger = workbook._paper_ledger
    graph = load_workbook_pivot_graph(workbook)
    staged = _find_staged(ledger, handle)
    if staged is not None:
        _validate_published_cells(workbook, staged)
    _assert_name_available(name, graph, ledger, ignore_name=state.name)
    node = _node_for(handle, graph) if staged is None else None
    payloads = _renamed_payloads(workbook, staged, node, name)
    spec = replace(state.projection.spec, name=name)
    allocation = staged.allocation if staged is not None else \
        _allocation_from_handle(workbook, worksheet, handle)
    replacing_create = staged is not None and staged.kind == "create" \
        and not staged.replace_existing
    baseline_spec = None if replacing_create else (
        staged.baseline_spec if staged is not None
        and staged.baseline_spec is not None else state.projection.spec)
    operation = PivotCreateOperation(
        kind="create" if replacing_create else "rename",
        session_id=staged.session_id if replacing_create else uuid.uuid4().hex,
        name=name,
        sheet=worksheet.title,
        spec=spec,
        source_identity=staged.source_identity if staged is not None else "",
        allocation=allocation,
        payloads=payloads,
        output_range=staged.output_range if staged is not None
        else state.projection.output_range,
        output_cells=staged.output_cells if staged is not None else (),
        owned_coordinates=staged.owned_coordinates if staged is not None else (),
        clear_coordinates=staged.clear_coordinates if staged is not None else (),
        replace_existing=False if replacing_create else True,
        relationship_id=staged.relationship_id if staged is not None
        else handle._identity.relationship_id,
        cache_rebuild=staged.cache_rebuild if staged is not None else False,
        validate_source_identity=(
            staged.validate_source_identity if staged is not None else False),
        source_snapshot=staged.source_snapshot if staged is not None else None,
        published_cell_payloads=(
            staged.published_cell_payloads if staged is not None else ()),
        rollback_cells=staged.rollback_cells if staged is not None else (),
        rollback_dirty_coordinates=(
            staged.rollback_dirty_coordinates if staged is not None else ()),
        rollback_value_overwrites=(
            staged.rollback_value_overwrites if staged is not None else ()),
        rollback_number_formats=(
            staged.rollback_number_formats if staged is not None else ()),
        rollback_cell_styles=(
            staged.rollback_cell_styles if staged is not None else ()),
        qualification=state.qualification,
        worksheet=worksheet,
        semantic_effects=(
            ("create",) if replacing_create else _net_semantic_effects(
                staged, baseline_spec, spec, "rename")
        ),
        baseline_spec=baseline_spec,
        **_publication_fields(staged),
    )
    operations = dict(ledger.pivot_operations)
    if staged is not None:
        operations.pop(staged.session_id, None)
    operations[operation.session_id] = operation
    ledger.pivot_operations = operations
    invalidate_pivot_overlay(workbook)
    identity = _staged_identity_for_operation(operation, handle)
    return PivotTable(worksheet, identity, _session_for(workbook).generation)


def delete_pivot(handle):
    state = handle._state()
    _prepare_mutation(state, handle, "delete", "can_delete")
    _assert_consumers_allow(handle._worksheet.parent, handle, "delete")
    worksheet = handle._worksheet
    workbook = worksheet.parent
    ledger = workbook._paper_ledger
    staged = _find_staged(ledger, handle)
    old_cells = _owned_cells(state, staged, worksheet)
    snapshots = _cell_snapshots(
        worksheet,
        [_Cell(row, column, value, "value")
         for row, column, value, _role in old_cells],
    )
    dirty_before = set(ledger.dirty_coordinates(worksheet))
    overwrites_before = set(ledger.value_overwrites.get(worksheet, ()))
    ops_before = dict(ledger.pivot_operations)
    try:
        if staged is not None and staged.kind == "create" and not staged.replace_existing:
            _drop_in_session_create(worksheet, ledger, staged)
        else:
            _clear_cells(worksheet, old_cells)
            clear_coordinates = {
                (row, column) for row, column, _value, _role in old_cells
            }
            if staged is not None:
                clear_coordinates.update(staged.clear_coordinates)
            clear_coordinates = tuple(sorted(clear_coordinates))
            allocation = staged.allocation if staged is not None else _allocation_from_handle(
                workbook, worksheet, handle)
            operation = PivotCreateOperation(
                kind="delete",
                session_id=uuid.uuid4().hex,
                name=state.name,
                sheet=worksheet.title,
                spec=state.projection.spec,
                source_identity=staged.source_identity if staged else "",
                allocation=allocation,
                payloads=None,
                output_range=state.projection.output_range,
                output_cells=(),
                owned_coordinates=(),
                clear_coordinates=clear_coordinates,
                replace_existing=True,
                relationship_id=handle._identity.relationship_id,
                validate_source_identity=False,
                published_cell_payloads=_capture_cell_payloads(
                    worksheet, (),
                ),
                qualification=state.qualification,
                worksheet=worksheet,
                semantic_effects=("delete",),
                baseline_spec=(
                    staged.baseline_spec if staged is not None
                    and staged.baseline_spec is not None
                    else state.projection.spec),
                **_publication_fields(staged, final_action="delete"),
            )
            ops = dict(ledger.pivot_operations)
            if staged is not None:
                ops.pop(staged.session_id, None)
            ops[operation.session_id] = operation
            ledger.pivot_operations = ops
    except Exception:
        _restore_cells(worksheet, snapshots)
        _restore_ledger_cells(ledger, worksheet, dirty_before, overwrites_before)
        ledger.pivot_operations = ops_before
        raise
    from openpyxl.pivot.api import invalidate_pivot_overlay
    invalidate_pivot_overlay(workbook)
    return None


def _rebuild(handle, kind, spec=None, allow_self_overlap=False):
    from openpyxl.pivot.api import (
        PivotTable,
        invalidate_pivot_overlay,
        require_pivot_inspection,
        _session_for,
    )

    worksheet = handle._worksheet
    require_pivot_inspection(worksheet, api="PivotTable.%s" % kind)
    state = handle._state()
    capability = _REBUILD_CAPABILITIES.get(kind)
    if capability is not None:
        _prepare_mutation(
            state, handle, kind, capability,
            refuse_shared=kind != "rename")
    spec = spec or state.projection.spec
    workbook = worksheet.parent
    if kind == "move":
        _assert_consumers_allow(workbook, handle, "move")
    elif kind == "update" and state.projection.spec is not None:
        verbs = []
        if spec.destination != state.projection.spec.destination:
            verbs.append("move")
        if spec.name != state.name:
            verbs.append("rename")
        if verbs:
            _assert_consumers_allow(workbook, handle, *verbs)
    ledger = workbook._paper_ledger
    staged = _find_staged(ledger, handle)
    if staged is not None:
        _validate_published_cells(workbook, staged)
    graph = load_workbook_pivot_graph(workbook)
    snapshot = _snapshot_for_rebuild(
        kind, workbook, state, staged, handle, spec)
    plan = plan_pivot(spec, snapshot)
    old_cells = _owned_cells(state, staged, worksheet)
    _assert_output_legal(
        worksheet, plan, snapshot, graph, ledger,
        ignore_coordinates=tuple((row, column) for row, column, _v, _r in old_cells)
        if allow_self_overlap else (),
        ignore_name=state.name if allow_self_overlap or kind != "create" else None,
    )
    if spec.name.casefold() != (state.name or "").casefold():
        _assert_name_available(spec.name, graph, ledger, ignore_name=state.name)
    noop = _is_noop(
        kind, state, plan, worksheet, spec, snapshot, staged, workbook, handle)
    if noop and kind == "refresh":
        return handle
    allocation = staged.allocation if staged is not None else _allocation_from_handle(
        workbook, worksheet, handle)
    replace_existing = staged is None or staged.replace_existing
    relationship_id = handle._identity.relationship_id
    if staged is not None and staged.kind == "create" and not staged.replace_existing:
        replace_existing = False
        relationship_id = staged.relationship_id

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
        _checkpoint("planned", workbook)
        payloads = build_pivot_payloads(
            plan,
            staged.allocation.cache_id if staged is not None else int(
                _node_for(handle, graph).cache_id),
            workbook,
            records_relationship_id=allocation.records_relationship_id,
            pivot_cache_relationship_id=allocation.pivot_cache_relationship_id,
        )
        _write_output_cells(worksheet, plan.output.cells)
        _clear_obsolete(worksheet, old_cells, plan)
        clear_coordinates = set(_clear_coordinates(old_cells, plan))
        if staged is not None:
            clear_coordinates.update(staged.clear_coordinates)
        clear_coordinates = tuple(sorted(clear_coordinates))
        rollback_cells = ()
        rollback_dirty = ()
        rollback_overwrites = ()
        if staged is not None and staged.kind == "create" \
                and not staged.replace_existing:
            rollback_cells = _merge_snapshots(
                staged.rollback_cells, cell_snapshots, old_snapshots)
            rollback_dirty = staged.rollback_dirty_coordinates
            rollback_overwrites = staged.rollback_value_overwrites
        cache_rebuild = kind in ("refresh", "repoint", "update")
        if staged is not None:
            cache_rebuild = cache_rebuild or staged.cache_rebuild
        replacing_create = staged is not None and staged.kind == "create" \
            and not staged.replace_existing
        baseline_spec = None if replacing_create else (
            staged.baseline_spec if staged is not None
            and staged.baseline_spec is not None else state.projection.spec)
        operation = PivotCreateOperation(
            kind="create" if (staged is not None and staged.kind == "create"
                              and not staged.replace_existing) else kind,
            session_id=(staged.session_id if staged is not None
                        and staged.kind == "create" and not staged.replace_existing
                        else uuid.uuid4().hex),
            name=spec.name,
            sheet=worksheet.title,
            spec=spec,
            source_identity=_content_identity(snapshot),
            allocation=allocation,
            payloads=payloads,
            output_range=plan.output.ref,
            output_cells=tuple(
                (cell.row, cell.column, cell.value, cell.role)
                for cell in plan.output.cells
            ),
            owned_coordinates=tuple(
                (cell.row, cell.column) for cell in plan.output.cells
            ),
            clear_coordinates=clear_coordinates,
            noop=False,
            replace_existing=replace_existing,
            relationship_id=relationship_id,
            cache_rebuild=cache_rebuild,
            validate_source_identity=cache_rebuild or not replace_existing,
            source_snapshot=snapshot,
            published_cell_payloads=_capture_cell_payloads(
                worksheet,
                tuple((cell.row, cell.column) for cell in plan.output.cells),
            ),
            rollback_cells=rollback_cells,
            rollback_dirty_coordinates=rollback_dirty,
            rollback_value_overwrites=rollback_overwrites,
            rollback_number_formats=(
                staged.rollback_number_formats if staged is not None else ()),
            rollback_cell_styles=(
                staged.rollback_cell_styles if staged is not None else ()),
            qualification=state.qualification,
            worksheet=worksheet,
            semantic_effects=(
                ("create",) if replacing_create
                else _net_semantic_effects(
                    staged, baseline_spec, spec, kind)
            ),
            baseline_spec=baseline_spec,
            **_publication_fields(staged, final_action=kind),
        )
        ops = dict(ledger.pivot_operations)
        if staged is not None:
            ops.pop(staged.session_id, None)
        ops[operation.session_id] = operation
        ledger.pivot_operations = ops
        _checkpoint("ledger", workbook)
    except Exception:
        _restore_cells(worksheet, cell_snapshots + old_snapshots)
        _restore_ledger_cells(ledger, worksheet, dirty_before, overwrites_before)
        ledger.pivot_operations = ops_before
        _restore_styles(workbook, formats_before, styles_before)
        raise

    invalidate_pivot_overlay(workbook)
    session = _session_for(workbook)
    from openpyxl.pivot.create import _staged_identity
    identity = _staged_identity(operation) if not replace_existing \
        else handle._identity
    if replace_existing:
        identity = PivotIdentity(
            worksheet_part=allocation.worksheet_part,
            pivot_part=allocation.pivot_part,
            relationship_id=relationship_id,
            name=spec.name,
        )
    return PivotTable(worksheet, identity, session.generation)


class _Cell:
    def __init__(self, row, column, value, role):
        self.row = row
        self.column = column
        self.value = value
        self.role = role


_REBUILD_CAPABILITIES = {
    "refresh": "can_headless_refresh",
    "repoint": "can_repoint_source",
    "move": "can_move",
    "update": "can_edit_layout",
    "rename": "can_rename",
}


def _prepare_mutation(state, handle, verb, capability, refuse_shared=True):
    if refuse_shared and state.qualification.origin == "paper":
        _refuse_shared_cache(state, handle, verb)
    _require_capability(state, capability)


def _require_capability(state, name):
    caps = state.qualification.capabilities
    if getattr(caps, name, False):
        return
    if state.qualification.origin == "foreign":
        raise UnsupportedStructureError(
            "this foreign pivot does not support %s; call adopt() first. "
            "Nothing was changed." % name,
            kind="unsupported-pivot-operation",
            options=["adopt", name],
        )
    raise UnsupportedStructureError(
        "this pivot does not support %s. Nothing was changed." % name,
        kind="unsupported-pivot-operation",
        options=[name],
    )


def _refuse_shared_cache(state, handle, verb):
    if not state.qualification.cache_shared:
        return
    scope = list(state.qualification.refresh_on_open_scope)
    raise RelationshipPolicyError(
        "cannot %s a pivot that shares its cache; siblings are %s. "
        "Nothing was changed." % (verb, ", ".join(scope) or "unknown"),
        kind="pivot-cache-shared",
        options=scope,
        anchor=handle._identity.pivot_part,
    )


def _find_staged(ledger, handle):
    identity = handle._identity
    for operation in getattr(ledger, "pivot_operations", {}).values():
        if operation.allocation.pivot_part == identity.pivot_part:
            return operation
        if identity.relationship_id == "staged:%s" % operation.session_id:
            return operation
    return None


def _owned_cells(state, staged, worksheet):
    if staged is not None and staged.output_cells:
        _validate_published_cells(worksheet.parent, staged)
        return tuple(
            cell for cell in staged.output_cells if cell[2] is not None)
    from openpyxl.pivot.qualify import _reconstruct_owned_output

    workbook = worksheet.parent
    graph = load_workbook_pivot_graph(workbook)
    node = graph.pivots_by_identity.get(state.identity)
    if node is None:
        for item in graph.pivots:
            if item.identity.pivot_part == state.identity.pivot_part:
                node = item
                break
    owned = None if node is None else _reconstruct_owned_output(
        workbook, node, state.projection)
    if owned is None:
        raise BoundaryViolationError(
            "cannot prove ownership of pivot output before clearing",
            kind="pivot-output-collision",
            anchor="%s!%s" % (
                worksheet.title, state.projection.output_range or state.name),
        )
    return owned


def _drop_in_session_create(worksheet, ledger, staged):
    """Cancel an unsaved create so save is a true package no-op."""
    _validate_published_cells(worksheet.parent, staged)
    _restore_cells(worksheet, staged.rollback_cells)
    _restore_styles(
        worksheet.parent,
        staged.rollback_number_formats,
        staged.rollback_cell_styles,
    )
    dirty = set(staged.rollback_dirty_coordinates)
    overwrites = set(staged.rollback_value_overwrites)
    if dirty:
        ledger.cells[worksheet] = dirty
    else:
        ledger.cells.pop(worksheet, None)
    if overwrites:
        ledger.value_overwrites[worksheet] = overwrites
    else:
        ledger.value_overwrites.pop(worksheet, None)
    remaining = dict(ledger.pivot_operations)
    remaining.pop(staged.session_id, None)
    ledger.pivot_operations = remaining


def _new_output_values(plan):
    return {(cell.row, cell.column): cell.value for cell in plan.output.cells}


def _clear_coordinates(old_cells, plan):
    new_values = _new_output_values(plan)
    return tuple(
        (row, column) for row, column, _value, _role in old_cells
        if new_values.get((row, column), None) is None
    )


def _clear_obsolete(worksheet, old_cells, plan):
    new_values = _new_output_values(plan)
    remaining = [
        (row, column, value, role)
        for row, column, value, role in old_cells
        if new_values.get((row, column), None) is None
    ]
    _clear_cells(worksheet, remaining)


def _clear_cells(worksheet, cells):
    store = getattr(worksheet, "_cells", {})
    for row, column, expected, _role in cells:
        existing = store.get((row, column))
        if existing is None:
            continue
        if not _same_typed_value(existing.value, expected):
            raise BoundaryViolationError(
                "cannot clear unowned or externally changed pivot output "
                "at %s" % existing.coordinate,
                kind="pivot-output-collision",
                anchor="%s!%s" % (worksheet.title, existing.coordinate),
            )
        if getattr(existing, "comment", None) is not None:
            raise BoundaryViolationError(
                "pivot output at %s has a comment and cannot be cleared"
                % existing.coordinate,
                kind="pivot-output-collision",
                anchor="%s!%s" % (worksheet.title, existing.coordinate),
            )
        if getattr(existing, "_hyperlink", None) is not None:
            raise BoundaryViolationError(
                "pivot output at %s has a hyperlink and cannot be cleared"
                % existing.coordinate,
                kind="pivot-output-collision",
                anchor="%s!%s" % (worksheet.title, existing.coordinate),
            )
        existing.value = None
        existing._style = None


def _is_noop(kind, state, plan, worksheet, spec, snapshot, staged, workbook, handle):
    if kind != "refresh":
        return False
    if state.projection.output_range != plan.output.ref:
        return False
    current_spec = state.projection.spec
    if current_spec is None or current_spec.to_dict() != spec.to_dict():
        return False
    live_identity = _content_identity(snapshot)
    if staged is not None:
        if staged.source_identity != live_identity:
            return False
    else:
        from openpyxl.pivot.qualify import _snapshot_from_cache_package
        from openpyxl.preserve.pivots import source_impacts

        graph = load_workbook_pivot_graph(workbook)
        node = _node_for(handle, graph)
        try:
            persisted = _snapshot_from_cache_package(
                workbook._paper_source, node, state.projection)
        except Exception:
            return False
        if _content_identity(persisted) != live_identity:
            return False
        allocation = _allocation_from_handle(workbook, worksheet, handle)
        if any(impact["part"] == allocation.cache_part
               for impact in source_impacts(workbook, workbook._paper_ledger)):
            return False
    store = getattr(worksheet, "_cells", {})
    for cell in plan.output.cells:
        if cell.value is None:
            continue
        existing = store.get((cell.row, cell.column))
        actual = None if existing is None else existing.value
        if not _same_typed_value(actual, cell.value):
            return False
    return True


def _snapshot_for_rebuild(kind, workbook, state, staged, handle, spec):
    if kind != "move":
        return snapshot_for_pivot(workbook, spec.source)
    if staged is not None and staged.source_snapshot is not None:
        return staged.source_snapshot
    from openpyxl.pivot.qualify import _snapshot_from_cache_package

    graph = load_workbook_pivot_graph(workbook)
    node = _node_for(handle, graph)
    return _snapshot_from_cache_package(
        workbook._paper_source, node, state.projection)


def _merge_snapshots(*groups):
    merged = {}
    for group in groups:
        for row, column, state in group:
            merged.setdefault((row, column), state)
    return tuple(
        (row, column, merged[(row, column)])
        for row, column in sorted(merged)
    )


def _same_typed_value(left, right):
    return type(left) is type(right) and left == right


def _renamed_payloads(workbook, staged, node, name):
    from openpyxl.preserve import crosspart
    from openpyxl.preserve.xmlscan import ScanRefusal

    if staged is not None:
        payloads = staged.payloads
        pivot_xml = payloads.pivot_table
    else:
        try:
            with zipfile.ZipFile(io.BytesIO(workbook._paper_source)) as archive:
                pivot_xml = archive.read(node.identity.pivot_part)
        except (KeyError, OSError, zipfile.BadZipFile) as exc:
            raise UnsupportedStructureError(
                "cannot read the pivot definition for a targeted rename. "
                "Nothing was changed.",
                kind="invalid-pivot-graph",
                anchor=node.identity.pivot_part,
            ) from exc
        payloads = PivotPayloads(b"", b"", pivot_xml)
    try:
        root = crosspart.scan_small(
            pivot_xml, "pivotTableDefinition", max_depth=0,
            allow_prefixed_root=True)
        edit = crosspart._patch_attr(pivot_xml, root, "name", name)
        renamed = crosspart.apply_edits(pivot_xml, [edit])
    except (ScanRefusal, UnsupportedStructureError) as exc:
        raise UnsupportedStructureError(
            "cannot patch this pivot name without reserializing its "
            "definition. Nothing was changed.",
            kind="unsupported-pivot-operation",
            anchor=None if node is None else node.identity.pivot_part,
        ) from exc
    return replace(payloads, pivot_table=renamed)


def _staged_identity_for_operation(operation, handle):
    if operation.kind == "create" and not operation.replace_existing:
        from openpyxl.pivot.create import _staged_identity
        return _staged_identity(operation)
    return PivotIdentity(
        worksheet_part=operation.allocation.worksheet_part,
        pivot_part=operation.allocation.pivot_part,
        relationship_id=operation.relationship_id,
        name=operation.name,
    )


def _allocation_from_handle(workbook, worksheet, handle):
    graph = load_workbook_pivot_graph(workbook)
    node = _node_for(handle, graph)
    cache = graph.caches_by_part.get(node.cache_definition_part)
    if cache is None:
        raise RelationshipPolicyError(
            "cannot locate the dedicated cache for this pivot",
            kind="invalid-pivot-graph",
            anchor=handle._identity.pivot_part,
        )
    if (not node.cache_relationship_id or not cache.records_relationship_id
            or not cache.records_part):
        raise RelationshipPolicyError(
            "cannot resolve the pivot's internal cache relationships",
            kind="invalid-pivot-graph",
            anchor=handle._identity.pivot_part,
        )
    return PivotAllocation(
        pivot_part=node.identity.pivot_part,
        cache_part=node.cache_definition_part,
        records_part=cache.records_part,
        cache_id=int(node.cache_id),
        worksheet_part=node.identity.worksheet_part,
        workbook_part=graph.workbook_part,
        pivot_cache_relationship_id=node.cache_relationship_id,
        records_relationship_id=cache.records_relationship_id,
    )


def _node_for(handle, graph):
    node = graph.pivots_by_identity.get(handle._identity)
    if node is None:
        for item in graph.pivots:
            if item.identity.pivot_part == handle._identity.pivot_part:
                return item
    if node is None:
        raise TargetNotFoundError(
            "pivot %r is no longer in the package graph" % handle._identity.name,
            kind="stale-pivot-handle",
            anchor=handle._identity.pivot_part,
        )
    return node


def _spec_with_changes(spec, changes):
    from openpyxl.pivot.create import (
        _coerce_filters, _coerce_rows, _coerce_values, _validate_public_spec,
    )
    known = {
        "rows", "columns", "filters", "values", "layout", "values_axis",
        "row_grand_totals", "column_grand_totals", "subtotals", "style",
        "destination", "source", "name",
    }
    unknown = sorted(set(changes) - known)
    if unknown:
        raise TypeError(
            "PivotTable.update() got unexpected keyword argument(s): %s"
            % ", ".join(unknown))
    values = {
        "name": spec.name,
        "source": spec.source,
        "destination": spec.destination,
        "rows": spec.rows,
        "columns": spec.columns,
        "filters": spec.filters,
        "values": spec.values,
        "layout": spec.layout,
        "values_axis": spec.values_axis,
        "row_grand_totals": spec.row_grand_totals,
        "column_grand_totals": spec.column_grand_totals,
        "subtotals": spec.subtotals,
        "style": spec.style,
    }
    if "source" in changes:
        values["source"] = PivotSource.parse(changes["source"])
    if "rows" in changes:
        values["rows"] = _coerce_rows(changes["rows"])
    if "columns" in changes:
        values["columns"] = _coerce_rows(changes["columns"])
    if "filters" in changes:
        values["filters"] = _coerce_filters(changes["filters"])
    if "values" in changes:
        values["values"] = _coerce_values(changes["values"])
    for key in ("layout", "values_axis", "row_grand_totals",
                "column_grand_totals", "subtotals", "style",
                "destination", "name"):
        if key in changes:
            values[key] = changes[key]
    spec = PivotSpec(**values)
    _validate_public_spec(spec)
    return spec


_EFFECT_ORDER = ("refresh", "repoint", "move", "update", "rename")
_UPDATE_FIELDS = (
    "rows", "columns", "filters", "values", "layout", "values_axis",
    "row_grand_totals", "column_grand_totals", "subtotals", "style",
)


def _owned_from_plan_cells(cells):
    return tuple(
        (cell.row, cell.column, cell.value, cell.role)
        for cell in cells
        if cell.value is not None
    )


def _publication_fields(staged, final_action=None):
    if staged is None:
        return {
            "origin_before": "paper",
            "publication_strategy": None,
            "final_action": final_action,
            "original_payload_hashes": (),
            "original_cache_id": None,
            "persisted_cache_identity": None,
        }
    return {
        "origin_before": getattr(staged, "origin_before", "paper"),
        "publication_strategy": getattr(staged, "publication_strategy", None),
        "final_action": final_action or getattr(staged, "final_action", None),
        "original_payload_hashes": getattr(
            staged, "original_payload_hashes", ()),
        "original_cache_id": getattr(staged, "original_cache_id", None),
        "persisted_cache_identity": getattr(
            staged, "persisted_cache_identity", None),
    }


def _assert_consumers_allow(workbook, handle, *verbs):
    wanted = set()
    if "rename" in verbs:
        wanted.add("can_rename")
    if "delete" in verbs:
        wanted.add("can_delete")
    if "move" in verbs:
        wanted.add("can_move")
    if not wanted:
        return
    from openpyxl.pivot.adopt_depend import scan_consumer_constraints

    state = handle._state()
    ledger = getattr(workbook, "_paper_ledger", None)
    staged = None if ledger is None else _find_staged(ledger, handle)
    if staged is not None and staged.output_cells:
        footprint = {
            (row, column) for row, column, value, _role in staged.output_cells
            if value is not None
        }
    else:
        try:
            footprint = {
                (row, column)
                for row, column, _value, _role in _owned_cells(
                    state, staged, handle._worksheet)
            }
        except Exception:
            footprint = set()
    for reason in scan_consumer_constraints(
            workbook, state.name, handle._worksheet.title, footprint):
        if reason.capability in wanted:
            raise UnsupportedStructureError(
                "this operation would invalidate a dependent reference "
                "(%s). Nothing was changed." % reason.code,
                kind="unsupported-pivot-operation",
                options=[reason.code],
            )


def _net_semantic_effects(staged, baseline, after, verb):
    prior = set(getattr(staged, "semantic_effects", ()) or ())
    if "create" in prior:
        return ("create",)
    if "adopt" in prior:
        return ("adopt",)
    effects = {"refresh"} if "refresh" in prior else set()
    if baseline.source != after.source:
        effects.add("repoint")
    if baseline.destination != after.destination:
        effects.add("move")
    if baseline.name != after.name:
        effects.add("rename")
    if any(getattr(baseline, name) != getattr(after, name)
           for name in _UPDATE_FIELDS):
        effects.add("update")
    if verb == "refresh":
        effects.add("refresh")
    elif not effects and staged is None and verb in _EFFECT_ORDER:
        effects.add(verb)
    return tuple(name for name in _EFFECT_ORDER if name in effects)
