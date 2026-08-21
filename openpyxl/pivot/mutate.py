# paper-xlsx: refresh, repoint, move, and later lifecycle rebuilds

"""Rebuild a Paper-managed dedicated-cache pivot from a complete spec.

Shared-cache and foreign pivots refuse before any cell or ledger mutation.
In-session creates are updated in place so save still adds one graph.
"""

from __future__ import annotations

import uuid
from dataclasses import replace

from openpyxl.errors import (
    BoundaryViolationError,
    RelationshipPolicyError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from openpyxl.pivot.api_types import PivotSource, PivotSpec
from openpyxl.pivot.build import build_pivot_payloads
from openpyxl.pivot.calculate import snapshot_for_pivot
from openpyxl.pivot.create import (
    PivotCreateOperation,
    _assert_name_available,
    _assert_output_legal,
    _cell_snapshots,
    _checkpoint,
    _content_identity,
    _restore_cells,
    _restore_ledger_cells,
    _restore_styles,
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
    spec = _spec_with_changes(state.projection.spec, changes)
    return _rebuild(handle, kind="update", spec=spec, allow_self_overlap=True)


def rename_pivot(handle, name):
    state = handle._state()
    _prepare_mutation(state, handle, "rename", "can_rename", refuse_shared=False)
    if name == state.name:
        return handle
    spec = replace(state.projection.spec, name=name)
    return _rebuild(handle, kind="rename", spec=spec, allow_self_overlap=True)


def delete_pivot(handle):
    state = handle._state()
    _prepare_mutation(state, handle, "delete", "can_delete")
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
                clear_coordinates=tuple((row, column) for row, column, _v, _r in old_cells),
                replace_existing=True,
                relationship_id=handle._identity.relationship_id,
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
    ledger = workbook._paper_ledger
    graph = load_workbook_pivot_graph(workbook)
    snapshot = snapshot_for_pivot(workbook, spec.source)
    plan = plan_pivot(spec, snapshot)
    staged = _find_staged(ledger, handle)
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
        )
        _write_output_cells(worksheet, plan.output.cells)
        _clear_obsolete(worksheet, old_cells, plan)
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
            clear_coordinates=_clear_coordinates(old_cells, plan),
            noop=False,
            replace_existing=replace_existing,
            relationship_id=relationship_id,
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
    if not getattr(caps, name, False):
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
    store = getattr(worksheet, "_cells", {})
    dirty = ledger.cells.get(worksheet)
    overwrites = ledger.value_overwrites.get(worksheet)
    coords = staged.owned_coordinates or tuple(
        (row, column) for row, column, _value, _role in staged.output_cells)
    for row, column in coords:
        store.pop((row, column), None)
        if dirty is not None:
            dirty.discard((row, column))
        if overwrites is not None:
            overwrites.discard((row, column))
    if dirty is not None and not dirty:
        ledger.cells.pop(worksheet, None)
    if overwrites is not None and not overwrites:
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
        if existing.value != expected:
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
        from openpyxl.preserve.pivots import source_impacts

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
        if actual != cell.value:
            return False
    return True


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
    return PivotAllocation(
        pivot_part=node.identity.pivot_part,
        cache_part=node.cache_definition_part,
        records_part=cache.records_part,
        cache_id=int(node.cache_id),
        worksheet_part=node.identity.worksheet_part,
        workbook_part=graph.workbook_part,
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
