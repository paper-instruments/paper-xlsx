# paper-xlsx: atomic preserve-mode pivot creation

"""Stage one Paper-owned pivot create. Mutation of worksheet cells and the
ledger happens only after the full plan and package payloads exist. Any
exception rolls both back.
"""

from __future__ import annotations

import uuid
from dataclasses import dataclass

from openpyxl.errors import (
    AmbiguousTargetError,
    BoundaryViolationError,
    UnsupportedStructureError,
)
from openpyxl.pivot.api_types import (
    PivotAxisField,
    PivotMeasure,
    PivotSource,
    PivotSpec,
)
from openpyxl.pivot.build import build_pivot_payloads
from openpyxl.pivot.graph import load_workbook_pivot_graph
from openpyxl.pivot.inspect import PivotProjection
from openpyxl.pivot.plan import plan_pivot
from openpyxl.pivot.qualify import PivotCapabilities, PivotQualification
from openpyxl.pivot.source import _source_identity, snapshot_from_workbook
from openpyxl.preserve.pivotgraph import allocate_create
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula


# Tests may assign a callable ``(name, workbook)`` to inject refusals.
CREATE_CHECKPOINT = None


@dataclass(frozen=True)
class PivotCreateOperation:
    kind: str
    session_id: str
    name: str
    sheet: str
    spec: PivotSpec
    source_identity: str
    allocation: object
    payloads: object
    output_range: str
    output_cells: tuple
    owned_coordinates: tuple


def create_pivot(worksheet, name, source, destination, rows, values,
                 columns=None, filters=None, layout="tabular",
                 values_axis="columns", row_grand_totals=True,
                 column_grand_totals=True, subtotals=False, style=None):
    """Validate, plan, build, then stage one PR 4 pivot create."""
    from openpyxl.pivot.api import (
        invalidate_pivot_overlay,
        require_pivot_inspection,
        _session_for,
        PivotTable,
    )

    require_pivot_inspection(worksheet, api="Worksheet.pivots.create")
    workbook = worksheet.parent
    ledger = workbook._paper_ledger
    _checkpoint("start", workbook)
    _refuse_unsupported_breadth(
        columns, filters, values, layout, values_axis, style, rows)
    spec = PivotSpec(
        name=name,
        source=PivotSource.parse(source),
        destination=destination,
        rows=_coerce_rows(rows),
        columns=(),
        filters=(),
        values=_coerce_values(values),
        layout=layout,
        values_axis=values_axis,
        row_grand_totals=row_grand_totals,
        column_grand_totals=column_grand_totals,
        subtotals=subtotals,
        style=style,
    )
    if spec.source.kind not in ("table", "range"):
        raise UnsupportedStructureError(
            "PR 4 pivot creation accepts a table or sheet-qualified range "
            "source only. Nothing was changed.",
            kind="unsupported-pivot-source",
            options=[spec.source.kind],
        )
    graph = load_workbook_pivot_graph(workbook)
    _assert_name_available(name, graph, ledger)
    _checkpoint("validated", workbook)

    snapshot = snapshot_from_workbook(workbook, spec.source)
    if snapshot.formula_coordinates:
        raise UnsupportedStructureError(
            "formula-backed pivot sources are not supported on the PR 4 "
            "create spine. Nothing was changed.",
            kind="unsupported-pivot-source",
            options=list(snapshot.formula_coordinates),
        )
    plan = plan_pivot(spec, snapshot)
    _assert_output_legal(worksheet, plan, snapshot, graph, ledger)
    _checkpoint("planned", workbook)

    allocation = allocate_create(workbook, worksheet, graph, ledger)
    payloads = build_pivot_payloads(plan, allocation.cache_id)
    _checkpoint("built", workbook)

    dirty_before = set(ledger.dirty_coordinates(worksheet))
    overwrites_before = set(ledger.value_overwrites.get(worksheet, ()))
    ops_before = dict(ledger.pivot_operations)
    cell_snapshots = _cell_snapshots(worksheet, plan.output.cells)
    try:
        _write_output_cells(worksheet, plan.output.cells)
        _checkpoint("cells", workbook)
        operation = PivotCreateOperation(
            kind="create",
            session_id=uuid.uuid4().hex,
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
        )
        staged = dict(ledger.pivot_operations)
        staged[operation.session_id] = operation
        ledger.pivot_operations = staged
        _checkpoint("ledger", workbook)
    except Exception:
        _restore_cells(worksheet, cell_snapshots)
        _restore_ledger_cells(ledger, worksheet, dirty_before, overwrites_before)
        ledger.pivot_operations = ops_before
        raise

    invalidate_pivot_overlay(workbook)
    session = _session_for(workbook)
    identity = _staged_identity(operation)
    return PivotTable(worksheet, identity, session.generation)


def validate_create_freshness(workbook, ledger):
    """Refuse a save when a staged create's source changed after planning."""
    for operation in getattr(ledger, "pivot_operations", {}).values():
        if operation.kind != "create":
            continue
        snapshot = snapshot_from_workbook(workbook, operation.spec.source)
        if _content_identity(snapshot) != operation.source_identity:
            raise UnsupportedStructureError(
                "the source of staged pivot %r changed after create; "
                "saving would publish a stale dedicated cache. Recreate "
                "the pivot from the current source. Nothing was written."
                % operation.name,
                kind="stale-pivot",
                anchor="%s!%s" % (operation.sheet, operation.name),
            )


def staged_projection(operation):
    spec = operation.spec
    return PivotProjection(
        complete=True,
        spec=spec,
        source=spec.source,
        destination=spec.destination,
        output_range=operation.output_range,
        rows=spec.rows,
        columns=spec.columns,
        filters=spec.filters,
        values=spec.values,
    )


def staged_qualification():
    return PivotQualification(
        valid=True,
        origin="paper",
        capabilities=PivotCapabilities(
            can_refresh_on_open=True,
            can_headless_refresh=True,
            can_rebuild_cache=True,
            can_edit_layout=True,
            can_repoint_source=True,
            can_move=True,
            can_rename=True,
            can_delete=True,
        ),
        refresh_on_open_scope=("this-cache",),
        reasons=(),
        source_supported=True,
        cache_shared=False,
        extensions=(),
    )


def _staged_identity(operation):
    from openpyxl.pivot.graph import PivotIdentity
    return PivotIdentity(
        worksheet_part=operation.allocation.worksheet_part,
        pivot_part=operation.allocation.pivot_part,
        relationship_id="staged:%s" % operation.session_id,
        name=operation.name,
    )


def iter_staged_states(workbook):
    from openpyxl.pivot.api import _PivotState
    ledger = getattr(workbook, "_paper_ledger", None)
    if ledger is None:
        return ()
    states = []
    for operation in getattr(ledger, "pivot_operations", {}).values():
        if operation.kind != "create":
            continue
        identity = _staged_identity(operation)
        states.append(_PivotState(
            identity=identity,
            sheet_title=operation.sheet,
            name=operation.name,
            projection=staged_projection(operation),
            qualification=staged_qualification(),
        ))
    return tuple(states)


def _refuse_unsupported_breadth(columns, filters, values, layout,
                                values_axis, style, rows):
    if columns:
        raise UnsupportedStructureError(
            "PR 4 pivot creation does not accept column fields. "
            "Nothing was changed.",
            kind="unsupported-pivot-feature",
            options=["columns"],
        )
    if filters:
        raise UnsupportedStructureError(
            "PR 4 pivot creation does not accept filters. "
            "Nothing was changed.",
            kind="unsupported-pivot-feature",
            options=["filters"],
        )
    if layout != "tabular":
        raise UnsupportedStructureError(
            "PR 4 pivot creation only supports tabular layout. "
            "Nothing was changed.",
            kind="unsupported-pivot-feature",
            options=[layout],
        )
    if values_axis != "columns":
        raise UnsupportedStructureError(
            "PR 4 pivot creation only places values on columns. "
            "Nothing was changed.",
            kind="unsupported-pivot-feature",
            options=[values_axis],
        )
    if style is not None:
        raise UnsupportedStructureError(
            "PR 4 pivot creation does not accept a style override. "
            "Nothing was changed.",
            kind="unsupported-pivot-feature",
            options=["style"],
        )
    coerced_values = list(values or ())
    if len(coerced_values) != 1:
        raise UnsupportedStructureError(
            "PR 4 pivot creation requires exactly one sum measure. "
            "Nothing was changed.",
            kind="unsupported-pivot-feature",
            options=["values"],
        )
    measure = coerced_values[0]
    if isinstance(measure, str):
        aggregate = "sum"
    elif isinstance(measure, dict):
        aggregate = measure.get("aggregate", "sum")
    else:
        aggregate = getattr(measure, "aggregate", None)
    if aggregate != "sum":
        raise UnsupportedStructureError(
            "PR 4 pivot creation only supports the sum aggregate. "
            "Nothing was changed.",
            kind="unsupported-pivot-feature",
            options=[aggregate],
        )
    if getattr(measure, "number_format", None):
        raise UnsupportedStructureError(
            "PR 4 pivot creation does not accept a measure number format. "
            "Nothing was changed.",
            kind="unsupported-pivot-feature",
            options=["number_format"],
        )
    coerced_rows = list(rows or ())
    if len(coerced_rows) != 1:
        raise UnsupportedStructureError(
            "PR 4 pivot creation requires exactly one row field. "
            "Nothing was changed.",
            kind="unsupported-pivot-feature",
            options=["rows"],
        )
    row = coerced_rows[0]
    if isinstance(row, PivotAxisField) and row.items:
        raise UnsupportedStructureError(
            "PR 4 pivot creation does not accept selected row items. "
            "Nothing was changed.",
            kind="unsupported-pivot-feature",
            options=["rows.items"],
        )


def _coerce_rows(rows):
    out = []
    for item in rows:
        if isinstance(item, str):
            out.append(PivotAxisField(item))
        elif isinstance(item, PivotAxisField):
            out.append(item)
        else:
            raise TypeError("rows must contain field names or PivotAxisField")
    return tuple(out)


def _coerce_values(values):
    out = []
    for item in values:
        if isinstance(item, PivotMeasure):
            out.append(item)
        elif isinstance(item, str):
            out.append(PivotMeasure(item, aggregate="sum"))
        else:
            raise TypeError("values must contain field names or PivotMeasure")
    return tuple(out)


def _assert_name_available(name, graph, ledger):
    folded = name.casefold()
    matches = [
        node.identity.name for node in graph.pivots
        if node.identity.name and node.identity.name.casefold() == folded
    ]
    for operation in getattr(ledger, "pivot_operations", {}).values():
        if operation.name.casefold() == folded:
            matches.append(operation.name)
    if len(matches) == 1:
        raise AmbiguousTargetError(
            "pivot name %r already exists" % name,
            kind="ambiguous-pivot",
            options=matches,
        )
    if matches:
        raise AmbiguousTargetError(
            "pivot name %r is ambiguous" % name,
            kind="ambiguous-pivot",
            options=matches,
        )


def _assert_output_legal(worksheet, plan, snapshot, graph, ledger):
    output = range_boundaries(plan.output.ref)
    source_bounds = snapshot.bounds
    if source_bounds[1] is not None and source_bounds[0] == worksheet.title:
        source = (
            source_bounds[1], source_bounds[2],
            source_bounds[3], source_bounds[4],
        )
        if _ranges_overlap(output, source):
            raise BoundaryViolationError(
                "pivot output %s overlaps its source" % plan.output.ref,
                kind="pivot-source-output-overlap",
                anchor="%s!%s" % (worksheet.title, plan.output.ref),
            )
    for table in dict.values(getattr(worksheet, "tables", {}) or {}):
        try:
            bounds = range_boundaries(table.ref)
        except (TypeError, ValueError):
            continue
        if _ranges_overlap(output, bounds):
            raise BoundaryViolationError(
                "pivot output %s overlaps table %r"
                % (plan.output.ref, getattr(table, "displayName", table.ref)),
                kind="pivot-output-collision",
                anchor="%s!%s" % (worksheet.title, plan.output.ref),
            )
    for merged in worksheet.merged_cells.ranges:
        bounds = (merged.min_col, merged.min_row, merged.max_col, merged.max_row)
        if _ranges_overlap(output, bounds):
            raise BoundaryViolationError(
                "pivot output %s overlaps a merged range" % plan.output.ref,
                kind="pivot-output-collision",
                anchor="%s!%s" % (worksheet.title, plan.output.ref),
            )
    for node in graph.pivots:
        if node.sheet_title != worksheet.title or not node.output_range:
            continue
        try:
            bounds = range_boundaries(node.output_range)
        except (TypeError, ValueError):
            continue
        if _ranges_overlap(output, bounds):
            raise BoundaryViolationError(
                "pivot output %s overlaps existing pivot %r"
                % (plan.output.ref, node.identity.name),
                kind="pivot-output-collision",
                anchor="%s!%s" % (worksheet.title, plan.output.ref),
            )
    for operation in getattr(ledger, "pivot_operations", {}).values():
        if operation.sheet != worksheet.title:
            continue
        bounds = range_boundaries(operation.output_range)
        if _ranges_overlap(output, bounds):
            raise BoundaryViolationError(
                "pivot output %s overlaps staged pivot %r"
                % (plan.output.ref, operation.name),
                kind="pivot-output-collision",
                anchor="%s!%s" % (worksheet.title, plan.output.ref),
            )
    cells = getattr(worksheet, "_cells", {})
    for cell in plan.output.cells:
        existing = cells.get((cell.row, cell.column))
        if existing is None:
            continue
        if existing.value is not None:
            raise BoundaryViolationError(
                "pivot output collides with a nonblank cell at %s"
                % existing.coordinate,
                kind="pivot-output-collision",
                anchor="%s!%s" % (worksheet.title, existing.coordinate),
            )
        value = existing._value
        if isinstance(value, (ArrayFormula, DataTableFormula)):
            raise BoundaryViolationError(
                "pivot output collides with an array or data-table formula "
                "at %s" % existing.coordinate,
                kind="pivot-output-collision",
                anchor="%s!%s" % (worksheet.title, existing.coordinate),
            )


def _ranges_overlap(left, right):
    lmin_col, lmin_row, lmax_col, lmax_row = left
    rmin_col, rmin_row, rmax_col, rmax_row = right
    return not (
        lmax_col < rmin_col or rmax_col < lmin_col
        or lmax_row < rmin_row or rmax_row < lmin_row
    )


def _cell_snapshots(worksheet, cells):
    snapshots = []
    store = getattr(worksheet, "_cells", {})
    for cell in cells:
        existing = store.get((cell.row, cell.column))
        if existing is None:
            snapshots.append((cell.row, cell.column, None))
        else:
            snapshots.append((
                cell.row, cell.column,
                (existing._value, existing.data_type, existing._style),
            ))
    return tuple(snapshots)


def _write_output_cells(worksheet, cells):
    for item in cells:
        if item.value is None:
            continue
        worksheet.cell(item.row, item.column, value=item.value)


def _restore_cells(worksheet, snapshots):
    store = getattr(worksheet, "_cells", {})
    for row, column, state in snapshots:
        if state is None:
            store.pop((row, column), None)
            continue
        cell = store.get((row, column))
        if cell is None:
            continue
        cell._value, cell.data_type, cell._style = state


def _restore_ledger_cells(ledger, worksheet, dirty_before, overwrites_before):
    if dirty_before:
        ledger.cells[worksheet] = set(dirty_before)
    else:
        ledger.cells.pop(worksheet, None)
    if overwrites_before:
        ledger.value_overwrites[worksheet] = set(overwrites_before)
    else:
        ledger.value_overwrites.pop(worksheet, None)


def _content_identity(snapshot):
    """Identity of source values only — not the workbook dirty generation."""
    return _source_identity(
        snapshot.source, snapshot.fields, snapshot.records,
        snapshot.formula_coordinates, generation=None)


def _checkpoint(name, workbook):
    hook = CREATE_CHECKPOINT
    if hook is not None:
        hook(name, workbook)
