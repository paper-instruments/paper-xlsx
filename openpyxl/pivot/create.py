# paper-xlsx: atomic preserve-mode pivot creation

"""Stage one Paper-owned pivot create. Mutation of worksheet cells and the
ledger happens only after the full plan and package payloads exist. Any
exception rolls both back.
"""

from __future__ import annotations

import uuid
from copy import copy
from dataclasses import dataclass, replace

from openpyxl.errors import (
    AmbiguousTargetError,
    BoundaryViolationError,
    UnsupportedStructureError,
)
from openpyxl.pivot.api_types import (
    PivotAxisField,
    PivotItemFilter,
    PivotMeasure,
    PivotSource,
    PivotSpec,
)
from openpyxl.pivot.build import BUILT_IN_PIVOT_STYLES, build_pivot_payloads
from openpyxl.pivot.graph import load_workbook_pivot_graph
from openpyxl.pivot.inspect import PivotProjection
from openpyxl.pivot.plan import plan_pivot
from openpyxl.pivot.qualify import PivotCapabilities, PivotQualification
from openpyxl.pivot.calculate import snapshot_for_pivot
from openpyxl.pivot.source import _source_identity
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
    clear_coordinates: tuple = ()
    noop: bool = False
    replace_existing: bool = False
    relationship_id: str | None = None
    cache_rebuild: bool = False
    validate_source_identity: bool = True
    source_snapshot: object = None
    published_cell_payloads: tuple = ()
    rollback_cells: tuple = ()
    rollback_dirty_coordinates: tuple = ()
    rollback_value_overwrites: tuple = ()
    rollback_number_formats: tuple = ()
    rollback_cell_styles: tuple = ()


def create_pivot(worksheet, name, source, destination, rows, values,
                 columns=None, filters=None, layout="tabular",
                 values_axis="columns", row_grand_totals=True,
                 column_grand_totals=True, subtotals=False, style=None,
                 **kwargs):
    """Validate, plan, build, then stage one v1 pivot create."""
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
    if kwargs:
        raise TypeError(
            "Worksheet.pivots.create() got unexpected keyword argument(s): %s"
            % ", ".join(sorted(kwargs)))
    spec = PivotSpec(
        name=name,
        source=PivotSource.parse(source),
        destination=destination,
        rows=_coerce_rows(rows),
        columns=_coerce_rows(columns),
        filters=_coerce_filters(filters),
        values=_coerce_values(values),
        layout=layout,
        values_axis=values_axis,
        row_grand_totals=row_grand_totals,
        column_grand_totals=column_grand_totals,
        subtotals=subtotals,
        style=style,
    )
    _validate_public_spec(spec)
    if spec.source.kind not in ("table", "range"):
        raise UnsupportedStructureError(
            "pivot creation accepts a table or sheet-qualified range "
            "source only. Nothing was changed.",
            kind="unsupported-pivot-source",
            options=[spec.source.kind],
        )
    graph = load_workbook_pivot_graph(workbook)
    _assert_name_available(name, graph, ledger)
    _checkpoint("validated", workbook)

    snapshot = snapshot_for_pivot(workbook, spec.source)
    spec = replace(spec, filters=_resolve_filters(spec.filters, snapshot))
    plan = plan_pivot(spec, snapshot)
    _assert_output_legal(worksheet, plan, snapshot, graph, ledger)
    _checkpoint("planned", workbook)

    allocation = allocate_create(workbook, worksheet, graph, ledger)
    dirty_before = set(ledger.dirty_coordinates(worksheet))
    overwrites_before = set(ledger.value_overwrites.get(worksheet, ()))
    ops_before = dict(ledger.pivot_operations)
    cell_snapshots = _cell_snapshots(worksheet, plan.output.cells)
    formats_before = tuple(workbook._number_formats)
    styles_before = tuple(workbook._cell_styles)
    try:
        payloads = build_pivot_payloads(plan, allocation.cache_id, workbook)
        _checkpoint("built", workbook)
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
            cache_rebuild=True,
            source_snapshot=snapshot,
            published_cell_payloads=_capture_cell_payloads(
                worksheet,
                ((cell.row, cell.column) for cell in plan.output.cells),
            ),
            rollback_cells=cell_snapshots,
            rollback_dirty_coordinates=tuple(sorted(dirty_before)),
            rollback_value_overwrites=tuple(sorted(overwrites_before)),
            rollback_number_formats=formats_before,
            rollback_cell_styles=styles_before,
        )
        staged = dict(ledger.pivot_operations)
        staged[operation.session_id] = operation
        ledger.pivot_operations = staged
        _checkpoint("ledger", workbook)
    except Exception:
        _restore_cells(worksheet, cell_snapshots)
        _restore_ledger_cells(ledger, worksheet, dirty_before, overwrites_before)
        ledger.pivot_operations = ops_before
        _restore_styles(workbook, formats_before, styles_before)
        raise

    invalidate_pivot_overlay(workbook)
    session = _session_for(workbook)
    identity = _staged_identity(operation)
    return PivotTable(worksheet, identity, session.generation)


def validate_create_freshness(workbook, ledger):
    """Validate staged pivot sources and their exact worksheet payloads."""
    _validate_managed_output_mutations(workbook, ledger)
    for operation in getattr(ledger, "pivot_operations", {}).values():
        if getattr(operation, "noop", False):
            continue
        _validate_published_cells(workbook, operation)
        if operation.kind == "delete" \
                or not getattr(operation, "validate_source_identity", True):
            continue
        snapshot = snapshot_for_pivot(workbook, operation.spec.source)
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
    rid = operation.relationship_id or ("staged:%s" % operation.session_id)
    return PivotIdentity(
        worksheet_part=operation.allocation.worksheet_part,
        pivot_part=operation.allocation.pivot_part,
        relationship_id=rid,
        name=operation.name,
    )


def iter_staged_states(workbook):
    from openpyxl.pivot.api import _PivotState
    ledger = getattr(workbook, "_paper_ledger", None)
    if ledger is None:
        return ()
    states = []
    for operation in getattr(ledger, "pivot_operations", {}).values():
        if operation.kind == "delete" or operation.noop:
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


def _validate_public_spec(spec):
    if spec.style is not None and spec.style not in BUILT_IN_PIVOT_STYLES:
        raise UnsupportedStructureError(
            "pivot style %r is not a supported built-in PivotTable style. "
            "Nothing was changed." % spec.style,
            kind="unsupported-pivot-feature",
            options=sorted(BUILT_IN_PIVOT_STYLES)[:8],
        )
    from openpyxl.pivot.aggregate import _default_caption
    captions = []
    for measure in spec.values:
        caption = measure.caption or _default_caption(measure)
        if caption in captions:
            raise BoundaryViolationError(
                "measure captions must be unique; %r is repeated" % caption,
                kind="invalid-pivot-source",
                options=captions,
            )
        captions.append(caption)


def _resolve_filters(filters, snapshot):
    from openpyxl.pivot.source import typed_value

    resolved = []
    for item in filters:
        if item.field not in snapshot.field_index:
            raise BoundaryViolationError(
                "filter field %r is not in the source" % item.field,
                kind="invalid-pivot-source",
                options=list(snapshot.fields),
            )
        shared = snapshot.shared_items[snapshot.field_index[item.field]]
        shared_set = set(shared)
        if item.include is not None:
            selected = []
            selected_typed = set()
            for value in item.include:
                typed = typed_value(value)
                if typed not in shared_set:
                    raise BoundaryViolationError(
                        "filter value %r is not in field %r"
                        % (value, item.field),
                        kind="invalid-pivot-source",
                        options=[item.field],
                    )
                if typed in selected_typed:
                    continue
                selected_typed.add(typed)
                selected.append(_plain(typed))
            resolved.append(PivotItemFilter(item.field, include=selected))
            continue
        excluded = {typed_value(value) for value in (item.exclude or ())}
        include = [_plain(value) for value in shared if value not in excluded]
        if not include:
            raise BoundaryViolationError(
                "filters excluded every source row",
                kind="invalid-pivot-source",
            )
        resolved.append(PivotItemFilter(item.field, include=include))
    return tuple(resolved)


def _plain(value):
    if getattr(value, "kind", None) == "blank":
        return None
    return getattr(value, "value", value)


def _coerce_rows(rows):
    if not rows:
        return ()
    out = []
    for item in rows:
        if isinstance(item, str):
            out.append(PivotAxisField(item))
        elif isinstance(item, PivotAxisField):
            out.append(item)
        else:
            raise TypeError("rows/columns must contain field names or PivotAxisField")
    return tuple(out)


def _coerce_filters(filters):
    if not filters:
        return ()
    out = []
    for item in filters:
        if isinstance(item, PivotItemFilter):
            out.append(item)
        else:
            raise TypeError("filters must contain PivotItemFilter values")
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


def hidden_pivot_parts(ledger):
    """Package parts superseded by staged replace or delete operations."""
    parts = set()
    if ledger is None:
        return parts
    for operation in getattr(ledger, "pivot_operations", {}).values():
        if getattr(operation, "noop", False):
            continue
        if operation.kind == "delete" or operation.replace_existing:
            part = getattr(operation.allocation, "pivot_part", None)
            if part:
                parts.add(part)
    return parts


def _assert_name_available(name, graph, ledger, ignore_name=None):
    folded = name.casefold()
    ignore = None if not ignore_name else ignore_name.casefold()
    hidden = hidden_pivot_parts(ledger)
    matches = [
        node.identity.name for node in graph.pivots
        if node.identity.name
        and node.identity.name.casefold() == folded
        and node.identity.pivot_part not in hidden
        and (ignore is None or node.identity.name.casefold() != ignore)
    ]
    for operation in getattr(ledger, "pivot_operations", {}).values():
        if operation.kind == "delete" or getattr(operation, "noop", False):
            continue
        if operation.name.casefold() != folded:
            continue
        if ignore is not None and operation.name.casefold() == ignore:
            continue
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


def _assert_output_legal(worksheet, plan, snapshot, graph, ledger,
                         ignore_coordinates=(), ignore_name=None):
    occupied = _occupied_bounds(plan)
    source_bounds = snapshot.bounds
    if source_bounds[1] is not None and source_bounds[0] == worksheet.title:
        source = (
            source_bounds[1], source_bounds[2],
            source_bounds[3], source_bounds[4],
        )
        if _ranges_overlap(occupied, source):
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
        if _ranges_overlap(occupied, bounds):
            raise BoundaryViolationError(
                "pivot output %s overlaps table %r"
                % (plan.output.ref, getattr(table, "displayName", table.ref)),
                kind="pivot-output-collision",
                anchor="%s!%s" % (worksheet.title, plan.output.ref),
            )
    for merged in worksheet.merged_cells.ranges:
        bounds = (merged.min_col, merged.min_row, merged.max_col, merged.max_row)
        if _ranges_overlap(occupied, bounds):
            raise BoundaryViolationError(
                "pivot output %s overlaps a merged range" % plan.output.ref,
                kind="pivot-output-collision",
                anchor="%s!%s" % (worksheet.title, plan.output.ref),
            )
    hidden = hidden_pivot_parts(ledger)
    for node in graph.pivots:
        if node.sheet_title != worksheet.title or not node.output_range:
            continue
        if node.identity.pivot_part in hidden:
            continue
        if ignore_name and node.identity.name == ignore_name:
            continue
        try:
            bounds = range_boundaries(node.output_range)
        except (TypeError, ValueError):
            continue
        if _ranges_overlap(occupied, bounds):
            raise BoundaryViolationError(
                "pivot output %s overlaps existing pivot %r"
                % (plan.output.ref, node.identity.name),
                kind="pivot-output-collision",
                anchor="%s!%s" % (worksheet.title, plan.output.ref),
            )
    for operation in getattr(ledger, "pivot_operations", {}).values():
        if operation.sheet != worksheet.title:
            continue
        if operation.kind == "delete" or getattr(operation, "noop", False):
            continue
        if ignore_name and operation.name == ignore_name:
            continue
        if not operation.output_range:
            continue
        bounds = range_boundaries(operation.output_range)
        if _ranges_overlap(occupied, bounds):
            raise BoundaryViolationError(
                "pivot output %s overlaps staged pivot %r"
                % (plan.output.ref, operation.name),
                kind="pivot-output-collision",
                anchor="%s!%s" % (worksheet.title, plan.output.ref),
            )
    cells = getattr(worksheet, "_cells", {})
    ignored = set(ignore_coordinates)
    for cell in plan.output.cells:
        if (cell.row, cell.column) in ignored:
            continue
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
        if existing.has_style or existing._comment is not None \
                or existing._hyperlink is not None:
            raise BoundaryViolationError(
                "pivot output collides with formatted or annotated cell %s"
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


def _occupied_bounds(plan):
    columns = [cell.column for cell in plan.output.cells]
    rows = [cell.row for cell in plan.output.cells]
    return min(columns), min(rows), max(columns), max(rows)


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
                (
                    existing._value,
                    existing.data_type,
                    copy(existing._style) if existing._style is not None else None,
                    existing._hyperlink,
                    existing._comment,
                ),
            ))
    return tuple(snapshots)


def _write_output_cells(worksheet, cells):
    for item in cells:
        if item.value is None:
            continue
        cell = worksheet.cell(item.row, item.column, value=item.value)
        if item.number_format:
            cell.number_format = item.number_format


def _restore_cells(worksheet, snapshots):
    store = getattr(worksheet, "_cells", {})
    for row, column, state in snapshots:
        if state is None:
            store.pop((row, column), None)
            continue
        cell = store.get((row, column))
        if cell is None:
            cell = worksheet.cell(row, column)
        value, data_type, style, hyperlink, comment = state
        cell._value = value
        cell._data_type = data_type
        cell._style = copy(style) if style is not None else None
        cell._hyperlink = hyperlink
        if hyperlink is not None:
            hyperlink.ref = cell.coordinate
        cell._comment = comment
        if comment is not None:
            comment._parent = cell


def _capture_cell_payloads(worksheet, coordinates):
    """Freeze the exact cell payload a staged pivot intends to publish."""
    store = getattr(worksheet, "_cells", {})
    return tuple(
        (row, column, _cell_payload(store.get((row, column))))
        for row, column in sorted(set(coordinates))
    )


def _cell_payload(cell):
    if cell is None:
        return None
    value = cell._value
    value_type = (type(value).__module__, type(value).__qualname__)
    style = tuple(cell._style) if cell._style is not None else None
    comment = None
    if cell._comment is not None:
        comment = (
            cell._comment.text,
            cell._comment.author,
            cell._comment.height,
            cell._comment.width,
        )
    hyperlink = None
    if cell._hyperlink is not None:
        hyperlink = (
            cell._hyperlink.target,
            cell._hyperlink.location,
            cell._hyperlink.tooltip,
            cell._hyperlink.display,
        )
    return (value_type, value, cell.data_type, style, comment, hyperlink)


def _validate_published_cells(workbook, operation):
    worksheet = workbook[operation.sheet]
    store = getattr(worksheet, "_cells", {})
    for row, column, expected in getattr(
            operation, "published_cell_payloads", ()):
        actual = _cell_payload(store.get((row, column)))
        if actual != expected:
            coordinate = worksheet.cell(row, column).coordinate
            raise BoundaryViolationError(
                "pivot output changed after the operation was staged at %s; "
                "saving would publish worksheet cells that disagree with the "
                "pivot graph. Nothing was written." % coordinate,
                kind="pivot-output-collision",
                anchor="%s!%s" % (worksheet.title, coordinate),
            )


def _validate_managed_output_mutations(workbook, ledger):
    """Reject direct writes inside loaded PivotTable output rectangles."""
    from openpyxl.pivot.graph import load_workbook_pivot_graph
    from openpyxl.utils import get_column_letter

    graph = load_workbook_pivot_graph(workbook)
    operations = {
        operation.allocation.pivot_part: operation
        for operation in getattr(ledger, "pivot_operations", {}).values()
    }
    worksheets = {worksheet.title: worksheet for worksheet in workbook.worksheets}
    for node in graph.pivots:
        if not node.output_range or node.sheet_title not in worksheets:
            continue
        try:
            min_col, min_row, max_col, max_row = range_boundaries(
                node.output_range)
        except (TypeError, ValueError):
            continue
        worksheet = worksheets[node.sheet_title]
        managed = {
            (row, column)
            for row in range(min_row, max_row + 1)
            for column in range(min_col, max_col + 1)
        }
        cache = graph.caches_by_part.get(node.cache_definition_part)
        try:
            from openpyxl.pivot.inspect import project_pivot
            from openpyxl.pivot.plan import plan_pivot
            from openpyxl.pivot.qualify import _snapshot_from_cache_package

            projection = project_pivot(
                node, cache, workbook._paper_source, workbook)
            if projection.complete:
                snapshot = _snapshot_from_cache_package(
                    workbook._paper_source, node, projection)
                plan = plan_pivot(projection.spec, snapshot)
                managed = {
                    (cell.row, cell.column) for cell in plan.output.cells
                }
        except Exception:
            pass
        dirty = {
            (row, column)
            for row, column in ledger.dirty_coordinates(worksheet)
            if (row, column) in managed
        }
        if not dirty:
            continue
        operation = operations.get(node.identity.pivot_part)
        allowed = set()
        if operation is not None:
            allowed = {
                (row, column)
                for row, column, _payload in getattr(
                    operation, "published_cell_payloads", ())
            }
        unexpected = sorted(dirty - allowed)
        if not unexpected:
            continue
        row, column = unexpected[0]
        coordinate = "%s%s" % (get_column_letter(column), row)
        raise BoundaryViolationError(
            "cell %s is inside PivotTable %r output and was edited outside "
            "its pivot operation. Nothing was written."
            % (coordinate, node.identity.name),
            kind="pivot-output-collision",
            anchor="%s!%s" % (worksheet.title, coordinate),
        )


def _restore_ledger_cells(ledger, worksheet, dirty_before, overwrites_before):
    if dirty_before:
        ledger.cells[worksheet] = set(dirty_before)
    else:
        ledger.cells.pop(worksheet, None)
    if overwrites_before:
        ledger.value_overwrites[worksheet] = set(overwrites_before)
    else:
        ledger.value_overwrites.pop(worksheet, None)


def _restore_styles(workbook, formats_before, styles_before):
    from openpyxl.utils.indexed_list import IndexedList

    workbook._number_formats = IndexedList(formats_before)
    workbook._cell_styles = IndexedList(styles_before)


def _content_identity(snapshot):
    """Identity of source values only — not the workbook dirty generation."""
    return _source_identity(
        snapshot.source, snapshot.fields, snapshot.records,
        snapshot.formula_coordinates, generation=None)


def _checkpoint(name, workbook):
    hook = CREATE_CHECKPOINT
    if hook is not None:
        hook(name, workbook)
