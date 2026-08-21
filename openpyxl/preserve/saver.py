# paper-xlsx: the preserve-mode save

"""Save dispatch target for preserve-mode workbooks.

Ordered-stream splice: untouched parts raw-copy from the retained bytes
(byte-identical by construction); touched worksheet parts are spliced;
cross-part edits (new sheets, styles append, calcChain cascade, workbook.xml
elements, hyperlink relationships, content types) are targeted byte edits
against the original payloads. Everything is validated BEFORE the first
output byte, so every refusal is atomic.

Still refused in v0 (typed, never silent): comment changes on loaded sheets;
table add/remove; charts/images/comments/tables on ADDED sheets (partially
deferred); custom-property part creation; workbook.xml
elements outside `{sheets, definedNames, calcPr, bookViews}`; chartsheet
changes; unsupported edits to preserved object parts.
"""

import io
import os
import re
import zipfile
from dataclasses import dataclass, field
from types import MappingProxyType

from openpyxl.errors import UnsupportedStructureError
from openpyxl.xml.constants import ARC_CORE, ARC_CUSTOM, ARC_THEME, ARC_STYLE, REL_NS, WORKSHEET_TYPE
from openpyxl.xml.functions import tostring

from . import crosspart, zipio
from . import drawings as drawings_mod
from . import ledger as ledger_mod
from .ledger import render_core_model, render_custom_model, _render_chartsheet
from .regions import (
    CT_ORDER_INDEX,
    diff_regions,
    diff_row_attrs,
    hyperlink_signatures,
    render_cf_for_write,
    render_hyperlinks_for_write,
)
from .splice import resolve_dirty_cells, splice_sheet
from .xmlscan import ScanRefusal, scan_sheet

_CALC_CHAIN = "xl/calcChain.xml"
_CUSTOM_REGIONS = ("conditionalFormatting", "hyperlinks", "tableParts")


def _refuse(msg):
    raise UnsupportedStructureError(msg + " Nothing was written.")


def _copy_ledger_value(value):
    if isinstance(value, dict):
        return {key: _copy_ledger_value(item) for key, item in value.items()}
    if isinstance(value, set):
        return set(value)
    if isinstance(value, list):
        return list(value)
    if isinstance(value, tuple):
        return tuple(value)
    return value


class _PlanningState:
    """Save-time serializer mutations, restored on success or refusal."""

    def __init__(self, workbook):
        self.workbook = workbook
        self.calc_mode = workbook.calculation.calcMode
        self.calc_flag = workbook.calculation.fullCalcOnLoad
        self.force_full_calc = workbook.calculation.forceFullCalc
        self.registries = []
        for name in ("_fonts", "_fills", "_borders", "_alignments",
                     "_protections", "_number_formats", "_cell_styles"):
            registry = getattr(workbook, name)
            self.registries.append((registry, list(registry),
                                    registry.clean, dict(registry._dict)))
        self.dxfs = list(workbook._differential_styles.styles)
        self.attrs = []
        led = workbook._paper_ledger
        for ws in workbook.worksheets:
            self.attrs.append((ws, "_id", ws._id))
            self.attrs.append((ws, "_hyperlinks", ws._hyperlinks))
            self.attrs.append((ws, "_comments", ws._comments))
            self.attrs.append((ws.sheet_format, "outlineLevelCol",
                               ws.sheet_format.outlineLevelCol))
            link_coords = set()
            if led is not None:
                link_coords.update(led.region_snapshots.get(ws, {}).get(
                    "hyperlinks", {}))
                link_coords.update(led.dirty_coordinates(ws))
            for coord in link_coords:
                cell = ws._cells.get(coord)
                if cell is None:
                    continue
                link = getattr(cell, "_hyperlink", None)
                if link is not None:
                    self.attrs.append((link, "id", link.id))
            for cf in ws.conditional_formatting:
                for rule in cf.rules:
                    self.attrs.append((rule, "dxfId", rule.dxfId))
            for chart in getattr(ws, "_charts", ()):
                self.attrs.append((chart, "_id", chart._id))
            for image in getattr(ws, "_images", ()):
                self.attrs.append((image, "_id", image._id))
                self.attrs.append((image, "format", image.format))
            for table in ws.tables.values():
                self.attrs.append((table, "id", table.id))
        self.ledger = led
        self.ledger_state = {
            slot: _copy_ledger_value(getattr(led, slot))
            for slot in led.__slots__
        } if led is not None else None

    def restore(self):
        self.workbook.calculation.calcMode = self.calc_mode
        self.workbook.calculation.fullCalcOnLoad = self.calc_flag
        self.workbook.calculation.forceFullCalc = self.force_full_calc
        for registry, values, clean, index in self.registries:
            registry[:] = values
            registry.clean = clean
            registry._dict = index
        self.workbook._differential_styles.styles[:] = self.dxfs
        for obj, name, value in self.attrs:
            setattr(obj, name, value)
        if self.ledger_state is not None:
            for slot, value in self.ledger_state.items():
                setattr(self.ledger, slot, value)


@dataclass(frozen=True)
class _PreservePlan:
    """Immutable handoff from complete preflight to archive delivery."""

    changed_parts: tuple
    added_parts: tuple
    dropped_parts: tuple
    build: object = field(repr=False, compare=False)
    source: bytes = field(repr=False, compare=False)
    source_identity: object = field(repr=False, compare=False)
    expected_identity: object = field(repr=False, compare=False)
    target_is_source: bool = field(repr=False, compare=False)
    crosscheck: bool = field(repr=False, compare=False)
    dirty_by_part: object = field(repr=False, compare=False)
    baselines: object = field(repr=False, compare=False)
    region_claims: object = field(repr=False, compare=False)
    row_claims: object = field(repr=False, compare=False)


@dataclass(frozen=True)
class _PlanContext:
    """Stable package inputs shared by every planning phase."""

    workbook: object
    ledger: object
    source: bytes
    archive: object = field(repr=False, compare=False)
    names: object = field(repr=False, compare=False)
    part_plan: object = field(repr=False, compare=False)
    workbook_part: str
    sheet_parts: object = field(repr=False, compare=False)
    workbook_rels_part: str
    translator: object = field(repr=False, compare=False)


@dataclass(frozen=True)
class _AddedSheetsPlan:
    sheet_parts: tuple
    rels_parts: tuple
    sheet_entries: tuple
    content_type_appends: tuple
    workbook_rels_appends: tuple


@dataclass(frozen=True)
class _SheetObjectPlan:
    armed_snapshot: object = field(repr=False, compare=False)
    table_changes: tuple
    new_drawables: tuple


@dataclass(frozen=True)
class _SheetChanges:
    ledger_dirty: object = field(repr=False, compare=False)
    cache_writes: object = field(repr=False, compare=False)
    regions: object = field(repr=False, compare=False)
    rows: object = field(repr=False, compare=False)
    comments_changed: bool
    shifts: tuple
    table_lifecycle: bool
    changed: bool


@dataclass(frozen=True)
class _LoadedSheetSource:
    part: str
    original: bytes = field(repr=False)
    scan: object = field(repr=False, compare=False)
    table_parts: object = field(repr=False, compare=False)
    legacy_drawing: object = field(repr=False, compare=False)


@dataclass(frozen=True)
class _SheetEdits:
    dirty: object = field(repr=False, compare=False)
    cache_invalidations: object = field(repr=False, compare=False)
    regions: object = field(repr=False, compare=False)
    rows: object = field(repr=False, compare=False)


@dataclass(frozen=True)
class _LoadedSheetsPlan:
    parts: object = field(repr=False, compare=False)
    dirty_by_part: object = field(repr=False, compare=False)
    region_claims: object = field(repr=False, compare=False)
    row_claims: object = field(repr=False, compare=False)
    baselines: object = field(repr=False, compare=False)
    sheet_rels_updates: object = field(repr=False, compare=False)
    need_styles_part: bool


@dataclass(frozen=True)
class _WorkbookPartsPlan:
    workbook_xml: object = field(repr=False, compare=False)
    workbook_rels: object = field(repr=False, compare=False)
    content_types: object = field(repr=False, compare=False)
    extra_rels: object = field(repr=False, compare=False)
    styles: object = field(repr=False, compare=False)
    core_changed: bool
    custom_render: object = field(repr=False, compare=False)
    custom_changed: bool
    theme_changed: bool


_UNSET = object()


def _frozen_mapping(mapping, *, set_values=False):
    values = {
        key: frozenset(value) if set_values else value
        for key, value in mapping.items()
    }
    return MappingProxyType(values)


def _commit_preserve_plan(save_plan, target):
    """Perform the only output-producing phase after planning succeeds."""
    def validate_source():
        if save_plan.source_identity is not None:
            zipio._assert_path_identity(save_plan.source_identity)

    if save_plan.crosscheck:
        data = zipio.build_archive_bytes(save_plan.build)
        from .crosscheck import verify_splice

        verify_splice(
            save_plan.source, data, save_plan.dirty_by_part,
            baselines=save_plan.baselines,
            region_claims=save_plan.region_claims,
            row_claims=save_plan.row_claims)
        return zipio.deliver(
            data, target, expected_identity=save_plan.expected_identity,
            precommit=validate_source,
            postcommit=(None if save_plan.target_is_source
                        else validate_source))

    return zipio.build_and_deliver(
        save_plan.build, target,
        expected_identity=save_plan.expected_identity,
        precommit=validate_source,
        postcommit=(None if save_plan.target_is_source
                    else validate_source))


def save_preserved(workbook, target, *, allow_formula_loss=False,
                   _skip_pivot_freshness=False):
    """Plan and deliver without retaining serializer side effects."""
    zipio.validate_target(target)
    source_identity = workbook._paper_source_identity
    if source_identity is not None:
        zipio._assert_path_identity(source_identity)
    expected_identity = _expected_delivery_identity(workbook, target)
    target_is_source = source_identity is not None \
        and expected_identity is not None \
        and (expected_identity.requested == source_identity.requested
             or zipio._same_occupant(expected_identity, source_identity))
    state = _PlanningState(workbook)
    try:
        save_plan = _plan_preserved(
            workbook, allow_formula_loss=allow_formula_loss,
            expected_identity=expected_identity,
            target_is_source=target_is_source,
            skip_pivot_freshness=_skip_pivot_freshness)
        committed_identity = _commit_preserve_plan(save_plan, target)
        if source_identity is not None and expected_identity is not None \
                and (expected_identity.requested == source_identity.requested
                     or zipio._same_occupant(
                         expected_identity, source_identity)):
            workbook._paper_source_identity = committed_identity
        return True
    finally:
        state.restore()


def validate_preserved(workbook, *, allow_formula_loss=False):
    """Run the exact preserve save planner without assembling an archive.

    Serializer helpers used during planning can mutate model registries, so
    validation uses the same planning-state guard as delivery and restores the
    workbook before returning or raising.

    :param workbook: Preserve-mode workbook to validate.
    :type workbook: openpyxl.workbook.workbook.Workbook
    :param allow_formula_loss: Allow edited cached values to replace formulas.
    :type allow_formula_loss: bool
    :return: ``None``.
    :rtype: None
    """
    source_identity = workbook._paper_source_identity
    if source_identity is not None:
        zipio._assert_path_identity(source_identity)
    state = _PlanningState(workbook)
    try:
        _plan_preserved(workbook, allow_formula_loss=allow_formula_loss)
    finally:
        state.restore()


def _expected_delivery_identity(workbook, target):
    """Capture the destination, retaining load-time custody for aliases."""
    if hasattr(target, "write"):
        if type(target) is io.BytesIO:
            return None
        path = zipio._path_backed_target(target)
    else:
        path = target
    requested = os.path.abspath(os.fspath(path))
    source = workbook._paper_source_identity
    if source is not None and requested == source.requested:
        return source
    current = zipio.path_identity(requested, allow_missing=True)
    if source is not None and zipio._same_occupant(current, source):
        return current
    return current


def _validate_preserve_model(workbook, led, allow_formula_loss,
                             skip_pivot_freshness=False):
    """Validate workbook-wide invariants and arm recalculation metadata."""
    if workbook.data_only and not allow_formula_loss:
        _refuse(
            "this workbook was loaded with data_only=True: its cells hold "
            "cached values, not formulas, so every cell you edited would "
            "have its formula replaced by a literal (untouched cells keep "
            "their formulas in the preserved bytes). Reload without "
            "data_only=True to edit formulas safely, or pass "
            "wb.save(path, allow_formula_loss=True) to accept the loss for "
            "the edited cells.")

    led.check_style_registry(workbook)
    from openpyxl.pivot.create import validate_create_freshness
    from .pivots import source_impacts, validate_source_freshness

    if skip_pivot_freshness:
        # Disposable formula-source candidates are not publication artifacts.
        # Stale-cache refusal would otherwise block the oracle from seeing
        # current source edits that a headless refresh is trying to consume.
        pivot_impacts = source_impacts(workbook, led)
    else:
        validate_create_freshness(workbook, led)
        pivot_impacts = validate_source_freshness(workbook, led)
    force_calcpr = led.formulas_changed \
        or _dirty_feeds_formulas(workbook, led) \
        or any(impact.get("formula_binding_changed")
               for impact in pivot_impacts)
    if force_calcpr:
        # A human opener must compute fresh numbers. This applies both to
        # formula text edits and writes to cells that formulas read.
        workbook.calculation.calcMode = "auto"
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True

    for cs, snap in led.chartsheet_snapshots.items():
        if _render_chartsheet(cs) != snap:
            _refuse("chartsheet {0!r} changed; chartsheet splicing is not "
                    "supported in v0.".format(cs.title))
        cs_objects = ledger_mod.diff_objects(
            cs, led.object_snapshots.get(cs))
        if cs_objects:
            _refuse(
                "loaded object(s) were modified in-session on chartsheet "
                "{0!r}: {1}. Their backing parts are preserved verbatim, "
                "so the edits cannot be saved faithfully.".format(
                    cs.title, "; ".join(
                        "{0} {1!r}".format(kind, key)
                        for kind, key in cs_objects)))
    if bool(workbook.template) != led.template_flag:
        _refuse("wb.template changed; rewriting the workbook content type "
                "under preserve mode is not supported in v0.")
    if ledger_mod._external_links_snapshot(workbook) \
            != led.external_links_snapshot:
        _refuse("external workbook links were modified in-session; their "
                "parts are preserved verbatim, so the edits cannot be "
                "saved faithfully. Reopen without preserve=True to "
                "rewrite the workbook lossily instead.")
    return force_calcpr


def _plan_added_sheets(workbook, led, zin, names, wb_part, wb_rels_part,
                       translator, part_plan):
    """Plan model-owned parts and registrations for newly added sheets."""
    added = [ws for ws in workbook._sheets if ws in led.added_sheets]
    if added:
        tail = workbook._sheets[-len(added):]
        if set(tail) != set(added):
            _refuse("sheets added in-session must come after all loaded "
                    "sheets (insertion at other positions would reorder the "
                    "preserved sheet list).")
    new_sheet_parts = []
    new_rels_parts = []
    new_sheet_entries = []
    ct_appends = []
    wb_rels_appends = []
    if not added:
        return (new_sheet_parts, new_rels_parts, new_sheet_entries,
                ct_appends, wb_rels_appends)

    original_wb_rels = zin.read(wb_rels_part)
    next_part_num = _next_sheet_number(names)
    next_sheet_id = _next_sheet_id(zin.read(wb_part))
    for index, ws in enumerate(added):
        _check_added_sheet_supported(ws)
        part_name = "xl/worksheets/sheet{0}.xml".format(
            next_part_num + index)
        ws._id = next_part_num + index
        payload, rel_entries = _generate_sheet_part(ws)
        payload = _rewrite_added_sheet_styles(payload, workbook, translator)
        rel_entries = drawings_mod.plan_added_sheet_drawing(
            workbook, ws, part_plan, names, rel_entries)
        sheet_rels = crosspart.render_rels_document(rel_entries) \
            if rel_entries else None
        sheet_rels = _plan_added_sheet_comments(
            workbook, ws, part_plan, names, sheet_rels)
        new_sheet_parts.append((part_name, payload))
        if sheet_rels is not None:
            new_rels_parts.append((_rels_path(part_name), sheet_rels))
        rid = part_plan.reserve_rid(wb_rels_part, original_wb_rels)
        new_sheet_entries.append(
            (ws.title, next_sheet_id + index, rid, ws.sheet_state))
        wb_rels_appends.append(
            (rid, "{0}/{1}".format(REL_NS, ws._rel_type),
             _relative_target(wb_part, part_name), None))
        ct_appends.append((part_name, WORKSHEET_TYPE))
    return (new_sheet_parts, new_rels_parts, new_sheet_entries,
            ct_appends, wb_rels_appends)


def _plan_removed_sheets(led, zin, names, sheet_parts, wb_rels_part,
                         part_plan):
    """Register the exclusive package-part cascade for removed sheets."""
    for removed_title in led.removed_sheets:
        removed_part = sheet_parts.get(removed_title)
        if removed_part is None or removed_part not in names:
            _refuse("cannot locate the package part for removed sheet "
                    "{0!r}.".format(removed_title))
        closure = _exclusive_closure(zin, names, removed_part)
        part_plan.remove_part(
            removed_part, referencing_rels=[(wb_rels_part, removed_part)])
        for child_part in closure:
            if child_part not in part_plan.dropped:
                part_plan.remove_part(child_part)


def _plan_chart_rename_cascade(workbook, led, zin, plan):
    """Compose all sheet-title changes over affected chart parts once."""
    rename_map = {orig: ws_obj.title
                  for ws_obj, orig in led.renames.items()
                  if ws_obj.title != orig}
    if not rename_map:
        return
    from .chartpatch import patch_chart_renames
    from .structural import _charts_referencing

    chart_targets = set()
    for original_title in rename_map:
        chart_targets |= set(_charts_referencing(
            workbook, original_title))
    for chart_part in sorted(chart_targets):
        payload = plan.get(chart_part, zin.read(chart_part))
        patched = patch_chart_renames(payload, rename_map)
        if patched is not None:
            plan[chart_part] = patched


def _plan_workbook_metadata(workbook, led, zin, names, wb_part,
                            new_sheet_entries, force_calcpr, part_plan):
    """Plan workbook XML plus modeled core, custom, and theme metadata."""
    force_tags = ["calcPr"] if force_calcpr else []
    order_now = [led.renames.get(sheet, sheet.title)
                 for sheet in workbook._sheets]
    removed = set(led.removed_sheets)
    armed_minus_removed = [title for title in led.sheet_order
                           if title not in removed]
    loaded_titles = set(led.sheet_order)
    loaded_now = [title for title in order_now if title in loaded_titles]
    if led.removed_sheets or loaded_now != armed_minus_removed:
        if workbook.chartsheets and any(
                ws.defined_names for ws in workbook.worksheets):
            _refuse("sheet removal/reorder on a workbook with chartsheets "
                    "AND sheet-scoped defined names would mis-scope the "
                    "names (writer numbering skew).")
        force_tags += ["definedNames", "bookViews"]
    wb_xml_plan = crosspart.plan_workbook_xml(
        workbook, led, zin.read(wb_part), new_sheet_entries,
        force_tags=tuple(force_tags))

    core_changed = render_core_model(workbook) != led.core_snapshot
    if core_changed and ARC_CORE not in names:
        _refuse("document properties changed but the package has no "
                "docProps/core.xml part; part creation is not supported "
                "in v0.")
    custom_render = render_custom_model(workbook)
    custom_delta = custom_render != led.custom_snapshot
    custom_changed = (custom_delta and ARC_CUSTOM in names
                      and custom_render is not None)
    if custom_delta and custom_render is not None and ARC_CUSTOM not in names:
        from openpyxl.xml.constants import CPROPS_TYPE

        part_plan.add_part(
            ARC_CUSTOM, custom_render, content_type=CPROPS_TYPE,
            relate_from="", rel_type=REL_NS + "/custom-properties")
    if custom_delta and custom_render is None and ARC_CUSTOM in names:
        part_plan.remove_part(
            ARC_CUSTOM,
            referencing_rels=[("_rels/.rels", "docProps/custom.xml")])

    theme_changed = workbook.loaded_theme is not None \
        and ARC_THEME in names \
        and workbook.loaded_theme != zin.read(ARC_THEME)
    return (wb_xml_plan, core_changed, custom_render, custom_changed,
            theme_changed)


def _compose_package_registries(zin, names, wb_rels_part, wb_rels_appends,
                                ct_appends, part_plan, plan,
                                sheet_rels_updates):
    """Compose lifecycle registrations after every part planner has run."""
    engine_rels = part_plan.touched_rels_parts()
    wb_rels_plan = None
    if wb_rels_appends or wb_rels_part in engine_rels:
        payload = zin.read(wb_rels_part)
        payload = part_plan.apply_rels(wb_rels_part, payload)
        if wb_rels_appends:
            payload = crosspart.rels_append(payload, wb_rels_appends)
        wb_rels_plan = payload
    extra_rels_updates = {}
    for rels_part in engine_rels:
        if rels_part == wb_rels_part:
            continue
        if rels_part in sheet_rels_updates:
            existing = sheet_rels_updates.pop(rels_part)
        else:
            existing = plan.pop(rels_part, None)
            if existing is None:
                existing = zin.read(rels_part) if rels_part in names else None
        extra_rels_updates[rels_part] = part_plan.apply_rels(
            rels_part, existing)

    ct_plan = None
    if ct_appends or part_plan:
        payload = zin.read("[Content_Types].xml")
        payload = part_plan.apply_content_types(payload)
        if ct_appends:
            payload = crosspart.ct_append_overrides(payload, ct_appends)
        ct_plan = payload
    return wb_rels_plan, extra_rels_updates, ct_plan


def _prepare_plan_context(workbook):
    ledger = workbook._paper_ledger
    source = workbook._paper_source
    if ledger is None or source is None:
        _refuse("preserve-mode save requires a workbook loaded with "
                "preserve=True.")
    archive = zipfile.ZipFile(io.BytesIO(source))
    names = set(archive.namelist())
    from . import lifecycle

    part_plan = lifecycle.PartPlan(names)
    workbook_part, sheet_parts = _package_info(archive)
    workbook_rels_part = _rels_path(workbook_part)
    translator = None
    if ARC_STYLE in names:
        from .styletrans import StyleTranslator

        translator = StyleTranslator(workbook, archive.read(ARC_STYLE))
    return _PlanContext(
        workbook=workbook,
        ledger=ledger,
        source=source,
        archive=archive,
        names=names,
        part_plan=part_plan,
        workbook_part=workbook_part,
        sheet_parts=sheet_parts,
        workbook_rels_part=workbook_rels_part,
        translator=translator,
    )


def _plan_new_sheets(context):
    result = _plan_added_sheets(
        context.workbook, context.ledger, context.archive, context.names,
        context.workbook_part, context.workbook_rels_part,
        context.translator, context.part_plan)
    return _AddedSheetsPlan(*[tuple(value) for value in result])


def _plan_requested_parts(context, parts):
    ledger = context.ledger
    if ledger.image_replacements:
        from .images import plan_replacements

        plan_replacements(
            context.archive, ledger.image_replacements, context.part_plan,
            context.names, parts)
    if ledger.pivot_refresh_requests:
        from .pivots import plan_refresh

        plan_refresh(
            context.archive, ledger.pivot_refresh_requests, parts)
    if getattr(ledger, "pivot_operations", None):
        from .pivotgraph import plan_creates

        plan_creates(context)
    if not ledger.shifts:
        return
    from .chartpatch import plan_chart_updates

    for sheet in context.workbook.worksheets:
        for operation, index, amount in ledger.shifts.get(sheet, ()):
            chart_parts, blockers = plan_chart_updates(
                context.workbook,
                ledger.renames.get(sheet, sheet.title),
                operation, index, amount, overrides=parts)
            if blockers:
                _refuse("chart parts referencing sheet {0!r} cannot be "
                        "patched: {1}.".format(
                            sheet.title, "; ".join(blockers)))
            parts.update(chart_parts)


def _plan_sheet_objects(context, sheet, parts):
    ledger = context.ledger
    changed = ledger_mod.diff_objects(
        sheet, ledger.object_snapshots.get(sheet))
    armed = ledger.object_snapshots.get(sheet) or {}
    armed_tables = set(armed.get("table", {}))
    table_changes = tuple(
        key for kind, key in changed
        if kind == "table" and key in armed_tables
        and key in getattr(sheet, "tables", {}))
    changed = [(kind, key) for kind, key in changed if kind != "table"]
    new_drawables = tuple(
        (kind, key) for kind, key in changed
        if kind in ("chart", "image")
        and key not in armed.get(kind, {})
        and key < len(getattr(sheet, "_" + kind + "s", []) or []))
    changed = [item for item in changed if item not in new_drawables]
    chart_mutations = tuple(
        key for kind, key in changed
        if kind == "chart" and key in armed.get("chart", {})
        and key < len(getattr(sheet, "_charts", []) or []))
    if chart_mutations:
        _plan_chart_property_edits(
            context, sheet, armed, chart_mutations, parts)
        changed = [
            (kind, key) for kind, key in changed
            if not (kind == "chart" and key in chart_mutations)]
    if changed:
        details = "; ".join(
            "{0} {1!r} ({2})".format(
                kind, key, ledger_mod._OBJECT_UNLOCKS[kind])
            for kind, key in changed)
        _refuse(
            "loaded object(s) were modified in-session on sheet {0!r}: "
            "{1}. Their backing parts are preserved verbatim, so the edits "
            "cannot be saved faithfully — editing loaded objects of these "
            "kinds is not supported yet. Reopen without preserve=True to "
            "rewrite the workbook lossily instead.".format(
                sheet.title, details))
    return _SheetObjectPlan(
        armed_snapshot=armed,
        table_changes=table_changes,
        new_drawables=new_drawables,
    )


def _plan_chart_property_edits(context, sheet, armed, mutations, parts):
    from . import chartpatch as chartpatch_mod

    chart_parts = {}
    for key in mutations:
        chart = sheet._charts[key]
        armed_render, armed_anchor = armed["chart"][key][:2]
        part_name = getattr(chart, "_paper_part", None)
        if part_name is None or part_name not in context.names:
            _refuse("chart {0} on sheet {1!r} was modified but its package "
                    "part could not be located; the edit cannot be "
                    "expressed. Reopen without preserve=True to rewrite "
                    "the workbook lossily.".format(key, sheet.title))
        if ledger_mod._anchor_fingerprint(chart) != armed_anchor:
            _refuse("chart {0} on sheet {1!r}: the anchor (position/size) "
                    "changed; anchors live in the preserved drawing part "
                    "and cannot be patched. Only title text and series "
                    "ranges are editable on loaded charts.".format(
                        key, sheet.title))
        current_render, settled = ledger_mod._settled(
            lambda chart=chart: tostring(chart._write()))
        if not settled:
            _refuse("chart {0} on sheet {1!r}: its serializer is impure, "
                    "so the edit cannot be expressed faithfully.".format(
                        key, sheet.title))
        if current_render == armed_render:
            _refuse("chart {0} on sheet {1!r} was modified, but its "
                    "serializer does not represent the changed property; "
                    "the edit cannot be saved faithfully.".format(
                        key, sheet.title))
        base = chart_parts.get(part_name)
        composed = base is not None or part_name in parts
        if base is None:
            base = parts.get(part_name)
        if base is None:
            base = context.archive.read(part_name)
        chart_parts[part_name] = chartpatch_mod.plan_property_edits(
            context.workbook, sheet, key, armed_render, current_render, base,
            allow_composed_formula_baseline=composed)
    parts.update(chart_parts)


def _inspect_sheet_changes(context, sheet, object_plan):
    ledger = context.ledger
    ledger_dirty = ledger.dirty_coordinates(sheet)
    cache_writes = ledger.cache_writes.get(sheet, {})
    regions = diff_regions(
        sheet, ledger.region_snapshots.get(sheet, {}))
    pinned = sorted(
        ledger.pinned_regions.get(sheet, set()) & set(regions))
    if pinned:
        _refuse(
            "region(s) {0} changed on sheet {1!r}, but their serializers "
            "are impure (arm-time renders disagreed with themselves), so "
            "the model render cannot be spliced faithfully. Reopen without "
            "preserve=True to rewrite the sheet lossily.".format(
                ", ".join(pinned), sheet.title))
    rows = diff_row_attrs(
        sheet, ledger.row_attr_snapshots.get(sheet, {}))
    comments_changed = _comments_changed(sheet, ledger)
    shifts = tuple(ledger.shifts.get(sheet, ()))
    changed = bool(
        ledger_dirty or regions or rows or comments_changed or shifts
        or ledger.rich_text_mode or object_plan.table_changes
        or object_plan.new_drawables or cache_writes)
    return _SheetChanges(
        ledger_dirty=ledger_dirty,
        cache_writes=cache_writes,
        regions=regions,
        rows=rows,
        comments_changed=comments_changed,
        shifts=shifts,
        table_lifecycle="tableParts" in regions,
        changed=changed,
    )


def _prepare_loaded_sheet_source(context, sheet, object_plan, changes, parts,
                                 baselines, force_calcpr):
    ledger = context.ledger
    original_title = ledger.renames.get(sheet, sheet.title)
    part = context.sheet_parts.get(original_title)
    if part is None or part not in context.names:
        _refuse("cannot locate the package part for sheet {0!r} via the "
                "workbook relationships.".format(sheet.title))
    original = context.archive.read(part)
    if object_plan.table_changes:
        from . import tables as tables_mod

        tables_mod.plan_table_mutations(
            context.workbook, sheet, part, context.archive,
            object_plan.table_changes, parts,
            armed_tables=object_plan.armed_snapshot.get("table", {}))
    legacy_drawing = _plan_loaded_sheet_comments(
        context, sheet, part, changes.comments_changed)
    table_parts = _UNSET
    if changes.table_lifecycle:
        from . import tables as tables_mod

        table_parts = tables_mod.plan_table_lifecycle(
            context.workbook, sheet, part, context.archive,
            ledger.region_snapshots.get(sheet, {}).get("tableParts", ()),
            parts, context.part_plan, context.names)
    if changes.shifts:
        from .structural import apply_shift_to_bytes

        for operation, index, amount in changes.shifts:
            original = apply_shift_to_bytes(
                original, operation, index, amount)
        baselines[part] = original
    try:
        scan = scan_sheet(original)
    except ScanRefusal:
        if force_calcpr and not changes.changed:
            return None
        raise
    return _LoadedSheetSource(
        part=part,
        original=original,
        scan=scan,
        table_parts=table_parts,
        legacy_drawing=legacy_drawing,
    )


def _plan_loaded_sheet_comments(context, sheet, part, changed):
    if not changed:
        return None
    from . import comments as comments_mod

    kind = comments_mod.comment_machinery_kind(
        context.archive, part, context.names)
    if kind == "comments":
        _refuse("comments changed on sheet {0!r}, which already carries "
                "comment parts; editing preserved comment/VML machinery is "
                "not supported yet (comment CREATION on comment-free sheets "
                "is).".format(sheet.title))
    if kind == "other-vml":
        _refuse("comments changed on sheet {0!r}, which carries non-comment "
                "VML (such as controls or header/footer drawing state); "
                "adding comment VML without rewriting that machinery is "
                "unsupported.".format(sheet.title))
    if context.ledger.comment_snapshots.get(sheet):
        _refuse("internal: comment snapshot mismatch on a sheet without "
                "comment machinery ({0!r}).".format(sheet.title))
    return comments_mod.plan_comment_creation(
        context.workbook, sheet, part, context.archive, context.part_plan,
        context.names)


def _initial_sheet_edits(context, sheet, changes, source, force_calcpr):
    dirty = resolve_dirty_cells(
        sheet, changes.ledger_dirty, source.scan,
        value_overwrites=context.ledger.value_overwrites.get(sheet, set()))
    cache_invalidations = _formula_cache_invalidations(source.scan) \
        if force_calcpr else set()
    if changes.cache_writes:
        cache_invalidations -= set(changes.cache_writes)
    regions = {
        tag: rendered for tag, rendered in changes.regions.items()
        if tag not in _CUSTOM_REGIONS}
    if changes.table_lifecycle and source.table_parts is not _UNSET:
        regions["tableParts"] = source.table_parts
    if source.legacy_drawing is not None:
        regions["legacyDrawing"] = source.legacy_drawing
    rows = changes.rows
    if changes.shifts:
        for members in source.scan.shared_members.values():
            dirty |= members
    return _SheetEdits(
        dirty=dirty,
        cache_invalidations=cache_invalidations,
        regions=regions,
        rows=rows,
    )


def _plan_sheet_drawables(context, sheet, object_plan, source, regions, parts):
    if not object_plan.new_drawables:
        return
    new_charts = [
        sheet._charts[key] for kind, key in object_plan.new_drawables
        if kind == "chart"]
    new_images = [
        sheet._images[key] for kind, key in object_plan.new_drawables
        if kind == "image"]
    drawing_part, drawing_rid = drawings_mod._existing_drawing_part(
        context.archive, context.names, source.part)
    if drawing_part is None:
        if source.scan.regions.get("drawing"):
            _refuse("sheet {0!r} carries a drawing element whose "
                    "relationship target cannot be resolved; adding "
                    "charts/images to it is not possible.".format(
                        sheet.title))
        rels_part = _rels_path(source.part)
        original_rels = context.archive.read(rels_part) \
            if rels_part in context.names else None
        regions["drawing"] = drawings_mod.plan_fresh_drawing(
            context.workbook, sheet, context.part_plan, context.names,
            source.part, original_rels, new_charts, new_images)
        return
    drawing_base = parts.get(drawing_part)
    if drawing_base is None:
        drawing_base = context.archive.read(drawing_part)
    drawing_rels = _rels_path(drawing_part)
    parts[drawing_part] = drawings_mod.plan_drawing_append(
        context.workbook, sheet, context.part_plan, context.names,
        drawing_part, drawing_base,
        context.archive.read(drawing_rels)
        if drawing_rels in context.names else None,
        new_charts, new_images)
    if not source.scan.regions.get("drawing"):
        regions["drawing"] = (
            b'<drawing xmlns:r="%s" r:id="%s"/>'
            % (REL_NS.encode("ascii"), drawing_rid.encode("ascii")))


def _translate_and_patch_regions(context, sheet, source, edits):
    rows = edits.rows
    regions = edits.regions
    if rows:
        rows = _translate_row_styles(sheet, rows, context.translator)
    if "cols" in regions and regions["cols"]:
        regions["cols"] = _translate_col_styles(
            sheet, regions["cols"], context.translator)
    from .lexical import patch_xml

    baselines = context.ledger.region_snapshots.get(sheet, {})
    for tag, rendered in tuple(regions.items()):
        spans = source.scan.regions.get(tag, ())
        baseline = baselines.get(tag)
        if tag == "cols" or len(spans) != 1 \
                or not isinstance(baseline, (bytes, bytearray)) \
                or not isinstance(rendered, (bytes, bytearray)) \
                or not baseline or not rendered:
            continue
        span = spans[0]
        patched = patch_xml(
            source.original[span.start:span.end], baseline, rendered, tag)
        if patched is not None:
            regions[tag] = patched
    return _SheetEdits(
        dirty=edits.dirty,
        cache_invalidations=edits.cache_invalidations,
        regions=regions,
        rows=rows,
    )


def _plan_conditional_formatting(context, sheet, source, all_regions,
                                 regions):
    if "conditionalFormatting" not in all_regions:
        return None
    from . import x14
    from .lexical import patch_xml

    scan = source.scan
    original = source.original
    baselines = context.ledger.region_snapshots.get(sheet, {})
    if x14.sheet_has_cf_twins(scan, original):
        replacement, extension = x14.plan_cf_composed(
            context.workbook, sheet, scan, original,
            baselines.get("conditionalFormatting", ()))
        if extension is not None:
            regions["extLst"] = extension
        return replacement
    render_cf_for_write(sheet)
    current_blocks = tuple(
        tostring(cf.to_tree()) for cf in sheet.conditional_formatting)
    armed_blocks = baselines.get("conditionalFormatting", ())
    spans = scan.regions.get("conditionalFormatting", ())
    if len(spans) == len(armed_blocks) == len(current_blocks) and spans:
        replacements = []
        for span, baseline, current in zip(
                spans, armed_blocks, current_blocks):
            patched = patch_xml(
                original[span.start:span.end], baseline, current,
                "conditionalFormatting")
            if patched is None:
                replacements = []
                break
            replacements.append((span, patched))
        if replacements:
            run_start = spans[0].start
            run_end = spans[-1].end
            return crosspart.apply_edits(
                original[run_start:run_end],
                [(span.start - run_start, span.end - run_start, patched)
                 for span, patched in replacements])
    return render_cf_for_write(sheet)


def _plan_sheet_hyperlinks(context, sheet, source, changes,
                           sheet_rels_updates):
    replacement = None
    if changes.shifts and "hyperlinks" not in changes.regions \
            and (hyperlink_signatures(sheet)
                 or source.scan.regions.get("hyperlinks")):
        replacement = render_hyperlinks_for_write(sheet)
    if "hyperlinks" in changes.regions:
        replacement, rels_update = _plan_hyperlinks(
            context.workbook, sheet, context.ledger, context.archive,
            source.part, context.names, context.part_plan)
        if rels_update is not None:
            sheet_rels_updates[rels_update[0]] = rels_update[1]
    return replacement


def _write_loaded_sheet_plan(context, sheet, source, changes, edits,
                             cf_replacement, hyperlink_replacement, parts,
                             dirty_by_part, region_claims, row_claims):
    if not (edits.dirty or edits.regions or edits.rows or changes.shifts
            or cf_replacement is not None
            or hyperlink_replacement is not None
            or changes.cache_writes or edits.cache_invalidations):
        return False
    need_styles_part = False
    if context.translator is None:
        styles_needed = any(
            sheet._cells[(row, column)]._style is not None
            for row, column in edits.dirty
            if (row, column) in sheet._cells)
        need_styles_part = styles_needed
        resolver = _model_style_resolver
    else:
        resolver = context.translator.resolver()
    payload = splice_sheet(
        sheet, source.original, edits.dirty, edits.regions, edits.rows,
        scan=source.scan,
        cf_replacement=cf_replacement,
        hyperlinks_replacement=hyperlink_replacement,
        style_resolver=resolver,
        value_overwrites=context.ledger.value_overwrites.get(
            sheet, frozenset()),
        cache_writes=changes.cache_writes,
        cache_invalidations=edits.cache_invalidations)
    parts[source.part], dimension_changed = _sync_dimension(
        payload, sheet, source.scan, edits.dirty)
    dirty_by_part[source.part] = edits.dirty \
        | set(changes.cache_writes) | set(edits.cache_invalidations)
    claims = set(edits.regions)
    if cf_replacement is not None:
        claims.add("conditionalFormatting")
    if hyperlink_replacement is not None:
        claims.add("hyperlinks")
    if dimension_changed:
        claims.add("dimension")
    region_claims[source.part] = claims
    row_claims[source.part] = set(edits.rows)
    return need_styles_part


def _plan_loaded_sheet(context, sheet, force_calcpr, parts, dirty_by_part,
                       region_claims, row_claims, baselines,
                       sheet_rels_updates):
    object_plan = _plan_sheet_objects(context, sheet, parts)
    changes = _inspect_sheet_changes(context, sheet, object_plan)
    if not (changes.changed or force_calcpr):
        return False
    source = _prepare_loaded_sheet_source(
        context, sheet, object_plan, changes, parts, baselines,
        force_calcpr)
    if source is None:
        return False
    edits = _initial_sheet_edits(
        context, sheet, changes, source, force_calcpr)
    _plan_sheet_drawables(
        context, sheet, object_plan, source, edits.regions, parts)
    edits = _translate_and_patch_regions(
        context, sheet, source, edits)
    cf_replacement = _plan_conditional_formatting(
        context, sheet, source, changes.regions, edits.regions)
    if "dataValidations" in edits.regions:
        from . import x14

        x14.check_dv_coexistence(sheet, source.scan, source.original)
    hyperlink_replacement = _plan_sheet_hyperlinks(
        context, sheet, source, changes, sheet_rels_updates)
    return _write_loaded_sheet_plan(
        context, sheet, source, changes, edits, cf_replacement,
        hyperlink_replacement, parts, dirty_by_part, region_claims,
        row_claims)


def _plan_loaded_sheets(context, force_calcpr, parts):
    dirty_by_part = {}
    region_claims = {}
    row_claims = {}
    baselines = {}
    sheet_rels_updates = {}
    need_styles_part = False
    for sheet in context.workbook.worksheets:
        if sheet in context.ledger.added_sheets:
            continue
        need_styles_part = _plan_loaded_sheet(
            context, sheet, force_calcpr, parts, dirty_by_part,
            region_claims, row_claims, baselines,
            sheet_rels_updates) or need_styles_part
    return _LoadedSheetsPlan(
        parts=parts,
        dirty_by_part=dirty_by_part,
        region_claims=region_claims,
        row_claims=row_claims,
        baselines=baselines,
        sheet_rels_updates=sheet_rels_updates,
        need_styles_part=need_styles_part,
    )


def _plan_workbook_parts(context, added, loaded, force_calcpr):
    ledger = context.ledger
    _plan_removed_sheets(
        ledger, context.archive, context.names, context.sheet_parts,
        context.workbook_rels_part, context.part_plan)
    _plan_chart_rename_cascade(
        context.workbook, ledger, context.archive, loaded.parts)
    if ledger.formulas_changed and _CALC_CHAIN in context.names:
        context.part_plan.remove_part(
            _CALC_CHAIN,
            referencing_rels=[
                (context.workbook_rels_part, _CALC_CHAIN)])

    styles = None
    if context.translator is not None:
        styles = crosspart.plan_styles_xml(
            context.workbook, ledger, context.archive.read(ARC_STYLE),
            context.translator)
    else:
        from .ledger import _style_fingerprint

        lengths, _fingerprint = _style_fingerprint(context.workbook)
        if loaded.need_styles_part or lengths != ledger._style_lengths \
                or len(context.workbook._cell_styles) \
                != ledger.orig_cell_styles_len:
            from openpyxl.styles.stylesheet import write_stylesheet

            context.part_plan.add_part(
                ARC_STYLE, tostring(write_stylesheet(context.workbook)),
                content_type="application/vnd.openxmlformats-"
                             "officedocument.spreadsheetml.styles+xml",
                relate_from=context.workbook_part,
                rel_type=REL_NS + "/styles")

    (workbook_xml, core_changed, custom_render, custom_changed,
     theme_changed) = _plan_workbook_metadata(
         context.workbook, ledger, context.archive, context.names,
         context.workbook_part, added.sheet_entries, force_calcpr,
         context.part_plan)
    from .pivotgraph import apply_workbook_cache_registry

    workbook_xml = apply_workbook_cache_registry(context, workbook_xml)
    workbook_rels, extra_rels, content_types = \
        _compose_package_registries(
            context.archive, context.names, context.workbook_rels_part,
            added.workbook_rels_appends, added.content_type_appends,
            context.part_plan, loaded.parts, loaded.sheet_rels_updates)
    return _WorkbookPartsPlan(
        workbook_xml=workbook_xml,
        workbook_rels=workbook_rels,
        content_types=content_types,
        extra_rels=extra_rels,
        styles=styles,
        core_changed=core_changed,
        custom_render=custom_render,
        custom_changed=custom_changed,
        theme_changed=theme_changed,
    )


def _archive_builder(context, added, loaded, workbook_parts):
    def build(zout):
        with zipfile.ZipFile(io.BytesIO(context.source)) as archive:
            for info in archive.infolist():
                name = info.filename
                if name in context.part_plan.dropped:
                    continue
                if name in context.part_plan.replaced:
                    zipio.write_entry(
                        zout, name, context.part_plan.replaced[name])
                elif name in workbook_parts.extra_rels:
                    zipio.write_entry(
                        zout, name, workbook_parts.extra_rels[name])
                elif name in loaded.parts:
                    zipio.write_entry(zout, name, loaded.parts[name])
                elif name in loaded.sheet_rels_updates:
                    zipio.write_entry(
                        zout, name, loaded.sheet_rels_updates[name])
                elif name == context.workbook_part \
                        and workbook_parts.workbook_xml is not None:
                    zipio.write_entry(
                        zout, name, workbook_parts.workbook_xml)
                elif name == context.workbook_rels_part \
                        and workbook_parts.workbook_rels is not None:
                    zipio.write_entry(
                        zout, name, workbook_parts.workbook_rels)
                elif name == "[Content_Types].xml" \
                        and workbook_parts.content_types is not None:
                    zipio.write_entry(
                        zout, name, workbook_parts.content_types)
                elif name == ARC_STYLE and workbook_parts.styles is not None:
                    zipio.write_entry(zout, name, workbook_parts.styles)
                elif name == ARC_CORE and workbook_parts.core_changed:
                    zipio.write_entry(
                        zout, name, render_core_model(context.workbook))
                elif name == ARC_CUSTOM and workbook_parts.custom_changed:
                    zipio.write_entry(
                        zout, name, workbook_parts.custom_render)
                elif name == ARC_THEME and workbook_parts.theme_changed:
                    zipio.write_entry(
                        zout, name, context.workbook.loaded_theme)
                else:
                    zipio.copy_entry(archive, info, zout)
            for part_name, payload in added.sheet_parts:
                zipio.write_entry(zout, part_name, payload)
            for part_name, payload in added.rels_parts:
                zipio.write_entry(zout, part_name, payload)
            for part_name, payload in context.part_plan.added.items():
                zipio.write_entry(zout, part_name, payload)
            for part_name, payload in workbook_parts.extra_rels.items():
                if part_name not in context.names:
                    zipio.write_entry(zout, part_name, payload)
            for part_name, payload in loaded.sheet_rels_updates.items():
                if part_name not in context.names:
                    zipio.write_entry(zout, part_name, payload)
    return build


def _make_preserve_plan(context, added, loaded, workbook_parts,
                        expected_identity, target_is_source):
    changed_parts = set(loaded.parts) \
        | set(loaded.sheet_rels_updates) \
        | set(workbook_parts.extra_rels) \
        | set(context.part_plan.replaced)
    for part_name, replacement in (
            (context.workbook_part, workbook_parts.workbook_xml),
            (context.workbook_rels_part, workbook_parts.workbook_rels),
            ("[Content_Types].xml", workbook_parts.content_types),
            (ARC_STYLE, workbook_parts.styles)):
        if replacement is not None:
            changed_parts.add(part_name)
    if workbook_parts.core_changed:
        changed_parts.add(ARC_CORE)
    if workbook_parts.custom_changed:
        changed_parts.add(ARC_CUSTOM)
    if workbook_parts.theme_changed:
        changed_parts.add(ARC_THEME)
    added_parts = set(context.part_plan.added)
    added_parts.update(name for name, _payload in added.sheet_parts)
    added_parts.update(name for name, _payload in added.rels_parts)
    return _PreservePlan(
        changed_parts=tuple(sorted(changed_parts)),
        added_parts=tuple(sorted(added_parts)),
        dropped_parts=tuple(sorted(context.part_plan.dropped)),
        build=_archive_builder(context, added, loaded, workbook_parts),
        source=context.source,
        source_identity=context.workbook._paper_source_identity,
        expected_identity=expected_identity,
        target_is_source=target_is_source,
        crosscheck=(os.environ.get("PAPER_LEDGER_CROSSCHECK") == "1"
                    and bool(loaded.parts)),
        dirty_by_part=_frozen_mapping(
            loaded.dirty_by_part, set_values=True),
        baselines=_frozen_mapping(loaded.baselines),
        region_claims=_frozen_mapping(
            loaded.region_claims, set_values=True),
        row_claims=_frozen_mapping(
            loaded.row_claims, set_values=True),
    )


def _plan_preserved(workbook, *, allow_formula_loss=False,
                    expected_identity=None, target_is_source=False,
                    skip_pivot_freshness=False):
    """Run every non-writing phase and return one immutable save plan."""
    context = _prepare_plan_context(workbook)
    try:
        force_calcpr = _validate_preserve_model(
            workbook, context.ledger, allow_formula_loss,
            skip_pivot_freshness=skip_pivot_freshness)
        added = _plan_new_sheets(context)
        requested_parts = {}
        _plan_requested_parts(context, requested_parts)
        loaded = _plan_loaded_sheets(
            context, force_calcpr, requested_parts)
        workbook_parts = _plan_workbook_parts(
            context, added, loaded, force_calcpr)
        return _make_preserve_plan(
            context, added, loaded, workbook_parts,
            expected_identity, target_is_source)
    finally:
        context.archive.close()


def _model_style_resolver(cell):
    """Style resolver for freshly CREATED styles.xml parts: the part is
    generated whole from the model, so cells write model indices (the two
    numberings coincide by construction)."""
    if cell._style is None:
        return None
    return cell.style_id


def _sync_dimension(payload, ws, original_scan, dirty):
    """Synchronize ``dimension`` to the final XML cell/merge extent."""
    from openpyxl.utils import get_column_letter

    emitted_dirty = set()
    for coord in dirty:
        cell = ws._cells.get(coord)
        if cell is not None and (
                cell._value is not None or cell.has_style
                or getattr(cell, "_hyperlink", None) is not None
                or getattr(cell, "_comment", None) is not None):
            emitted_dirty.add(coord)
    min_row = min_col = None
    max_row = max_col = None
    for row, span in original_scan.rows.items():
        for col in span.cells:
            coord = (row, col)
            if coord in dirty and coord not in emitted_dirty:
                continue
            min_row = row if min_row is None else min(min_row, row)
            max_row = row if max_row is None else max(max_row, row)
            min_col = col if min_col is None else min(min_col, col)
            max_col = col if max_col is None else max(max_col, col)
    for row, col in emitted_dirty:
        min_row = row if min_row is None else min(min_row, row)
        max_row = row if max_row is None else max(max_row, row)
        min_col = col if min_col is None else min(min_col, col)
        max_col = col if max_col is None else max(max_col, col)
    for merged in ws.merged_cells.ranges:
        min_row = merged.min_row if min_row is None else min(
            min_row, merged.min_row)
        max_row = merged.max_row if max_row is None else max(
            max_row, merged.max_row)
        min_col = merged.min_col if min_col is None else min(
            min_col, merged.min_col)
        max_col = merged.max_col if max_col is None else max(
            max_col, merged.max_col)
    if min_row is not None:
        ref = "{0}{1}:{2}{3}".format(
            get_column_letter(min_col), min_row,
            get_column_letter(max_col), max_row)
    else:
        ref = "A1:A1"
    rendered = b'<dimension ref="%s"/>' % ref.encode("ascii")
    spans = original_scan.regions.get("dimension", [])
    if len(spans) > 1:
        _refuse("worksheet carries multiple dimension elements")
    if spans:
        span = spans[0]
        original = original_scan.data[span.start:span.end]
        if original == rendered:
            return payload, False
        offset = payload.find(original)
        if offset < 0:
            _refuse("worksheet dimension moved unexpectedly during planning")
        return (payload[:offset] + rendered + payload[offset + len(original):],
                True)
    rank = CT_ORDER_INDEX["dimension"]
    offset = None
    for tag, span in original_scan.region_order:
        if CT_ORDER_INDEX.get(tag, len(CT_ORDER_INDEX)) > rank:
            original = original_scan.data[span.start:span.end]
            found = payload.find(original)
            if found >= 0:
                offset = found
                break
    if offset is None:
        match = re.search(br"<(?:[A-Za-z_][\w.-]*:)?sheetData\b", payload)
        if match is not None:
            offset = match.start()
    if offset is None:
        _refuse("worksheet has no dimension insertion point")
    return payload[:offset] + rendered + payload[offset:], True


def _namelist(source):
    with zipfile.ZipFile(io.BytesIO(source)) as z:
        return set(z.namelist())


def _formula_cache_invalidations(scan):
    targets = set()
    array_bounds = scan.array_bounds
    for row_index, row_span in scan.rows.items():
        for col, cell_span in row_span.cells.items():
            array_member = bool(array_bounds) and any(
                min_row <= row_index <= max_row
                and min_col <= col <= max_col
                for min_row, min_col, max_row, max_col in array_bounds)
            if cell_span.has_formula or array_member:
                targets.add((row_index, col))
    return targets


def _dirty_feeds_formulas(workbook, led):
    """True when any value-overwritten cell intersects a reference some formula
    makes: the saved file's caches for those formulas are
    stale, and the human opener must recompute. Structured/table and
    unresolvable references count as always-intersecting
    (conservative)."""
    if not any(led.value_overwrites.values()):
        return False
    from .perception import dependency_sketch

    sketch = dependency_sketch(workbook)
    if not sketch.references and not sketch.unresolved:
        return False
    if sketch.unresolved and any(led.value_overwrites.values()):
        return True
    for ws, dirty in led.value_overwrites.items():
        if not dirty:
            continue
        title = ws.title.casefold()      # Excel sheet names: case-insensitive
        for refs in sketch.references.values():
            for ref_sheet, bounds, _raw in refs:
                if ref_sheet.casefold() != title:
                    continue
                min_col, min_row, max_col, max_row = bounds
                if min_col is None:
                    min_col, max_col = 1, 1 << 20
                if min_row is None:
                    min_row, max_row = 1, 1 << 22
                for (row, col) in dirty:
                    if row is None or col is None:
                        continue
                    if min_row <= row <= max_row \
                            and min_col <= col <= max_col:
                        return True
    return False


def _comments_changed(ws, led):
    from .ledger import _comment_snapshot

    return _comment_snapshot(ws) != led.comment_snapshots.get(ws, {})


def _rels_path(part_name):
    folder, _, base = part_name.rpartition("/")
    return "{0}/_rels/{1}.rels".format(folder, base) if folder \
        else "_rels/{0}.rels".format(base)


def _relative_target(wb_part, part_name):
    """Target of ``part_name`` relative to the workbook part's folder."""
    base = wb_part.rsplit("/", 1)[0] + "/" if "/" in wb_part else ""
    if part_name.startswith(base):
        return part_name[len(base):]
    return "/" + part_name


def _next_sheet_number(names):
    highest = 0
    for name in names:
        m = re.match(r"xl/worksheets/sheet(\d+)\.xml$", name)
        if m:
            highest = max(highest, int(m.group(1)))
    return highest + 1


def _next_sheet_id(wb_xml):
    root = crosspart.scan_small(wb_xml, "workbook", max_depth=2)
    highest = 0
    for child in root.children:
        if child.local() == "sheets":
            for sheet in child.children:
                try:
                    highest = max(highest, int(sheet.attrs.get("sheetId", 0)))
                except ValueError:
                    pass
    return highest + 1


def _check_added_sheet_supported(ws):
    # charts/images on added sheets generate via the drawing machinery
    # (preserve/drawings.py — stock writer output through the engine)
    if getattr(ws, "_pivots", None):
        _refuse("sheet {0!r} was added with pivot tables; not supported in "
                "v0.".format(ws.title))
    if ws.tables:
        _refuse("sheet {0!r} was added with tables; table-part generation "
                "is not supported in v0.".format(ws.title))
    # comments on added sheets generate via the comment-creation machinery (the
    # stock writer emits <legacyDrawing r:id="anysvml"/> whenever the
    # sheet has comments; the saver adds the matching parts + rels)


def _exclusive_closure(zin, names, root_part):
    """Parts reachable ONLY through ``root_part``'s relationship tree —
    the deletion cascade set (drawings, charts, comments, tables, their
    auxiliaries). Shared parts (referenced from any surviving rels part)
    are conservatively kept."""
    from . import lifecycle as _lc

    def rels_of(part):
        rp = _rels_path(part)
        return rp if rp in names else None

    def targets(rels_part):
        out = []
        root = crosspart.scan_small(zin.read(rels_part), "Relationships",
                                    max_depth=1)
        owner = _lc._owner_of_rels(rels_part)
        for child in root.children:
            if child.local() != "Relationship":
                continue
            if child.attrs.get("TargetMode") == "External":
                continue
            out.append(_lc._resolve_target(owner,
                                           child.attrs.get("Target", "")))
        return out

    # closure through the removed tree
    closure = set()
    frontier = [root_part]
    while frontier:
        part = frontier.pop()
        rp = rels_of(part)
        if rp is None:
            continue
        for target in targets(rp):
            if target in names and target not in closure \
                    and target != root_part:
                closure.add(target)
                frontier.append(target)

    # reference counting: anything reachable from a SURVIVING rels part
    # stays (conservative — orphans are worse than shared-part deletion)
    surviving_refs = set()
    for name in names:
        if not name.endswith(".rels"):
            continue
        owner = _lc._owner_of_rels(name)
        if owner == root_part or owner in closure:
            continue
        for target in targets(name):
            surviving_refs.add(target)
    return sorted(closure - surviving_refs)


def _plan_added_sheet_comments(workbook, ws, part_plan, names, sheet_rels):
    """Comments on an ADDED sheet (enabling copy_worksheet
    of commented sheets): the stock writer already emitted
    <legacyDrawing r:id="anysvml"/> and collected CommentRecords during
    generation; add the parts via the engine and the two rels the stock
    archive writer would have added."""
    if not ws._comments:
        return sheet_rels
    from openpyxl.comments.comment_sheet import CommentSheet
    from openpyxl.xml.functions import tostring

    from . import comments as comments_mod
    from .comments import COMMENTS_CONTENT_TYPE, VML_CONTENT_TYPE

    for record in ws._comments:
        for text in (record.text.t or "",):
            from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

            if ILLEGAL_CHARACTERS_RE.search(text or ""):
                _refuse("comment on added sheet {0!r} contains characters "
                        "that cannot be written to XML.".format(ws.title))
    cs = CommentSheet.from_comments(ws._comments)
    payload = tostring(cs.to_tree())
    if not payload.startswith(b"<?xml"):
        payload = (b'<?xml version="1.0" encoding="UTF-8" '
                   b'standalone="yes"?>\n' + payload)
    vml = cs.write_shapes(None)

    all_names = set(names) | set(part_plan.added)
    number = comments_mod._next_number(
        all_names, r"xl/comments/comment(\d+)\.xml$")
    comments_part = "xl/comments/comment{0}.xml".format(number)
    vml_part = "xl/drawings/commentsDrawing{0}.vml".format(
        comments_mod._next_number(
            all_names, r"xl/drawings/commentsDrawing(\d+)\.vml$"))
    part_plan.add_part(comments_part, payload,
                       content_type=COMMENTS_CONTENT_TYPE)
    part_plan.add_part(vml_part, vml)
    part_plan.add_default("vml", VML_CONTENT_TYPE)

    entries = [
        ("comments", REL_NS + "/comments", "/" + comments_part, None),
        ("anysvml", REL_NS + "/vmlDrawing", "/" + vml_part, None),
    ]
    if sheet_rels is None:
        return crosspart.render_rels_document(entries)
    return crosspart.rels_append(sheet_rels, entries)


def _generate_sheet_part(ws):
    """Generate a NEW sheet's part payload with the stock writer (the sheet
    exists only in the model — there is nothing to splice against). Returns
    (payload, rel_entries) — entries as (rid, type, target, mode) tuples so
    downstream planners (drawings) can fill targets before rendering."""
    from openpyxl.worksheet._writer import WorksheetWriter

    writer = WorksheetWriter(ws, out=io.BytesIO())
    writer.write()
    payload = writer.read()
    entries = [(rel.Id, rel.Type, rel.Target,
                rel.TargetMode or None) for rel in writer._rels]
    return payload, entries


def _translate_row_styles(ws, row_changes, translator):
    """Row display attrs carry the MODEL style index in 's'; translate to
    the FILE xf numbering (allocating the appended xf)."""
    out = {}
    for idx, attrs in row_changes.items():
        attrs = dict(attrs)
        if "s" in attrs:
            dim = ws.row_dimensions.get(idx)
            style_array = getattr(dim, "_style", None) if dim is not None \
                else None
            if translator is None or style_array is None:
                from openpyxl.errors import UnsupportedStructureError

                raise UnsupportedStructureError(
                    "row {0} carries a style that cannot be translated to "
                    "the original stylesheet. Nothing was written.".format(
                        idx))
            attrs["s"] = str(translator.resolve(style_array))
            attrs.setdefault("customFormat", "1")
        out[idx] = attrs
    return out


def _translate_col_styles(ws, rendered_cols, translator):
    """The cols element render carries MODEL style indices in style
    attributes; rewrite each through the translator."""
    if b"style=" not in rendered_cols:
        return rendered_cols
    if translator is None:
        from openpyxl.errors import UnsupportedStructureError

        raise UnsupportedStructureError(
            "column styles cannot be written: the package has no "
            "xl/styles.xml part. Nothing was written.")
    table = translator.model_to_file_table()

    def _sub(match):
        model_idx = int(match.group(1))
        file_idx = table.get(model_idx)
        if file_idx is None:
            return match.group(0)
        return b'style="%d"' % file_idx

    return re.sub(br'style="(\d+)"', _sub, rendered_cols)


def _rewrite_added_sheet_styles(payload, workbook, translator):
    """A freshly generated (added) sheet part carries MODEL style indices in
    its s attributes; rewrite them into FILE xf indices via the translator.
    Cells without an s attribute keep the implicit 0 — file xf 0
    by construction, since loaded entries keep their positions."""
    if translator is None or b' s="' not in payload:
        return payload
    table = translator.model_to_file_table()
    scan = scan_sheet(payload)
    edits = []
    for row in scan.rows.values():
        for cell in row.cells.values():
            s = cell.attrs.get("s")
            if s is None:
                continue
            file_idx = table.get(int(s))
            if file_idx is None or str(file_idx) == s:
                continue
            head_end = payload.index(b">", cell.start) + 1
            head = payload[cell.start:head_end]
            new_head = head.replace(
                b' s="%s"' % s.encode("ascii"),
                b' s="%d"' % file_idx, 1)
            edits.append((cell.start, head_end, new_head))
    if not edits:
        return payload
    return crosspart.apply_edits(payload, edits)


def _plan_hyperlinks(workbook, ws, led, zin, sheet_part, names,
                     part_plan):
    """Hyperlink ADDITIONS on a loaded sheet: allocate relationship ids,
    render the new hyperlinks element, and return the updated sheet-rels
    payload. Removals/changes refuse (dangling or rewritten relationships).
    Ids come from the ENGINE's shared per-rels-part allocator — an
    independent next_rid computation collides with any other planner
    touching the same rels part in one save (duplicate rId
    with a fresh drawing)."""
    arm = led.region_snapshots.get(ws, {}).get("hyperlinks", {})
    now = hyperlink_signatures(ws)
    removed = set(arm) - set(now)
    changed = {k for k in set(arm) & set(now) if arm[k] != now[k]}
    if changed and led.renames:
        from .rewrite import rename_sheets_in_formula_fragment

        rename_map = {original: sheet.title
                      for sheet, original in led.renames.items()
                      if original != sheet.title}
        derived = set()
        for coordinate in changed:
            old_target, old_location, old_tooltip, old_display = arm[coordinate]
            new_target, new_location, new_tooltip, new_display = now[coordinate]
            if (old_target, old_tooltip, old_display) != \
                    (new_target, new_tooltip, new_display) \
                    or not isinstance(old_location, str):
                continue
            expected, did_change = rename_sheets_in_formula_fragment(
                old_location, rename_map)
            if did_change and expected == new_location:
                derived.add(coordinate)
        changed -= derived
    if removed or changed:
        from openpyxl.errors import RelationshipPolicyError

        raise RelationshipPolicyError(
            "hyperlinks were removed or modified on sheet {0!r}; only "
            "hyperlink ADDITION is supported in v0 (removal would leave "
            "dangling or rewritten preserved relationships). Nothing was "
            "written.".format(ws.title))
    added = set(now) - set(arm)
    if not added:
        return render_hyperlinks_for_write(ws), None

    rels_part = _rels_path(sheet_part)
    if rels_part in names:
        rels_payload = zin.read(rels_part)
    else:
        rels_payload = None

    entries = []
    for (row, col) in sorted(added):
        cell = ws._cells[(row, col)]
        link = cell._hyperlink
        if link.target:
            rid = part_plan.reserve_rid(rels_part, rels_payload)
            link.id = rid
            entries.append((rid, _HYPERLINK_REL, link.target, "External"))
        else:
            link.id = None     # internal (location-only) links carry no rel

    rendered = render_hyperlinks_for_write(ws)
    if not entries:
        return rendered, None
    if rels_payload is not None:
        return rendered, (rels_part, crosspart.rels_append(rels_payload,
                                                           entries))
    return rendered, (rels_part, crosspart.render_rels_document(entries))


_HYPERLINK_REL = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                  "relationships/hyperlink")


def _package_info(zin):
    """(workbook part name, `{sheet title -> part name}`), rels-driven: via [Content_Types] -> workbook part -> workbook rels -> targets.
    Never pattern-matches canonical paths."""
    from openpyxl.packaging.manifest import Manifest
    from openpyxl.packaging.relationship import get_dependents, get_rels_path
    from openpyxl.xml.functions import fromstring
    from openpyxl.xml.constants import ARC_CONTENT_TYPES, XLSM, XLSX, XLTM, XLTX

    package = Manifest.from_tree(fromstring(zin.read(ARC_CONTENT_TYPES)))
    wb_part = None
    for ct in (XLTM, XLTX, XLSM, XLSX):
        part = package.find(ct)
        if part:
            wb_part = part.PartName[1:]
            break
    if wb_part is None:
        _refuse("cannot locate the workbook part in [Content_Types].xml.")

    rels = get_dependents(zin, get_rels_path(wb_part))
    id_to_target = {rel.Id: rel.target for rel in rels
                    if rel.TargetMode != "External"}

    mapping = {}
    root = fromstring(zin.read(wb_part))
    ns_main = root.tag.split("}")[0].strip("{")
    rel_ns = ("http://schemas.openxmlformats.org/officeDocument/2006/"
              "relationships")
    for sheet_el in root.iter("{%s}sheet" % ns_main):
        name = sheet_el.get("name")
        rid = sheet_el.get("{%s}id" % rel_ns)
        if name and rid and rid in id_to_target:
            mapping[name] = id_to_target[rid]
    return wb_part, mapping
