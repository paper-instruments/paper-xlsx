# paper-xlsx: the preserve-mode save (CONVENTIONS §3.4/§3.5; PR-0 §3/§6)

"""Save dispatch target for preserve-mode workbooks.

Ordered-stream splice: untouched parts raw-copy from the retained bytes
(byte-identical by construction); touched worksheet parts are spliced;
cross-part edits (new sheets, styles append, calcChain cascade, workbook.xml
elements, hyperlink relationships, content types) are targeted byte edits
against the original payloads. Everything is validated BEFORE the first
output byte, so every refusal is atomic.

Still refused in v0 (typed, never silent): comment changes on loaded sheets;
table add/remove; charts/images/comments/tables on ADDED sheets (D9 partial
deferral, recorded in PAPER.md); custom-property part creation; workbook.xml
elements outside {sheets, definedNames, calcPr, bookViews}; chartsheet
changes; mark_dirty on non-worksheet parts.
"""

import io
import os
import re
import zipfile

from openpyxl.errors import UnsupportedStructureError
from openpyxl.xml.constants import ARC_CORE, ARC_CUSTOM, ARC_THEME, ARC_STYLE, REL_NS, WORKSHEET_TYPE

from . import crosspart, zipio
from . import drawings as drawings_mod
from . import ledger as ledger_mod
from .ledger import render_core_model, render_custom_model, _render_chartsheet
from .regions import (
    diff_regions,
    diff_row_attrs,
    hyperlink_signatures,
    render_cf_for_write,
    render_hyperlinks_for_write,
)
from .splice import resolve_dirty_cells, splice_sheet
from .xmlscan import scan_sheet

_CALC_CHAIN = "xl/calcChain.xml"
_CUSTOM_REGIONS = ("conditionalFormatting", "hyperlinks", "tableParts")


def _refuse(msg):
    raise UnsupportedStructureError(msg + " Nothing was written.")


def save_preserved(workbook, target, *, allow_formula_loss=False):
    """Save a preserve-mode workbook to ``target`` (path or binary
    file-like). Validates fully, then writes atomically."""
    led = workbook._paper_ledger
    source = workbook._paper_source
    if led is None or source is None:
        _refuse("preserve-mode save requires a workbook loaded with "
                "preserve=True.")

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

    force_calcpr = False
    if led.formulas_changed or _dirty_feeds_formulas(workbook, led):
        # honesty organ (PLAN Phase 3, widened by PLAN-v0.1 1.2): a human
        # opener's Excel must always compute fresh numbers — stale cached
        # values can never masquerade as current. That holds for formula
        # TEXT edits and equally for the most common agent edit of all: a
        # VALUE write into cells that formulas read. The model's
        # CalcProperties defaults the flag to True, so the arm-vs-save
        # diff cannot see this change: calcPr is forced into the
        # workbook.xml plan (sanctioned collateral, PR-0 D2) and
        # re-rendered from the fully-modeled object.
        workbook.calculation.fullCalcOnLoad = True
        force_calcpr = True

    if led.parts:
        for part in led.parts:
            _refuse("mark_dirty({0!r}): part-level re-serialization of "
                    "non-worksheet parts is not supported in v0 (the part "
                    "has no faithful model source). For raw byte swaps of "
                    "unmanaged parts (media), use "
                    "wb.replace_part(name, payload).".format(part))
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

    zin = zipfile.ZipFile(io.BytesIO(source))
    names = set(zin.namelist())
    from . import lifecycle
    part_plan = lifecycle.PartPlan(names)

    wb_part, sheet_parts = _package_info(zin)
    wb_rels_part = _rels_path(wb_part)

    # model style indices drift from the file's on non-openpyxl producers
    # (numFmt normalization, Normal-style bootstrap): every emitted s
    # attribute goes through the translator (PR-0 D2)
    translator = None
    if ARC_STYLE in names:
        from .styletrans import StyleTranslator
        translator = StyleTranslator(workbook, zin.read(ARC_STYLE))

    # ---- added sheets ----------------------------------------------------
    added = [ws for ws in workbook._sheets if ws in led.added_sheets]
    if added:
        tail = workbook._sheets[-len(added):]
        if set(tail) != set(added):
            _refuse("sheets added in-session must come after all loaded "
                    "sheets (insertion at other positions would reorder the "
                    "preserved sheet list).")
    new_sheet_parts = []      # [(part_name, payload)]
    new_rels_parts = []       # [(part_name, payload)]
    new_sheet_entries = []    # [(title, sheetId, rId, state)]
    ct_appends = []
    wb_rels_appends = []
    if added:
        original_wb_rels = zin.read(wb_rels_part)
        next_part_num = _next_sheet_number(names)
        next_sheet_id = _next_sheet_id(zin.read(wb_part))
        for i, ws in enumerate(added):
            _check_added_sheet_supported(ws)
            part_name = "xl/worksheets/sheet{0}.xml".format(next_part_num + i)
            ws._id = next_part_num + i    # keeps ws.path consistent
            payload, rel_entries = _generate_sheet_part(ws)
            payload = _rewrite_added_sheet_styles(payload, workbook,
                                                  translator)
            rel_entries = drawings_mod.plan_added_sheet_drawing(
                workbook, ws, part_plan, names, rel_entries)
            sheet_rels = crosspart.render_rels_document(rel_entries) \
                if rel_entries else None
            sheet_rels = _plan_added_sheet_comments(
                workbook, ws, part_plan, names, sheet_rels)
            new_sheet_parts.append((part_name, payload))
            if sheet_rels is not None:
                new_rels_parts.append((_rels_path(part_name), sheet_rels))
            # rIds reserved through the ENGINE's shared allocator: an
            # engine append to workbook rels in the same save (styles.xml
            # creation) must never collide (Batch-2 gate: duplicate rId4)
            rid = part_plan.reserve_rid(wb_rels_part, original_wb_rels)
            new_sheet_entries.append(
                (ws.title, next_sheet_id + i, rid, ws.sheet_state))
            wb_rels_appends.append(
                (rid, "{0}/{1}".format(REL_NS, ws._rel_type),
                 _relative_target(wb_part, part_name), None))
            ct_appends.append((part_name, WORKSHEET_TYPE))

    # ---- loaded-sheet plans ----------------------------------------------
    plan = {}
    dirty_by_part = {}
    region_claims = {}        # part -> region tags knowingly rewritten
    row_claims = {}           # part -> row indices with claimed attr edits
    baselines = {}            # part -> shifted baseline bytes (Phase 6b)
    sheet_rels_updates = {}   # part_name -> new payload
    need_styles_part = False  # styled writes into a styles-less package
    for ws in workbook.worksheets:
        if ws in led.added_sheets:
            continue
        changed_objects = ledger_mod.diff_objects(
            ws, led.object_snapshots.get(ws))
        armed_snap = led.object_snapshots.get(ws) or {}
        _armed_tables = set(armed_snap.get("table", {}))
        table_changes = [key for kind, key in changed_objects
                         if kind == "table"
                         and key in _armed_tables
                         and key in getattr(ws, "tables", {})]
        changed_objects = [(kind, key) for kind, key in changed_objects
                           if kind != "table"]
        # Batch 4 (PR-1 §3): chart/image ADDITIONS on loaded sheets become
        # new parts through the engine; per-property chart MUTATIONS
        # become byte patches when chartpatch can express them
        new_drawables = [
            (kind, key) for kind, key in changed_objects
            if kind in ("chart", "image")
            and key not in armed_snap.get(kind, {})
            and key < len(getattr(ws, "_" + kind + "s", []) or [])]
        changed_objects = [co for co in changed_objects
                           if co not in new_drawables]
        chart_prop_parts = {}     # chart part -> patched payload
        chart_mutations = [
            key for kind, key in changed_objects
            if kind == "chart" and key in armed_snap.get("chart", {})
            and key < len(getattr(ws, "_charts", []) or [])]
        if chart_mutations:
            from openpyxl.xml.functions import tostring

            from . import chartpatch as chartpatch_mod
            from .structural import _charts_referencing

            # a shift rewrites chart <c:f> texts itself; composing a
            # property edit on top would shift the NEW range too (silent
            # double-shift). Refuse — but ONLY for chart parts a shift
            # actually patches (Batch-4 gate: a shift on an unrelated
            # sheet false-refused every chart edit)
            shift_affected = set()
            for shifted_ws, ops in led.shifts.items():
                if ops:
                    shift_affected |= set(_charts_referencing(
                        workbook,
                        led.renames.get(shifted_ws, shifted_ws.title)))
            for key in chart_mutations:
                chart = ws._charts[key]
                armed_render, armed_anchor = armed_snap["chart"][key]
                part_name = getattr(chart, "_paper_part", None)
                if part_name is None or part_name not in names:
                    _refuse("chart {0} on sheet {1!r} was modified but its "
                            "package part could not be located; the edit "
                            "cannot be expressed. Reopen without "
                            "preserve=True to rewrite the workbook "
                            "lossily.".format(key, ws.title))
                if part_name in shift_affected:
                    _refuse("chart {0} on sheet {1!r} was edited in the "
                            "same session as a row/column shift that "
                            "patches the same chart part; the two "
                            "rewrites cannot be composed faithfully — do "
                            "these edits in separate sessions.".format(
                                key, ws.title))
                if ledger_mod._anchor_fingerprint(chart) != armed_anchor:
                    _refuse("chart {0} on sheet {1!r}: the anchor "
                            "(position/size) changed; anchors live in the "
                            "preserved drawing part and cannot be patched. "
                            "Only title text and series ranges are "
                            "editable on loaded charts.".format(
                                key, ws.title))
                current_render, settled = ledger_mod._settled(
                    lambda c=chart: tostring(c._write()))
                if not settled:
                    _refuse("chart {0} on sheet {1!r}: its serializer is "
                            "impure, so the edit cannot be expressed "
                            "faithfully.".format(key, ws.title))
                base = chart_prop_parts.get(part_name)
                if base is None:
                    # compose over an earlier shift's chart patch, never
                    # over the raw source (the Phase-6b overrides lesson)
                    base = plan.get(part_name)
                if base is None:
                    base = zin.read(part_name)
                chart_prop_parts[part_name] = \
                    chartpatch_mod.plan_property_edits(
                        workbook, ws, key, armed_render, current_render,
                        base)
            plan.update(chart_prop_parts)
            changed_objects = [
                (kind, k) for kind, k in changed_objects
                if not (kind == "chart" and k in chart_mutations)]
        if changed_objects:
            # the boundary class (PLAN-v0.1 1.1): these objects' backing
            # parts are preserved bytes the splice never re-serializes —
            # an accepted-but-unsaved edit is the forbidden fourth outcome
            details = "; ".join(
                "{0} {1!r} ({2})".format(
                    kind, key, ledger_mod._OBJECT_UNLOCKS[kind])
                for kind, key in changed_objects)
            _refuse(
                "loaded object(s) were modified in-session on sheet "
                "{0!r}: {1}. Their backing parts are preserved verbatim, "
                "so the edits cannot be saved faithfully — {2}. Reopen "
                "without preserve=True to rewrite the workbook lossily "
                "instead.".format(
                    ws.title, details,
                    "editing loaded objects of these kinds is not "
                    "supported yet"))
        ledger_dirty = led.dirty_coordinates(ws)
        cache_writes = led.cache_writes.get(ws, {})
        all_region_changes = diff_regions(ws, led.region_snapshots.get(ws, {}))
        pinned_hits = sorted(
            led.pinned_regions.get(ws, set()) & set(all_region_changes))
        if pinned_hits:
            # the region's serializer disagreed with itself at arm time
            # (render-time side effects): its output cannot be trusted to
            # express only the user's edit, so splicing it risks silent
            # drift — refuse instead (PLAN-v0.1 0.3)
            _refuse(
                "region(s) {0} changed on sheet {1!r}, but their "
                "serializers are impure (arm-time renders disagreed with "
                "themselves), so the model render cannot be spliced "
                "faithfully. Reopen without preserve=True to rewrite the "
                "sheet lossily.".format(", ".join(pinned_hits), ws.title))
        row_changes = diff_row_attrs(ws, led.row_attr_snapshots.get(ws, {}))
        comments_changed = _comments_changed(ws, led)
        shift_ops = led.shifts.get(ws, [])
        if not (ledger_dirty or all_region_changes or row_changes
                or comments_changed or shift_ops or led.rich_text_mode
                or table_changes or new_drawables or cache_writes):
            continue
        table_lifecycle = "tableParts" in all_region_changes

        # renamed sheets are still keyed by their ORIGINAL title in the
        # original workbook.xml (PLAN-v0.1 3.2)
        original_title = led.renames.get(ws, ws.title)
        part = sheet_parts.get(original_title)
        if part is None or part not in names:
            _refuse("cannot locate the package part for sheet {0!r} via "
                    "the workbook relationships.".format(ws.title))
        original = zin.read(part)
        if table_changes:
            # Batch 2 (PR-1 1.2): loaded-table mutations re-render the
            # table part from the fully-modeled Table via the engine
            from . import tables as tables_mod

            tables_mod.plan_table_mutations(
                workbook, ws, part, zin, table_changes, plan)
        legacy_drawing_bytes = None
        if comments_changed:
            from . import comments as comments_mod

            if comments_mod.sheet_has_comment_machinery(zin, part, names):
                _refuse("comments changed on sheet {0!r}, which already "
                        "carries comment parts; editing preserved comment/"
                        "VML machinery is not supported yet (comment "
                        "CREATION on comment-free sheets is).".format(
                            ws.title))
            if led.comment_snapshots.get(ws):
                _refuse("internal: comment snapshot mismatch on a sheet "
                        "without comment machinery ({0!r}).".format(
                            ws.title))
            legacy_drawing_bytes = comments_mod.plan_comment_creation(
                workbook, ws, part, zin, part_plan, names)
        table_parts_bytes = _UNSET = object()
        if table_lifecycle:
            # Batch 2 (PR-1 1.2): table ADD/REMOVE via the lifecycle
            # engine; the crafted tableParts bytes ride the region splice
            from . import tables as tables_mod

            table_parts_bytes = tables_mod.plan_table_lifecycle(
                workbook, ws, part, zin,
                led.region_snapshots.get(ws, {}).get("tableParts", ()),
                plan, part_plan, names)
        if shift_ops:
            # Phase 6b: the byte renumber runs first (deleted rows cut,
            # shifted r attributes rewritten, all other bytes verbatim);
            # the standard splice then treats the shifted bytes as its
            # baseline
            from .structural import apply_shift_to_bytes
            from .chartpatch import plan_chart_updates
            for op, op_idx, op_amount in shift_ops:
                original = apply_shift_to_bytes(original, op, op_idx,
                                                op_amount)
                # overrides: a chart part already patched for another
                # sheet's shift must be patched incrementally, never
                # re-planned from the source (the earlier rewrite would be
                # silently discarded)
                chart_plans, chart_blockers = plan_chart_updates(
                    workbook, led.renames.get(ws, ws.title), op, op_idx,
                    op_amount, overrides=plan)
                if chart_blockers:
                    _refuse("chart parts referencing sheet {0!r} cannot be "
                            "patched: {1}.".format(
                                ws.title, "; ".join(chart_blockers)))
                plan.update(chart_plans)
            baselines[part] = original
        scan = scan_sheet(original)
        dirty = resolve_dirty_cells(ws, ledger_dirty, scan)

        region_changes = {tag: rendered
                          for tag, rendered in all_region_changes.items()
                          if tag not in _CUSTOM_REGIONS}
        if table_lifecycle and table_parts_bytes is not _UNSET:
            region_changes["tableParts"] = table_parts_bytes
        if legacy_drawing_bytes is not None:
            region_changes["legacyDrawing"] = legacy_drawing_bytes
        if new_drawables:
            # Batch 4 (PR-1 §3): new chart/image objects become fresh
            # parts through the engine; the sheet gains ONE spliced
            # <drawing r:id> element, or the anchors land in the sheet's
            # EXISTING drawing part (anchor-only originals only)
            new_charts = [ws._charts[k] for kind, k in new_drawables
                          if kind == "chart"]
            new_images = [ws._images[k] for kind, k in new_drawables
                          if kind == "image"]
            existing_drawing, existing_drawing_rid = \
                drawings_mod._existing_drawing_part(zin, names, part)
            if existing_drawing is None:
                if scan.regions.get("drawing"):
                    _refuse("sheet {0!r} carries a drawing element whose "
                            "relationship target cannot be resolved; "
                            "adding charts/images to it is not "
                            "possible.".format(ws.title))
                rels_part = _rels_path(part)
                original_rels = zin.read(rels_part) \
                    if rels_part in names else None
                region_changes["drawing"] = drawings_mod.plan_fresh_drawing(
                    workbook, ws, part_plan, names, part, original_rels,
                    new_charts, new_images)
            else:
                drawing_base = plan.get(existing_drawing)
                if drawing_base is None:
                    drawing_base = zin.read(existing_drawing)
                drels = _rels_path(existing_drawing)
                plan[existing_drawing] = drawings_mod.plan_drawing_append(
                    workbook, ws, part_plan, names, existing_drawing,
                    drawing_base,
                    zin.read(drels) if drels in names else None,
                    new_charts, new_images)
                if not scan.regions.get("drawing"):
                    # legal-but-odd package: the drawing rel and part
                    # exist, the sheet never references them — without
                    # this element the appended objects are invisible
                    # (Batch-4 gate: orphan drawing rel)
                    region_changes["drawing"] = (
                        b'<drawing xmlns:r="%s" r:id="%s"/>'
                        % (REL_NS.encode("ascii"),
                           existing_drawing_rid.encode("ascii")))

        # PR-0 D2 applies to ROW and COLUMN styles too: dict(RowDimension)
        # and ColumnDimension.to_tree() carry MODEL style indices — every
        # one must be translated to the FILE's xf numbering (and the xf
        # appended) before touching the spliced bytes
        if row_changes:
            row_changes = _translate_row_styles(ws, row_changes, translator)
        if "cols" in region_changes and region_changes["cols"]:
            region_changes["cols"] = _translate_col_styles(
                ws, region_changes["cols"], translator)

        cf_replacement = None
        if "conditionalFormatting" in all_region_changes:
            from . import x14

            if x14.sheet_has_cf_twins(scan, original):
                # Batch 3 (PR-1 2.1): twin-bearing CF is COMPOSED from
                # original bytes (the model drops <x14:id> pointers on
                # re-render) with the extLst twins patched in lockstep
                cf_replacement, ext_crafted = x14.plan_cf_composed(
                    workbook, ws, scan, original,
                    led.region_snapshots.get(ws, {}).get(
                        "conditionalFormatting", ()))
                if ext_crafted is not None:
                    region_changes["extLst"] = ext_crafted
            else:
                cf_replacement = render_cf_for_write(ws)
        if "dataValidations" in region_changes:
            from . import x14

            # classic DV edits coexist with verbatim x14 DVs unless the
            # ranges overlap (Batch 3 lifts the blanket D15 refusal)
            x14.check_dv_coexistence(ws, scan, original)

        if shift_ops:
            # a shift splits/renumbers shared groups whose members would
            # otherwise re-derive positionally from a stale host: dissolve
            # every group on the sheet (members re-emit as plain formulas
            # from the correctly rewritten model)
            for members in scan.shared_members.values():
                dirty |= members

        hyperlinks_replacement = None
        if shift_ops and "hyperlinks" not in all_region_changes \
                and (hyperlink_signatures(ws)
                     or scan.regions.get("hyperlinks")):
            # anchors moved with their cells (or the anchored row was
            # deleted): re-render the element — possibly to nothing — so the
            # original refs can never reattach to shifted rows. Relationship
            # ids on surviving link objects are unchanged.
            hyperlinks_replacement = render_hyperlinks_for_write(ws)
        if "hyperlinks" in all_region_changes:
            hyperlinks_replacement, rels_update = _plan_hyperlinks(
                workbook, ws, led, zin, part, names, part_plan)
            if rels_update is not None:
                sheet_rels_updates[rels_update[0]] = rels_update[1]

        if not (dirty or region_changes or row_changes or shift_ops
                or cf_replacement is not None
                or hyperlinks_replacement is not None
                or cache_writes):
            continue
        if translator is None:
            # no styles.xml in the package: the part is CREATED from the
            # model via the lifecycle engine (PR-1 1.4), and cells write
            # MODEL indices (a fresh part shares the model's numbering)
            styles_needed = any(
                ws._cells[(r, c)]._style is not None
                for (r, c) in dirty if (r, c) in ws._cells)
            if styles_needed:
                need_styles_part = True
            resolver = _model_style_resolver
        else:
            resolver = translator.resolver()
        plan[part] = splice_sheet(
            ws, original, dirty, region_changes, row_changes, scan=scan,
            cf_replacement=cf_replacement,
            hyperlinks_replacement=hyperlinks_replacement,
            style_resolver=resolver,
            value_overwrites=led.value_overwrites.get(ws, frozenset()),
            cache_writes=cache_writes)
        # cache-written cells are CLAIMED changes: the crosscheck verifies
        # them exactly like dirty cells (PLAN-v0.1 5.3)
        dirty_by_part[part] = dirty | set(cache_writes)
        claims = set(region_changes)
        if cf_replacement is not None:
            claims.add("conditionalFormatting")
        if hyperlinks_replacement is not None:
            claims.add("hyperlinks")
        region_claims[part] = claims
        row_claims[part] = set(row_changes)

    # ---- removed sheets: the part cascade (3.2) ----------------------------
    for removed_title in led.removed_sheets:
        removed_part = sheet_parts.get(removed_title)
        if removed_part is None or removed_part not in names:
            _refuse("cannot locate the package part for removed sheet "
                    "{0!r}.".format(removed_title))
        closure = _exclusive_closure(zin, names, removed_part)
        part_plan.remove_part(
            removed_part, referencing_rels=[(wb_rels_part, removed_part)])
        for child_part in closure:
            if child_part in part_plan.dropped:
                continue
            part_plan.remove_part(child_part)

    # ---- rename cascade: chart parts, ONE simultaneous mapping ------------
    # (sequential pairwise patching merges reference classes on title
    # SWAPS — Batch-3 gate)
    rename_map = {orig: ws_obj.title
                  for ws_obj, orig in led.renames.items()
                  if ws_obj.title != orig}
    if rename_map:
        from .chartpatch import patch_chart_renames
        from .structural import _charts_referencing

        chart_targets = set()
        for orig in rename_map:
            chart_targets |= set(_charts_referencing(workbook, orig))
        for chart_part in sorted(chart_targets):
            payload = plan.get(chart_part, zin.read(chart_part))
            patched = patch_chart_renames(payload, rename_map)
            if patched is not None:
                plan[chart_part] = patched

    # ---- calcChain cascade (D13), on the lifecycle engine ------------------
    drop_calcchain = led.formulas_changed and _CALC_CHAIN in names
    if drop_calcchain:
        part_plan.remove_part(
            _CALC_CHAIN,
            referencing_rels=[(wb_rels_part, _CALC_CHAIN)])

    # ---- styles append (runs AFTER splices: resolution allocates new xfs) --
    styles_plan = None
    if translator is not None:
        styles_plan = crosspart.plan_styles_xml(workbook, led,
                                                zin.read(ARC_STYLE),
                                                translator)
    else:
        from .ledger import _style_fingerprint
        lengths, _fp = _style_fingerprint(workbook)
        if need_styles_part or lengths != led._style_lengths \
                or len(workbook._cell_styles) != led.orig_cell_styles_len:
            # the package has no styles.xml: generate it whole from the
            # model (nothing to preserve) via the lifecycle engine
            from openpyxl.styles.stylesheet import write_stylesheet
            from openpyxl.xml.functions import tostring

            part_plan.add_part(
                ARC_STYLE, tostring(write_stylesheet(workbook)),
                content_type="application/vnd.openxmlformats-"
                             "officedocument.spreadsheetml.styles+xml",
                relate_from=wb_part,
                rel_type=REL_NS + "/styles")

    # ---- workbook.xml plan -------------------------------------------------
    force_tags = ["calcPr"] if force_calcpr else []
    order_now = []
    for sheet_obj in workbook._sheets:
        order_now.append(led.renames.get(sheet_obj, sheet_obj.title))
    armed_minus_removed = [t for t in led.sheet_order
                           if t not in set(led.removed_sheets)]
    loaded_now = [t for t in order_now if t in set(led.sheet_order)]
    if led.removed_sheets or loaded_now != armed_minus_removed:
        # localSheetId and activeTab are position-derived: re-render both
        # workbook elements whenever positions changed (PLAN-v0.1 3.2)
        if workbook.chartsheets and any(
                ws_.defined_names for ws_ in workbook.worksheets):
            # upstream's writer numbers localSheetId over WORKSHEETS only
            # while readers index the full sheet list: a forced re-render
            # on a chartsheet-bearing book mis-scopes every local name
            # (Batch-3 gate) — refuse until the writer is fixed
            _refuse("sheet removal/reorder on a workbook with chartsheets "
                    "AND sheet-scoped defined names would mis-scope the "
                    "names (writer numbering skew).")
        force_tags += ["definedNames", "bookViews"]
    wb_xml_plan = crosspart.plan_workbook_xml(
        workbook, led, zin.read(wb_part), new_sheet_entries,
        force_tags=tuple(force_tags))

    # ---- workbook rels + content types -------------------------------------
    core_changed = render_core_model(workbook) != led.core_snapshot
    if core_changed and ARC_CORE not in names:
        _refuse("document properties changed but the package has no "
                "docProps/core.xml part; part creation is not supported "
                "in v0.")
    custom_render = render_custom_model(workbook)
    custom_delta = custom_render != led.custom_snapshot
    custom_changed = (custom_delta and ARC_CUSTOM in names
                      and custom_render is not None)
    if custom_delta and custom_render is not None \
            and ARC_CUSTOM not in names:
        from openpyxl.xml.constants import CPROPS_TYPE

        part_plan.add_part(
            ARC_CUSTOM, custom_render, content_type=CPROPS_TYPE,
            relate_from="",
            rel_type=REL_NS + "/custom-properties")
    if custom_delta and custom_render is None and ARC_CUSTOM in names:
        part_plan.remove_part(
            ARC_CUSTOM,
            referencing_rels=[("_rels/.rels", "docProps/custom.xml")])

    theme_changed = False
    if workbook.loaded_theme is not None and ARC_THEME in names:
        theme_changed = workbook.loaded_theme != zin.read(ARC_THEME)

    # ---- ct/rels composition: AFTER every engine registration ------------
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
        # compose ON TOP of the hyperlink planner's output when both touch
        # one rels part (Batch-2 gate: the engine payload shadowed the
        # hyperlink rel — dangling r:id in the saved sheet)
        if rels_part in sheet_rels_updates:
            existing = sheet_rels_updates.pop(rels_part)
        else:
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


    # ---- assemble -----------------------------------------------------------
    def build(zout):
        for info in zin.infolist():
            name = info.filename
            if name in part_plan.dropped:
                continue
            if name in led.replaced_parts:
                zipio.write_entry(zout, name, led.replaced_parts[name])
                continue
            if name in extra_rels_updates:
                zipio.write_entry(zout, name, extra_rels_updates[name])
                continue
            if name in plan:
                zipio.write_entry(zout, name, plan[name])
            elif name in sheet_rels_updates:
                zipio.write_entry(zout, name, sheet_rels_updates[name])
            elif name == wb_part and wb_xml_plan is not None:
                zipio.write_entry(zout, name, wb_xml_plan)
            elif name == wb_rels_part and wb_rels_plan is not None:
                zipio.write_entry(zout, name, wb_rels_plan)
            elif name == "[Content_Types].xml" and ct_plan is not None:
                zipio.write_entry(zout, name, ct_plan)
            elif name == ARC_STYLE and styles_plan is not None:
                zipio.write_entry(zout, name, styles_plan)
            elif name == ARC_CORE and core_changed:
                zipio.write_entry(zout, name, render_core_model(workbook))
            elif name == ARC_CUSTOM and custom_changed:
                zipio.write_entry(zout, name, custom_render)
            elif name == ARC_THEME and theme_changed:
                zipio.write_entry(zout, name, workbook.loaded_theme)
            else:
                zipio.copy_entry(zin, info, zout)
        for part_name, payload in new_sheet_parts:
            zipio.write_entry(zout, part_name, payload)
        for part_name, payload in new_rels_parts:
            zipio.write_entry(zout, part_name, payload)
        for part_name, payload in part_plan.added.items():
            zipio.write_entry(zout, part_name, payload)
        for part_name, payload in extra_rels_updates.items():
            if part_name not in names:
                zipio.write_entry(zout, part_name, payload)
        # rels parts created for LOADED sheets that had none (first
        # hyperlink on a rels-less sheet): they exist only in the plan
        for part_name, payload in sheet_rels_updates.items():
            if part_name not in names:
                zipio.write_entry(zout, part_name, payload)

    # contradictory combos refuse rather than resolve silently
    # (Batch-2 gate: replace_part payloads vanished under drops/re-renders)
    for name in led.replaced_parts:
        if name in part_plan.dropped:
            _refuse("replace_part({0!r}) conflicts with this save removing "
                    "the same part; drop one of the two edits.".format(name))
        if name in plan:
            _refuse("replace_part({0!r}) conflicts with model edits that "
                    "re-render the same part; drop one of the two "
                    "edits.".format(name))

    data = zipio.build_archive_bytes(build)

    if os.environ.get("PAPER_LEDGER_CROSSCHECK") == "1" and plan:
        from .crosscheck import verify_splice
        verify_splice(source, data, dirty_by_part, baselines=baselines,
                      region_claims=region_claims, row_claims=row_claims)

    zipio.deliver(data, target)
    return True


def _model_style_resolver(cell):
    """Style resolver for freshly CREATED styles.xml parts: the part is
    generated whole from the model, so cells write model indices (the two
    numberings coincide by construction)."""
    if cell._style is None:
        return None
    return cell.style_id


def _namelist(source):
    with zipfile.ZipFile(io.BytesIO(source)) as z:
        return set(z.namelist())


def _dirty_feeds_formulas(workbook, led):
    """True when any ledger-dirty cell intersects a reference some formula
    makes (PLAN-v0.1 1.2): the saved file's caches for those formulas are
    stale, and the human opener must recompute. Structured/table and
    unresolvable references count as always-intersecting (conservative,
    like the Phase-6 guards)."""
    if not any(led.cells.values()):
        return False
    from .perception import dependency_sketch

    sketch = dependency_sketch(workbook)
    if not sketch.references and not sketch.unresolved:
        return False
    if sketch.unresolved and any(led.cells.values()):
        return True
    for ws, dirty in led.cells.items():
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
    # charts/images on added sheets generate via the Batch-4 machinery
    # (preserve/drawings.py — stock writer output through the engine)
    if getattr(ws, "_pivots", None):
        _refuse("sheet {0!r} was added with pivot tables; not supported in "
                "v0.".format(ws.title))
    if ws.tables:
        _refuse("sheet {0!r} was added with tables; table-part generation "
                "is not supported in v0.".format(ws.title))
    # comments on added sheets generate via the Batch-2 machinery (the
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
    """Comments on an ADDED sheet (PLAN-v0.1 3.2, enabling copy_worksheet
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
    the FILE xf numbering (allocating the appended xf) — PR-0 D2."""
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
    attributes; rewrite each through the translator (PR-0 D2)."""
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
    its s attributes; rewrite them into FILE xf indices via the translator
    (PR-0 D2). Cells without an s attribute keep the implicit 0 — file xf 0
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
    touching the same rels part in one save (Batch-4 gate: duplicate rId
    with a fresh drawing)."""
    arm = led.region_snapshots.get(ws, {}).get("hyperlinks", {})
    now = hyperlink_signatures(ws)
    removed = set(arm) - set(now)
    changed = {k for k in set(arm) & set(now) if arm[k] != now[k]}
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
    """(workbook part name, {sheet title -> part name}), rels-driven (PR-0
    D11): via [Content_Types] -> workbook part -> workbook rels -> targets.
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
