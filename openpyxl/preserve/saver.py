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
    if led.formulas_changed:
        # honesty organ (PLAN Phase 3): a human opener's Excel must always
        # compute fresh numbers — stale cached values can never masquerade
        # as current. The model's CalcProperties defaults the flag to True,
        # so the arm-vs-save diff cannot see this change: calcPr is forced
        # into the workbook.xml plan (sanctioned collateral for formula
        # edits, PR-0 D2) and re-rendered from the fully-modeled object.
        workbook.calculation.fullCalcOnLoad = True
        force_calcpr = True

    if led.parts:
        for part in led.parts:
            _refuse("mark_dirty({0!r}): part-level re-serialization of "
                    "non-worksheet parts is not supported in v0 (the part "
                    "has no faithful model source).".format(part))
    if render_custom_model(workbook) != led.custom_snapshot:
        if ARC_CUSTOM not in _namelist(source):
            _refuse("custom document properties changed but the package has "
                    "no docProps/custom.xml part; part creation is not "
                    "supported in v0.")
    for cs, snap in led.chartsheet_snapshots.items():
        if _render_chartsheet(cs) != snap:
            _refuse("chartsheet {0!r} changed; chartsheet splicing is not "
                    "supported in v0.".format(cs.title))

    zin = zipfile.ZipFile(io.BytesIO(source))
    names = set(zin.namelist())

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
        next_rid = crosspart.rels_next_rid(original_wb_rels)
        next_part_num = _next_sheet_number(names)
        next_sheet_id = _next_sheet_id(zin.read(wb_part))
        for i, ws in enumerate(added):
            _check_added_sheet_supported(ws)
            part_name = "xl/worksheets/sheet{0}.xml".format(next_part_num + i)
            ws._id = next_part_num + i    # keeps ws.path consistent
            payload, sheet_rels = _generate_sheet_part(ws)
            payload = _rewrite_added_sheet_styles(payload, workbook,
                                                  translator)
            new_sheet_parts.append((part_name, payload))
            if sheet_rels is not None:
                new_rels_parts.append((_rels_path(part_name), sheet_rels))
            rid = "rId{0}".format(next_rid + i)
            new_sheet_entries.append(
                (ws.title, next_sheet_id + i, rid, ws.sheet_state))
            wb_rels_appends.append(
                (rid, "{0}/{1}".format(REL_NS, ws._rel_type),
                 _relative_target(wb_part, part_name), None))
            ct_appends.append((part_name, WORKSHEET_TYPE))

    # ---- loaded-sheet plans ----------------------------------------------
    plan = {}
    dirty_by_part = {}
    baselines = {}            # part -> shifted baseline bytes (Phase 6b)
    sheet_rels_updates = {}   # part_name -> new payload
    for ws in workbook.worksheets:
        if ws in led.added_sheets:
            continue
        ledger_dirty = led.dirty_coordinates(ws)
        all_region_changes = diff_regions(ws, led.region_snapshots.get(ws, {}))
        row_changes = diff_row_attrs(ws, led.row_attr_snapshots.get(ws, {}))
        comments_changed = _comments_changed(ws, led)
        shift_ops = led.shifts.get(ws, [])
        if not (ledger_dirty or all_region_changes or row_changes
                or comments_changed or shift_ops or led.rich_text_mode):
            continue
        if comments_changed:
            _refuse("comments changed on sheet {0!r}; comment-part editing "
                    "is not supported in v0.".format(ws.title))
        if "tableParts" in all_region_changes:
            _refuse("tables were added or removed on sheet {0!r}; table-part "
                    "lifecycle is not supported in v0.".format(ws.title))

        part = sheet_parts.get(ws.title)
        if part is None or part not in names:
            _refuse("cannot locate the package part for sheet {0!r} via "
                    "the workbook relationships.".format(ws.title))
        original = zin.read(part)
        if shift_ops:
            # Phase 6b: the byte renumber runs first (deleted rows cut,
            # shifted r attributes rewritten, all other bytes verbatim);
            # the standard splice then treats the shifted bytes as its
            # baseline
            from .structural import apply_shift_to_bytes
            for op, op_idx, op_amount in shift_ops:
                original = apply_shift_to_bytes(original, op, op_idx,
                                                op_amount)
            baselines[part] = original
        scan = scan_sheet(original)
        dirty = resolve_dirty_cells(ws, ledger_dirty, scan)

        region_changes = {tag: rendered
                          for tag, rendered in all_region_changes.items()
                          if tag not in _CUSTOM_REGIONS}

        cf_replacement = None
        if "conditionalFormatting" in all_region_changes:
            cf_replacement = render_cf_for_write(ws)

        hyperlinks_replacement = None
        if shift_ops and "hyperlinks" not in all_region_changes \
                and hyperlink_signatures(ws):
            # anchors moved with their cells: re-render the element (the
            # relationship ids on the link objects are unchanged)
            hyperlinks_replacement = render_hyperlinks_for_write(ws)
        if "hyperlinks" in all_region_changes:
            hyperlinks_replacement, rels_update = _plan_hyperlinks(
                workbook, ws, led, zin, part, names)
            if rels_update is not None:
                sheet_rels_updates[rels_update[0]] = rels_update[1]

        if not (dirty or region_changes or row_changes or shift_ops
                or cf_replacement is not None
                or hyperlinks_replacement is not None):
            continue
        if translator is None and any(
                ws._cells[(r, c)]._style is not None
                for (r, c) in dirty if (r, c) in ws._cells):
            _refuse("styled cells cannot be written: the package has no "
                    "xl/styles.xml part and part creation is not supported "
                    "in v0.")
        plan[part] = splice_sheet(
            ws, original, dirty, region_changes, row_changes, scan=scan,
            cf_replacement=cf_replacement,
            hyperlinks_replacement=hyperlinks_replacement,
            style_resolver=translator.resolver() if translator else None)
        dirty_by_part[part] = dirty

    # ---- calcChain cascade (D13) ------------------------------------------
    drop_calcchain = led.formulas_changed and _CALC_CHAIN in names

    # ---- styles append (runs AFTER splices: resolution allocates new xfs) --
    styles_plan = None
    if translator is not None:
        styles_plan = crosspart.plan_styles_xml(workbook, led,
                                                zin.read(ARC_STYLE),
                                                translator)
    else:
        # a package without styles.xml cannot take style appends
        from .ledger import _style_fingerprint
        lengths, _fp = _style_fingerprint(workbook)
        if (lengths != led._style_lengths
                or len(workbook._cell_styles) != led.orig_cell_styles_len):
            _refuse("styles were added but the package has no xl/styles.xml "
                    "part; part creation is not supported in v0.")

    # ---- workbook.xml plan -------------------------------------------------
    wb_xml_plan = crosspart.plan_workbook_xml(
        workbook, led, zin.read(wb_part), new_sheet_entries,
        force_tags=("calcPr",) if force_calcpr else ())

    # ---- workbook rels + content types -------------------------------------
    wb_rels_plan = None
    if wb_rels_appends or drop_calcchain:
        payload = zin.read(wb_rels_part)
        if drop_calcchain:
            payload = crosspart.rels_remove_by_target_suffix(
                payload, "calcChain.xml")
        if wb_rels_appends:
            payload = crosspart.rels_append(payload, wb_rels_appends)
        wb_rels_plan = payload

    ct_plan = None
    if ct_appends or drop_calcchain:
        payload = zin.read("[Content_Types].xml")
        if drop_calcchain:
            payload = crosspart.ct_remove_override(payload, _CALC_CHAIN)
        if ct_appends:
            payload = crosspart.ct_append_overrides(payload, ct_appends)
        ct_plan = payload

    core_changed = render_core_model(workbook) != led.core_snapshot
    if core_changed and ARC_CORE not in names:
        _refuse("document properties changed but the package has no "
                "docProps/core.xml part; part creation is not supported "
                "in v0.")
    custom_changed = (render_custom_model(workbook) != led.custom_snapshot
                      and ARC_CUSTOM in names)

    theme_changed = False
    if workbook.loaded_theme is not None and ARC_THEME in names:
        theme_changed = workbook.loaded_theme != zin.read(ARC_THEME)

    # ---- assemble -----------------------------------------------------------
    def build(zout):
        for info in zin.infolist():
            name = info.filename
            if name == _CALC_CHAIN and drop_calcchain:
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
                zipio.write_entry(zout, name, render_custom_model(workbook))
            elif name == ARC_THEME and theme_changed:
                zipio.write_entry(zout, name, workbook.loaded_theme)
            else:
                zipio.copy_entry(zin, info, zout)
        for part_name, payload in new_sheet_parts:
            zipio.write_entry(zout, part_name, payload)
        for part_name, payload in new_rels_parts:
            zipio.write_entry(zout, part_name, payload)
        # rels parts created for LOADED sheets that had none (first
        # hyperlink on a rels-less sheet): they exist only in the plan
        for part_name, payload in sheet_rels_updates.items():
            if part_name not in names:
                zipio.write_entry(zout, part_name, payload)

    data = zipio.build_archive_bytes(build)

    if os.environ.get("PAPER_LEDGER_CROSSCHECK") == "1" and plan:
        from .crosscheck import verify_splice
        verify_splice(source, data, dirty_by_part, baselines=baselines)

    zipio.deliver(data, target)
    return True


def _namelist(source):
    with zipfile.ZipFile(io.BytesIO(source)) as z:
        return set(z.namelist())


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
    if getattr(ws, "_charts", None) or getattr(ws, "_images", None):
        _refuse("sheet {0!r} was added with charts or images; generating "
                "drawing parts under preserve mode is not supported in v0 "
                "(PR-0 D9 partial deferral).".format(ws.title))
    if getattr(ws, "_pivots", None):
        _refuse("sheet {0!r} was added with pivot tables; not supported in "
                "v0.".format(ws.title))
    if ws.tables:
        _refuse("sheet {0!r} was added with tables; table-part generation "
                "is not supported in v0.".format(ws.title))
    for cell in ws._cells.values():
        if cell._comment is not None:
            _refuse("sheet {0!r} was added with comments; comment-part "
                    "generation is not supported in v0.".format(ws.title))


def _generate_sheet_part(ws):
    """Generate a NEW sheet's part payload with the stock writer (the sheet
    exists only in the model — there is nothing to splice against). Returns
    (payload, rels_payload_or_None)."""
    from openpyxl.worksheet._writer import WorksheetWriter

    writer = WorksheetWriter(ws, out=io.BytesIO())
    writer.write()
    payload = writer.read()
    rels_payload = None
    if len(writer._rels):
        entries = [(rel.Id, rel.Type, rel.Target,
                    rel.TargetMode or None) for rel in writer._rels]
        rels_payload = crosspart.render_rels_document(entries)
    return payload, rels_payload


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


def _plan_hyperlinks(workbook, ws, led, zin, sheet_part, names):
    """Hyperlink ADDITIONS on a loaded sheet: allocate relationship ids,
    render the new hyperlinks element, and return the updated sheet-rels
    payload. Removals/changes refuse (dangling or rewritten relationships)."""
    arm = led.region_snapshots.get(ws, {}).get("hyperlinks", {})
    now = hyperlink_signatures(ws)
    removed = set(arm) - set(now)
    changed = {k for k in set(arm) & set(now) if arm[k] != now[k]}
    if removed or changed:
        _refuse("hyperlinks were removed or modified on sheet {0!r}; only "
                "hyperlink ADDITION is supported in v0 (removal would leave "
                "or rewrite preserved relationships).".format(ws.title))
    added = set(now) - set(arm)
    if not added:
        return render_hyperlinks_for_write(ws), None

    rels_part = _rels_path(sheet_part)
    if rels_part in names:
        rels_payload = zin.read(rels_part)
        next_rid = crosspart.rels_next_rid(rels_payload)
    else:
        rels_payload = None
        next_rid = 1

    entries = []
    counter = 0
    for (row, col) in sorted(added):
        cell = ws._cells[(row, col)]
        link = cell._hyperlink
        if link.target:
            rid = "rId{0}".format(next_rid + counter)
            counter += 1
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
