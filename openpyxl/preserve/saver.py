# paper-xlsx: the preserve-mode save (CONVENTIONS §3.4/§3.5; PR-0 §3/§6)

"""Save dispatch target for preserve-mode workbooks.

Ordered-stream splice: untouched parts raw-copy from the retained bytes
(byte-identical by construction); touched worksheet parts are spliced;
everything is validated BEFORE the first output byte, so every refusal is
atomic. Build stage: Phase 2c — cell/region splice live; cross-part edits
(sheet additions, styles append, calcChain cascade, comments, hyperlinks,
conditional formatting, workbook.xml, mark_dirty parts) still refuse with
typed errors and land in Phase 2d.
"""

import io
import os
import zipfile

from openpyxl.errors import UnsupportedStructureError
from openpyxl.xml.constants import ARC_CORE, ARC_THEME

from . import zipio
from .ledger import render_core_model, render_custom_model, render_workbook_model
from .regions import diff_regions, diff_row_attrs
from .splice import resolve_dirty_cells, splice_sheet
from .xmlscan import scan_sheet


def _refuse(msg):
    raise UnsupportedStructureError(msg + " Nothing was written.")


def save_preserved(workbook, target):
    """Save a preserve-mode workbook to ``target`` (path or binary
    file-like). Validates fully, then writes atomically."""
    led = workbook._paper_ledger
    source = workbook._paper_source
    if led is None or source is None:
        _refuse("preserve-mode save requires a workbook loaded with "
                "preserve=True.")

    if workbook.data_only:
        _refuse(
            "saving a workbook loaded with data_only=True would write "
            "cached values over formulas for every edited cell (formulas "
            "were never loaded). Reload without data_only=True to edit.")

    # Tier-3 guard: in-place mutation of shared interned styles
    led.check_style_registry(workbook)

    # ---- build-stage refusals (all lifted in Phase 2d) -----------------
    if led.added_sheets:
        _refuse("saving with sheets added in-session requires new-part "
                "generation plus workbook.xml/rels/[Content_Types] edits, "
                "which land in Phase 2d.")
    if led.parts:
        _refuse("mark_dirty() part-level re-serialization lands in "
                "Phase 2d.")
    if render_workbook_model(workbook) != led.workbook_snapshot:
        _refuse("workbook-level state changed (sheet state/order, defined "
                "names, calculation properties, book views, protection or "
                "code name); splicing workbook.xml lands in Phase 2d.")
    if render_custom_model(workbook) != led.custom_snapshot:
        _refuse("custom document properties changed; writing them lands in "
                "Phase 2d.")
    for cs, snap in led.chartsheet_snapshots.items():
        from .ledger import _render_chartsheet
        if _render_chartsheet(cs) != snap:
            _refuse("chartsheet {0!r} changed; chartsheet splicing is not "
                    "supported in v0.".format(cs.title))

    zin = zipfile.ZipFile(io.BytesIO(source))
    names = set(zin.namelist())

    core_changed = render_core_model(workbook) != led.core_snapshot
    if core_changed and ARC_CORE not in names:
        _refuse("document properties changed but the package has no "
                "docProps/core.xml part; adding parts lands in Phase 2d.")

    theme_changed = False
    if workbook.loaded_theme is not None and ARC_THEME in names:
        theme_changed = workbook.loaded_theme != zin.read(ARC_THEME)

    # ---- per-sheet plan -------------------------------------------------
    sheet_parts = _sheet_part_map(zin)
    plan = {}
    dirty_by_part = {}
    for ws in workbook.worksheets:
        if ws in led.added_sheets:
            continue  # unreachable (refused above); defensive
        ledger_dirty = led.dirty_coordinates(ws)
        region_changes = diff_regions(ws, led.region_snapshots.get(ws, {}))
        row_changes = diff_row_attrs(ws, led.row_attr_snapshots.get(ws, {}))
        comments_changed = _comments_changed(ws, led)
        maybe_rich = led.rich_text_mode
        if not (ledger_dirty or region_changes or row_changes
                or comments_changed or maybe_rich):
            continue
        if comments_changed:
            _refuse("comments changed on sheet {0!r}; comment-part editing "
                    "lands in Phase 2d.".format(ws.title))
        part = sheet_parts.get(ws.title)
        if part is None or part not in names:
            _refuse("cannot locate the package part for sheet {0!r} via "
                    "the workbook relationships.".format(ws.title))
        original = zin.read(part)
        scan = scan_sheet(original)
        dirty = resolve_dirty_cells(ws, ledger_dirty, scan)
        if not (dirty or region_changes or row_changes):
            continue
        _check_no_new_styles(workbook, ws, dirty, led)
        plan[part] = splice_sheet(ws, original, dirty, region_changes,
                                  row_changes, scan=scan)
        dirty_by_part[part] = dirty

    if led.formulas_changed and "xl/calcChain.xml" in names:
        _refuse("formulas changed and the package carries xl/calcChain.xml; "
                "the calcChain deletion cascade lands in Phase 2d.")

    # ---- assemble -------------------------------------------------------
    def build(zout):
        for info in zin.infolist():
            name = info.filename
            if name in plan:
                zipio.write_entry(zout, name, plan[name])
            elif name == ARC_CORE and core_changed:
                zipio.write_entry(zout, name, render_core_model(workbook))
            elif name == ARC_THEME and theme_changed:
                zipio.write_entry(zout, name, workbook.loaded_theme)
            else:
                zipio.copy_entry(zin, info, zout)

    data = zipio.build_archive_bytes(build)

    if os.environ.get("PAPER_LEDGER_CROSSCHECK") == "1" and plan:
        from .crosscheck import verify_splice
        verify_splice(source, data, dirty_by_part)

    zipio.deliver(data, target)
    return True


def _comments_changed(ws, led):
    from .ledger import _comment_snapshot

    return _comment_snapshot(ws) != led.comment_snapshots.get(ws, {})


def _check_no_new_styles(workbook, ws, dirty, led):
    """Phase-2c stage guard: dirty cells must reuse xf entries that already
    exist in the original stylesheet (new-style appends land in 2d).

    Membership is checked by linear scan over the original prefix — never
    IndexedList.index(), whose duplicate-handling returns wrong indices
    (measured upstream bug)."""
    limit = led.orig_cell_styles_len
    originals = list(workbook._cell_styles)[:limit]
    for (row, col) in sorted(dirty):
        cell = ws._cells.get((row, col))
        if cell is None or cell._style is None:
            continue
        if not any(existing == cell._style for existing in originals):
            _refuse(
                "cell {0} on sheet {1!r} uses a style that does not exist "
                "in the original stylesheet; appending new styles lands in "
                "Phase 2d.".format(cell.coordinate, ws.title))


def _sheet_part_map(zin):
    """Map sheet titles to their package part names, rels-driven (PR-0
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

    rels_path = get_rels_path(wb_part)
    rels = get_dependents(zin, rels_path)
    # get_dependents already resolves each rel's .target to an absolute
    # normalized part name (relationship.py:106-129)
    id_to_target = {rel.Id: rel.target for rel in rels
                    if rel.TargetMode != "External"}

    mapping = {}
    root = fromstring(zin.read(wb_part))
    ns_main = root.tag.split("}")[0].strip("{")
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    for sheet_el in root.iter("{%s}sheet" % ns_main):
        name = sheet_el.get("name")
        rid = sheet_el.get("{%s}id" % rel_ns)
        if name and rid and rid in id_to_target:
            mapping[name] = id_to_target[rid]
    return mapping


