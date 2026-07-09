# paper-xlsx: the structural-edit guard (PLAN Phase 6a; PR-0 §8)

"""Analyze what a row/column shift would strand.

The scariest damage in the model is here: ``insert_rows`` moves cells while
updating NOTHING — not formulas, not defined names, not chart ranges — so
one inserted row silently corrupts every SUM below it with numbers that
look plausible (measured: LibreOffice computes 1100/6399/5400 where the
correct answers are 7499/6500 — OPEN-QUESTIONS Q11). Under preserve mode
the shift refuses with the precise list of what would break; the stock path
keeps stock behavior plus a loud warning.
"""

import io
import re
import zipfile

from openpyxl.errors import UnsupportedStructureError

MAX_COL = 1 << 20
MAX_ROW = 1 << 22

# hard sheet limits (ECMA-376): shifting occupied cells past these is a
# BoundaryViolationError, not a silent wrap or drop
EXCEL_MAX_ROW = 1048576
EXCEL_MAX_COL = 16384


class AddressRemap:
    """How one structural edit moved addresses (CONVENTIONS §2, pinned):
    every pre-edit address must be remapped through this, never reused.

    ``map('Model!B12') -> 'Model!B13'``; addresses whose cells the edit
    deleted map to ``None``; addresses on other sheets (or untouched by
    the shift) come back unchanged. Accepts bare cells, ranges, and
    sheet-qualified forms; ``$`` markers are kept positionally, matching
    the rewriter's Excel semantics."""

    def __init__(self, sheet_title, operation, index, amount):
        self.sheet_title = sheet_title
        self.operation = operation
        self.index = index
        self.amount = amount

    def map(self, address):
        from .rewrite import REF_ERROR, shift_ref

        sheet, ref = None, address
        m = _SHEET_PREFIX_RE.match(address)
        if m:
            sheet = (m.group(1) or m.group(2)).replace("''", "'")
            ref = m.group(3)
        if sheet is not None \
                and sheet.casefold() != self.sheet_title.casefold():
            return address
        axis = "rows" if self.operation.endswith("_rows") else "cols"
        is_delete = self.operation.startswith("delete")
        shifted = shift_ref(ref, axis, self.index, self.amount, is_delete)
        if shifted == REF_ERROR:
            return None
        if sheet is None:
            return shifted
        return address[:m.end(0) - len(m.group(3))] + shifted

    def __repr__(self):
        return "AddressRemap({0!r}, {1}, index={2}, amount={3})".format(
            self.sheet_title, self.operation, self.index, self.amount)


_SHEET_PREFIX_RE = re.compile(r"^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$")


def shift_bounds(kind, index):
    """The cell region a shift at ``index`` moves or destroys: everything
    at or after the index (deletes also destroy the range itself)."""
    if kind in ("insert_rows", "delete_rows"):
        return (1, index, MAX_COL, MAX_ROW)
    return (index, 1, MAX_COL, MAX_ROW)


def _intersects(bounds, min_col, min_row, max_col, max_row):
    b_min_col, b_min_row, b_max_col, b_max_row = bounds
    return not (b_max_col < min_col or b_min_col > max_col
                or b_max_row < min_row or b_min_row > max_row)


def analyze_shift(ws, kind, index):
    """Everything a shift would strand, as human-readable impact lines."""
    from openpyxl.utils.cell import range_boundaries

    from .perception import dependency_sketch

    wb = ws.parent
    bounds = shift_bounds(kind, index)
    min_col, min_row, max_col, max_row = bounds
    impacts = []

    # formulas (cross-sheet included) whose references intersect the region
    sketch = dependency_sketch(wb)
    formulas = sketch.cells_referencing(ws.title, bounds)
    if formulas:
        shown = ", ".join(formulas[:8])
        if len(formulas) > 8:
            shown += ", ... ({0} total)".format(len(formulas))
        impacts.append(
            "formulas referencing the shifted cells would keep their old "
            "ranges and silently compute wrong numbers: {0}".format(shown))

    # defined names pointing into the region
    stranded_names = []
    name_sources = [(None, wb.defined_names)]
    name_sources += [(sheet.title, sheet.defined_names)
                     for sheet in wb.worksheets]
    for scope, names in name_sources:
        for name, dn in names.items():
            try:
                for dest_sheet, dest_ref in dn.destinations:
                    if dest_sheet != ws.title:
                        continue
                    dest_bounds = range_boundaries(dest_ref.replace("$", ""))
                    if _intersects(dest_bounds, *_norm(bounds)):
                        stranded_names.append(name)
                        break
            except Exception:
                continue
    if stranded_names:
        impacts.append(
            "defined name(s) {0} would keep pointing at the old "
            "cells".format(", ".join(sorted(set(stranded_names)))))

    # merged ranges
    merged = [str(r) for r in ws.merged_cells.ranges
              if _ranges_hit(r, bounds)]
    if merged:
        impacts.append("merged range(s) {0} would not move with their "
                       "content".format(", ".join(sorted(merged))))

    # conditional formatting and data validation ranges
    cf_hit = []
    for cf in ws.conditional_formatting:
        for rng in getattr(cf.sqref, "ranges", []):
            if _ranges_hit(rng, bounds):
                cf_hit.append(str(cf.sqref))
                break
    if cf_hit:
        impacts.append("conditional formatting on {0} would keep the old "
                       "ranges".format(", ".join(sorted(set(cf_hit)))))
    dv_hit = []
    if ws.data_validations:
        for dv in ws.data_validations.dataValidation:
            for rng in getattr(dv.sqref, "ranges", []):
                if _ranges_hit(rng, bounds):
                    dv_hit.append(str(dv.sqref))
                    break
    if dv_hit:
        impacts.append("data validation on {0} would keep the old "
                       "ranges".format(", ".join(sorted(set(dv_hit)))))

    # tables
    tables_hit = [name for name, ref in ws.tables.items()
                  if _ref_hit(ref, bounds)]
    if tables_hit:
        impacts.append("table(s) {0} would keep their old extents".format(
            ", ".join(sorted(tables_hit))))

    # preserved charts: their XML is raw retained bytes — a shift makes them
    # point at wrong rows with no error anywhere (PR-0 §8: refusal is the
    # only honest v0 answer on chart-referenced sheets)
    led_ref = getattr(wb, "_paper_ledger", None)
    chart_title = led_ref.renames.get(ws, ws.title) if led_ref else ws.title
    charts = _charts_referencing(wb, chart_title)
    if charts:
        impacts.append(
            "preserved chart(s) ({0}) reference this sheet; their series "
            "ranges live in preserved bytes and would point at the wrong "
            "rows".format(", ".join(sorted(charts))))

    return impacts


def _norm(bounds):
    return bounds


def _ranges_hit(cell_range, bounds):
    return _intersects((cell_range.min_col, cell_range.min_row,
                        cell_range.max_col, cell_range.max_row), *bounds)


def _ref_hit(ref, bounds):
    from openpyxl.utils.cell import range_boundaries

    try:
        return _intersects(range_boundaries(ref.replace("$", "")), *bounds)
    except Exception:
        return True   # unparseable: assume affected (conservative)


def _title_needles(sheet_title):
    """Byte needles for a sheet title inside preserved XML: raw and
    XML-escaped forms, quoted variants, all lower-cased for case-insensitive
    search (Excel resolves sheet names case-insensitively)."""
    escaped = (sheet_title.replace("&", "&amp;").replace("<", "&lt;")
               .replace(">", "&gt;"))
    variants = {sheet_title, escaped,
                "'{0}'".format(sheet_title.replace("'", "''")),
                "'{0}'".format(escaped.replace("'", "''")),
                "'{0}'".format(escaped.replace("'", "&apos;&apos;"))}
    return [v.encode("utf-8").lower() for v in variants]


def _charts_referencing(wb, sheet_title):
    """Names of retained chart parts whose XML mentions the sheet."""
    source = getattr(wb, "_paper_source", None)
    if not source:
        return []
    needles = _title_needles(sheet_title)
    hits = []
    with zipfile.ZipFile(io.BytesIO(source)) as z:
        for name in z.namelist():
            if (name.startswith("xl/charts/") and name.endswith(".xml")
                    and "/_rels/" not in name):
                payload = z.read(name).lower()
                if any(needle in payload for needle in needles):
                    hits.append(name)
    return hits


STOCK_WARNING = (
    "{0}() moves cells but updates NOTHING that points at them: formulas "
    "keep their old ranges, defined names and chart series keep their old "
    "cells — the numbers that come out will look plausible and be wrong. "
    "Open the file with preserve=True to get a safety analysis instead."
)


# ---------------------------------------------------------------------
# Phase 6b: performing the shift (model fixups + byte renumber)

def shift_blockers(ws, operation, index, amount=1):
    """Content that makes a shift unsafe to REWRITE in v0 — anything whose
    references live outside the fully-modeled set. A non-empty result means
    refusal (with analyze_shift providing the victim list)."""
    wb = ws.parent
    blockers = []
    source = getattr(wb, "_paper_source", None)
    led = getattr(wb, "_paper_ledger", None)
    # multiple shifts per session compose: model fixups and snapshot
    # rebases run at edit time in order, the byte renumber replays the
    # recorded ops in order at save (PLAN-v0.1 3.3 retired the
    # one-shift-per-session refusal)
    led_ref = getattr(wb, "_paper_ledger", None)
    lookup_title = led_ref.renames.get(ws, ws.title) if led_ref else ws.title
    # in-session charts are model-rendered at save: a delete that removes
    # their charted cells has no honest rewrite — block BEFORE any cell
    # moves (Batch-4 gate)
    if led_ref is not None and operation.startswith("delete"):
        from .rewrite import shift_name_value

        axis_ = "rows" if "rows" in operation else "cols"
        for sheet in wb.worksheets:
            armed_charts = (led_ref.object_snapshots.get(sheet) or {}).get(
                "chart", {})
            for i, chart in enumerate(getattr(sheet, "_charts", []) or []):
                if i in armed_charts:
                    continue
                for f in _chart_source_refs(chart):
                    new_f, chg = shift_name_value(f, ws.title, axis_,
                                                  index, amount, True)
                    if chg and "#REF" in new_f and "#REF" not in f:
                        blockers.append(
                            "an in-session chart charts {0!r}, which this "
                            "delete removes".format(f))
    part_payload = _sheet_payload(wb, lookup_title)
    if part_payload is None:
        blockers.append("the sheet's package part could not be located")
        return blockers
    if b"extLst" in part_payload:
        blockers.append(
            "the sheet carries extension content (extLst: sparklines, x14 "
            "rules, ...) whose cell ranges live in unmodeled bytes")
    if b"t=\"array\"" in part_payload or b"t='array'" in part_payload:
        blockers.append("the sheet carries array formulas (ref rewriting "
                        "for spill ranges is not supported in v0)")
    if b"t=\"dataTable\"" in part_payload \
            or b"t='dataTable'" in part_payload:
        blockers.append("the sheet carries what-if data tables; their "
                        "ref/r1/r2 inputs live in unmodeled bytes and "
                        "would silently mis-shift (Batch-3 gate)")
    if any(cell._comment is not None for cell in ws._cells.values()):
        blockers.append("the sheet carries comments; their anchors live in "
                        "comment/VML parts the shift cannot rewrite")
    if b"<legacyDrawing" in part_payload:
        blockers.append("the sheet references a legacy (VML) drawing")
    charts = _charts_referencing(wb, ws.title)
    if charts:
        # Phase 6c: chart series references and drawing anchors are
        # patchable in place — dry-run the planner; only its blockers
        # (extension machinery, would-be #REF!) refuse now
        from .chartpatch import plan_chart_updates

        _plans, chart_blockers = plan_chart_updates(
            wb, ws.title, operation, index, amount)
        blockers.extend(chart_blockers)
    pivots = _pivots_referencing(wb, ws.title)
    if pivots:
        blockers.append(
            "preserved pivot part(s) {0} reference this sheet".format(
                ", ".join(sorted(pivots))))
    if ws.tables:
        blockers.append("the sheet carries table(s) {0}; table-part "
                        "rewriting is not supported in v0".format(
                            ", ".join(sorted(ws.tables))))
    if ws.row_breaks or ws.col_breaks:
        blockers.append("the sheet carries manual page breaks, which "
                        "anchor to row/column numbers (not rewritten in v0)")
    for other in wb.worksheets:
        for (_r, _c), cell in other._cells.items():
            if cell.data_type == "f" and isinstance(cell._value, str) \
                    and "[" in cell._value and ws.title in cell._value:
                blockers.append(
                    "formula {0}!{1} uses a structured/external reference "
                    "mentioning this sheet".format(other.title,
                                                   cell.coordinate))
                break
    return blockers


def _pivots_referencing(wb, sheet_title):
    source = getattr(wb, "_paper_source", None)
    if not source:
        return []
    needles = _title_needles(sheet_title)
    hits = []
    with zipfile.ZipFile(io.BytesIO(source)) as z:
        for name in z.namelist():
            if name.startswith(("xl/pivotTables/", "xl/pivotCache/")) \
                    and name.endswith(".xml"):
                payload = z.read(name).lower()
                if any(needle in payload for needle in needles):
                    hits.append(name)
    return hits


def _sheet_payload(wb, title):
    source = getattr(wb, "_paper_source", None)
    if not source:
        return None
    import zipfile as _zf

    from .saver import _package_info

    with _zf.ZipFile(io.BytesIO(source)) as z:
        _wb_part, mapping = _package_info(z)
        part = mapping.get(title)
        if part is None:
            return None
        return z.read(part)


def _chart_source_refs(chart):
    """The live reference objects of a model chart's data sources —
    (yields each Ref with a string .f)."""
    for ser in getattr(chart, "series", []) or []:
        for src_name in ("val", "yVal", "xVal", "bubbleSize", "cat", "tx"):
            src = getattr(ser, src_name, None)
            if src is None:
                continue
            for ref_name in ("numRef", "strRef", "multiLvlStrRef"):
                ref = getattr(src, ref_name, None)
                if ref is not None and isinstance(getattr(ref, "f", None),
                                                  str):
                    yield ref.f


def _shift_added_chart_refs(chart, target_title, axis, index, amount,
                            is_delete, shift_name_value):
    """Rewrite one model chart's data-source references for a shift on
    ``target_title`` (deletes that would strand a chart were already
    blocked pre-move by shift_blockers)."""
    for ser in getattr(chart, "series", []) or []:
        for src_name in ("val", "yVal", "xVal", "bubbleSize", "cat", "tx"):
            src = getattr(ser, src_name, None)
            if src is None:
                continue
            for ref_name in ("numRef", "strRef", "multiLvlStrRef"):
                ref = getattr(src, ref_name, None)
                if ref is None or not isinstance(getattr(ref, "f", None),
                                                 str):
                    continue
                new_f, changed = shift_name_value(
                    ref.f, target_title, axis, index, amount, is_delete)
                if not changed:
                    continue
                if "#REF" in new_f and "#REF" not in ref.f:
                    raise UnsupportedStructureError(
                        "internal: a delete stranding an in-session "
                        "chart ({0!r}) escaped the pre-move blocker; the "
                        "model may be partially shifted — do not "
                        "save.".format(ref.f))
                ref.f = new_f


def apply_model_shift(ws, operation, index, amount):
    """All the reference updates stock openpyxl skips, applied to the MODEL
    after the cells moved (Excel insert/delete semantics via rewrite.py).
    Also rebases the positional arm snapshots so pure moves are not
    mis-detected as user changes."""
    from .ledger import _armed_ledger_for_ws
    from .rewrite import (
        row_mapping,
        shift_cell_range,
        shift_formula,
        shift_name_value,
    )

    wb = ws.parent
    led = _armed_ledger_for_ws(ws)
    axis = "rows" if "rows" in operation else "cols"
    is_delete = operation.startswith("delete")
    mapper = row_mapping(operation, index, amount)

    # 0. rebase positional ledger/arm state FIRST: pre-shift dirty marks and
    # snapshots move with the cells; everything the fixups below record is
    # already in post-shift coordinates and must NOT be remapped again
    if led is not None:
        _rebase_snapshots(led, ws, mapper, axis)

    # 1. formulas everywhere in the workbook that reference this sheet
    for other in wb.worksheets:
        for (row, col), cell in list(other._cells.items()):
            if cell.data_type != "f" or not isinstance(cell._value, str):
                continue
            new_formula, changed = shift_formula(
                cell._value, other.title, ws.title, axis, index, amount,
                is_delete)
            if changed:
                cell.value = new_formula     # through the chokepoint: dirty

    # 2. defined names (workbook- and sheet-scoped) and print settings
    for names in [wb.defined_names] + [s.defined_names
                                       for s in wb.worksheets]:
        for name in list(names):
            dn = names[name]
            if not isinstance(dn.value, str):
                continue
            new_value, changed = shift_name_value(
                dn.value, ws.title, axis, index, amount, is_delete)
            if changed:
                dn.attr_text = new_value

    # 2b. charts ADDED this session are model-rendered at save, so their
    # data-source references must follow the shift like every other model
    # reference (loaded charts' parts are byte-patched by
    # plan_chart_updates instead) — Batch-4 gate: an added chart's range
    # silently pointed at the pre-shift cells
    if led is not None:
        for sheet in wb.worksheets:
            armed_charts = (led.object_snapshots.get(sheet) or {}).get(
                "chart", {})
            for i, chart in enumerate(getattr(sheet, "_charts", []) or []):
                if i in armed_charts:
                    continue
                _shift_added_chart_refs(chart, ws.title, axis, index,
                                        amount, is_delete,
                                        shift_name_value)

    # 3. sheet-internal regions (fully modeled; the splice re-renders them)
    for rng in list(ws.merged_cells.ranges):
        if shift_cell_range(rng, axis, index, amount, is_delete) == "deleted":
            ws.merged_cells.ranges.remove(rng)
    for cf in list(ws.conditional_formatting):
        for rng in list(getattr(cf.sqref, "ranges", [])):
            if shift_cell_range(rng, axis, index, amount,
                                is_delete) == "deleted":
                cf.sqref.ranges.remove(rng)
    if ws.data_validations:
        for dv in list(ws.data_validations.dataValidation):
            for rng in list(getattr(dv.sqref, "ranges", [])):
                if shift_cell_range(rng, axis, index, amount,
                                    is_delete) == "deleted":
                    dv.sqref.ranges.remove(rng)
    if ws.auto_filter and ws.auto_filter.ref:
        from openpyxl.worksheet.cell_range import CellRange

        cr = CellRange(ws.auto_filter.ref)
        if shift_cell_range(cr, axis, index, amount, is_delete) == "changed":
            ws.auto_filter.ref = cr.coord

    if axis == "rows":
        # 4. row display attributes move with their rows
        new_dims = {}
        for r, dim in list(ws.row_dimensions.items()):
            new_row = mapper(r)
            if new_row is None:
                continue
            dim.index = new_row
            new_dims[new_row] = dim
        ws.row_dimensions.clear()
        ws.row_dimensions.update(new_dims)

    # 5. hyperlink anchors track their cells (both axes)
    for (_r, _c), cell in ws._cells.items():
        link = getattr(cell, "_hyperlink", None)
        if link is not None:
            link.ref = cell.coordinate

    if led is not None:
        led.shifts.setdefault(ws, []).append((operation, index, amount))
        led.formulas_changed = True


def _rebase_snapshots(led, ws, mapper, axis):
    """Every POSITIONAL piece of arm/ledger state (dirty cells, row attrs,
    hyperlink anchors, comments) must follow a pure move, or the save would
    lose pre-shift edits and mis-read the move as user changes."""
    if axis == "rows":
        def map_key(row, col):
            new_row = mapper(row)
            return None if new_row is None else (new_row, col)
    else:
        def map_key(row, col):
            new_col = mapper(col)   # the caller passes the column mapper
            return None if new_col is None else (row, new_col)

    dirty = led.cells.get(ws)
    if dirty:
        led.cells[ws] = {mapped for (r, c) in dirty
                         for mapped in (map_key(r, c),) if mapped is not None}
    links = led.region_snapshots.get(ws, {}).get("hyperlinks")
    if links:
        led.region_snapshots[ws]["hyperlinks"] = {
            mapped: sig for (r, c), sig in links.items()
            for mapped in (map_key(r, c),) if mapped is not None}
    comments = led.comment_snapshots.get(ws)
    if comments:
        led.comment_snapshots[ws] = {
            mapped: sig for (r, c), sig in comments.items()
            for mapped in (map_key(r, c),) if mapped is not None}
    if axis == "rows":
        rows = led.row_attr_snapshots.get(ws)
        if rows:
            led.row_attr_snapshots[ws] = {
                mapper(r): attrs for r, attrs in rows.items()
                if mapper(r) is not None}


def apply_shift_to_bytes(original, operation, index, amount):
    """The byte-level renumber pre-transform: deleted rows are cut; shifted
    rows get their r attributes (row and cells) rewritten; every other byte
    — cell contents, unmodeled attributes, spans — is copied verbatim. The
    result becomes the baseline the standard splice runs against."""
    from .rewrite import row_mapping
    from .xmlscan import scan_sheet

    axis = "rows" if "rows" in operation else "cols"
    scan = scan_sheet(original)
    edits = []

    if axis == "rows":
        mapper = row_mapping(operation, index, amount)
        for row_index, row_span in scan.rows.items():
            new_row = mapper(row_index)
            if new_row is None:
                edits.append((row_span.start, row_span.end, b""))
                continue
            if new_row == row_index:
                continue
            head_end = (row_span.content_start
                        if not row_span.self_closing else row_span.end)
            head = original[row_span.start:head_end]
            new_head = re.sub(
                br'(<row[^>]*?\sr=")%d(")' % row_index,
                br"\g<1>%d\g<2>" % new_row, head, 1)
            edits.append((row_span.start, head_end, new_head))
            for col, cell_span in row_span.cells.items():
                cell_head_end = original.index(b">", cell_span.start) + 1
                cell_head = original[cell_span.start:cell_head_end]
                old_ref = cell_span.attrs["r"].encode("ascii")
                letters = old_ref.rstrip(b"0123456789")
                new_ref = letters + str(new_row).encode("ascii")
                new_cell_head = cell_head.replace(
                    b'r="%s"' % old_ref, b'r="%s"' % new_ref, 1)
                edits.append((cell_span.start, cell_head_end, new_cell_head))
    else:
        from openpyxl.utils import column_index_from_string, get_column_letter

        from .rewrite import _shift_span
        is_delete = operation.startswith("delete")
        for row_index, row_span in scan.rows.items():
            for col, cell_span in sorted(row_span.cells.items()):
                span = _shift_span(col, col, index, amount, is_delete)
                if span is None:
                    edits.append((cell_span.start, cell_span.end, b""))
                    continue
                if span[0] == col:
                    continue
                cell_head_end = original.index(b">", cell_span.start) + 1
                cell_head = original[cell_span.start:cell_head_end]
                old_ref = cell_span.attrs["r"].encode("ascii")
                digits = old_ref.lstrip(
                    b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz")
                new_ref = get_column_letter(span[0]).encode("ascii") + digits
                edits.append((cell_span.start, cell_head_end,
                              cell_head.replace(b'r="%s"' % old_ref,
                                                b'r="%s"' % new_ref, 1)))

    if not edits:
        return original
    from .crosspart import apply_edits

    return apply_edits(original, edits)
