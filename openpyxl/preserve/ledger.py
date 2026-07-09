# paper-xlsx: the dirty ledger (CONVENTIONS §3.3; PLAN §B; PR-0 D5/D7/D8)

"""Records what the object model changed, so the splice save can apply
exactly those edits to the retained package and nothing else.

The ledger is load-bearing, not an optimization: a compare-based patch-save
is impossible here because stock serialization of a whole sheet is the lossy
act — there is nothing faithful to compare. (Fully-modeled satellite
*elements* are the sanctioned exception: their arm-time model serializations
are snapshotted below and compared against save-time re-serializations —
self-consistent, so USER changes are detected with zero producer-quirk
noise.)

What the ledger holds:

- per-cell dirt, keyed by worksheet object: the coordinates whose model
  state must be spliced into the original sheet XML (a dirty coordinate
  absent from the model means the cell was deleted);
- ``formulas_changed`` — drives the calcChain deletion cascade and the
  recalc-on-load flag;
- sheets added in-session (generated whole at save; also exempt from the
  rename/remove refusals that protect loaded sheets);
- explicit part-level dirt from :meth:`Workbook.mark_dirty`;
- a fingerprint of the interned style components taken when the ledger
  armed, so the save can refuse if a shared style object was mutated in
  place (the StyleProxy nested-leak — silent fan-out corruption upstream).

The ledger ARMS only after load completes: the reader itself fires cell
binds, ``create_sheet`` and style writes while building the model, and pure
reads may materialize cells and dimensions afterwards — the ledger keys on
semantic mutation through public setters, never on materialization.
"""

from openpyxl.errors import (
    TargetNotFoundError,
    UnsupportedStructureError,
)


class DirtyLedger:

    __slots__ = ("armed", "cells", "parts", "formulas_changed",
                 "added_sheets", "loaded_sheet_titles", "_style_lengths",
                 "_style_fingerprint", "region_snapshots", "row_attr_snapshots",
                 "comment_snapshots", "workbook_snapshot", "core_snapshot",
                 "custom_snapshot", "chartsheet_snapshots", "pinned_regions",
                 "object_snapshots", "external_links_snapshot",
                 "protection_warned", "replaced_parts", "renames",
                 "sheet_order", "removed_sheets", "value_overwrites",
                 "orig_cell_styles_len", "rich_text_mode",
                 "sheet_states", "dxfs_len", "named_styles_len", "shifts",
                 "template_flag", "cache_writes")

    def __init__(self):
        self.armed = False
        self.cells = {}                # ws object -> set[(row, col)]
        self.parts = set()             # part names marked via mark_dirty
        self.formulas_changed = False
        self.added_sheets = set()      # ws objects created after arming
        self.loaded_sheet_titles = frozenset()
        self._style_lengths = ()
        self._style_fingerprint = ()
        # arm-time model serializations: comparing them against save-time
        # re-serializations detects USER changes with no producer-quirk
        # noise (PR-0 D5 Tier 2 realized as snapshot-vs-snapshot)
        self.region_snapshots = {}     # ws -> {tag: rendered}
        self.row_attr_snapshots = {}   # ws -> {row: attr tuple}
        self.comment_snapshots = {}    # ws -> {(row, col): (text, author)}
        self.workbook_snapshot = None  # workbook.xml rendered from the model
        self.core_snapshot = None
        self.custom_snapshot = None
        self.chartsheet_snapshots = {} # chartsheet -> rendered
        self.pinned_regions = {}       # ws -> {tag}: impure serializers
        self.object_snapshots = {}     # ws -> preserved-part-backed objects
        self.external_links_snapshot = ()
        self.protection_warned = set()   # sheets warned once (1.6)
        self.replaced_parts = {}         # raw byte swaps (PR-1 1.4)
        self.renames = {}                # ws -> ORIGINAL title (3.2)
        self.sheet_order = []            # _sheets titles at arm (3.2)
        self.removed_sheets = []         # ORIGINAL titles removed (3.2)
        self.value_overwrites = {}       # ws -> coords whose VALUE changed
        self.orig_cell_styles_len = 0
        self.rich_text_mode = False
        self.sheet_states = {}         # title -> state at arm (all sheets)
        self.dxfs_len = 0
        self.named_styles_len = 0
        self.shifts = {}               # ws -> [(operation, index, amount)]
        self.template_flag = False
        self.cache_writes = {}         # ws -> {(row, col): computed value}
                                       # (oracle write-back, PLAN-v0.1 5.3)

    # -- arming --------------------------------------------------------

    @classmethod
    def arm(cls, wb, rich_text=False):
        from .regions import snapshot_regions, snapshot_row_attrs

        led = cls()
        led.loaded_sheet_titles = frozenset(wb.sheetnames)
        led._style_lengths, led._style_fingerprint = _style_fingerprint(wb)
        for ws in wb.worksheets:
            # double-render (PLAN-v0.1 0.3): a serializer with render-time
            # side effects disagrees with itself across passes. Regions
            # where that happens are PINNED — the snapshot keeps the
            # settled second pass, and the saver refuses edits to them
            # rather than trusting an untrustworthy render (an impure
            # serializer must land in "pinned", never in "false dirty")
            first = snapshot_regions(ws)
            settled = snapshot_regions(ws)
            led.pinned_regions[ws] = {
                tag for tag, rendered in settled.items()
                if first.get(tag) != rendered}
            led.region_snapshots[ws] = settled
            led.row_attr_snapshots[ws] = snapshot_row_attrs(ws)
            led.comment_snapshots[ws] = _comment_snapshot(ws)
            led.object_snapshots[ws] = _object_snapshot(ws)
        led.external_links_snapshot = _external_links_snapshot(wb)
        for cs in wb.chartsheets:
            led.chartsheet_snapshots[cs] = _render_chartsheet(cs)
            # chartsheet-anchored charts are the same preserved-part-backed
            # boundary (Batch-1 gate: they were outside it entirely)
            led.object_snapshots[cs] = _object_snapshot(cs)
        from .crosspart import render_workbook_elements

        led.workbook_snapshot = render_workbook_elements(wb)
        led.core_snapshot = render_core_model(wb)
        led.custom_snapshot = render_custom_model(wb)
        led.orig_cell_styles_len = len(wb._cell_styles)
        led.rich_text_mode = rich_text
        led.sheet_states = {s.title: s.sheet_state for s in wb._sheets}
        led.sheet_order = [s.title for s in wb._sheets]
        led.template_flag = bool(wb.template)
        led.dxfs_len = len(wb._differential_styles.styles)
        led.named_styles_len = len(wb._named_styles)
        led.armed = True
        return led

    # -- recording ------------------------------------------------------

    def mark_cell(self, ws, row, column):
        self.cells.setdefault(ws, set()).add((row, column))

    def dirty_coordinates(self, ws):
        return self.cells.get(ws, set())

    def is_loaded_sheet(self, ws):
        return ws not in self.added_sheets and ws.title in self.loaded_sheet_titles

    # -- style in-place mutation check (PR-0 D5 Tier 3a) ----------------

    def check_style_registry(self, wb):
        """Refuse if any interned style component that existed at arm time
        was mutated in place (the StyleProxy nested-object leak): such a
        mutation silently restyles every aliased cell and cannot be
        expressed as an append-only styles.xml edit."""
        lengths, fingerprint = _style_fingerprint(
            wb, limits=self._style_lengths)
        if fingerprint != self._style_fingerprint:
            raise UnsupportedStructureError(
                "a shared style object was mutated in place after loading "
                "(e.g. cell.font.color.rgb = ...): this silently restyles "
                "every cell using that style and corrupts the style "
                "registry. Reassign a copied style instead: "
                "cell.font = cell.font.copy(color=Color(rgb=...)). "
                "Nothing was written."
            )


def _style_fingerprint(wb, limits=None):
    """Serialize the interned style components (optionally only the first
    ``limits`` entries per collection — the entries that existed at arm
    time; later appends are legal)."""
    from openpyxl.xml.functions import tostring

    collections = (wb._fonts, wb._fills, wb._borders, wb._alignments,
                   wb._protections, wb._number_formats)
    lengths = tuple(len(c) for c in collections)
    caps = limits if limits is not None else lengths
    rendered = []
    for coll, cap in zip(collections, caps):
        for obj in list(coll)[:cap]:
            if hasattr(obj, "to_tree"):
                rendered.append(tostring(obj.to_tree()))
            else:
                rendered.append(repr(obj).encode())
    return lengths, tuple(rendered)


def _comment_snapshot(ws):
    snap = {}
    for (row, col), cell in ws._cells.items():
        comment = getattr(cell, "_comment", None)
        if comment is not None:
            # height/width included: a resize on a machinery-carrying
            # sheet must refuse, not vanish (Batch-2 gate)
            snap[(row, col)] = (comment.text, comment.author,
                                comment.height, comment.width)
    return snap


def _settled(render):
    """Render twice, keep the second (the 0.3 discipline: a serializer
    with render-time side effects settles after one pass; comparing
    settled-vs-settled is producer-quirk-free). Returns (settled bytes,
    self_consistent flag)."""
    first = render()
    second = render()
    return second, first == second


def _object_snapshot(ws):
    """Serialized fingerprints of the model objects whose backing parts
    are PRESERVED BYTES — the review's named threat class: loaded tables,
    charts, images, pivots are live and mutable, but the splice never
    re-serializes their parts, so an in-session edit would vanish
    silently. Snapshot at arm, compare at save, refuse on drift
    (PLAN-v0.1 1.1: refusal is fully acceptable; silence is not)."""
    from openpyxl.xml.functions import tostring

    snap = {"unstable": set()}

    tables = {}
    ws_tables = getattr(ws, "tables", None) or {}
    for name in ws_tables:
        tbl = ws_tables[name]
        rendered, ok = _settled(lambda t=tbl: tostring(t.to_tree()))
        tables[name] = rendered
        if not ok:
            snap["unstable"].add(("table", name))
    snap["table"] = tables

    charts = {}
    for i, chart in enumerate(getattr(ws, "_charts", []) or []):
        # the anchor is NOT part of chart._write() (it lives in the
        # preserved drawing part) — snapshot it too, or a chart move
        # vanishes silently (Batch-1 gate)
        rendered, ok = _settled(lambda c=chart: tostring(c._write()))
        charts[i] = (rendered, _anchor_fingerprint(chart))
        if not ok:
            snap["unstable"].add(("chart", i))
    snap["chart"] = charts

    images = {}
    for i, image in enumerate(getattr(ws, "_images", []) or []):
        anchor = getattr(image, "anchor", None)
        if anchor is not None and hasattr(anchor, "to_tree"):
            rendered, ok = _settled(lambda a=anchor: tostring(a.to_tree()))
        else:
            rendered, ok = repr(anchor).encode("utf-8"), True
        images[i] = (rendered, getattr(image, "path", None),
                     _image_data_digest(image))
        if not ok:
            snap["unstable"].add(("image", i))
    snap["image"] = images

    pivots = {}
    for i, pivot in enumerate(getattr(ws, "_pivots", []) or []):
        rendered, ok = _settled(lambda p=pivot: tostring(p.to_tree()))
        pivots[i] = rendered
        if not ok:
            snap["unstable"].add(("pivot", i))
    snap["pivot"] = pivots

    return snap


_OBJECT_UNLOCKS = {
    "table": "table editing lands with the v0.1 lifecycle engine "
             "(Batch 2)",
    "chart": "only title/axis text and series ranges are editable on "
             "loaded charts (chartpatch, v0.1 Batch 4); this edit is "
             "outside that set",
    "image": "loaded images are preserved verbatim and not editable; "
             "ADDING images is supported (v0.1 Batch 4)",
    "pivot": "pivot editing is out of scope (preservation and "
             "refresh-on-load cover brownfield pivots)",
}


def _anchor_fingerprint(obj):
    from openpyxl.xml.functions import tostring

    anchor = getattr(obj, "anchor", None)
    if anchor is None:
        return None
    if hasattr(anchor, "to_tree"):
        try:
            return tostring(anchor.to_tree())
        except Exception:
            pass
    return repr(anchor).encode("utf-8")


def _image_data_digest(image):
    # a data swap with identical anchor+path must not vanish silently
    # (Batch-1 gate): fingerprint the backing bytes. NEVER via
    # image._data() — it closes the ref stream (a destructive read that
    # would break the second snapshot and mutate the model at arm).
    import hashlib

    ref = getattr(image, "ref", None)
    try:
        if hasattr(ref, "getvalue"):           # BytesIO: non-destructive
            data = ref.getvalue()
        elif hasattr(ref, "read"):
            pos = ref.tell()
            data = ref.read()
            ref.seek(pos)
        elif isinstance(ref, str):
            with open(ref, "rb") as f:
                data = f.read()
        else:
            return repr(type(ref))
        return hashlib.sha256(data).hexdigest()
    except Exception:
        return None


def diff_objects(ws, armed):
    """(kind, key) pairs whose settled serialization drifted since arm —
    in-session mutations of preserved-part-backed objects."""
    if not armed:
        return []
    current = _object_snapshot(ws)
    unstable = armed.get("unstable", set()) | current.get("unstable", set())
    changed = []
    for kind in ("table", "chart", "image", "pivot"):
        before = armed.get(kind, {})
        after = current.get(kind, {})
        for key in set(before) | set(after):
            if (kind, key) in unstable:
                # a serializer that disagrees with itself cannot express
                # edits: skip the compare (no false refusals on no-ops) —
                # the 0.3 pin discipline applied to objects. No stable
                # real-world instance is known; if one appears, its edits
                # are untrackable and this skip is the documented limit.
                continue
            if before.get(key) != after.get(key):
                changed.append((kind, key))
    return sorted(changed)


def _external_links_snapshot(wb):
    from openpyxl.xml.functions import tostring

    return tuple(tostring(link.to_tree())
                 for link in wb._external_links or [])


def render_core_model(wb):
    from openpyxl.xml.functions import tostring

    return tostring(wb.properties.to_tree())


def render_custom_model(wb):
    from openpyxl.xml.functions import tostring

    props = wb.custom_doc_props
    if not len(props):
        return None
    return tostring(props.to_tree())


def _render_chartsheet(cs):
    from openpyxl.xml.functions import tostring

    return tostring(cs.to_tree())


# ---------------------------------------------------------------------
# hook helpers: every call site bails in two attribute lookups when the
# workbook is not an armed preserve-mode workbook

def _armed_ledger_for_wb(wb):
    led = getattr(wb, "_paper_ledger", None)
    if led is not None and led.armed:
        return led
    return None


def _armed_ledger_for_ws(ws):
    wb = getattr(ws, "parent", None)
    if wb is None:
        return None
    return _armed_ledger_for_wb(wb)


def mark_cell_dirty(cell, formula_involved=False, value_change=False):
    """Called from Cell mutation chokepoints (value bind, style set,
    hyperlink/comment/data_type assignment). ``value_change`` marks the
    coordinate as a VALUE overwrite — the only case where a cell's cm/vm
    rich-value metadata may drop (Batch-3 gate: style-only re-emissions
    and dissolution re-emits must carry it)."""
    ws = cell.parent
    if ws is None or cell.row is None or cell.column is None:
        # standalone cells (write-only compatibility) have no coordinates
        # yet; append() marks them once they are placed
        return
    led = _armed_ledger_for_ws(ws)
    if led is None:
        return
    led.mark_cell(ws, cell.row, cell.column)
    if value_change:
        led.value_overwrites.setdefault(ws, set()).add(
            (cell.row, cell.column))
    if formula_involved:
        led.formulas_changed = True


def check_protection(cell):
    """Protection awareness (PLAN-v0.1 1.6): we report protection, we
    never enforce or bypass it. Called BEFORE the value binds (a strict
    refusal must be atomic): a write to a locked cell of a protected
    sheet warns once per sheet — or refuses under wb.strict_protection.
    Scope: value writes (the chokepoint agents hit); style/comment edits
    to locked cells are not protection-checked in v0.1."""
    ws = cell.parent
    if ws is None:
        return
    led = _armed_ledger_for_ws(ws)
    if led is None:
        return
    try:
        protected = bool(ws.protection.sheet)
    except AttributeError:
        return
    if not protected:
        return
    if not cell.protection.locked:
        return
    wb = ws.parent
    if getattr(wb, "strict_protection", False):
        raise UnsupportedStructureError(
            "cell {0} on sheet {1!r} is locked and the sheet is protected; "
            "this workbook has strict_protection enabled, so the write is "
            "refused. Nothing was changed. Unlock the cell, unprotect the "
            "sheet, or set wb.strict_protection = False to warn "
            "instead.".format(cell.coordinate, ws.title))
    if ws not in led.protection_warned:
        led.protection_warned.add(ws)
        import warnings

        from openpyxl.errors import ProtectedWriteWarning

        warnings.warn(ProtectedWriteWarning(
            "writing to locked cell(s) on protected sheet {0!r} (first: "
            "{1}). The write proceeds — protection is reported, never "
            "enforced — but the sheet's author expected these cells to be "
            "read-only. Set wb.strict_protection = True to refuse such "
            "writes.".format(ws.title, cell.coordinate)), stacklevel=3)


def mark_styleable_dirty(instance):
    """Called from the style descriptors; ``instance`` is a Cell or a
    row/column dimension (both carry ``.parent`` = worksheet)."""
    ws = getattr(instance, "parent", None)
    if ws is None:
        return
    led = _armed_ledger_for_ws(ws)
    if led is None:
        return
    row = getattr(instance, "row", None)
    column = getattr(instance, "column", None)
    if row is not None and column is not None:
        led.mark_cell(ws, row, column)
    # dimensions carry no coordinate; their state serializes via the cols
    # element / row attributes, which the splice syncs from the model


def mark_deleted_cell(ws, row, column, was_formula):
    led = _armed_ledger_for_ws(ws)
    if led is None:
        return
    led.mark_cell(ws, row, column)
    led.value_overwrites.setdefault(ws, set()).add((row, column))
    if was_formula:
        led.formulas_changed = True


def mark_sheet_added(wb, ws):
    led = _armed_ledger_for_wb(wb)
    if led is None:
        return
    led.added_sheets.add(ws)


def allow_sheet_removal(wb, ws):
    """Removing a sheet ADDED in this session is a net no-op and allowed;
    removing a loaded sheet is refused (returns False)."""
    led = _armed_ledger_for_wb(wb)
    if led is None:
        return True
    if ws in led.added_sheets:
        led.added_sheets.discard(ws)
        led.cells.pop(ws, None)
        return True
    return False


def begin_structural_edit(ws, operation, index, amount):
    """Gate for insert/delete rows/cols under preserve (Phase 6b): shifts
    on fully-modeled sheets PROCEED (reference rewriting + byte renumber);
    anything with unmodeled range-bearing content refuses with the precise
    blocker and victim list. Returns True when the caller must run the
    model fixups after mutating."""
    led = _armed_ledger_for_ws(ws)
    if led is None:
        return False
    if ws in led.added_sheets:
        return False
    from .structural import (
        EXCEL_MAX_COL,
        EXCEL_MAX_ROW,
        analyze_shift,
        shift_blockers,
    )

    if operation == "insert_rows":
        # occupancy includes row dimensions and merged/CF anchors, not just
        # cells; and only content AT/AFTER the insert index moves (Batch-1
        # gate: dimension-only floor rows evaded; inserts beyond content
        # false-refused)
        occupied = _max_occupied_row(ws)
        if index <= occupied and occupied + amount > EXCEL_MAX_ROW:
            from openpyxl.errors import BoundaryViolationError

            raise BoundaryViolationError(
                "insert_rows({0}, {1}) would shift occupied content past "
                "row {2}, the sheet's hard limit. Nothing was "
                "changed.".format(index, amount, EXCEL_MAX_ROW))
    if operation == "insert_cols":
        occupied = _max_occupied_col(ws)
        if index <= occupied and occupied + amount > EXCEL_MAX_COL:
            from openpyxl.errors import BoundaryViolationError

            raise BoundaryViolationError(
                "insert_cols({0}, {1}) would shift occupied content past "
                "column {2} (XFD), the sheet's hard limit. Nothing was "
                "changed.".format(index, amount, EXCEL_MAX_COL))

    _check_sheet_protection_for_shift(ws, led, operation)

    blockers = shift_blockers(ws, operation, index, amount)
    if blockers:
        impacts = analyze_shift(ws, operation, index)
        lines = "".join("\n  - " + b for b in blockers)
        victim_lines = "".join("\n  - " + i for i in impacts)
        raise UnsupportedStructureError(
            "{0}() on preserved sheet {1!r} cannot be rewritten safely:"
            "{2}\nWhat the shift would otherwise corrupt:{3}\n"
            "Nothing was changed. Options: restructure the edit to avoid "
            "shifting, or perform it without preserve=True and accept "
            "stock behavior (references are NOT updated).".format(
                operation, ws.title, lines,
                victim_lines or "\n  - (no intersecting references found)")
        )
    return True


def _max_occupied_row(ws):
    candidates = [ws.max_row]
    if ws.row_dimensions:
        candidates.append(max(ws.row_dimensions))
    for rng in ws.merged_cells.ranges:
        candidates.append(rng.max_row)
    return max(candidates)


def _max_occupied_col(ws):
    candidates = [ws.max_column]
    if ws.column_dimensions:
        from openpyxl.utils import column_index_from_string

        candidates.append(max(
            dim.max or column_index_from_string(dim.index)
            for dim in ws.column_dimensions.values()))
    for rng in ws.merged_cells.ranges:
        candidates.append(rng.max_col)
    return max(candidates)


def _check_sheet_protection_for_shift(ws, led, operation):
    """Excel blocks row/column structural edits on protected sheets by
    default: warn (strict: refuse) — same 1.6 discipline as cell writes
    (Batch-1 gate: shifts evaded the protection check entirely)."""
    try:
        protected = bool(ws.protection.sheet)
    except AttributeError:
        return
    if not protected:
        return
    wb = ws.parent
    if getattr(wb, "strict_protection", False):
        raise UnsupportedStructureError(
            "{0}() on protected sheet {1!r}: this workbook has "
            "strict_protection enabled, so the structural edit is "
            "refused. Nothing was changed.".format(operation, ws.title))
    if ws not in led.protection_warned:
        led.protection_warned.add(ws)
        import warnings

        from openpyxl.errors import ProtectedWriteWarning

        warnings.warn(ProtectedWriteWarning(
            "{0}() on protected sheet {1!r}: the edit proceeds — "
            "protection is reported, never enforced — but Excel itself "
            "would block it. Set wb.strict_protection = True to refuse "
            "instead.".format(operation, ws.title)), stacklevel=4)


def finish_structural_edit(ws, operation, index, amount):
    """Model-side reference fixups + snapshot rebasing, after the cells
    moved (see structural.apply_model_shift). Returns the pinned
    AddressRemap (CONVENTIONS §2): pre-edit addresses must be remapped
    through it, never reused."""
    from .structural import AddressRemap, apply_model_shift

    apply_model_shift(ws, operation, index, amount)
    return AddressRemap(ws.title, operation, index, amount)


def refuse_structural_edit(ws, operation, index=None):
    """Row/column shifts under preserve are refused in v0, with the precise
    list of what the shift would strand (PLAN Phase 6a): formulas (cross-
    sheet included), defined names, CF/DV ranges, merges, tables, and
    series ranges inside preserved chart bytes. Raised BEFORE any mutation;
    Phase 6b upgrades refusal to a correct rewrite."""
    led = _armed_ledger_for_ws(ws)
    if led is None:
        return
    if ws in led.added_sheets:
        # sheets created in-session are generated whole at save: the
        # original package holds nothing they could corrupt
        return
    impacts = []
    if index is not None:
        from .structural import analyze_shift
        impacts = analyze_shift(ws, operation, index)
    lines = "".join("\n  - " + line for line in impacts) or (
        "\n  - (no intersecting references found, but row/column shifts "
        "also renumber every cell address below/right of the edit)")
    raise UnsupportedStructureError(
        "{0}() on preserved sheet {1!r} would silently corrupt:{2}\n"
        "Nothing was changed. Options: restructure the edit to avoid "
        "shifting (e.g. write into empty rows), or perform it without "
        "preserve=True and accept stock behavior (references are NOT "
        "updated — the numbers will look plausible and be wrong).".format(
            operation, ws.title, lines)
    )


def refuse_sheet_lifecycle(wb, operation, detail):
    led = _armed_ledger_for_wb(wb)
    if led is None:
        return
    raise UnsupportedStructureError(
        "{0} is not supported in preserve mode: {1} Nothing was changed. "
        "Reopen without preserve=True to accept stock behavior.".format(
            operation, detail)
    )


def refuse_chart_or_image_add(ws, what):
    """Call-time gate for add_chart/add_image under preserve (PLAN-v0.1
    4.2, battery 22): additions are supported on added sheets, on loaded
    sheets without drawing machinery (fresh drawing part + one spliced
    element), and on loaded sheets whose existing drawing is anchor-only
    (anchors appended into the original part). A drawing carrying anything
    else refuses NOW — atomically, before the object joins the model."""
    led = _armed_ledger_for_ws(ws)
    if led is None:
        return
    if ws in led.added_sheets:
        return
    wb = ws.parent
    source = getattr(wb, "_paper_source", None)
    if source is None:
        return
    import io
    import zipfile

    from . import drawings as drawings_mod
    from .saver import _package_info

    with zipfile.ZipFile(io.BytesIO(source)) as zin:
        names = set(zin.namelist())
        _wb_part, mapping = _package_info(zin)
        part = mapping.get(led.renames.get(ws, ws.title))
        if part is None:
            raise UnsupportedStructureError(
                "add_{0}(): the package part for sheet {1!r} could not be "
                "located. Nothing was changed.".format(what, ws.title))
        drawing_part, _rid = drawings_mod._existing_drawing_part(
            zin, names, part)
        if drawing_part is not None \
                and not drawings_mod._anchor_only(zin.read(drawing_part)):
            raise UnsupportedStructureError(
                "add_{0}() on sheet {1!r}: the sheet's existing drawing "
                "carries content other than plain chart/image anchors "
                "(shapes, alternate content, ...); appending into it is "
                "not supported. Nothing was changed.".format(
                    what, ws.title))


def record_rename(sheet_child, new_title):
    """Renaming a LOADED sheet (PLAN-v0.1 3.2, battery 8): the cascade
    rewrite. Model formulas and defined names are rewritten NOW (upstream
    rewrites nothing — the model must stay coherent in-session); chart
    parts referencing the old name are byte-patched at save; the sheets
    entry in workbook.xml gets a name patch keyed by the ORIGINAL title.
    Refuses (atomic, before any mutation): textual references the rewrite
    cannot see (the old title inside formula STRING literals — the
    INDIRECT class), and pivot parts referencing the old name."""
    wb = getattr(sheet_child, "parent", None)
    if wb is None:
        return
    led = _armed_ledger_for_wb(wb)
    if led is None:
        return
    if sheet_child in led.added_sheets:
        return
    old_title = sheet_child.title
    if old_title not in led.loaded_sheet_titles:
        return

    from .rewrite import rename_sheet_in_formula, title_in_string_literals
    from .structural import _pivots_referencing

    # guards first — the refusal must precede every mutation
    textual = []
    for ws in wb.worksheets:
        for (row, col), cell in ws._cells.items():
            if cell.data_type == "f" and isinstance(cell._value, str) \
                    and title_in_string_literals(cell._value, old_title):
                textual.append("{0}!{1}".format(ws.title, cell.coordinate))
    if textual:
        raise UnsupportedStructureError(
            "renaming sheet {0!r} cannot rewrite textual references to it "
            "inside formula strings (INDIRECT-style) at: {1}. Rewrite "
            "those formulas first. Nothing was changed.".format(
                old_title, ", ".join(sorted(textual)[:8])))
    if _pivots_referencing(wb, old_title):
        raise UnsupportedStructureError(
            "renaming sheet {0!r} is not supported while pivot parts "
            "reference it (pivot cacheSource rewriting is out of scope). "
            "Nothing was changed.".format(old_title))

    # model-side cascade: formulas everywhere + defined names. The
    # rewrites are DERIVED from already-accepted formulas and reference
    # the new title before it lands on the sheet object — the lint
    # chokepoint must not judge them (Batch-6 gate: the cascade tripped
    # unknown-sheet, and refuse mode would have refused the rename)
    _saved_lint = getattr(wb, "formula_lint", "warn")
    wb.formula_lint = "off"
    try:
        for ws in wb.worksheets:
            for (row, col), cell in sorted(ws._cells.items()):
                if cell.data_type != "f" \
                        or not isinstance(cell._value, str):
                    continue
                new_formula, changed = rename_sheet_in_formula(
                    cell._value, old_title, new_title)
                if changed:
                    cell.value = new_formula    # public setter: ledgered
    finally:
        wb.formula_lint = _saved_lint
    _rename_defined_names(wb, old_title, new_title)
    for scoped in wb.worksheets:
        _rename_defined_names(scoped, old_title, new_title)
    # in-session charts are model-rendered at save: their data-source
    # references follow the rename like every other model reference
    # (loaded charts' parts are byte-patched at save instead) — Batch-4
    # gate: an added chart's part kept the old, now-nonexistent title
    from .structural import _chart_source_ref_objects

    for sheet in wb.worksheets:
        armed_charts = (led.object_snapshots.get(sheet) or {}).get(
            "chart", {})
        for i, chart in enumerate(getattr(sheet, "_charts", []) or []):
            if i in armed_charts:
                continue
            for ref in _chart_source_ref_objects(chart):
                rewritten, changed = rename_sheet_in_formula(
                    "=" + ref.f, old_title, new_title)
                if changed:
                    ref.f = rewritten[1:]

    # ledger bookkeeping: the sheet keeps counting as LOADED under its
    # new name, state patches re-key, and the save patches the name attr
    # of the <sheet> entry still carrying the ORIGINAL bytes
    original = led.renames.get(sheet_child, old_title)
    led.renames[sheet_child] = original
    led.loaded_sheet_titles = frozenset(
        (led.loaded_sheet_titles - {old_title}) | {new_title})
    if old_title in led.sheet_states:
        led.sheet_states[new_title] = led.sheet_states.pop(old_title)
    led.formulas_changed = True


def _rename_defined_names(holder, old_title, new_title):
    from .rewrite import rename_sheet_in_formula

    for name in list(holder.defined_names):
        dn = holder.defined_names[name]
        if not dn.value:
            continue
        rewritten, changed = rename_sheet_in_formula(
            "=" + dn.value, old_title, new_title)
        if changed:
            dn.value = rewritten[1:]


# ---------------------------------------------------------------------
# Workbook.mark_dirty

def mark_dirty_target(wb, target):
    """Implementation of ``Workbook.mark_dirty(target)`` (PR-0 §4)."""
    led = _armed_ledger_for_wb(wb)
    if led is None:
        raise ValueError(
            "mark_dirty() is only meaningful on a workbook loaded with "
            "preserve=True")
    if not isinstance(target, str) or not target:
        raise TypeError("mark_dirty() takes a sheet-qualified A1 range "
                        "('Model!B7:C9') or a package part name "
                        "('xl/media/image1.png')")
    if "!" in target:
        title, bounds = _parse_sheet_range(target)
        for ws in wb.worksheets:
            if ws.title == title:
                min_col, min_row, max_col, max_row = bounds
                # open-ended (whole-row/column) bounds clamp to the model's
                # populated extent — those are the only splice-able cells
                if min_row is None:
                    min_row, max_row = ws.min_row or 1, ws.max_row or 1
                if min_col is None:
                    min_col, max_col = ws.min_column or 1, ws.max_column or 1
                for row in range(min_row, max_row + 1):
                    for col in range(min_col, max_col + 1):
                        led.mark_cell(ws, row, col)
                return
        raise TargetNotFoundError(
            "mark_dirty: no worksheet named {0!r}".format(title))
    # part-name form
    import io
    import zipfile

    source = getattr(wb, "_paper_source", None)
    names = set()
    if source:
        with zipfile.ZipFile(io.BytesIO(source)) as z:
            names = set(z.namelist())
    if target not in names:
        raise TargetNotFoundError(
            "mark_dirty: no part named {0!r} in the retained package "
            "(part names are exact, e.g. 'xl/media/image1.png')".format(target))
    led.parts.add(target)


def _parse_sheet_range(target):
    """Parse sheet-qualified A1 (pinned addressing), fixing the two upstream
    warts: doubled-quote un-escaping and '$' tolerance."""
    from openpyxl.utils.cell import range_to_tuple

    title, bounds = range_to_tuple(target)
    # upstream keeps escaped quotes in the title group; undo that
    title = title.replace("''", "'")
    return title, bounds


class RemovalReport:
    """What a sheet deletion removed and remapped (PR-1 §2.2, pinned)."""

    def __init__(self, removed_parts, remapped_names):
        self.removed_parts = list(removed_parts)
        self.remapped_names = remapped_names

    def to_dict(self):
        return {"schema": "removal_report", "version": 1,
                "removed_parts": list(self.removed_parts),
                "remapped_names": self.remapped_names}

    def __repr__(self):
        return "RemovalReport({0} parts, {1} names)".format(
            len(self.removed_parts), self.remapped_names)


def _victim_exclusive_parts(wb, original_title):
    """Package parts that die WITH the sheet if it is removed (its
    exclusive relationship closure) — references from these parts back at
    the sheet strand nothing."""
    source = getattr(wb, "_paper_source", None)
    if not source:
        return frozenset()
    import io
    import zipfile

    from .saver import _exclusive_closure, _package_info

    with zipfile.ZipFile(io.BytesIO(source)) as zin:
        names = set(zin.namelist())
        _wb_part, mapping = _package_info(zin)
        part = mapping.get(original_title)
        if part is None:
            return frozenset()
        return frozenset(_exclusive_closure(zin, names, part))


def audit_sheet_removal(wb, ws):
    """The reference audit before a LOADED sheet may be removed
    (PLAN-v0.1 3.2): anything on ANOTHER sheet pointing at the victim
    refuses with the full enumeration — formulas (3-D endpoints and
    textual/INDIRECT included), defined names, chart parts, pivot parts."""
    from .rewrite import (
        rename_sheet_in_formula,
        title_in_string_literals,
    )
    from .structural import _charts_referencing, _pivots_referencing

    title = ws.title
    led = _armed_ledger_for_wb(wb)

    def _refs_victim(formula):
        probe = formula if formula.startswith("=") else "=" + formula
        _, refs_it = rename_sheet_in_formula(probe, title, title + "_")
        return refs_it

    victims = []
    for other in wb.worksheets:
        if other is ws:
            continue
        for (row, col), cell in sorted(other._cells.items()):
            if cell.data_type != "f" or not isinstance(cell._value, str):
                continue
            # a reference the rename machinery could rewrite is exactly a
            # reference the delete would strand
            if (_refs_victim(cell._value)
                    or title_in_string_literals(cell._value, title)):
                victims.append("{0}!{1}".format(other.title,
                                                cell.coordinate))
        # sheet-scoped names, CF rule formulas and DV formulas on a
        # SURVIVING sheet all strand exactly like cell formulas do
        # (Batch-3 gate: the audit only walked workbook-level names)
        for name in list(other.defined_names):
            dn = other.defined_names[name]
            if dn.value and _refs_victim(dn.value):
                victims.append("defined name {0!r} (scoped to sheet "
                               "{1!r})".format(name, other.title))
        for cf in other.conditional_formatting:
            for rule in cf.rules:
                for f in (rule.formula or []):
                    if isinstance(f, str) and _refs_victim(f):
                        victims.append(
                            "conditional-formatting rule on {0}!{1}".format(
                                other.title, cf.sqref))
                        break
        for dv in other.data_validations.dataValidation:
            for f in (dv.formula1, dv.formula2):
                if isinstance(f, str) and f and _refs_victim(f):
                    victims.append(
                        "data validation on {0}!{1}".format(
                            other.title, dv.sqref))
                    break
    for name in list(wb.defined_names):
        dn = wb.defined_names[name]
        if dn.value and _refs_victim(dn.value):
            victims.append("defined name {0!r}".format(name))
    # byte-level searches run against the ORIGINAL package, so they must
    # use the sheet's original title when it was renamed this session
    original_title = led.renames.get(ws, ws.title) if led else ws.title
    own_parts = _victim_exclusive_parts(wb, original_title)
    for part in _charts_referencing(wb, original_title):
        # a chart anchored ON the victim dies with it in the exclusive
        # closure — its self-references strand nothing (Batch-3 gate:
        # own-exclusive-chart false refusal)
        if part not in own_parts:
            victims.append("chart part {0}".format(part))
    if _pivots_referencing(wb, original_title):
        victims.append("pivot parts")
    if victims:
        raise UnsupportedStructureError(
            "removing sheet {0!r} would strand references to it:\n  - "
            "{1}\nNothing was changed. Delete or rewrite those references "
            "first.".format(title, "\n  - ".join(victims[:12])))


def record_sheet_removal(wb, ws):
    """Called by Workbook.remove for a LOADED sheet, AFTER the audit."""
    led = _armed_ledger_for_wb(wb)
    if led is None:
        return
    original_title = led.renames.pop(ws, ws.title)
    led.removed_sheets.append(original_title)
    led.loaded_sheet_titles = led.loaded_sheet_titles - {ws.title}
    led.sheet_states.pop(ws.title, None)
    led.cells.pop(ws, None)
    led.object_snapshots.pop(ws, None)
    led.region_snapshots.pop(ws, None)
    led.row_attr_snapshots.pop(ws, None)
    led.comment_snapshots.pop(ws, None)
    led.pinned_regions.pop(ws, None)
    # calcChain carries positional sheet indexes: it dies with the sheet
    led.formulas_changed = True


def begin_move_range(ws, move_spec):
    """move_range under preserve (PLAN-v0.1 3.3): expressed as tracked
    cell edits (source cleared + destination written — no rows shift, so
    no byte renumber). Guards refuse what the move cannot keep coherent:
    merges/CF/DV/tables intersecting either rectangle, and formulas
    OUTSIDE the moved block referencing the source (Excel's cut-paste
    would follow them; we do not rewrite them in this wave)."""
    led = _armed_ledger_for_ws(ws)
    if led is None or ws in led.added_sheets:
        return
    from openpyxl.utils.cell import range_boundaries

    from .structural import _intersects
    from .perception import dependency_sketch

    cell_range, rows, cols, _translate = move_spec
    if isinstance(cell_range, str):
        src = range_boundaries(cell_range)
    else:
        src = (cell_range.min_col, cell_range.min_row,
               cell_range.max_col, cell_range.max_row)
    dst = (src[0] + cols, src[1] + rows, src[2] + cols, src[3] + rows)
    if dst[0] < 1 or dst[1] < 1:
        raise UnsupportedStructureError(
            "move_range would move cells before column A / row 1. "
            "Nothing was changed.")
    from .structural import EXCEL_MAX_COL, EXCEL_MAX_ROW

    if dst[2] > EXCEL_MAX_COL or dst[3] > EXCEL_MAX_ROW:
        from openpyxl.errors import BoundaryViolationError

        raise BoundaryViolationError(
            "move_range would move cells past the sheet limits "
            "(XFD/1048576). Nothing was changed.")
    _check_sheet_protection_for_shift(ws, led, "move_range")

    problems = []
    for rng in ws.merged_cells.ranges:
        b = (rng.min_col, rng.min_row, rng.max_col, rng.max_row)
        if _intersects(b, *src) or _intersects(b, *dst):
            problems.append("merged range {0}".format(rng))
    for cf in ws.conditional_formatting:
        for rng in getattr(cf.sqref, "ranges", []):
            b = (rng.min_col, rng.min_row, rng.max_col, rng.max_row)
            if _intersects(b, *src) or _intersects(b, *dst):
                problems.append("conditional formatting {0}".format(rng))
    if ws.data_validations:
        for dv in ws.data_validations.dataValidation:
            for rng in getattr(dv.sqref, "ranges", []):
                b = (rng.min_col, rng.min_row, rng.max_col, rng.max_row)
                if _intersects(b, *src) or _intersects(b, *dst):
                    problems.append("data validation {0}".format(rng))
    for name, ref in getattr(ws, "tables", {}).items():
        from .structural import _ref_hit

        if _ref_hit(ref, src) or _ref_hit(ref, dst):
            problems.append("table {0!r}".format(name))

    from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

    from .structural import _charts_referencing

    original_title = led.renames.get(ws, ws.title)
    if _charts_referencing(ws.parent, original_title):
        problems.append("preserved chart(s) reference this sheet; their "
                        "series ranges cannot follow a rectangular move")
    for r in range(src[1], src[3] + 1):
        for c in range(src[0], src[2] + 1):
            cell = ws._cells.get((r, c))
            if cell is not None and isinstance(
                    cell._value, (ArrayFormula, DataTableFormula)):
                problems.append(
                    "array/data-table formula at {0} inside the moved "
                    "block".format(cell.coordinate))
    # defined names pointing INTO either rectangle (Excel cut-paste
    # follows them; we do not rewrite names on moves — Batch-3 gate)
    from openpyxl.utils.cell import range_boundaries as _rb

    holders = [ws.parent.defined_names] + [
        w.defined_names for w in ws.parent.worksheets]
    for names_holder in holders:
        for nm in list(names_holder):
            dn = names_holder[nm]
            try:
                for dest_sheet, dest_ref in dn.destinations:
                    if dest_sheet is None \
                            or dest_sheet.casefold() != ws.title.casefold():
                        continue
                    b = _rb(dest_ref.replace("$", ""))
                    if _intersects(b, *src) or _intersects(b, *dst):
                        problems.append(
                            "defined name {0!r} points into the moved/"
                            "target block".format(nm))
            except Exception:
                continue

    sketch = dependency_sketch(ws.parent)
    inside = set()
    for r in range(src[1], src[3] + 1):
        for c in range(src[0], src[2] + 1):
            inside.add((r, c))
    from openpyxl.utils.cell import coordinate_to_tuple

    for address in (set(sketch.cells_referencing(ws.title, src))
                    | set(sketch.cells_referencing(ws.title, dst))):
        title, _, coord = address.rpartition("!")
        bare = title.strip("'").replace("''", "'")
        try:
            rc = coordinate_to_tuple(coord)
        except Exception:
            problems.append("formula {0}".format(address))
            continue
        if bare.casefold() != ws.title.casefold() or rc not in inside:
            problems.append("formula {0} references the moved or target "
                            "block".format(address))
    if problems:
        raise UnsupportedStructureError(
            "move_range on sheet {0!r} cannot be kept coherent:\n  - {1}\n"
            "Nothing was changed. Restructure the edit or move the "
            "referencing content first.".format(
                ws.title, "\n  - ".join(sorted(set(problems))[:10])))

    # the move is plain cell edits: mark every source AND destination
    # coordinate dirty (upstream mutates the model right after this)
    for r in range(src[1], src[3] + 1):
        for c in range(src[0], src[2] + 1):
            led.mark_cell(ws, r, c)
    for r in range(dst[1], dst[3] + 1):
        for c in range(dst[0], dst[2] + 1):
            led.mark_cell(ws, r, c)
            led.value_overwrites.setdefault(ws, set()).add((r, c))
    for r in range(src[1], src[3] + 1):
        for c in range(src[0], src[2] + 1):
            led.value_overwrites.setdefault(ws, set()).add((r, c))
    led.formulas_changed = True          # moved formulas re-anchor
