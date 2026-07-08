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
                 "orig_cell_styles_len", "rich_text_mode",
                 "sheet_states", "dxfs_len", "named_styles_len", "shifts",
                 "template_flag")

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
        self.orig_cell_styles_len = 0
        self.rich_text_mode = False
        self.sheet_states = {}         # title -> state at arm (all sheets)
        self.dxfs_len = 0
        self.named_styles_len = 0
        self.shifts = {}               # ws -> [(operation, index, amount)]
        self.template_flag = False

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
        for cs in wb.chartsheets:
            led.chartsheet_snapshots[cs] = _render_chartsheet(cs)
        from .crosspart import render_workbook_elements

        led.workbook_snapshot = render_workbook_elements(wb)
        led.core_snapshot = render_core_model(wb)
        led.custom_snapshot = render_custom_model(wb)
        led.orig_cell_styles_len = len(wb._cell_styles)
        led.rich_text_mode = rich_text
        led.sheet_states = {s.title: s.sheet_state for s in wb._sheets}
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
            snap[(row, col)] = (comment.text, comment.author)
    return snap


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


def mark_cell_dirty(cell, formula_involved=False):
    """Called from Cell mutation chokepoints (value bind, style set,
    hyperlink/comment/data_type assignment)."""
    ws = cell.parent
    if ws is None or cell.row is None or cell.column is None:
        # standalone cells (write-only compatibility) have no coordinates
        # yet; append() marks them once they are placed
        return
    led = _armed_ledger_for_ws(ws)
    if led is None:
        return
    led.mark_cell(ws, cell.row, cell.column)
    if formula_involved:
        led.formulas_changed = True


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
    from .structural import analyze_shift, shift_blockers

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


def finish_structural_edit(ws, operation, index, amount):
    """Model-side reference fixups + snapshot rebasing, after the cells
    moved (see structural.apply_model_shift)."""
    from .structural import apply_model_shift

    apply_model_shift(ws, operation, index, amount)


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
    """Adding charts/images under preserve refuses in v0 (PR-0 D9 as
    amended in Phase 2d): drawing-part generation against the preserved
    package is out of scope, and silently dropping the object at save is
    the forbidden outcome."""
    led = _armed_ledger_for_ws(ws)
    if led is None:
        return
    raise UnsupportedStructureError(
        "add_{0}() is not supported in preserve mode (v0): generating "
        "drawing parts alongside the preserved package is not implemented, "
        "and the {0} would otherwise be silently absent from the saved "
        "file. Build charts in a separate stock-mode workbook. Nothing was "
        "changed.".format(what)
    )


def refuse_rename(sheet_child):
    """Renaming a LOADED sheet under preserve is refused: every formula,
    defined name and chart series referencing the old name — including
    inside preserved-bytes parts — would silently dangle."""
    wb = getattr(sheet_child, "parent", None)
    if wb is None:
        return
    led = _armed_ledger_for_wb(wb)
    if led is None:
        return
    if sheet_child in led.added_sheets:
        return
    if sheet_child.title in led.loaded_sheet_titles:
        raise UnsupportedStructureError(
            "renaming sheet {0!r} is not supported in preserve mode: "
            "formulas, defined names and chart references to the old name "
            "(including inside preserved charts) would silently break. "
            "Nothing was changed.".format(sheet_child.title)
        )


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
