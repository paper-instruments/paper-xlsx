# paper-xlsx: the splice writer (CONVENTIONS §3.4; PR-0 D6/D7/D15)

"""Apply ledger-recorded edits to one original worksheet part, byte-wise.

The original bytes are the source of truth: untouched ranges are copied
verbatim; dirty cells are replaced/inserted/deleted at their scanned spans
(two sorted streams — rows and cells are coordinate-ordered); user-changed
satellite regions are replaced whole from the model. Everything unmodeled
(extLst, mc:AlternateContent, drawing references, foreign attributes) passes
through because it is never interpreted.
"""

from openpyxl.errors import UnsupportedStructureError
from openpyxl.utils.cell import range_boundaries

from . import emit
from .regions import CT_ORDER_INDEX, REGION_BY_TAG, DETECT_ONLY_REGIONS
from .xmlscan import scan_sheet


class SpliceRefusal(UnsupportedStructureError):
    pass


def _cells_in_ref(ref):
    min_col, min_row, max_col, max_row = range_boundaries(ref)
    return {(r, c) for r in range(min_row, max_row + 1)
            for c in range(min_col, max_col + 1)}


def resolve_dirty_cells(ws, ledger_dirty, scan):
    """The effective dirty-coordinate set for one sheet (PR-0 D7).

    - rich-text cells are always dirty (in-place edits bypass every hook);
    - a dirty cell intersecting a shared-formula group dissolves the WHOLE
      group: every member re-emits as a plain formula from the model (the
      model already holds the expanded formulas);
    - a dirty cell intersecting an array formula refuses;
    - a dirty cell carrying cm/vm metadata or unexpected children refuses.
    """
    from openpyxl.cell.rich_text import CellRichText

    dirty = set(ledger_dirty)
    for (row, col), cell in ws._cells.items():
        if isinstance(cell._value, CellRichText):
            dirty.add((row, col))

    if not dirty:
        return dirty

    # array formulas: refuse on intersection
    for ref in scan.array_refs:
        hit = dirty & _cells_in_ref(ref)
        if hit:
            raise SpliceRefusal(
                "cannot edit cell(s) {0} on sheet {1!r}: they intersect the "
                "array formula range {2}. Editing array formulas is not "
                "supported in v0; nothing was written.".format(
                    sorted(hit), ws.title, ref))

    # shared-formula groups: dissolve on touch
    for si, ref in scan.shared_groups.items():
        members = _cells_in_ref(ref) | scan.shared_members.get(si, set())
        if dirty & members:
            dirty |= members
    # a follower whose host/ref we never saw cannot be dissolved safely
    for si, members in scan.shared_members.items():
        if si not in scan.shared_groups and (dirty & members):
            raise SpliceRefusal(
                "cannot edit shared-formula follower cell(s) {0} on sheet "
                "{1!r}: the group host (si={2}) was not found in the sheet. "
                "Nothing was written.".format(
                    sorted(dirty & members), ws.title, si))

    # per-cell guards on the originals being replaced
    for (row, col) in sorted(dirty):
        row_span = scan.rows.get(row)
        cell_span = row_span.cells.get(col) if row_span else None
        if cell_span is None:
            continue
        if "cm" in cell_span.attrs or "vm" in cell_span.attrs:
            raise SpliceRefusal(
                "cannot edit cell {0}{1} on sheet {2!r}: it carries cell "
                "metadata (cm/vm — dynamic-array or rich-value metadata) "
                "that would go stale on the new value. Nothing was "
                "written.".format(_col_letter(col), row, ws.title))
        if cell_span.has_extlst:
            raise SpliceRefusal(
                "cannot edit cell {0}{1} on sheet {2!r}: it carries a "
                "cell-level extLst the replacement cannot preserve. "
                "Nothing was written.".format(_col_letter(col), row, ws.title))
    return dirty


def _col_letter(col):
    letters = ""
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def splice_sheet(ws, original, dirty_cells, region_changes, row_attr_changes,
                 scan=None):
    """Return the new part payload for one worksheet.

    ``dirty_cells``: resolved coordinate set (see resolve_dirty_cells).
    ``region_changes``: {tag: serialized bytes or None} — user-changed
    satellite regions (None = region now absent).
    ``row_attr_changes``: {row_index: {attr: value}} — changed row display
    attributes.
    ``scan``: a SheetScan of ``original`` if the caller already has one.
    """
    if scan is None:
        scan = scan_sheet(original)

    for tag in region_changes:
        if tag in DETECT_ONLY_REGIONS:
            raise SpliceRefusal(
                "changes to {0} on sheet {1!r} are not writable at this "
                "build stage (cross-part coordination lands in Phase 2d); "
                "nothing was written.".format(tag, ws.title))
        if tag not in REGION_BY_TAG:
            raise SpliceRefusal(
                "internal: unexpected region change {0!r}".format(tag))
        spans = scan.regions.get(tag, [])
        for span in spans:
            if b"extLst" in original[span.start:span.end]:
                raise SpliceRefusal(
                    "cannot rewrite the {0} element on sheet {1!r}: the "
                    "original carries an extLst extension the model cannot "
                    "re-serialize. Nothing was written.".format(tag, ws.title))
        if len(spans) > 1:
            raise SpliceRefusal(
                "cannot rewrite {0} on sheet {1!r}: multiple original "
                "elements. Nothing was written.".format(tag, ws.title))

    # x14 twin gates (PR-0 D15): a DV change with an x14 dataValidations
    # block in the sheet extLst desyncs the twins
    if "dataValidations" in region_changes:
        ext_spans = scan.regions.get("extLst", [])
        for span in ext_spans:
            if b"dataValidations" in original[span.start:span.end]:
                raise SpliceRefusal(
                    "cannot change data validations on sheet {0!r}: the "
                    "sheet carries x14 data validations in its extLst; "
                    "editing the classic element alone would desync them. "
                    "Nothing was written.".format(ws.title))

    if dirty_cells:
        new_rows = {r for (r, c) in dirty_cells} - set(scan.rows)
        if new_rows and not scan.rows_monotonic:
            raise SpliceRefusal(
                "cannot insert rows into sheet {0!r}: its rows are not in "
                "ascending order. Nothing was written.".format(ws.title))

    # ------- assemble the edit list: (start, end, replacement) -----------
    edits = []

    # 1. region replacements / removals / insertions
    for tag, rendered in region_changes.items():
        spans = scan.regions.get(tag, [])
        if spans:
            span = spans[0]
            edits.append((span.start, span.end, rendered or b""))
        elif rendered:
            insert_at = _region_insert_offset(scan, tag)
            edits.append((insert_at, insert_at, rendered))

    # 2. cell and row edits inside sheetData
    edits.extend(_sheetdata_edits(ws, scan, dirty_cells, row_attr_changes))

    # ------- apply (sorted, non-overlapping by construction) -------------
    edits.sort(key=lambda e: (e[0], e[1]))
    for (s1, e1), (s2, e2) in zip(
            [(s, e) for s, e, _ in edits], [(s, e) for s, e, _ in edits][1:]):
        if e1 > s2:
            raise SpliceRefusal(
                "internal: overlapping splice edits ({0},{1}) and "
                "({2},{3})".format(s1, e1, s2, e2))

    out = []
    pos = 0
    for start, end, replacement in edits:
        out.append(original[pos:start])
        out.append(replacement)
        pos = end
    out.append(original[pos:])
    return b"".join(out)


def _region_insert_offset(scan, tag):
    """Where a region element absent from the original must be inserted,
    per the CT_Worksheet child sequence."""
    order = CT_ORDER_INDEX[tag]
    for other_tag, span in scan.region_order:
        if CT_ORDER_INDEX.get(other_tag, len(CT_ORDER_INDEX)) > order:
            return span.start
    if scan.root_end_offset is None:
        raise SpliceRefusal("internal: no insertion point found")
    return scan.root_end_offset


def _sheetdata_edits(ws, scan, dirty_cells, row_attr_changes):
    edits = []
    by_row = {}
    for (row, col) in dirty_cells:
        by_row.setdefault(row, set()).add(col)

    touched_rows = set(by_row) | set(row_attr_changes)
    if not touched_rows:
        return edits

    if scan.sheetdata_content is None and scan.sheetdata is not None:
        # self-closing <sheetData/>: expand it if we need content
        needs_content = any(r not in scan.rows for r in touched_rows)
        if needs_content:
            span = scan.sheetdata
            new_rows_bytes = _emit_rows_block(
                ws, sorted(r for r in touched_rows), by_row, row_attr_changes)
            edits.append((span.start, span.end,
                          b"<sheetData>" + new_rows_bytes + b"</sheetData>"))
            return edits

    pending_new_rows = sorted(r for r in touched_rows if r not in scan.rows)

    # existing rows: per-row cell edits and attribute sync
    for row_index in sorted(r for r in touched_rows if r in scan.rows):
        row_span = scan.rows[row_index]
        row_edits = _row_edits(ws, row_span, by_row.get(row_index, set()),
                               row_attr_changes.get(row_index))
        edits.extend(row_edits)

    # new rows: insert each before the first existing row with a larger index
    ordered_existing = [r for r in scan.row_order]
    for new_index in pending_new_rows:
        insert_before = None
        for existing in ordered_existing:
            if existing > new_index:
                insert_before = scan.rows[existing]
                break
        cells_bytes = _emit_row_cells(ws, new_index,
                                      sorted(by_row.get(new_index, set())))
        attrs = row_attr_changes.get(new_index, {})
        row_bytes = emit.emit_new_row(ws, new_index, cells_bytes, attrs)
        if not cells_bytes and not attrs:
            continue
        if insert_before is not None:
            offset = insert_before.start
        else:
            # after the last existing row, before </sheetData>
            if scan.sheetdata_content is not None:
                offset = scan.sheetdata_content[1]
            else:
                raise SpliceRefusal(
                    "internal: nowhere to insert row {0}".format(new_index))
        edits.append((offset, offset, row_bytes))
    return edits


def _emit_row_cells(ws, row_index, columns):
    cells_bytes = []
    for col in columns:
        cell = ws._cells.get((row_index, col))
        if cell is None:
            continue
        rendered = emit.emit_cell(ws, cell)
        if rendered is not None:
            cells_bytes.append(rendered)
    return cells_bytes


def _emit_rows_block(ws, row_indices, by_row, row_attr_changes):
    parts = []
    for row_index in row_indices:
        cells_bytes = _emit_row_cells(ws, row_index,
                                      sorted(by_row.get(row_index, set())))
        attrs = row_attr_changes.get(row_index, {})
        if cells_bytes or attrs:
            parts.append(emit.emit_new_row(ws, row_index, cells_bytes, attrs))
    return b"".join(parts)


def _row_edits(ws, row_span, dirty_cols, new_attrs):
    """Edits inside one existing row: replace/insert/delete cells, sync the
    row start tag's attributes when they changed."""
    edits = []

    if new_attrs is not None:
        # rewrite the start tag, preserving the row's content
        attrs = dict(new_attrs)
        # keep original spans attribute verbatim (tolerated-stale, PR-0 D6)
        if "spans" in row_span.attrs:
            attrs.setdefault("spans", row_span.attrs["spans"])
        if row_span.self_closing:
            edits.append((row_span.start, row_span.end,
                          emit.row_start_tag(row_span.index, attrs,
                                             self_closing=True)))
            if dirty_cols:
                raise SpliceRefusal(
                    "internal: dirty cells in a self-closing row")
            return edits
        edits.append((row_span.start, row_span.content_start,
                      emit.row_start_tag(row_span.index, attrs)))

    if not dirty_cols:
        return edits

    if row_span.self_closing:
        # row exists but holds no cells: rebuild it whole
        cells_bytes = _emit_row_cells(ws, row_span.index, sorted(dirty_cols))
        attrs = dict(new_attrs) if new_attrs is not None else \
            {k: v for k, v in row_span.attrs.items() if k != "r"}
        edits = [(row_span.start, row_span.end,
                  emit.emit_new_row(ws, row_span.index, cells_bytes, attrs))]
        return edits

    existing_cols = sorted(row_span.cells)
    if existing_cols != sorted(set(existing_cols)):
        raise SpliceRefusal("internal: duplicate cells in row")
    monotonic = all(a < b for a, b in zip(existing_cols, existing_cols[1:]))

    inserts = sorted(c for c in dirty_cols if c not in row_span.cells)
    if inserts and not monotonic:
        raise SpliceRefusal(
            "cannot insert cells into row {0}: its cells are not in "
            "ascending column order. Nothing was written.".format(
                row_span.index))

    for col in sorted(dirty_cols):
        cell_span = row_span.cells.get(col)
        cell = ws._cells.get((row_span.index, col))
        rendered = emit.emit_cell(ws, cell) if cell is not None else None
        if rendered is not None and cell_span is not None:
            rendered = emit.carry_attributes(rendered, cell_span.attrs)

        if cell_span is not None:
            # replace or delete an existing cell element
            edits.append((cell_span.start, cell_span.end, rendered or b""))
        elif rendered is not None:
            # insert a new cell at its sorted position
            after = [c for c in existing_cols if c > col]
            if after:
                offset = row_span.cells[after[0]].start
            else:
                offset = row_span.content_end
            edits.append((offset, offset, rendered))
    return edits
