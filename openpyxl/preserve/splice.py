# paper-xlsx: the splice writer

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

from .regions import (CT_ORDER_INDEX, REGION_BY_TAG, DETECT_ONLY_REGIONS,
                      SAVER_CRAFTED_REGIONS)
from .xmlscan import scan_sheet


class SpliceRefusal(UnsupportedStructureError):
    pass


class _ChildSpan:
    __slots__ = ("name", "start", "end")

    def __init__(self, name, start, end):
        self.name = name
        self.start = start
        self.end = end


def _markup_name(data, start, end):
    index = start + 2 if data[start:start + 2] == b"</" else start + 1
    name_start = index
    while index < end and data[index] not in b" \t\r\n/>":
        index += 1
    if index == name_start:
        raise ValueError("XML markup has no element name")
    return data[name_start:index]


def _direct_cell_children(cell_bytes):
    """Return exact spans for direct children of an unprefixed cell."""
    tag_end, self_closing, _attrs = emit._start_tag_attributes(cell_bytes)
    if self_closing:
        return ()
    close_start = cell_bytes.rfind(b"</c>")
    if close_start < tag_end or cell_bytes[close_start:] != b"</c>":
        raise ValueError("cell fragment has no exact closing tag")
    children = []
    stack = []
    direct_start = direct_name = None
    position = tag_end + 1
    while position < close_start:
        start = cell_bytes.find(b"<", position, close_start)
        if start < 0:
            break
        if cell_bytes.startswith(b"<!--", start):
            end = cell_bytes.find(b"-->", start + 4, close_start)
            if end < 0:
                raise ValueError("unterminated comment")
            position = end + 3
            continue
        if cell_bytes.startswith(b"<?", start):
            end = cell_bytes.find(b"?>", start + 2, close_start)
            if end < 0:
                raise ValueError("unterminated processing instruction")
            position = end + 2
            continue
        if cell_bytes.startswith(b"<![CDATA[", start):
            end = cell_bytes.find(b"]]>", start + 9, close_start)
            if end < 0:
                raise ValueError("unterminated CDATA section")
            position = end + 3
            continue
        if cell_bytes.startswith(b"<!", start):
            raise ValueError("unsupported declaration in cell content")
        if cell_bytes.startswith(b"</", start):
            end = cell_bytes.find(b">", start + 2, close_start)
            if end < 0 or not stack:
                raise ValueError("malformed closing tag")
            if stack.pop() != _markup_name(cell_bytes, start, end):
                raise ValueError("mismatched child markup")
            if not stack:
                children.append(_ChildSpan(
                    direct_name, direct_start, end + 1))
                direct_start = direct_name = None
            position = end + 1
            continue
        relative_end = emit._start_tag_end(cell_bytes[start:close_start])
        end = start + relative_end
        name = _markup_name(cell_bytes, start, end)
        self_closing_child = cell_bytes[end - 1:end] == b"/"
        if not stack:
            direct_start, direct_name = start, name
        if self_closing_child:
            if not stack:
                children.append(_ChildSpan(name, start, end + 1))
                direct_start = direct_name = None
        else:
            stack.append(name)
        position = end + 1
    if stack:
        raise ValueError("unclosed child markup")
    return tuple(children)


def _direct_child(cell_bytes, name):
    matches = [child for child in _direct_cell_children(cell_bytes)
               if child.name == name]
    if len(matches) > 1:
        raise ValueError("duplicate child {0!r}".format(name))
    return matches[0] if matches else None


# row attributes the model owns (dict(RowDimension) keys + spans/r); anything
# else found on an original row is unmodeled and carried verbatim
_MODEL_ROW_ATTRS = frozenset((
    "ht", "customFormat", "customHeight", "s", "hidden", "outlineLevel",
    "collapsed", "thickTop", "thickBot", "spans",
))


def _range_bounds(ref):
    min_col, min_row, max_col, max_row = range_boundaries(ref)
    return min_row, min_col, max_row, max_col


def _coordinates_in_bounds(coordinates, bounds):
    min_row, min_col, max_row, max_col = bounds
    return {
        (row, col) for row, col in coordinates
        if min_row <= row <= max_row and min_col <= col <= max_col
    }


def _coordinate_in_bounds(coordinate, bounds):
    row, col = coordinate
    min_row, min_col, max_row, max_col = bounds
    return min_row <= row <= max_row and min_col <= col <= max_col


def resolve_dirty_cells(ws, ledger_dirty, scan, value_overwrites=frozenset()):
    """The effective dirty-coordinate set for one sheet.

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

    # array/spill formulas: refuse on intersection, naming the anchor
    # (the in_spill context — members of a dynamic-array
    # spill are blank cells in the file; the range lives on the anchor)
    for ref in scan.array_refs:
        hit = _coordinates_in_bounds(dirty, _range_bounds(ref))
        if hit:
            anchor = ref.split(":")[0]
            raise SpliceRefusal(
                "cannot edit cell(s) {0} on sheet {1!r}: they are inside "
                "the array/spill range {2} (in_spill; anchored at {3}). "
                "Array editing is not supported under preserve mode — "
                "reopen without preserve=True to restructure the array. "
                "Nothing was written.".format(
                    sorted(hit), ws.title, ref, anchor))

    # shared-formula groups: dissolve on touch. Membership comes from the
    # OBSERVED member cells (every member carries si); the host's ref
    # attribute can be stale after a shift and is used only as a widener
    # when it still parses
    for si, ref in scan.shared_groups.items():
        members = set(scan.shared_members.get(si, set()))
        try:
            bounds = _range_bounds(ref)
            members |= _coordinates_in_bounds(ws._cells, bounds)
        except Exception:
            pass
        if dirty & members:
            dirty |= scan.shared_members.get(si, set()) | (dirty & members)
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
        # cm/vm cell metadata (dynamic-array / rich-value): a plain
        # value overwrite is CORRECT with the attributes dropped — the
        # cell simply stops being a rich value/spill anchor; the metadata
        # part keeps unreferenced records (legal dead weight). The emit
        # carry excludes cm/vm.
        if cell_span.has_extlst:
            raise SpliceRefusal(
                "cannot edit cell {0}{1} on sheet {2!r}: it carries a "
                "cell-level extLst the replacement cannot preserve. "
                "Nothing was written.".format(_col_letter(col), row, ws.title))
        if cell_span.has_unowned_children \
                and ((row, col) in value_overwrites
                     or cell_span.shared_si is not None):
            raise SpliceRefusal(
                "cannot rewrite cell {0}{1} on sheet {2!r}: it carries "
                "unowned direct-child XML that cannot be preserved with "
                "this value/formula edit. Nothing was written.".format(
                    _col_letter(col), row, ws.title))
    return dirty


def _col_letter(col):
    letters = ""
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def splice_sheet(ws, original, dirty_cells, region_changes, row_attr_changes,
                 scan=None, cf_replacement=None, hyperlinks_replacement=None,
                 style_resolver=None,
                 value_overwrites=frozenset(), cache_writes=None,
                 cache_invalidations=None):
    """Return the new part payload for one worksheet.

    ``dirty_cells``: resolved coordinate set (see resolve_dirty_cells).
    ``region_changes``: {tag: serialized bytes or None} — user-changed
    satellite regions (None = region now absent).
    ``row_attr_changes``: {row_index: {attr: value}} — changed row display
    attributes.
    ``scan``: a SheetScan of ``original`` if the caller already has one.
    ``cf_replacement``: bytes replacing ALL conditionalFormatting elements
    (gated by the caller; may be b"" to remove them).
    ``hyperlinks_replacement``: bytes replacing the hyperlinks element.
    ``style_resolver``: cell -> FILE xf index (StyleTranslator) —
    model style indices must never reach the spliced bytes.
    ``cache_writes``: {(row, col): computed_value} — cached-value updates
    for UNTOUCHED formula cells (oracle write-back): the
    <f> bytes stay verbatim, only the cached <v> (and its t attribute)
    change.
    ``cache_invalidations``: {(row, col)} — formula cells and array followers
    whose cached <v> must be removed because results are uncertified.
    """
    if scan is None:
        scan = scan_sheet(original)
    if style_resolver is None:
        style_resolver = lambda cell: None  # noqa: E731 — styleless contexts only
    cache_invalidations = set(cache_invalidations or ())

    for tag in region_changes:
        if tag in DETECT_ONLY_REGIONS:
            raise SpliceRefusal(
                "changes to {0} on sheet {1!r} are not writable in v0 "
                "(table-part lifecycle); nothing was "
                "written.".format(tag, ws.title))
        if tag not in REGION_BY_TAG and tag not in SAVER_CRAFTED_REGIONS:
            raise SpliceRefusal(
                "internal: unexpected region change {0!r}".format(tag))
        spans = scan.regions.get(tag, [])
        for span in spans:
            if tag in SAVER_CRAFTED_REGIONS:
                # crafted bytes are composed FROM the original (extensions
                # preserved by construction) — the model-render guard
                # below does not apply
                continue
            if b"extLst" in original[span.start:span.end]:
                raise SpliceRefusal(
                    "cannot rewrite the {0} element on sheet {1!r}: the "
                    "original carries an extLst extension the model cannot "
                    "re-serialize. Nothing was written.".format(tag, ws.title))
        if len(spans) > 1:
            raise SpliceRefusal(
                "cannot rewrite {0} on sheet {1!r}: multiple original "
                "elements. Nothing was written.".format(tag, ws.title))

    # x14 twin coexistence is checked by the saver (preserve.x14);
    # the blanket gates moved there with the composed-CF machinery

    if dirty_cells:
        new_rows = {r for (r, c) in dirty_cells} - set(scan.rows)
        if new_rows and not scan.rows_monotonic:
            raise SpliceRefusal(
                "cannot insert rows into sheet {0!r}: its rows are not in "
                "ascending order. Nothing was written.".format(ws.title))

    # ------- assemble the edit list: (start, end, replacement, rank) -----
    # rank breaks ties between insertions landing at the same offset: they
    # must come out in CT_Worksheet child order (measured: DV added together
    # with CF previously emitted in schema-invalid order)
    edits = []

    # 1. region replacements / removals / insertions
    for tag, rendered in region_changes.items():
        spans = scan.regions.get(tag, [])
        if spans:
            span = spans[0]
            edits.append((span.start, span.end, rendered or b"",
                          CT_ORDER_INDEX.get(tag, 0)))
        elif rendered:
            insert_at = _region_insert_offset(scan, tag)
            edits.append((insert_at, insert_at, rendered,
                          CT_ORDER_INDEX.get(tag, 0)))

    # 1b. conditional formatting: replace the whole (possibly multi-element)
    # run with the freshly rendered blocks (dxf handling done by the caller)
    if cf_replacement is not None:
        # twin-bearing sheets reach here only with COMPOSED bytes (the
        # saver's x14 planner) — the orphaning gates moved
        # into that planner
        spans = scan.regions.get("conditionalFormatting", [])
        cf_rank = CT_ORDER_INDEX["conditionalFormatting"]
        if spans:
            edits.append((spans[0].start, spans[0].end, cf_replacement,
                          cf_rank))
            for extra in spans[1:]:
                edits.append((extra.start, extra.end, b"", cf_rank))
        elif cf_replacement:
            offset = _region_insert_offset(scan, "conditionalFormatting")
            edits.append((offset, offset, cf_replacement, cf_rank))

    # 1c. hyperlinks element
    if hyperlinks_replacement is not None:
        spans = scan.regions.get("hyperlinks", [])
        link_rank = CT_ORDER_INDEX["hyperlinks"]
        if spans:
            edits.append((spans[0].start, spans[0].end,
                          hyperlinks_replacement, link_rank))
        elif hyperlinks_replacement:
            offset = _region_insert_offset(scan, "hyperlinks")
            edits.append((offset, offset, hyperlinks_replacement, link_rank))

    # 2. cell and row edits inside sheetData (rank 0: sheetData precedes
    # every insertable region, so offsets never collide with them)
    edits.extend((s, e, r, 0) for (s, e, r)
                 in _sheetdata_edits(ws, scan, original, dirty_cells,
                                     row_attr_changes, style_resolver,
                                     value_overwrites=value_overwrites,
                                     cache_invalidations=cache_invalidations))
    if cache_writes:
        overlap = set(cache_writes) & set(dirty_cells)
        if overlap:
            raise SpliceRefusal(
                "internal: cache writes and dirty cells overlap at "
                "{0}".format(sorted(overlap)[:4]))
        invalidate_overlap = set(cache_writes) & cache_invalidations
        if invalidate_overlap:
            raise SpliceRefusal(
                "internal: cache writes and cache invalidations overlap at "
                "{0}".format(sorted(invalidate_overlap)[:4]))
        edits.extend((s, e, r, 0) for (s, e, r)
                     in _cache_value_edits(ws, scan, original, cache_writes))
    standalone_invalidations = cache_invalidations - set(dirty_cells)
    if standalone_invalidations:
        edits.extend((s, e, r, 0) for (s, e, r)
                     in _formula_cache_invalidation_edits(
                         ws, scan, original, standalone_invalidations))

    # ------- apply (sorted, non-overlapping by construction) -------------
    edits.sort(key=lambda e: (e[0], e[1], e[3]))
    for (s1, e1), (s2, e2) in zip(
            [(s, e) for s, e, _r, _k in edits],
            [(s, e) for s, e, _r, _k in edits][1:]):
        if e1 > s2:
            raise SpliceRefusal(
                "internal: overlapping splice edits ({0},{1}) and "
                "({2},{3})".format(s1, e1, s2, e2))

    out = []
    pos = 0
    for start, end, replacement, _rank in edits:
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


def _sheetdata_edits(ws, scan, original, dirty_cells, row_attr_changes, resolve,
                     value_overwrites=frozenset(),
                     cache_invalidations=frozenset()):
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
                ws, sorted(r for r in touched_rows), by_row, row_attr_changes,
                resolve)
            edits.append((span.start, span.end,
                          b"<sheetData>" + new_rows_bytes + b"</sheetData>"))
            return edits

    pending_new_rows = sorted(r for r in touched_rows if r not in scan.rows)

    # existing rows: per-row cell edits and attribute sync
    for row_index in sorted(r for r in touched_rows if r in scan.rows):
        row_span = scan.rows[row_index]
        row_edits = _row_edits(ws, row_span, original,
                               by_row.get(row_index, set()),
                               row_attr_changes.get(row_index), resolve,
                               value_overwrites=value_overwrites,
                               cache_invalidations=cache_invalidations)
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
                                      sorted(by_row.get(new_index, set())),
                                      resolve)
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


def _emit_row_cells(ws, row_index, columns, resolve):
    cells_bytes = []
    for col in columns:
        cell = ws._cells.get((row_index, col))
        if cell is None:
            continue
        rendered = emit.emit_cell(ws, cell, resolve(cell))
        if rendered is not None:
            cells_bytes.append(rendered)
    return cells_bytes


def _emit_rows_block(ws, row_indices, by_row, row_attr_changes, resolve):
    parts = []
    for row_index in row_indices:
        cells_bytes = _emit_row_cells(ws, row_index,
                                      sorted(by_row.get(row_index, set())),
                                      resolve)
        attrs = row_attr_changes.get(row_index, {})
        if cells_bytes or attrs:
            parts.append(emit.emit_new_row(ws, row_index, cells_bytes, attrs))
    return b"".join(parts)


def _row_edits(ws, row_span, original, dirty_cols, new_attrs, resolve,
               value_overwrites=frozenset(), cache_invalidations=frozenset()):
    """Edits inside one existing row: replace/insert/delete cells, sync the
    row start tag's attributes when they changed."""
    from openpyxl.cell.rich_text import CellRichText

    edits = []

    if new_attrs is not None:
        # rewrite the start tag, preserving the row's content; original
        # attributes the model does not know (x14ac:dyDescent, xr:uid, ...)
        # are carried verbatim — the attribute-carry rule applies to rows too
        attrs = dict(new_attrs)
        for key, value in row_span.attrs.items():
            if key not in _MODEL_ROW_ATTRS and key != "r":
                attrs.setdefault(key, value)
        if "spans" in row_span.attrs:
            attrs.setdefault("spans", row_span.attrs["spans"])
        if row_span.self_closing:
            if dirty_cols:
                cells_bytes = _emit_row_cells(ws, row_span.index,
                                              sorted(dirty_cols), resolve)
                edits.append((row_span.start, row_span.end,
                              emit.emit_new_row(ws, row_span.index,
                                                cells_bytes, attrs)))
            else:
                edits.append((row_span.start, row_span.end,
                              emit.row_start_tag(row_span.index, attrs,
                                                 self_closing=True)))
            return edits
        edits.append((row_span.start, row_span.content_start,
                      emit.row_start_tag(row_span.index, attrs)))

    if not dirty_cols:
        return edits

    if row_span.self_closing:
        # row exists but holds no cells: rebuild it whole
        cells_bytes = _emit_row_cells(ws, row_span.index, sorted(dirty_cols),
                                      resolve)
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
        coordinate = (row_span.index, col)
        original_cell = (original[cell_span.start:cell_span.end]
                         if cell_span is not None else None)
        preserve_cell_content = (
            cell is not None and original_cell is not None
            and coordinate not in value_overwrites
            and cell_span.shared_si is None
            and not isinstance(cell._value, CellRichText))
        if preserve_cell_content:
            rendered = emit.patch_cell_style(original_cell, resolve(cell))
        else:
            rendered = emit.emit_cell(ws, cell, resolve(cell)) \
                if cell is not None else None
        if rendered is not None and cell_span is not None \
                and not preserve_cell_content:
            rendered = emit.carry_attributes(
                rendered, original_cell,
                drop_metadata=coordinate in value_overwrites)
        if rendered is not None and cell_span is not None \
                and coordinate in cache_invalidations:
            formula_names = getattr(cell_span, "formula_names", ()) \
                if preserve_cell_content else ()
            cache_names = getattr(cell_span, "cache_names", ()) \
                if preserve_cell_content else ()
            rendered = _patch_formula_cache_invalidation(
                rendered,
                "{0}!r{1}c{2}".format(
                    ws.title, row_span.index, col),
                require_formula=False,
                formula_names=formula_names,
                cache_names=cache_names)

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


# ---------------------------------------------------------------------
# oracle write-back: cached-value updates on untouched
# formula cells — the <f> bytes verbatim, the <v> replaced

def _serialize_cached_value(value, epoch):
    """(t_attr_or_None, v_text) for a computed value; the mirror of what
    Excel itself writes for a formula cell's cache."""
    import datetime

    from openpyxl.utils.datetime import to_excel

    if isinstance(value, bool):
        return b"b", b"1" if value else b"0"
    if isinstance(value, (int, float)):
        return None, repr(value).encode("ascii")
    if isinstance(value, (datetime.datetime, datetime.date, datetime.time,
                          datetime.timedelta)):
        return None, repr(to_excel(value, epoch)).encode("ascii")
    if isinstance(value, str):
        from openpyxl.oracle import ERROR_TOKENS

        is_error = value in ERROR_TOKENS
        text = value.encode("utf-8")
        text = (text.replace(b"&", b"&amp;").replace(b"<", b"&lt;")
                .replace(b">", b"&gt;"))
        text = text.replace(b"\r", b"&#13;")
        if is_error:
            return b"e", text
        return b"str", text
    raise SpliceRefusal(
        "cache write value {0!r} has no cached-value serialization. "
        "Nothing was written.".format(value))


def _cache_value_edits(ws, scan, original, cache_writes):
    edits = []
    epoch = ws.parent.epoch
    array_bounds = [_range_bounds(ref) for ref in scan.array_refs]
    for (row, col), value in sorted(cache_writes.items()):
        label = "{0}!r{1}c{2}".format(ws.title, row, col)
        row_span = scan.rows.get(row)
        cell_span = row_span.cells.get(col) if row_span is not None else None
        if cell_span is None:
            raise SpliceRefusal(
                "cache write target {0} does not exist in the original "
                "bytes. Nothing was written.".format(label))
        cell_bytes = original[cell_span.start:cell_span.end]
        edits.append((cell_span.start, cell_span.end,
                      _patch_cached_value(
                          cell_bytes, value, epoch, label,
                          allow_cache_only=any(
                              _coordinate_in_bounds((row, col), bounds)
                              for bounds in array_bounds),
                          formula_names=getattr(
                              cell_span, "formula_names", ()),
                          cache_names=getattr(
                              cell_span, "cache_names", ()))))
    return edits


def _formula_cache_invalidation_edits(ws, scan, original, coordinates):
    edits = []
    array_bounds = [_range_bounds(ref) for ref in scan.array_refs]
    for row, col in sorted(coordinates):
        label = "{0}!r{1}c{2}".format(ws.title, row, col)
        row_span = scan.rows.get(row)
        cell_span = row_span.cells.get(col) if row_span is not None else None
        if cell_span is None:
            raise SpliceRefusal(
                "cache invalidation target {0} does not exist in the "
                "original bytes. Nothing was written.".format(label))
        cell_bytes = original[cell_span.start:cell_span.end]
        edits.append((cell_span.start, cell_span.end,
                      _patch_formula_cache_invalidation(
                          cell_bytes, label,
                          allow_cache_only=any(
                              _coordinate_in_bounds((row, col), bounds)
                              for bounds in array_bounds),
                          formula_names=getattr(
                              cell_span, "formula_names", ()),
                          cache_names=getattr(
                              cell_span, "cache_names", ()))))
    return edits


def _patch_formula_cache_invalidation(cell_bytes, label,
                                      require_formula=True,
                                      allow_cache_only=False,
                                      formula_names=(), cache_names=()):
    try:
        tag_end, self_closing, _attrs = emit._start_tag_attributes(cell_bytes)
    except ValueError as exc:
        raise SpliceRefusal(
            "cache invalidation target {0} is not a formula cell. Nothing "
            "was written.".format(label)) from exc
    head = cell_bytes[:tag_end + 1]
    if self_closing:
        if require_formula:
            raise SpliceRefusal(
                "cache invalidation target {0} carries no formula. Nothing "
                "was written.".format(label))
        return cell_bytes
    close_start = cell_bytes.rfind(b"</c>")
    if close_start < tag_end or cell_bytes[close_start:] != b"</c>":
        raise SpliceRefusal(
            "cache invalidation target {0} has malformed cell markup. "
            "Nothing was written.".format(label))
    body = cell_bytes[tag_end + 1:close_start]
    try:
        children = _direct_cell_children(cell_bytes)
        formula_names = set(formula_names) | {b"f"}
        cache_names = set(cache_names) | {b"v"}
        formulas = [child for child in children
                    if child.name in formula_names]
        caches = [child for child in children if child.name in cache_names]
    except ValueError as exc:
        raise SpliceRefusal(
            "cache invalidation target {0} has malformed child markup. "
            "Nothing was written.".format(label)) from exc
    if len(formulas) > 1:
        raise SpliceRefusal(
            "cache invalidation target {0} has duplicate formula "
            "markup. Nothing was written.".format(label))
    if not formulas:
        if require_formula and not allow_cache_only:
            raise SpliceRefusal(
                "cache invalidation target {0} carries no formula. "
                "Nothing was written.".format(label))
        if not allow_cache_only:
            return cell_bytes

    head = emit.patch_start_tag_attribute(head, b"t", None)
    for cached in reversed(caches):
        body_start = tag_end + 1
        start = cached.start - body_start
        end = cached.end - body_start
        body = body[:start] + body[end:]
    return head + body + b"</c>"


def _patch_cached_value(cell_bytes, value, epoch, label,
                        allow_cache_only=False,
                        formula_names=(), cache_names=()):
    try:
        tag_end, self_closing, _attrs = emit._start_tag_attributes(cell_bytes)
    except ValueError as exc:
        raise SpliceRefusal(
            "cache write target {0} is not a formula cell. Nothing was "
            "written.".format(label)) from exc
    head = cell_bytes[:tag_end + 1]
    if self_closing:
        body, close = b"", b""
    else:
        close_start = cell_bytes.rfind(b"</c>")
        if close_start < tag_end or cell_bytes[close_start:] != b"</c>":
            raise SpliceRefusal(
                "cache write target {0} has malformed cell markup. Nothing "
                "was written.".format(label))
        body = cell_bytes[tag_end + 1:close_start]
        close = b"</c>"
    try:
        children = _direct_cell_children(cell_bytes)
        formula_names = set(formula_names) | {b"f"}
        cache_names = set(cache_names) | {b"v"}
        formulas = [child for child in children
                    if child.name in formula_names]
        caches = [child for child in children if child.name in cache_names]
    except ValueError as exc:
        raise SpliceRefusal(
            "cache write target {0} has malformed child markup. Nothing "
            "was written.".format(label)) from exc
    if len(formulas) > 1 or len(caches) > 1:
        raise SpliceRefusal(
            "cache write target {0} has duplicate formula/cache markup. "
            "Nothing was written.".format(label))
    formula = formulas[0] if formulas else None
    cached = caches[0] if caches else None
    if formula is None and not allow_cache_only:
        raise SpliceRefusal(
            "cache write target {0} carries no formula. Nothing was "
            "written.".format(label))
    if any(child.name not in (b"f", b"v") for child in children):
        raise SpliceRefusal(
            "cache write target {0} carries content besides its formula "
            "and cached value; updating it is not supported. Nothing was "
            "written.".format(label))
    t_attr, v_text = _serialize_cached_value(value, epoch)
    head = emit.patch_start_tag_attribute(head, b"t", t_attr)
    v_open = b"<v>"
    if isinstance(value, str) and value != value.strip():
        v_open = b'<v xml:space="preserve">'
    rendered = v_open + v_text + b"</v>"
    body_start = tag_end + 1
    if cached is not None:
        start = cached.start - body_start
        end = cached.end - body_start
        body = body[:start] + rendered + body[end:]
    elif formula is not None:
        insert_at = formula.end - body_start
        body = body[:insert_at] + rendered + body[insert_at:]
    elif close:
        body = rendered + body
    else:
        head = head[:-2] + b">"
        body = rendered
        close = b"</c>"
    return head + body + close
