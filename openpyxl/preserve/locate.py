# paper-xlsx: label localization

"""Find the VALUE cell that belongs to a text label — "give me the cell
next to 'Growth rate'" — with typed honesty about ambiguity.

Matching is exact-first, then normalized (case/whitespace-insensitive).
Zero label matches raise TargetNotFoundError; more than one label match,
or no locatable value neighbour, raises AmbiguousTargetError listing
every candidate address (the pinned class, earning its keep)."""

from openpyxl.errors import AmbiguousTargetError, TargetNotFoundError


def _normalize(text):
    return " ".join(str(text).split()).casefold()


def _label_matches(ws, label):
    exact = []
    normalized = []
    want = _normalize(label)
    for (row, col), cell in sorted(ws._cells.items()):
        value = cell._value
        if not isinstance(value, str) or cell.data_type == "f":
            continue
        if value == label:
            exact.append((row, col))
        elif _normalize(value) == want:
            normalized.append((row, col))
    return exact if exact else normalized


def _value_neighbour(ws, row, col, prefer):
    """The nearest value cell in the preferred direction — REFUSE, never
    guess (every silent-guess branch was a lying
    instrument). Walk rules:

    - merged-range interiors are COVERED cells, never targets: skipped;
    - formulas, non-strings and error-typed cells are values: returned;
    - a materialized EMPTY cell is a fillable value slot: returned;
    - a STRING neighbour is genuinely two-faced (text value or another
      label): it is returned only when the walk ends right after it —
      if anything populated follows, the caller gets an
      AmbiguousTargetError listing both candidates (raised by locate);
    - an unmaterialized gap ends the walk.

    Returns (cell, competing_cell_or_None); (None, None) = no target."""
    from openpyxl.cell.cell import MergedCell

    step = (0, 1) if prefer == "right" else (1, 0)
    r, c = row, col
    fallback = None
    for _ in range(6):
        r += step[0]
        c += step[1]
        cell = ws._cells.get((r, c))
        if cell is None:
            return (fallback, None)          # gap ends the walk
        if isinstance(cell, MergedCell):
            continue                          # covered, never a target
        if cell._value is None:
            if fallback is not None:
                return (fallback, None)
            return (cell, None)               # fillable value slot
        if cell.data_type == "f" \
                or not isinstance(cell._value, str) \
                or cell.data_type == "e":
            if fallback is not None:
                return (fallback, cell)       # competition: ambiguous
            return (cell, None)
        # a string: candidate value OR another label
        if fallback is not None:
            return (fallback, cell)           # two strings: ambiguous
        fallback = cell
    return (fallback, None)


def locate(ws, label, *, prefer="right"):
    """Worksheet.locate implementation."""
    if not isinstance(label, str) or not label.strip():
        raise TypeError("locate() takes a non-empty label string")
    if prefer not in ("right", "below"):
        raise ValueError("prefer must be 'right' or 'below' "
                         "(got {0!r})".format(prefer))
    matches = _label_matches(ws, label)
    if not matches:
        raise TargetNotFoundError(
            "no cell on sheet {0!r} holds the label {1!r} (exact or "
            "normalized).".format(ws.title, label),
            kind="label-not-found",
            anchor="{0}!{1}".format(ws.title, label))
    if len(matches) > 1:
        addresses = ["{0}!{1}{2}".format(
            ws.title, _col_letter(c), r) for (r, c) in matches]
        raise AmbiguousTargetError(
            "label {0!r} appears {1} times on sheet {2!r}: {3}. Qualify "
            "the request (a cell address, or a more specific "
            "label).".format(label, len(matches), ws.title,
                             ", ".join(addresses)),
            kind="ambiguous-label",
            anchor=addresses[0],
            options=addresses)
    row, col = matches[0]
    cell, competitor = _value_neighbour(ws, row, col, prefer)
    if competitor is not None:
        candidates = ["{0}!{1}".format(ws.title, cell.coordinate),
                      "{0}!{1}".format(ws.title, competitor.coordinate)]
        raise AmbiguousTargetError(
            "label {0!r} has two plausible value cells {1}: the adjacent "
            "text could be the value or another label. Address the cell "
            "directly.".format(label, " and ".join(candidates)),
            kind="ambiguous-value-cell",
            anchor=candidates[0],
            options=candidates)
    if cell is None:
        label_addr = "{0}!{1}{2}".format(ws.title, _col_letter(col), row)
        other = "below" if prefer == "right" else "right"
        raise AmbiguousTargetError(
            "label {0!r} at {1} has no value cell {2} of it; try "
            "prefer={3!r} or address the cell directly.".format(
                label, label_addr, prefer, other),
            kind="no-value-neighbour",
            anchor=label_addr,
            options=[label_addr])
    return cell


def _col_letter(col):
    from openpyxl.utils import get_column_letter

    return get_column_letter(col)


# ---------------------------------------------------------------------
# data-validation vocabulary (Worksheet.allowed_values)

def allowed_values(ws, cell):
    """The list-type data-validation vocabulary covering ``cell``
    (a Cell or an address string), or None when no list DV covers it or
    its source cannot be read without evaluation."""
    from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries

    if hasattr(cell, "row") and hasattr(cell, "column"):
        row, col = cell.row, cell.column
    else:
        row, col = coordinate_to_tuple(str(cell).replace("$", ""))
    for dv in ws.data_validations.dataValidation:
        if dv.type != "list" or not dv.formula1:
            continue
        hit = False
        for rng in getattr(dv.sqref, "ranges", []):
            if (rng.min_row <= row <= rng.max_row
                    and rng.min_col <= col <= rng.max_col):
                hit = True
                break
        if not hit:
            continue
        source = dv.formula1.strip()
        if source.startswith("="):
            source = source[1:]
        if source.startswith('"') and source.endswith('"'):
            # literal vocabulary: "Yes,No,Maybe"
            return [item.strip() for item in source[1:-1].split(",")]
        # a range reference: read the cells (same-sheet or qualified)
        ref = source.replace("$", "")
        target_ws = ws
        if "!" in ref:
            title, ref = ref.rsplit("!", 1)
            title = title.strip("'").replace("''", "'")
            matches = [w for w in ws.parent.worksheets
                       if w.title.casefold() == title.casefold()]
            if not matches:
                return None
            target_ws = matches[0]
        try:
            min_col, min_row, max_col, max_row = range_boundaries(ref)
        except ValueError:
            return None            # a name/formula source: not readable
        # whole-column/row sources clamp to the populated extent
        # (range_boundaries hands back None bounds — raw
        # TypeError); reversed sources normalize (gate: silent [])
        min_row = 1 if min_row is None else min_row
        min_col = 1 if min_col is None else min_col
        max_row = (target_ws.max_row or 1) if max_row is None else max_row
        max_col = (target_ws.max_column or 1) if max_col is None \
            else max_col
        if min_row > max_row:
            min_row, max_row = max_row, min_row
        if min_col > max_col:
            min_col, max_col = max_col, min_col
        out = []
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                vcell = target_ws._cells.get((r, c))
                if vcell is not None and vcell._value is not None:
                    out.append(vcell._value)
        return out
    return None
