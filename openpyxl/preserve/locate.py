# paper-xlsx: label localization (PLAN-v0.1 Batch 6, PR-1 §5; battery 23)

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
    """The nearest non-label cell in the preferred direction: first the
    immediate neighbour; if that is a non-empty STRING (another label),
    keep walking up to a few cells for the first value-like cell."""
    if prefer == "right":
        step = (0, 1)
    elif prefer == "below":
        step = (1, 0)
    else:
        raise ValueError("prefer must be 'right' or 'below' "
                         "(got {0!r})".format(prefer))
    r, c = row, col
    for _ in range(4):
        r += step[0]
        c += step[1]
        cell = ws._cells.get((r, c))
        if cell is None or cell._value is None:
            # an empty cell is a legitimate TARGET (a value slot waiting
            # to be filled) only when materialized; unmaterialized gaps
            # end the walk
            if cell is not None:
                return cell
            return None
        if cell.data_type == "f" or not isinstance(cell._value, str):
            return cell
        # a string: another label or a text VALUE — treat a string
        # immediately adjacent as a text value only if nothing better
        # follows; keep it as fallback
        fallback = cell
        nxt = ws._cells.get((r + step[0], c + step[1]))
        if nxt is None or nxt._value is None:
            return fallback
    return None


def locate(ws, label, *, prefer="right"):
    """Worksheet.locate implementation (PR-1 §5)."""
    if not isinstance(label, str) or not label.strip():
        raise TypeError("locate() takes a non-empty label string")
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
    cell = _value_neighbour(ws, row, col, prefer)
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
# data-validation vocabulary (PR-1 §5: Worksheet.allowed_values)

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
        out = []
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                vcell = target_ws._cells.get((r, c))
                if vcell is not None and vcell._value is not None:
                    out.append(vcell._value)
        return out
    return None
