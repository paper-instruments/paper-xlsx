# paper-xlsx: explicit style helpers

"""Small preserve-aware helpers for explicit formatting operations."""

from copy import copy

from openpyxl.utils.cell import range_boundaries
from openpyxl.xml.constants import MAX_COLUMN, MAX_ROW


def _coordinate_state(mapping, ws, coordinate):
    bucket = mapping.get(ws)
    if bucket is None:
        return False, None, False
    return True, bucket, coordinate in bucket


def _restore_coordinate(mapping, ws, coordinate, state):
    existed, original, present = state
    current = mapping.get(ws)
    if existed:
        if current is not original:
            mapping[ws] = original
        bucket = original
    else:
        bucket = current
        if bucket is None:
            return
    bucket.add(coordinate) if present else bucket.discard(coordinate)
    if not existed and not bucket:
        mapping.pop(ws, None)


class _CopyFormatTransaction:
    """Range-local rollback for one explicit format copy."""

    def __init__(self, ws, coordinates):
        self.ws = ws
        self.coordinates = {}
        for coordinate in coordinates:
            cell = ws._cells.get(coordinate)
            self.coordinates[coordinate] = (
                cell is not None,
                cell,
                getattr(cell, "_style", None),
            )
        self.current_row = ws._current_row
        self.ledger = getattr(ws.parent, "_paper_ledger", None)
        if self.ledger is not None and not self.ledger.armed:
            self.ledger = None
        self.ledger_states = {}
        if self.ledger is not None:
            for coordinate in coordinates:
                self.ledger_states[coordinate] = _coordinate_state(
                    self.ledger.cells, ws, coordinate)
            self.was_warned = ws in self.ledger.protection_warned
        self.active = True

    def commit(self):
        self.active = False

    def rollback(self):
        if not self.active:
            return
        for coordinate, (existed, cell, style) in reversed(
                list(self.coordinates.items())):
            if not existed:
                self.ws._cells.pop(coordinate, None)
                continue
            self.ws._cells[coordinate] = cell
            cell._style = style
        self.ws._current_row = self.current_row
        if self.ledger is not None:
            for coordinate, state in self.ledger_states.items():
                _restore_coordinate(
                    self.ledger.cells, self.ws, coordinate, state)
            if self.was_warned:
                self.ledger.protection_warned.add(self.ws)
            else:
                self.ledger.protection_warned.discard(self.ws)
        self.active = False


def _copy_format_commit_point(_coordinate):
    """Test seam for failure injection after each complete cell edit."""


def _source_style(ws, src_cell):
    from openpyxl.cell.cell import MergedCell
    from openpyxl.utils.cell import coordinate_to_tuple

    if isinstance(src_cell, str):
        row, column = coordinate_to_tuple(src_cell.replace("$", ""))
        if not 1 <= row <= MAX_ROW or not 1 <= column <= MAX_COLUMN:
            raise ValueError("copy_format source exceeds worksheet boundaries")
        source = ws._cells.get((row, column))
        style = getattr(source, "_style", None)
        if isinstance(source, MergedCell):
            from openpyxl.errors import UnsupportedStructureError

            raise UnsupportedStructureError(
                "copy_format source {0!r} is a non-anchor merged cell. "
                "Nothing was changed.".format(src_cell),
                kind="merged-format-source",
                anchor="{0}!{1}".format(ws.title, src_cell.replace("$", "")),
            )
        return (row, column), style
    if getattr(src_cell, "parent", None) is not ws:
        raise ValueError(
            "copy_format source cell must belong to the supplied worksheet")
    if isinstance(src_cell, MergedCell):
        from openpyxl.errors import UnsupportedStructureError

        raise UnsupportedStructureError(
            "copy_format source is a non-anchor merged cell. Nothing was "
            "changed.",
            kind="merged-format-source",
            anchor="{0}!{1}".format(ws.title, src_cell.coordinate),
        )
    return (src_cell.row, src_cell.column), getattr(src_cell, "_style", None)


def _plan_copy_format(ws, src_cell, dst_range):
    from openpyxl.cell.cell import MergedCell
    from openpyxl.errors import UnsupportedStructureError

    min_col, min_row, max_col, max_row = range_boundaries(
        str(dst_range).replace("$", ""))
    if None in (min_col, min_row, max_col, max_row):
        raise ValueError(
            "copy_format destination must be a finite cell range")
    min_row, max_row = sorted((min_row, max_row))
    min_col, max_col = sorted((min_col, max_col))
    if min_row < 1 or min_col < 1 \
            or max_row > MAX_ROW or max_col > MAX_COLUMN:
        raise ValueError(
            "copy_format destination exceeds the worksheet boundary")
    source_coordinate, source_style = _source_style(ws, src_cell)
    coordinates = tuple(
        (row, column)
        for row in range(min_row, max_row + 1)
        for column in range(min_col, max_col + 1)
        if (row, column) != source_coordinate
    )
    for coordinate in coordinates:
        cell = ws._cells.get(coordinate)
        if isinstance(cell, MergedCell):
            raise UnsupportedStructureError(
                "copy_format destination {0} is a non-anchor merged cell. "
                "Nothing was changed.".format(cell.coordinate),
                kind="merged-format-destination",
                anchor="{0}!{1}".format(ws.title, cell.coordinate),
            )

    ledger = getattr(ws.parent, "_paper_ledger", None)
    if ledger is not None and ledger.armed and bool(ws.protection.sheet):
        locked = []
        for coordinate in coordinates:
            cell = ws._cells.get(coordinate)
            protection = getattr(cell, "protection", None)
            is_locked = True if protection is None \
                else protection.locked is not False
            if is_locked:
                locked.append(coordinate)
        if locked and getattr(ws.parent, "strict_protection", False):
            from openpyxl.utils.cell import get_column_letter

            row, column = locked[0]
            address = "{0}{1}".format(get_column_letter(column), row)
            raise UnsupportedStructureError(
                "copy_format would change locked cell {0} on protected "
                "sheet {1!r}; strict_protection refuses the complete range. "
                "Nothing was changed.".format(address, ws.title),
                kind="protected-format-copy",
                anchor="{0}!{1}".format(ws.title, address),
            )
        warn = bool(locked and ws not in ledger.protection_warned)
    else:
        warn = False
    return source_style, coordinates, warn


def copy_format(ws, src_cell, dst_range):
    """Atomically copy one cell's complete cell style onto a finite range.

    The copied style includes font, fill, border, alignment, number format,
    and protection. Values, formulas, comments, hyperlinks, validation, row
    heights, and column widths are not copied.
    """
    from openpyxl.preserve.ledger import mark_styleable_dirty

    style_array, coordinates, warn = _plan_copy_format(
        ws, src_cell, dst_range)
    transaction = _CopyFormatTransaction(ws, coordinates)
    try:
        if warn:
            import warnings

            from openpyxl.errors import ProtectedWriteWarning

            transaction.ledger.protection_warned.add(ws)
            warnings.warn(ProtectedWriteWarning(
                "copy_format is changing locked cell format(s) on protected "
                "sheet {0!r}. The operation proceeds, but the sheet's author "
                "expected those cells to be read-only. Set "
                "wb.strict_protection = True to refuse the complete range."
                .format(ws.title)), stacklevel=2)
        for coordinate in coordinates:
            row, column = coordinate
            cell = ws.cell(row=row, column=column)
            cell._style = copy(style_array) if style_array is not None \
                else None
            mark_styleable_dirty(cell)
            _copy_format_commit_point(coordinate)
    except BaseException:
        transaction.rollback()
        raise
    transaction.commit()
    return len(coordinates)
