# paper-xlsx: explicit style helpers

"""Small preserve-aware helpers for explicit formatting operations."""

from copy import copy

from openpyxl.utils.cell import range_boundaries


def copy_format(ws, src_cell, dst_range):
    """Copy one cell's complete format onto every cell in a range."""
    src = ws[src_cell.replace("$", "")] if isinstance(src_cell, str) \
        else src_cell
    style_array = getattr(src, "_style", None)
    min_col, min_row, max_col, max_row = range_boundaries(
        str(dst_range).replace("$", ""))
    count = 0
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            if (row, col) == (src.row, src.column):
                continue
            cell = ws.cell(row=row, column=col)
            cell._style = copy(style_array) if style_array is not None \
                else None
            from openpyxl.preserve.ledger import mark_styleable_dirty

            mark_styleable_dirty(cell)
            count += 1
    return count
