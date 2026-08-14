# paper-xlsx: data-validation vocabulary

"""Read explicit list data-validation values for a worksheet cell."""


def allowed_values(ws, cell):
    """Return the list validation vocabulary covering ``cell``, if any."""
    from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries

    if hasattr(cell, "row") and hasattr(cell, "column"):
        row, col = cell.row, cell.column
    else:
        row, col = coordinate_to_tuple(str(cell).replace("$", ""))
    for dv in ws.data_validations.dataValidation:
        if dv.type != "list" or not dv.formula1:
            continue
        if not any(rng.min_row <= row <= rng.max_row
                   and rng.min_col <= col <= rng.max_col
                   for rng in getattr(dv.sqref, "ranges", [])):
            continue
        source = dv.formula1.strip()
        if source.startswith("="):
            source = source[1:]
        if source.startswith('"') and source.endswith('"'):
            return [item.strip() for item in source[1:-1].split(",")]
        ref = source.replace("$", "")
        target_ws = ws
        if "!" in ref:
            title, ref = ref.rsplit("!", 1)
            title = title.strip("'").replace("''", "'")
            matches = [candidate for candidate in ws.parent.worksheets
                       if candidate.title.casefold() == title.casefold()]
            if not matches:
                return None
            target_ws = matches[0]
        try:
            min_col, min_row, max_col, max_row = range_boundaries(ref)
        except ValueError:
            return None
        min_row = 1 if min_row is None else min_row
        min_col = 1 if min_col is None else min_col
        max_row = (target_ws.max_row or 1) if max_row is None else max_row
        max_col = (target_ws.max_column or 1) if max_col is None else max_col
        min_row, max_row = sorted((min_row, max_row))
        min_col, max_col = sorted((min_col, max_col))
        out = []
        for item_row in range(min_row, max_row + 1):
            for item_col in range(min_col, max_col + 1):
                value_cell = target_ws._cells.get((item_row, item_col))
                if value_cell is not None and value_cell._value is not None:
                    out.append(value_cell._value)
        return out
    return None
