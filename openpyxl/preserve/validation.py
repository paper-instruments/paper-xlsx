# paper-xlsx: data-validation vocabulary

"""Read deterministic list data-validation values for a worksheet cell."""


def _refuse(ws, cell_coordinate, detail, *, kind):
    from openpyxl.errors import UnsupportedStructureError

    address = "{0}!{1}".format(ws.title, cell_coordinate)
    raise UnsupportedStructureError(
        "list validation for {0} {1}. Only literal lists and static, "
        "one-dimensional value ranges are supported.".format(
            address, detail),
        kind=kind,
        anchor=address,
    )


def _literal_values(ws, coordinate, source):
    """Decode the outer Excel string while preserving item whitespace."""
    if len(source) < 2 or source[-1] != '"':
        _refuse(ws, coordinate, "has an unterminated literal list",
                kind="invalid-validation-literal")
    inner = source[1:-1]
    values = [""]
    index = 0
    while index < len(inner):
        char = inner[index]
        if char == ",":
            values.append("")
            index += 1
            continue
        if char == '"':
            if index + 1 >= len(inner) or inner[index + 1] != '"':
                _refuse(ws, coordinate,
                        "contains an invalid quote in its literal list",
                        kind="invalid-validation-literal")
            values[-1] += '"'
            index += 2
            continue
        values[-1] += char
        index += 1
    return values


def _static_range(ws, coordinate, source):
    from openpyxl.errors import TargetNotFoundError
    from openpyxl.utils.cell import SHEETRANGE_RE, range_boundaries

    if any(marker in source for marker in ("[", "]", "#", "(", ")")):
        _refuse(ws, coordinate,
                "uses a structured, external, dynamic, or spill source",
                kind="unsupported-validation-source")

    target_ws = ws
    ref = source
    if "!" in source:
        match = SHEETRANGE_RE.match(source)
        if match is None or match.end() != len(source):
            _refuse(ws, coordinate,
                    "uses a multi-area or malformed range source",
                    kind="unsupported-validation-source")
        title = match.group("quoted") or match.group("notquoted")
        if match.group("quoted") is not None:
            title = title.replace("''", "'")
        ref = match.group("cells")
        matches = [candidate for candidate in ws.parent.worksheets
                   if candidate.title.casefold() == title.casefold()]
        if not matches:
            address = "{0}!{1}".format(ws.title, coordinate)
            raise TargetNotFoundError(
                "list validation for {0} references missing worksheet "
                "{1!r}.".format(address, title),
                kind="missing-validation-source-sheet",
                anchor=address,
            )
        target_ws = matches[0]
    try:
        min_col, min_row, max_col, max_row = range_boundaries(ref)
    except ValueError:
        _refuse(ws, coordinate,
                "uses a defined name, formula, or malformed range source",
                kind="unsupported-validation-source")

    # A bare column-like name such as IN is a defined name, not a range.
    if ":" not in ref and (min_col is None or min_row is None):
        _refuse(ws, coordinate,
                "uses a defined name rather than a direct cell range",
                kind="unsupported-validation-source")
    min_row = 1 if min_row is None else min_row
    min_col = 1 if min_col is None else min_col
    max_row = (target_ws.max_row or 1) if max_row is None else max_row
    max_col = (target_ws.max_column or 1) if max_col is None else max_col
    min_row, max_row = sorted((min_row, max_row))
    min_col, max_col = sorted((min_col, max_col))
    if min_row != max_row and min_col != max_col:
        _refuse(ws, coordinate,
                "uses a two-dimensional range {0!r}".format(source),
                kind="unsupported-validation-source")
    return target_ws, min_col, min_row, max_col, max_row


def allowed_values(ws, cell):
    """Return deterministic list-validation values covering ``cell``.

    ``None`` means no list validation covers the cell. A matching validation
    whose vocabulary cannot be reported exactly raises a typed refusal.
    Blank cells in a static source range are returned as ``None`` entries.

    :param ws: Worksheet containing the validated cell.
    :type ws: openpyxl.worksheet.worksheet.Worksheet
    :param cell: Cell object or coordinate to inspect.
    :type cell: openpyxl.cell.cell.Cell or str
    :return: Allowed values, or ``None`` when no list validation applies.
    :rtype: list or None
    """
    from openpyxl.cell.cell import MergedCell
    from openpyxl.errors import AmbiguousTargetError
    from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter

    if hasattr(cell, "row") and hasattr(cell, "column"):
        if getattr(cell, "parent", ws) is not ws:
            raise ValueError("allowed_values cell must belong to the worksheet")
        row, col = cell.row, cell.column
    else:
        row, col = coordinate_to_tuple(str(cell).replace("$", ""))
    coordinate = "{0}{1}".format(get_column_letter(col), row)
    matches = []
    for validation in ws.data_validations.dataValidation:
        if validation.type != "list":
            continue
        if any(rng.min_row <= row <= rng.max_row
               and rng.min_col <= col <= rng.max_col
               for rng in getattr(validation.sqref, "ranges", [])):
            matches.append(validation)
    if not matches:
        return None
    if isinstance(ws._cells.get((row, col)), MergedCell):
        _refuse(ws, coordinate, "targets a non-anchor merged cell",
                kind="merged-validation-target")
    if len(matches) != 1:
        address = "{0}!{1}".format(ws.title, coordinate)
        raise AmbiguousTargetError(
            "{0} is covered by {1} list validations; no single vocabulary "
            "can be reported safely.".format(address, len(matches)),
            kind="overlapping-list-validations",
            anchor=address,
        )

    formula = matches[0].formula1
    if formula is None or not formula.strip():
        _refuse(ws, coordinate, "has no source formula",
                kind="unsupported-validation-source")
    source = formula.strip()
    if source.startswith("="):
        source = source[1:].strip()
    if source.startswith('"'):
        return _literal_values(ws, coordinate, source)
    if '"' in source:
        _refuse(ws, coordinate, "contains malformed literal text",
                kind="invalid-validation-literal")

    target_ws, min_col, min_row, max_col, max_row = _static_range(
        ws, coordinate, source)
    values = []
    if min_row == max_row:
        coordinates = ((min_row, col_index)
                       for col_index in range(min_col, max_col + 1))
    else:
        coordinates = ((row_index, min_col)
                       for row_index in range(min_row, max_row + 1))
    for source_coordinate in coordinates:
        value_cell = target_ws._cells.get(source_coordinate)
        if isinstance(value_cell, MergedCell):
            _refuse(ws, coordinate,
                    "reads a non-anchor cell inside a merged range",
                    kind="unsupported-validation-source")
        if value_cell is None:
            values.append(None)
            continue
        if value_cell.data_type == "f":
            _refuse(ws, coordinate,
                    "reads formula cell {0}!{1}".format(
                        target_ws.title, value_cell.coordinate),
                    kind="formula-validation-source")
        if value_cell.data_type == "e":
            _refuse(ws, coordinate,
                    "reads error cell {0}!{1}".format(
                        target_ws.title, value_cell.coordinate),
                    kind="error-validation-source")
        values.append(value_cell._value)
    return values
