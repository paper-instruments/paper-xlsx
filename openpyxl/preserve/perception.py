# paper-xlsx: dependency sketch

"""Coarse formula-dependency analysis used by model maps and shift guards."""

import re

from openpyxl.utils.cell import range_boundaries

VOLATILE_NONDETERMINISTIC = ("NOW", "TODAY", "RAND", "RANDBETWEEN",
                             "RANDARRAY", "CELL", "INFO")
_VOLATILE_FUNCTIONS = frozenset(
    name + "(" for name in VOLATILE_NONDETERMINISTIC)
_CONTEXTUAL_FUNCTIONS = frozenset(("SUBTOTAL(", "AGGREGATE("))

_VOLATILE_RE = re.compile(
    r"\b(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|CELL|INFO|INDIRECT|OFFSET)\s*\(",
    re.IGNORECASE)


def _quoted(title):
    from openpyxl.utils.cell import quote_sheetname

    return quote_sheetname(title)



class DependencySketch:
    """Coarse formula-dependency map: which cells feed which.

    ``references`` maps each formula cell (sheet-qualified A1) to the list
    of references its formula makes, as (sheet_title, bounds, raw) tuples —
    bounds may contain None for open-ended (whole-row/column) ranges.
    Structured references that cannot be resolved to cells are listed in
    ``unresolved`` (treated as always-intersecting).
    """

    def __init__(self):
        self.references = {}      # "Model!B6" -> [(sheet, bounds, raw)]
        self.unresolved = {}      # "Model!B6" -> [raw operand]
        self.volatile = set()     # formula addresses with known volatility
        self.contextual = set()   # formulas affected by display/filter state

    def cells_referencing(self, sheet_title, bounds):
        """Formula cells whose references intersect ``bounds`` on the given
        sheet — plus every cell with an unresolved (structured/table)
        reference, reported conservatively."""
        min_col, min_row, max_col, max_row = bounds
        title = sheet_title.casefold()   # Excel: sheet names case-insensitive
        hits = []
        for address, refs in self.references.items():
            for ref_sheet, ref_bounds, _raw in refs:
                if ref_sheet.casefold() != title:
                    continue
                if _intersects(ref_bounds, min_col, min_row, max_col, max_row):
                    hits.append(address)
                    break
        hits.extend(self.unresolved)
        return sorted(set(hits))

    def to_dict(self):
        return {
            "schema": "dependency_sketch",
            "version": 1,
            "references": {
                address: sorted(raw for (_s, _b, raw) in refs)
                for address, refs in sorted(self.references.items())
            },
            "unresolved": {address: sorted(raws) for address, raws
                           in sorted(self.unresolved.items())},
        }


def _intersects(bounds, min_col, min_row, max_col, max_row):
    b_min_col, b_min_row, b_max_col, b_max_row = bounds
    if b_min_col is None:
        b_min_col, b_max_col = 1, 1 << 20
    if b_min_row is None:
        b_min_row, b_max_row = 1, 1 << 22
    return not (b_max_col < min_col or b_min_col > max_col
                or b_max_row < min_row or b_min_row > max_row)


_SHEET_REF_RE = re.compile(r"^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$")


def dependency_sketch(wb):
    """Build a :class:`DependencySketch` from every formula in the model
    (tokenizer-based; — coarse is fine)."""
    from openpyxl.formula import Tokenizer

    sketch = DependencySketch()
    token_cache = {}
    for ws in wb.worksheets:
        for (row, col), cell in sorted(ws._cells.items()):
            if cell.data_type != "f":
                continue
            address = "{0}!{1}".format(_quoted(ws.title), cell.coordinate)
            formula = cell._value
            formula_ref = None
            if not isinstance(formula, str):
                formula_ref = getattr(formula, "ref", None)
                formula = getattr(formula, "text", None)
            if not isinstance(formula, str):
                ref = getattr(cell._value, "ref", None)
                sketch.unresolved.setdefault(address, []).append(
                    "{0}:{1}".format(
                        getattr(cell._value, "t", "formula-object"), ref))
                continue
            cached = token_cache.get(formula)
            if cached is None:
                try:
                    tokens = Tokenizer(formula).items
                except Exception:
                    sketch.unresolved.setdefault(address, []).append(formula)
                    continue
                operands = [t.value for t in tokens
                            if t.type == "OPERAND" and t.subtype == "RANGE"]
                # INDIRECT/OFFSET with computed-string targets leave no
                # RANGE operand at all: the formula must
                # count as unresolved (always-intersecting), never as
                # invisible
                functions = []
                for token in tokens:
                    if token.type != "FUNC" or token.subtype != "OPEN":
                        continue
                    name = token.value.upper()
                    if name.startswith("_XLFN."):
                        name = name[6:]
                    functions.append(name)
                indirect = any(
                    name in ("INDIRECT(", "OFFSET(")
                    for name in functions)
                volatile = any(
                    name in _VOLATILE_FUNCTIONS for name in functions)
                contextual = any(
                    name in _CONTEXTUAL_FUNCTIONS for name in functions)
                cached = (operands, indirect, volatile, contextual)
                token_cache[formula] = cached
            operands, indirect, volatile, contextual = cached
            if indirect:
                sketch.unresolved.setdefault(address, []).append(formula)
            if volatile:
                sketch.volatile.add(address)
            if contextual:
                sketch.contextual.add(address)
            for raw in operands:
                row_is_exact = formula_ref in (None, cell.coordinate)
                _classify(sketch, wb, ws, row, col, row_is_exact,
                          address, raw)
    return sketch


def _classify(sketch, wb, ws, row, col, row_is_exact, address, raw):
    ref = raw
    sheet_title = ws.title
    m = _SHEET_REF_RE.match(ref)
    if m:
        sheet_title = (m.group(1) or m.group(2))
        if m.group(1):
            sheet_title = sheet_title.replace("''", "'")
        ref = m.group(3)

    sheets_by_name = {sheet.title.casefold(): sheet.title
                      for sheet in wb.worksheets}
    canonical_title = sheets_by_name.get(sheet_title.casefold())
    if canonical_title is None:
        sketch.unresolved.setdefault(address, []).append(raw)
        return
    sheet_title = canonical_title

    if "[" in raw or "]" in raw:
        ranges = _structured_reference_ranges(
            wb, ws, row, col, row_is_exact, ref,
            sheet_title if m else None)
        # a resolved structured reference is always a non-empty range
        # list; anything else stays unresolved (always-intersecting)
        if ranges:
            for range_sheet, bounds in ranges:
                sketch.references.setdefault(address, []).append(
                    (range_sheet, bounds, raw))
            return
        # unsupported structured/table or external-workbook reference: not
        # resolvable
        sketch.unresolved.setdefault(address, []).append(raw)
        return
    if ":" in sheet_title:
        # a 3-D span (Sheet1:Sheet3!A1) is not one sheet: classify it
        # conservatively as unresolved (always-intersecting) rather than
        # recording a phantom sheet name nothing can ever match
        # (the phantom key silently defeated the recalc
        # guard and certification taint)
        sketch.unresolved.setdefault(address, []).append(raw)
        return

    plain = ref.replace("$", "")
    # a pure-alphabetic token without ':' is NEVER a cell/column reference
    # in a formula (column refs need "IN:IN"; cells need a row number) —
    # range_boundaries would happily parse "IN" as a column and hand the
    # taint walk phantom bounds (a defined name shaped like
    # a column letter escaped the input taint)
    if ":" not in plain and not any(ch.isdigit() for ch in plain):
        name = _defined_name(wb, ws, raw)
        if name is None:
            sketch.unresolved.setdefault(address, []).append(raw)
            return
        if name.value and "[" in name.value:
            sketch.unresolved.setdefault(address, []).append(raw)
            return
        try:
            destinations = list(name.destinations)
            if not destinations:
                raise ValueError("defined name has no static destinations")
            for dest_sheet, dest_ref in destinations:
                dest_bounds = range_boundaries(dest_ref.replace("$", ""))
                canonical = sheets_by_name.get(dest_sheet.casefold())
                if canonical is None:
                    raise ValueError("defined name targets a missing sheet")
                sketch.references.setdefault(address, []).append(
                    (canonical, dest_bounds, raw))
        except Exception:
            sketch.unresolved.setdefault(address, []).append(raw)
        return
    try:
        bounds = range_boundaries(plain)
    except Exception:
        # not A1-shaped: a defined name — expand via its destinations
        name = _defined_name(wb, ws, raw)
        if name is None:
            sketch.unresolved.setdefault(address, []).append(raw)
            return
        if name.value and "[" in name.value:
            # external-workbook reference hiding behind the name: the
            # expansion would drop the external marker
            sketch.unresolved.setdefault(address, []).append(raw)
            return
        try:
            destinations = list(name.destinations)
            if not destinations:
                raise ValueError("defined name has no static destinations")
            for dest_sheet, dest_ref in destinations:
                dest_bounds = range_boundaries(dest_ref.replace("$", ""))
                canonical = sheets_by_name.get(dest_sheet.casefold())
                if canonical is None:
                    raise ValueError("defined name targets a missing sheet")
                sketch.references.setdefault(address, []).append(
                    (canonical, dest_bounds, raw))
        except Exception:
            sketch.unresolved.setdefault(address, []).append(raw)
        return
    sketch.references.setdefault(address, []).append(
        (sheet_title, bounds, raw))


def _defined_name(wb, ws, raw):
    """Excel name lookup: case-insensitive and worksheet-local first."""
    folded = raw.casefold()
    for names in (ws.defined_names, wb.defined_names):
        for key, value in names.items():
            if key.casefold() == folded:
                return value
    return None


_STRUCTURED_ROW_SELECTORS = {
    "#all": "all",
    "#data": "data",
    "#headers": "headers",
    "#totals": "totals",
}


def _structured_reference_ranges(wb, ws, row, col, row_is_exact, raw,
                                 qualified_sheet):
    parsed = _parse_structured_reference(raw)
    if parsed is None:
        return None
    table_name, selector = parsed
    if table_name is None:
        if qualified_sheet is not None or not selector["current"]:
            return None
        table_ws, table = _same_table_for_current_row(ws, row, col)
    else:
        table_ws, table = _table_by_name(wb, table_name, qualified_sheet)
    if table is None:
        return None

    try:
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
    except Exception:
        return None
    layout = _table_layout(table, min_row, max_row)
    if layout is None:
        return None

    if selector["current"]:
        if not row_is_exact or layout["data"] is None \
                or not (layout["data"][0] <= row <= layout["data"][1]):
            return None
        row_bounds = (row, row)
    else:
        row_bounds = layout[selector["row"]]

    col_bounds = (min_col, max_col)
    if selector["columns"] is not None:
        columns = _table_column_map(table, table_ws, min_col, min_row,
                                    max_col)
        if columns is None:
            return None
        start_name, end_name = selector["columns"]
        start_index = columns.get(start_name.casefold())
        end_index = columns.get(end_name.casefold())
        if start_index is None or end_index is None:
            return None
        if start_index > end_index:
            return None
        col_bounds = (min_col + start_index, min_col + end_index)

    if row_bounds is None:
        # the selected region does not exist (headerless table, no totals
        # row, empty data region): Excel yields #REF! and the table shape
        # is degenerate, so take the conservative unresolved path rather
        # than recording no dependency at all (an operand recorded nowhere
        # is invisible to the move guards and to certification taint)
        return None

    return [(table_ws.title, (col_bounds[0], row_bounds[0],
                             col_bounds[1], row_bounds[1]))]


def _parse_structured_reference(raw):
    first = raw.find("[")
    if first == -1:
        return None
    if first == 0:
        table_name = None
    else:
        table_name = raw[:first]
        if not table_name:
            return None
    spec = raw[first:]
    if not (spec.startswith("[") and spec.endswith("]")):
        return None
    if spec.startswith("[["):
        if not spec.endswith("]]"):
            return None
        parts = _split_top_level(spec[1:-1], ",")
        if not parts:
            return None
    else:
        inner = spec[1:-1]
        if not inner or "[" in inner or "]" in inner:
            return None
        parts = [inner]

    row_selector = None
    columns = None
    current = False
    for part in parts:
        parsed = _parse_structured_part(part.strip())
        if parsed is None:
            return None
        kind, value = parsed
        if kind == "row":
            if row_selector is not None or current:
                return None
            row_selector = value
        elif kind == "current":
            if row_selector is not None or columns is not None or current:
                return None
            current = True
            columns = (value, value)
        else:
            if columns is not None:
                return None
            columns = value

    if current:
        row_kind = "data"
    else:
        row_kind = row_selector or "data"
    return table_name, {
        "row": row_kind,
        "columns": columns,
        "current": current,
    }


def _parse_structured_part(part):
    if not part:
        return None
    range_parts = _split_top_level(part, ":")
    if range_parts and len(range_parts) == 2:
        left = _simple_bracket_value(range_parts[0].strip())
        right = _simple_bracket_value(range_parts[1].strip())
        if not left or not right or left.startswith(("#", "@")) \
                or right.startswith(("#", "@")):
            return None
        return "columns", (left, right)
    if range_parts is None:
        return None

    value = _simple_bracket_value(part)
    if value is None:
        value = part
        if "[" in value or "]" in value:
            return None
    if value.startswith("#"):
        row = _STRUCTURED_ROW_SELECTORS.get(value.casefold())
        if row is None:
            return None
        return "row", row
    if value.startswith("@"):
        name = value[1:]
        if not name:
            return None
        return "current", name
    return "columns", (value, value)


def _simple_bracket_value(part):
    if not (part.startswith("[") and part.endswith("]")):
        return None
    value = part[1:-1]
    if not value or "[" in value or "]" in value:
        return None
    return value


def _split_top_level(value, separator):
    parts = []
    depth = 0
    start = 0
    for index, char in enumerate(value):
        if char == "[":
            depth += 1
        elif char == "]":
            depth -= 1
            if depth < 0:
                return None
        elif char == separator and depth == 0:
            parts.append(value[start:index])
            start = index + 1
    if depth != 0:
        return None
    parts.append(value[start:])
    return parts


def _table_by_name(wb, name, qualified_sheet):
    folded = name.casefold()
    matches = []
    for table_ws in wb.worksheets:
        if qualified_sheet is not None \
                and table_ws.title.casefold() != qualified_sheet.casefold():
            continue
        for table in table_ws.tables.values():
            names = (getattr(table, "displayName", None),
                     getattr(table, "name", None))
            if any(candidate and candidate.casefold() == folded
                   for candidate in names):
                matches.append((table_ws, table))
    if len(matches) != 1:
        return None, None
    return matches[0]


def _same_table_for_current_row(ws, row, col):
    matches = []
    for table in ws.tables.values():
        try:
            min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        except Exception:
            continue
        layout = _table_layout(table, min_row, max_row)
        if layout is None:
            continue
        if layout["data"] is None:
            continue
        data_min, data_max = layout["data"]
        if min_col <= col <= max_col and data_min <= row <= data_max:
            matches.append((ws, table))
    if len(matches) != 1:
        return None, None
    return matches[0]


def _table_layout(table, min_row, max_row):
    header_rows = table.headerRowCount
    if header_rows is None:
        header_rows = 1
    totals_rows = table.totalsRowCount or 0
    if getattr(table, "totalsRowShown", False) and not totals_rows:
        totals_rows = 1
    if header_rows < 0 or totals_rows < 0 \
            or min_row + header_rows + totals_rows - 1 > max_row:
        return None
    headers = None
    if header_rows:
        headers = (min_row, min_row + header_rows - 1)
    totals = None
    if totals_rows:
        totals = (max_row - totals_rows + 1, max_row)
    data = (min_row + header_rows, max_row - totals_rows)
    if data[0] > data[1]:
        data = None
    return {
        "all": (min_row, max_row),
        "data": data,
        "headers": headers,
        "totals": totals,
    }


def _table_column_map(table, ws, min_col, min_row, max_col):
    width = max_col - min_col + 1
    if len(table.tableColumns):
        names = [column.name for column in table.tableColumns]
        if len(names) != width:
            return None
    elif table.headerRowCount is None or table.headerRowCount:
        names = [ws.cell(min_row, column).value
                 for column in range(min_col, max_col + 1)]
    else:
        return None
    if any(not isinstance(name, str) or not name for name in names):
        return None
    folded = [name.casefold() for name in names]
    if len(set(folded)) != len(folded):
        return None
    return {name.casefold(): index for index, name in enumerate(names)}
