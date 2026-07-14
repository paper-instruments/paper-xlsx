# paper-xlsx: dependency sketch

"""Coarse formula-dependency analysis used by model maps and shift guards."""

import re

from openpyxl.utils.cell import range_boundaries

VOLATILE_NONDETERMINISTIC = ("NOW", "TODAY", "RAND", "RANDBETWEEN",
                             "RANDARRAY")

_VOLATILE_RE = re.compile(
    r"\b(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|INDIRECT|OFFSET)\s*\(",
    re.IGNORECASE)


def _quoted(title):
    from openpyxl.utils.cell import quote_sheetname

    return quote_sheetname(title)



class DependencySketch:
    """Coarse formula-dependency map: which cells feed which.

    ``references`` maps each formula cell (sheet-qualified A1) to the list
    of references its formula makes, as (sheet_title, bounds, raw) tuples —
    bounds may contain None for open-ended (whole-row/column) ranges.
    Table/structured references cannot be resolved to cells and are listed
    in ``unresolved`` (treated as always-intersecting).
    """

    def __init__(self):
        self.references = {}      # "Model!B6" -> [(sheet, bounds, raw)]
        self.unresolved = {}      # "Model!B6" -> [raw operand]

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
            if not isinstance(formula, str):
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
                indirect = any(
                    t.type == "FUNC" and t.subtype == "OPEN"
                    and t.value.upper().lstrip("_XLFN.")
                    in ("INDIRECT(", "OFFSET(")
                    for t in tokens)
                cached = (operands, indirect)
                token_cache[formula] = cached
            operands, indirect = cached
            if indirect:
                sketch.unresolved.setdefault(address, []).append(formula)
            for raw in operands:
                _classify(sketch, wb, ws, address, raw)
    return sketch


def _classify(sketch, wb, ws, address, raw):
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
        # structured/table or external-workbook reference: not resolvable
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
            for dest_sheet, dest_ref in name.destinations:
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
            for dest_sheet, dest_ref in name.destinations:
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
