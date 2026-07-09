# paper-xlsx: the model map (PLAN-v0.1 Batch 6.2, PR-1 §5)

"""Classify every populated cell of formula-bearing sheets by its ROLE in
the model: inputs (no formula, referenced by formulas), calculations
(formula, referenced), outputs (formula, unreferenced), constants (no
formula, unreferenced). Measurements, never judgments — set_input()
(Batch 7) consumes this; nothing here decides anything.

Fill-color corroboration: when the sheet uses a consistent fill for its
classified inputs, the map records that convention (an agent can then
trust color as a secondary signal)."""

from openpyxl.utils import get_column_letter


class ModelMap:

    SCHEMA = "model_map"
    VERSION = 1

    def __init__(self, sheets, conventions):
        self.sheets = sheets            # title -> {role: [addresses]}
        self.conventions = conventions  # {"input_fill": rgb or None}

    def to_dict(self):
        return {
            "schema": self.SCHEMA,
            "version": self.VERSION,
            "sheets": {title: {role: list(addrs)
                               for role, addrs in roles.items()}
                       for title, roles in self.sheets.items()},
            "conventions": dict(self.conventions),
        }

    def inputs(self, title=None):
        """Flat input addresses (optionally for one sheet)."""
        out = []
        for sheet, roles in sorted(self.sheets.items()):
            if title is not None and sheet != title:
                continue
            out.extend(roles.get("inputs", []))
        return out

    def __repr__(self):
        counts = {role: sum(len(r.get(role, []))
                            for r in self.sheets.values())
                  for role in ("inputs", "calculations", "outputs",
                               "constants")}
        return "ModelMap({0})".format(counts)


def _referenced_coordinates(wb):
    """{(title, row, col)} referenced by any formula, via the dependency
    sketch (bounded ranges only; unresolved references make a cell's
    UNREFERENCED classification unreliable — recorded as a convention)."""
    from .perception import dependency_sketch

    sketch = dependency_sketch(wb)
    referenced = set()
    # cap expansion per reference so whole-column refs do not explode:
    # clamp to each sheet's populated extent
    extents = {}
    for ws in wb.worksheets:
        extents[ws.title] = (ws.max_row or 1, ws.max_column or 1)
    by_title = {ws.title: ws for ws in wb.worksheets}
    for _address, refs in sketch.references.items():
        for (title, bounds, _raw) in refs:
            ws = by_title.get(title)
            if ws is None:
                continue
            max_row, max_col = extents[title]
            min_c, min_r, max_c, max_r = bounds
            min_r = max(1, min_r or 1)
            min_c = max(1, min_c or 1)
            max_r = min(max_r or max_row, max_row)
            max_c = min(max_c or max_col, max_col)
            for r in range(min_r, max_r + 1):
                for c in range(min_c, max_c + 1):
                    referenced.add((title, r, c))
    return referenced, bool(sketch.unresolved)


def build_model_map(wb):
    referenced, has_unresolved = _referenced_coordinates(wb)
    sheets = {}
    input_fills = {}
    for ws in wb.worksheets:
        has_formulas = any(cell.data_type == "f"
                           for cell in ws._cells.values())
        if not has_formulas:
            continue
        roles = {"inputs": [], "calculations": [], "outputs": [],
                 "constants": []}
        for (row, col), cell in sorted(ws._cells.items()):
            if cell._value is None:
                continue
            address = "{0}{1}".format(get_column_letter(col), row)
            is_formula = cell.data_type == "f"
            is_referenced = (ws.title, row, col) in referenced
            if is_formula and is_referenced:
                roles["calculations"].append(address)
            elif is_formula:
                roles["outputs"].append(address)
            elif is_referenced:
                roles["inputs"].append(address)
                fill = getattr(cell, "fill", None)
                rgb = getattr(getattr(fill, "start_color", None), "rgb",
                              None)
                if isinstance(rgb, str):
                    input_fills[rgb] = input_fills.get(rgb, 0) + 1
            else:
                roles["constants"].append(address)
        sheets[ws.title] = roles

    conventions = {"input_fill": None,
                   "unresolved_references": has_unresolved}
    if input_fills:
        top, count = max(input_fills.items(), key=lambda kv: kv[1])
        total = sum(input_fills.values())
        if top != "00000000" and count >= 3 and count / total >= 0.8:
            conventions["input_fill"] = top
    return ModelMap(sheets, conventions)
