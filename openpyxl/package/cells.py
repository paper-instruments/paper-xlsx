# paper-xlsx: cell-level semantic diff (PLAN Phase 4; PR-0 §5)

"""``diff_cells(a, b)``: which cells changed between two workbooks, as
(address, old/new value, old/new formula) — the same machinery family the
ledger cross-check uses, packaged for agents and tests.

SCOPE (by design, PLAN Phase 4): values and formulas only. Style-only
changes are invisible here — use ``diff_package`` (part-level, semantic)
to see styling churn."""

from openpyxl.utils.cell import quote_sheetname


class CellsDiff:

    SCHEMA = "cells_diff"
    VERSION = 1

    def __init__(self, changes, sheets_added, sheets_removed):
        self.changes = changes              # list of dicts, deterministic order
        self.sheets_added = sheets_added
        self.sheets_removed = sheets_removed

    @property
    def clean(self):
        return not (self.changes or self.sheets_added or self.sheets_removed)

    def to_dict(self):
        return {
            "schema": self.SCHEMA,
            "version": self.VERSION,
            "changes": list(self.changes),
            "sheets_added": list(self.sheets_added),
            "sheets_removed": list(self.sheets_removed),
        }

    def __repr__(self):
        return "CellsDiff({0} changes, +{1}/-{2} sheets)".format(
            len(self.changes), len(self.sheets_added),
            len(self.sheets_removed))


def _snapshot(source):
    """{sheet: {(row, col): (value, formula)}} for one package.

    Two loads, mirroring how the ecosystem reads workbooks: the formula view
    (data_only=False) and the cached-value view (data_only=True).
    """
    from openpyxl.reader.excel import load_workbook

    wb_formulas = load_workbook(source, data_only=False)
    if hasattr(source, "seek"):
        source.seek(0)
    wb_values = load_workbook(source, data_only=True)

    out = {}
    for ws in wb_formulas.worksheets:
        ws_values = wb_values[ws.title]
        cells = {}
        for (row, col), cell in ws._cells.items():
            formula = None
            value = cell._value
            if cell.data_type == "f":
                formula = value if isinstance(value, str) else str(value)
                vcell = ws_values._cells.get((row, col))
                value = vcell._value if vcell is not None else None
            if value is None and formula is None:
                # style-only cells are out of scope (see module docstring):
                # including them one-sidedly made a style-only cell compare
                # equal to an absent one anyway
                continue
            cells[(row, col)] = (value, formula)
        out[ws.title] = cells
    return out


def diff_cells(a, b):
    """Cell-level semantic diff of two packages (paths, bytes, or binary
    file-likes). Deterministic order: sheet, then row, then column."""
    snap_a = _snapshot(a)
    snap_b = _snapshot(b)

    sheets_added = sorted(set(snap_b) - set(snap_a))
    sheets_removed = sorted(set(snap_a) - set(snap_b))
    changes = []
    for title in sorted(set(snap_a) & set(snap_b)):
        cells_a = snap_a[title]
        cells_b = snap_b[title]
        for (row, col) in sorted(set(cells_a) | set(cells_b)):
            old_value, old_formula = cells_a.get((row, col), (None, None))
            new_value, new_formula = cells_b.get((row, col), (None, None))
            if old_value == new_value and old_formula == new_formula:
                continue
            from openpyxl.utils import get_column_letter

            changes.append({
                "address": "{0}!{1}{2}".format(
                    quote_sheetname(title), get_column_letter(col), row),
                "old_value": old_value,
                "new_value": new_value,
                "old_formula": old_formula,
                "new_formula": new_formula,
            })
    return CellsDiff(changes, sheets_added, sheets_removed)
