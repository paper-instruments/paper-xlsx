# paper-xlsx: workbook diff report

"""Compare two workbook packages cell-wise, classifying differences as
content-changed vs shifted-by-structural-edit (via AddressRemap chains).
Measurements, never judgments."""

import io

from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple


class DiffReport:

    SCHEMA = "workbook_diff"
    VERSION = 1

    def __init__(self, changed, shifted, added_sheets, removed_sheets):
        self.changed = changed          # [{"address","before","after"}]
        self.shifted = shifted          # [{"from","to","value"}]
        self.added_sheets = added_sheets
        self.removed_sheets = removed_sheets

    def to_dict(self):
        return {
            "schema": self.SCHEMA,
            "version": self.VERSION,
            "changed": [dict(e) for e in self.changed],
            "shifted": [dict(e) for e in self.shifted],
            "added_sheets": list(self.added_sheets),
            "removed_sheets": list(self.removed_sheets),
        }

    def __repr__(self):
        return "DiffReport(changed={0}, shifted={1})".format(
            len(self.changed), len(self.shifted))


def _cells_of(wb):
    from openpyxl.package.cells import _formula_payload

    out = {}
    for ws in wb.worksheets:
        for (row, col), cell in ws._cells.items():
            if cell._value is None:
                continue
            value = (_formula_payload(cell._value)
                     if cell.data_type == "f" else cell._value)
            out[(ws.title, row, col)] = (value, cell.data_type)
    return out


def _remap_coordinate(title, row, col, remaps):
    """Push one A-side coordinate through the AddressRemap chain (in the
    order the edits were performed) to its B-side location, or None when
    a remap deleted it."""
    for remap in remaps:
        if getattr(remap, "sheet_title", None) is not None \
                and remap.sheet_title.casefold() != title.casefold():
            continue
        address = "{0}{1}".format(get_column_letter(col), row)
        mapped = remap.map(address)
        if mapped is None:
            return None
        row, col = coordinate_to_tuple(mapped.replace("$", ""))
    return (title, row, col)


def diff_workbooks(a, b, remaps=()):
    """A cell-level report of how package ``b`` differs from ``a``
    (paths, bytes, or file-likes). ``remaps``: AddressRemap chain from
    the structural edits performed between the two states — differences
    explained by a remap classify as "shifted", the rest as "changed"."""
    from openpyxl.reader.excel import load_workbook
    from openpyxl.preserve.limits import read_bounded
    from openpyxl.preserve.zipguard import validate_package_bytes

    def _load(source):
        payload = read_bounded(source, context="workbook diff input")
        validate_package_bytes(payload, context="workbook diff input")
        return load_workbook(io.BytesIO(payload))

    wb_a, wb_b = _load(a), _load(b)
    cells_a, cells_b = _cells_of(wb_a), _cells_of(wb_b)
    titles_a = set(wb_a.sheetnames)
    titles_b = set(wb_b.sheetnames)

    changed = []
    shifted = []
    consumed_b = set()
    for key in sorted(cells_a, key=lambda k: (k[0], k[1], k[2])):
        title, row, col = key
        if title not in titles_b:
            continue                    # covered by removed_sheets
        value_a, type_a = cells_a[key]
        target = _remap_coordinate(title, row, col, remaps) \
            if remaps else key
        address = "{0}!{1}{2}".format(title, get_column_letter(col), row)
        if target is None:
            changed.append({"address": address,
                            "before": _txt(value_a), "after": None,
                            "before_type": type_a, "after_type": None})
            continue
        value_b, type_b = cells_b.get(target, (None, None))
        if _same(value_a, value_b) and type_a == type_b:
            if target != key:
                t_title, t_row, t_col = target
                shifted.append({
                    "from": address,
                    "to": "{0}!{1}{2}".format(
                        t_title, get_column_letter(t_col), t_row),
                    "value": _txt(value_a)})
            consumed_b.add(target)
            continue
        changed.append({
            "address": address,
            "before": _txt(value_a),
            "after": _txt(value_b),
            "before_type": type_a,
            "after_type": type_b})
        consumed_b.add(target)
    for key in sorted(cells_b, key=lambda k: (k[0], k[1], k[2])):
        # skip only coordinates already CONSUMED as some A-cell's image:
        # under remaps an A-side coordinate may have been vacated and
        # REWRITTEN with new content, which must be reported
        if key in consumed_b:
            continue
        if not remaps and key in cells_a:
            continue
        title, row, col = key
        if title not in titles_a:
            continue
        changed.append({
            "address": "{0}!{1}{2}".format(title, get_column_letter(col),
                                           row),
            "before": None,
            "after": _txt(cells_b[key][0]),
            "before_type": None,
            "after_type": cells_b[key][1]})
    return DiffReport(changed, shifted,
                      sorted(titles_b - titles_a),
                      sorted(titles_a - titles_b))


def _same(a, b):
    # True == 1 in Python but not in a spreadsheet (t="b" vs numeric):
    # bool-aware comparison (1 -> TRUE reported unchanged)
    if isinstance(a, bool) != isinstance(b, bool):
        return False
    return a == b


def _txt(value):
    if value is None or isinstance(value, (str, int, float, bool,
                                            dict, list)):
        return value
    return str(value)
