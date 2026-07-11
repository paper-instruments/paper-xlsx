# paper-xlsx: Excel-insert-semantics reference rewriting

"""Rewrite references for row/column inserts and deletes the way EXCEL
does — not the way fill/copy translation does.

Upstream's ``Translator`` implements FILL semantics: ``$``-anchored parts
are pinned and every reference shifts unconditionally. Excel's INSERT
semantics differ on every axis that matters: references at
or below the insertion point shift *including* ``$B$2``-style absolutes,
references above it stay, and ranges spanning the point EXPAND. Both
behaviors fall out of one rule: shift each endpoint independently when it
sits at/after the edit index. Deletes are the mirror image, with endpoints
inside the deleted zone clamped and fully-deleted references becoming
``#REF!`` — exactly what Excel writes.

Only the Tokenizer is reused (reference isolation); ``Translator`` stays
untouched — it is load-bearing for shared-formula expansion.
"""

import re

from openpyxl.utils import column_index_from_string, get_column_letter

_CELL_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$")
_COL_RANGE_RE = re.compile(r"^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})$")
_ROW_RANGE_RE = re.compile(r"^(\$?)([0-9]+):(\$?)([0-9]+)$")
# Unquoted sheet names are not ASCII-only. Excel permits Unicode letters and
# digits without quoting; exclude only syntax that belongs to quoted names or
# 3-D spans. Structured/external references are screened before this regex.
_SHEET_PREFIX_RE = re.compile(r"^(?:'((?:[^']|'')+)'|([^'!:]+))!(.+)$")

REF_ERROR = "#REF!"
EXCEL_MAX_ROW = 1048576
EXCEL_MAX_COL = 16384


def _checked_span(span, axis):
    if span is None:
        return None
    limit = EXCEL_MAX_ROW if axis == "rows" else EXCEL_MAX_COL
    if span[0] < 1 or span[1] > limit:
        from openpyxl.errors import BoundaryViolationError

        label = "row 1048576" if axis == "rows" else "column XFD"
        raise BoundaryViolationError(
            "the structural edit would rewrite a reference past Excel's "
            "hard {0} limit. Nothing was changed.".format(label))
    return span


def row_mapping(operation, index, amount):
    """old_row -> new_row, or None when the row is deleted."""
    if operation.startswith("insert"):
        def mapper(row):
            return row + amount if row >= index else row
    else:
        def mapper(row):
            if index <= row < index + amount:
                return None
            return row - amount if row >= index + amount else row
    return mapper


def _shift_point(value, index, amount, is_delete, is_start):
    """One endpoint, per Excel semantics; None means the whole reference
    collapses to #REF! (single-point case handled by the caller)."""
    if not is_delete:
        return value + amount if value >= index else value
    if value >= index + amount:
        return value - amount
    if value >= index:
        # endpoint inside the deleted zone: clamp toward the survivor
        return index if is_start else index - 1
    return value


def _shift_span(start, end, index, amount, is_delete):
    """(new_start, new_end) or None for #REF! (span fully deleted)."""
    if is_delete and start >= index and end < index + amount:
        return None
    new_start = _shift_point(start, index, amount, is_delete, True)
    new_end = _shift_point(end, index, amount, is_delete, False)
    if new_start > new_end:
        return None
    return new_start, new_end


def _shift_reference_span(first, second, axis, index, amount, is_delete):
    """Shift a possibly reversed formula range without losing orientation."""
    reversed_order = first > second
    span = _checked_span(
        _shift_span(min(first, second), max(first, second),
                    index, amount, is_delete), axis)
    if span is None:
        return None
    return (span[1], span[0]) if reversed_order else span


def shift_ref(ref, axis, index, amount, is_delete):
    """Shift one bare A1 reference (no sheet prefix). Returns the new text,
    ``ref`` unchanged when unaffected, or ``#REF!``. ``$`` markers are kept
    positionally — insert/delete moves absolutes too (Excel semantics)."""
    m = _CELL_RE.match(ref)
    if m:
        cd, col, rd, row = (m.group(1), m.group(2).upper(), m.group(3),
                            int(m.group(4)))
        if axis == "rows":
            span = _checked_span(
                _shift_span(row, row, index, amount, is_delete), axis)
            if span is None:
                return REF_ERROR
            return "{0}{1}{2}{3}".format(cd, col, rd, span[0])
        col_idx = column_index_from_string(col)
        span = _checked_span(
            _shift_span(col_idx, col_idx, index, amount, is_delete), axis)
        if span is None:
            return REF_ERROR
        return "{0}{1}{2}{3}".format(cd, get_column_letter(span[0]), rd, row)

    if ":" in ref and _CELL_RE.match(ref.split(":", 1)[0]) \
            and _CELL_RE.match(ref.split(":", 1)[1]):
        first, second = ref.split(":", 1)
        m1, m2 = _CELL_RE.match(first), _CELL_RE.match(second)
        c1d, c1, r1d, r1 = (m1.group(1), m1.group(2).upper(), m1.group(3),
                            int(m1.group(4)))
        c2d, c2, r2d, r2 = (m2.group(1), m2.group(2).upper(), m2.group(3),
                            int(m2.group(4)))
        if axis == "rows":
            span = _shift_reference_span(
                r1, r2, axis, index, amount, is_delete)
            if span is None:
                return REF_ERROR
            return "{0}{1}{2}{3}:{4}{5}{6}{7}".format(
                c1d, c1, r1d, span[0], c2d, c2, r2d, span[1])
        i1, i2 = column_index_from_string(c1), column_index_from_string(c2)
        span = _shift_reference_span(
            i1, i2, axis, index, amount, is_delete)
        if span is None:
            return REF_ERROR
        return "{0}{1}{2}{3}:{4}{5}{6}{7}".format(
            c1d, get_column_letter(span[0]), r1d, r1,
            c2d, get_column_letter(span[1]), r2d, r2)

    m = _ROW_RANGE_RE.match(ref)
    if m and axis == "rows":
        r1, r2 = int(m.group(2)), int(m.group(4))
        span = _shift_reference_span(
            r1, r2, axis, index, amount, is_delete)
        if span is None:
            return REF_ERROR
        return "{0}{1}:{2}{3}".format(m.group(1), span[0], m.group(3), span[1])

    m = _COL_RANGE_RE.match(ref)
    if m and axis == "cols":
        i1 = column_index_from_string(m.group(2).upper())
        i2 = column_index_from_string(m.group(4).upper())
        span = _shift_reference_span(
            i1, i2, axis, index, amount, is_delete)
        if span is None:
            return REF_ERROR
        return "{0}{1}:{2}{3}".format(
            m.group(1), get_column_letter(span[0]),
            m.group(3), get_column_letter(span[1]))

    # whole-column refs under a row shift (and vice versa) are unaffected;
    # anything else (defined names, structured refs) is not ours to touch
    return ref


def _quote_if_needed(title):
    if re.match(r"^[A-Za-z_][A-Za-z0-9_.]*$", title):
        return title
    return "'{0}'".format(title.replace("'", "''"))


def shift_formula(formula, context_sheet, target_sheet, axis, index, amount,
                  is_delete):
    """Rewrite one formula for a shift on ``target_sheet``.

    ``context_sheet`` is the sheet the formula lives on (unprefixed
    references resolve to it). Returns (new_formula, changed).
    """
    from openpyxl.formula import Tokenizer

    if not formula.startswith("="):
        return formula, False
    try:
        tok = Tokenizer(formula)
    except Exception as exc:
        from openpyxl.errors import UnsupportedStructureError

        raise UnsupportedStructureError(
            "cannot safely rewrite formula {0!r} on sheet {1!r}: the "
            "formula tokenizer rejected it ({2}). Nothing was changed."
            .format(formula, context_sheet, exc),
            kind="unparseable-structural-formula",
            anchor=context_sheet,
        ) from exc

    changed = False
    for token in tok.items:
        if token.type != "OPERAND" or token.subtype != "RANGE":
            continue
        raw = token.value
        if "[" in raw:
            continue           # structured/external: refused upstream
        sheet = context_sheet
        ref = raw
        prefix = ""
        m = _SHEET_PREFIX_RE.match(raw)
        if m:
            sheet = m.group(1).replace("''", "'") if m.group(1) else m.group(2)
            ref = m.group(3)
            prefix = raw[:len(raw) - len(ref)]
        if sheet is None or sheet.casefold() != target_sheet.casefold():
            # Excel resolves sheet names case-insensitively; an unprefixed
            # operand with no context sheet (name values) is never ours
            continue
        new_ref = shift_ref(ref, axis, index, amount, is_delete)
        if new_ref == ref:
            continue
        changed = True
        if new_ref == REF_ERROR:
            token.value = REF_ERROR   # Excel drops the sheet prefix too
        else:
            token.value = prefix + new_ref
    if not changed:
        return formula, False
    return tok.render(), True


def shift_name_value(value, target_sheet, axis, index, amount, is_delete):
    """Defined-name / print-area values are formula fragments with explicit
    sheet prefixes; rewrite them with the same machinery."""
    new_formula, changed = shift_formula(
        "=" + value, None, target_sheet, axis, index, amount, is_delete)
    return (new_formula[1:], True) if changed else (value, False)


def shift_formula_fragment(value, context_sheet, target_sheet, axis, index,
                           amount, is_delete):
    """Rewrite a CF/DV formula, which may legally omit the leading ``=``."""
    if not isinstance(value, str) or not value:
        return value, False
    explicit = value.startswith("=")
    formula = value if explicit else "=" + value
    rewritten, changed = shift_formula(
        formula, context_sheet, target_sheet, axis, index, amount, is_delete)
    if not changed:
        return value, False
    return rewritten if explicit else rewritten[1:], True


def shift_cell_range(cell_range, axis, index, amount, is_delete):
    """Shift a CellRange in place. Returns 'changed', 'unchanged' or
    'deleted' (range fully inside a deleted zone — the caller removes it)."""
    if axis == "rows":
        span = _shift_span(cell_range.min_row, cell_range.max_row,
                           index, amount, is_delete)
        if span is None:
            return "deleted"
        if (span[0], span[1]) == (cell_range.min_row, cell_range.max_row):
            return "unchanged"
        cell_range.min_row, cell_range.max_row = span
        return "changed"
    span = _shift_span(cell_range.min_col, cell_range.max_col,
                       index, amount, is_delete)
    if span is None:
        return "deleted"
    if (span[0], span[1]) == (cell_range.min_col, cell_range.max_col):
        return "unchanged"
    cell_range.min_col, cell_range.max_col = span
    return "changed"


# rename accepts the 3-D span form (Sheet1:Sheet3!) the shift regex
# deliberately excludes
_RENAME_PREFIX_RE = re.compile(r"^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$")


def rename_sheets_in_formula(formula, mapping):
    """Simultaneous multi-title rewrite: every sheet component maps
    through ``mapping`` (casefold keys resolved per component) exactly
    once — a swap can never cascade."""
    folded = {k.casefold(): v for k, v in mapping.items()}

    from openpyxl.formula import Tokenizer
    from openpyxl.utils.cell import quote_sheetname

    if not formula.startswith("="):
        return formula, False
    try:
        tok = Tokenizer(formula)
    except Exception:
        return formula, False
    changed = False
    for token in tok.items:
        if token.type != "OPERAND" or token.subtype != "RANGE":
            continue
        raw = token.value
        if "[" in raw:
            continue
        m = _RENAME_PREFIX_RE.match(raw)
        if not m:
            continue
        sheet = m.group(1).replace("''", "'") if m.group(1) else m.group(2)
        ref = m.group(3)
        parts = sheet.split(":") if ":" in sheet else [sheet]
        new_parts = [folded.get(p.casefold(), p) for p in parts]
        if new_parts == parts:
            continue
        changed = True
        if len(new_parts) == 1:
            token.value = "{0}!{1}".format(quote_sheetname(new_parts[0]),
                                           ref)
        else:
            token.value = "{0}!{1}".format(
                quote_sheetname(":".join(new_parts)), ref)
    if not changed:
        return formula, False
    return tok.render(), True


def rename_sheet_in_formula(formula, old_title, new_title):
    """Rewrite sheet-prefixed references from ``old_title`` to
    ``new_title`` (case-insensitive, quote-aware, 3-D span endpoints
    included). Returns (new_formula, changed)."""
    from openpyxl.formula import Tokenizer
    from openpyxl.utils.cell import quote_sheetname

    if not formula.startswith("="):
        return formula, False
    try:
        tok = Tokenizer(formula)
    except Exception:
        return formula, False

    folded = old_title.casefold()
    changed = False
    for token in tok.items:
        if token.type != "OPERAND" or token.subtype != "RANGE":
            continue
        raw = token.value
        if "[" in raw:
            continue
        m = _RENAME_PREFIX_RE.match(raw)
        if not m:
            continue
        sheet = m.group(1).replace("''", "'") if m.group(1) else m.group(2)
        ref = m.group(3)
        # 3-D spans: 'Sheet1:Sheet3' — QUOTED spans ('My Data:Sheet3')
        # quote the whole span, so split unconditionally
        parts = sheet.split(":") if ":" in sheet else [sheet]
        new_parts = [new_title if p.casefold() == folded else p
                     for p in parts]
        if new_parts == parts:
            continue
        changed = True
        rebuilt = ":".join(new_parts)
        if len(new_parts) == 1:
            token.value = "{0}!{1}".format(quote_sheetname(rebuilt), ref)
        else:
            token.value = "{0}!{1}".format(quote_sheetname(rebuilt), ref)
    if not changed:
        return formula, False
    return tok.render(), True


def rename_sheets_in_formula_fragment(value, mapping):
    """Rename sheet references in a CF/DV formula with optional ``=``."""
    if not isinstance(value, str) or not value:
        return value, False
    explicit = value.startswith("=")
    formula = value if explicit else "=" + value
    rewritten, changed = rename_sheets_in_formula(formula, mapping)
    if not changed:
        return value, False
    return rewritten if explicit else rewritten[1:], True


def rename_sheet_in_formula_fragment(value, old_title, new_title):
    return rename_sheets_in_formula_fragment(value, {old_title: new_title})


def title_in_string_literals(formula, title):
    """True when a formula's STRING literals mention ``title`` — the
    textual (INDIRECT-style) references a rename cannot rewrite."""
    from openpyxl.formula import Tokenizer

    try:
        tokens = Tokenizer(formula).items
    except Exception:
        return title.casefold() in formula.casefold()
    folded = title.casefold()
    for token in tokens:
        if token.type == "OPERAND" and token.subtype == "TEXT" \
                and folded in token.value.casefold():
            return True
    return False
