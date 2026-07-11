# paper-xlsx: table support under preserve

"""Loaded-table mutation, and the table row discipline.

Table parts are fully modeled upstream (Table is Serialisable end to end),
so a mutated table re-renders whole from the model — located by
displayName in the ORIGINAL sheet rels, never by guessed part numbering.
Guards refuse geometry the discipline cannot keep coherent: the header row
moves, the data region vanishes, or the column count disagrees with
tableColumns.
"""

import re

from openpyxl.errors import UnsupportedStructureError
from openpyxl.utils.cell import range_boundaries
from openpyxl.xml.functions import tostring

from . import crosspart

_TABLE_REL_TYPE_SUFFIX = "/table"
_DISPLAY_NAME_RE = re.compile(
    br'displayName=(?:"([^"]*)"|\'([^\']*)\')')
_REF_RE = re.compile(br'<table[^>]*\sref=(?:"([^"]*)"|\'([^\']*)\')')


def _refuse(msg):
    raise UnsupportedStructureError(msg + " Nothing was written.")


def sheet_table_parts(zin, sheet_part):
    """{displayName: (part_name, original_bytes)} for one sheet, resolved
    through the ORIGINAL rels (producers number table parts arbitrarily —
    Table.path's id-derived guess is not trustworthy)."""
    rels_part = _rels_path(sheet_part)
    names = set(zin.namelist())
    if rels_part not in names:
        return {}
    out = {}
    root = crosspart.scan_small(zin.read(rels_part), "Relationships",
                                max_depth=1)
    for child in root.children:
        if child.local() != "Relationship":
            continue
        if not child.attrs.get("Type", "").endswith(_TABLE_REL_TYPE_SUFFIX):
            continue
        target = child.attrs.get("Target", "")
        part = _resolve_target(sheet_part, target)
        if part not in names:
            continue
        payload = zin.read(part)
        m = _DISPLAY_NAME_RE.search(payload)
        if m:
            raw = m.group(1) if m.group(1) is not None else m.group(2)
            out[_unescape(raw.decode("utf-8"))] = (part, payload)
    return out


def validate_table(tbl, original_ref):
    """Geometry guards, against the ORIGINAL ref."""
    try:
        min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
    except Exception:
        _refuse("table {0!r}: ref {1!r} is not a rectangular "
                "range.".format(tbl.displayName, tbl.ref))
    o_min_col, o_min_row, _oc, _or = range_boundaries(original_ref)
    header = tbl.headerRowCount if tbl.headerRowCount is not None else 1
    totals = tbl.totalsRowCount or 0
    if min_row != o_min_row or min_col != o_min_col:
        _refuse("table {0!r}: the resize moved the table's anchor "
                "({1} -> {2}); the header row must stay fixed (resize "
                "downward/rightward only).".format(
                    tbl.displayName, original_ref, tbl.ref))
    if max_row - min_row + 1 < header + totals + 1:
        _refuse("table {0!r}: ref {1!r} leaves no data row (header={2}, "
                "totals={3}).".format(tbl.displayName, tbl.ref, header,
                                      totals))
    n_cols = max_col - min_col + 1
    if len(tbl.tableColumns) and n_cols != len(tbl.tableColumns):
        _refuse("table {0!r}: ref {1!r} spans {2} columns but the table "
                "defines {3} tableColumns; add or remove the columns "
                "explicitly.".format(tbl.displayName, tbl.ref, n_cols,
                                     len(tbl.tableColumns)))
    if tbl.autoFilter is not None and tbl.autoFilter.ref:
        try:
            a_min_col, a_min_row, a_max_col, a_max_row = \
                range_boundaries(tbl.autoFilter.ref)
        except Exception:
            _refuse("table {0!r}: autoFilter ref {1!r} is not a "
                    "range.".format(tbl.displayName, tbl.autoFilter.ref))
        if (a_min_col < min_col or a_max_col > max_col
                or a_min_row < min_row or a_max_row > max_row):
            _refuse("table {0!r}: autoFilter ref {1!r} lies outside the "
                    "table ref {2!r}; sync it (the append_row verb does "
                    "this automatically).".format(
                        tbl.displayName, tbl.autoFilter.ref, tbl.ref))


def plan_table_mutations(wb, ws, sheet_part, zin, changed_names, plan):
    """Re-render each changed loaded table into its ORIGINAL part."""
    parts = sheet_table_parts(zin, sheet_part)
    for name in changed_names:
        if name not in ws.tables:
            # removal is the lifecycle path (planned separately)
            continue
        tbl = ws.tables[name]
        if name not in parts:
            _refuse("table {0!r} on sheet {1!r} has no resolvable part in "
                    "the original package (displayName not found in the "
                    "sheet rels).".format(name, ws.title))
        part_name, original = parts[name]
        if b"<extLst" in original or b"xr:uid" in original \
                or b"xmlns:xr" in original:
            _refuse("table {0!r} on sheet {1!r} carries extension content "
                    "(extLst / xr revision ids) the model cannot "
                    "re-serialize; editing it would silently drop that "
                    "content (e.g. alt text). Recreate the table or edit "
                    "without preserve=True.".format(name, ws.title))
        m = _REF_RE.search(original)
        if m:
            raw = m.group(1) if m.group(1) is not None else m.group(2)
            original_ref = raw.decode("ascii")
        else:
            # a table part whose ref we cannot locate cannot be guard-
            # checked: refuse rather than silently disabling the anchor
            # guard (single-quoted ref no-op'd it)
            _refuse("table {0!r}: the original part's ref attribute could "
                    "not be located; the geometry guards cannot "
                    "run.".format(name))
        validate_table(tbl, original_ref)
        _check_display_name(wb, ws, tbl, original_names=set(parts))
        payload = tostring(tbl.to_tree())
        if not payload.startswith(b"<?xml"):
            payload = (b'<?xml version="1.0" encoding="UTF-8" '
                       b'standalone="yes"?>\n' + payload)
        plan[part_name] = payload


_XML_UNESCAPES = (("&lt;", "<"), ("&gt;", ">"), ("&quot;", '"'),
                  ("&apos;", "'"), ("&amp;", "&"))


def _unescape(text):
    for entity, char in _XML_UNESCAPES:
        text = text.replace(entity, char)
    return text


def _check_display_name(wb, ws, tbl, original_names):
    """Table displayNames are workbook-unique and share a namespace with
    defined names (case-insensitive, Excel semantics)."""
    name = tbl.displayName
    folded = name.casefold()
    for other in wb.defined_names:
        if other.casefold() == folded:
            _refuse("table {0!r} collides with the defined name {1!r} "
                    "(Excel treats table and defined names as one "
                    "case-insensitive namespace).".format(name, other))
    for sheet in wb.worksheets:
        for other_name in getattr(sheet, "tables", {}):
            if sheet is ws and other_name == name:
                continue
            if other_name.casefold() == folded:
                _refuse("table {0!r} collides with table {1!r} on sheet "
                        "{2!r}; displayNames are workbook-"
                    "unique.".format(name, other_name, sheet.title))


def _validate_formula_grid(formula, table_name):
    """Refuse formula references outside Excel's physical grid."""
    from openpyxl.formula import Tokenizer
    from openpyxl.utils.cell import range_boundaries
    from openpyxl.errors import BoundaryViolationError

    try:
        tokens = Tokenizer(formula).items
    except Exception:
        return
    for token in tokens:
        if token.type != "OPERAND" or token.subtype != "RANGE" \
                or "[" in token.value:
            continue
        ref = token.value.rsplit("!", 1)[-1].replace("$", "")
        if ":" not in ref and not any(char.isdigit() for char in ref):
            continue
        try:
            bounds = range_boundaries(ref)
        except ValueError:
            continue
        cols = [value for value in (bounds[0], bounds[2])
                if value is not None]
        rows = [value for value in (bounds[1], bounds[3])
                if value is not None]
        if any(value > 16384 for value in cols) or any(
                value > 1048576 for value in rows):
            raise BoundaryViolationError(
                "append_row would generate formula {0!r} outside Excel's "
                "row/column limits for table {1!r}. Nothing was changed."
                .format(formula, table_name))


def _normalize_append_value(ws, row, col, value, table_name):
    """Validate one planned value without materializing its destination."""
    from openpyxl.cell.cell import (
        CellRichText,
        ILLEGAL_CHARACTERS_RE,
        _TYPES,
        get_type,
    )
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.exceptions import IllegalCharacterError

    if value is None:
        return None
    value_type = type(value)
    data_type = _TYPES.get(value_type)
    if data_type is None:
        data_type = get_type(value_type, value)
    if data_type is None:
        raise ValueError("Cannot convert {0!r} to Excel".format(value))

    if data_type == "s" and not isinstance(value, CellRichText):
        if not isinstance(value, str):
            value = str(value, ws.parent.encoding)
        value = str(value)[:32767]
        if ILLEGAL_CHARACTERS_RE.search(value):
            raise IllegalCharacterError(
                "{0} cannot be used in worksheets.".format(value))
    formula = value if isinstance(value, str) and value.startswith("=") \
        else getattr(value, "text", None)
    if isinstance(formula, str) and formula.startswith("="):
        _validate_formula_grid(formula, table_name)
        if getattr(ws.parent, "formula_lint", "warn") == "refuse":
            from openpyxl.formula.lint import lint_formula
            findings = lint_formula(formula, workbook=ws.parent, sheet=ws)
            if findings:
                from openpyxl.errors import UnsupportedStructureError

                summary = "; ".join(
                    "[{0}] {1}".format(item["code"], item["message"])
                    for item in findings[:6])
                raise UnsupportedStructureError(
                    "formula planned for {0}!{1}{2} failed the pre-flight "
                    "lint and wb.formula_lint is 'refuse': {3}. Nothing "
                    "was changed.".format(
                        ws.title, get_column_letter(col), row, summary))
    return value


def _preflight_append_protection(ws, coordinates):
    """Strict protection refusal before any destination cell is created."""
    if not bool(ws.protection.sheet) or not getattr(
            ws.parent, "strict_protection", False):
        return
    from openpyxl.errors import UnsupportedStructureError
    from openpyxl.utils import get_column_letter

    for row, col in sorted(set(coordinates)):
        cell = ws._cells.get((row, col))
        if cell is None or cell.protection.locked:
            raise UnsupportedStructureError(
                "append_row would write locked cell {0}{1} on protected "
                "sheet {2!r}, and strict_protection is enabled. Unlock the "
                "target cells or unprotect the sheet. Nothing was changed."
                .format(get_column_letter(col), row, ws.title))


def append_row(ws, table_name, values):
    """Append one row of ``values`` below the table's last data row: writes
    the cells, extends ``tbl.ref``,
    keeps the totals row last (its cells move down one row), re-derives
    calculated-column formulas, and syncs the table's autoFilter.

    ``values``: list (positional per column) or dict keyed by column name.
    Refuses: content below the table (moving it needs the structural
    machinery — restructure or use stock mode); calculated-column values
    that disagree with the column formula; non-preserve workbooks are
    fine too (the model edit works the same everywhere).
    """
    from openpyxl.formula.translate import Translator
    from openpyxl.utils import get_column_letter

    if table_name not in ws.tables:
        from openpyxl.errors import TargetNotFoundError

        raise TargetNotFoundError(
            "no table named {0!r} on sheet {1!r}.".format(
                table_name, ws.title))
    tbl = ws.tables[table_name]
    min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
    totals = tbl.totalsRowCount or 0
    new_data_row = max_row - totals + 1        # where the new data lands
    n_cols = max_col - min_col + 1

    led = getattr(ws.parent, "_paper_ledger", None)
    original_comments = led.comment_snapshots.get(ws, {}) if led else {}
    original_links = (led.region_snapshots.get(ws, {}).get("hyperlinks", {})
                      if led else {})
    if totals and any((max_row, col) in original_comments
                      for col in range(min_col, max_col + 1)):
        _refuse(
            "append_row: table {0!r} has a totals-row comment from the "
            "original package; preserve mode cannot rewrite its existing "
            "comment/VML anchor yet. Remove or relocate the comment in "
            "Excel before appending".format(table_name))
    if totals and any((max_row, col) in original_links
                      for col in range(min_col, max_col + 1)):
        _refuse(
            "append_row: table {0!r} has a totals-row hyperlink from the "
            "original package; preserve mode cannot move an existing "
            "hyperlink relationship. Remove or relocate the hyperlink in "
            "Excel before appending".format(table_name))

    if max_row >= 1048576:
        from openpyxl.errors import BoundaryViolationError

        raise BoundaryViolationError(
            "append_row would extend table {0!r} past row 1048576, Excel's "
            "hard sheet limit. Nothing was changed.".format(table_name))

    # normalize values
    if isinstance(values, dict):
        by_name = dict(values)
        cols = [c.name for c in tbl.tableColumns]
        unknown = set(by_name) - set(cols)
        if unknown:
            _refuse("append_row: unknown column(s) {0} for table "
                    "{1!r}.".format(sorted(unknown), table_name))
        row_values = [by_name.get(c) for c in cols]
    else:
        row_values = list(values)
        if len(row_values) > n_cols:
            _refuse("append_row: {0} values for a {1}-column "
                    "table.".format(len(row_values), n_cols))
        row_values += [None] * (n_cols - len(row_values))

    # content below the table cannot be shifted here (tables are shift
    # blockers) — refuse loudly
    below = max_row + 1
    for (r, c), cell in ws._cells.items():
        if r >= below and min_col <= c <= max_col \
                and (cell._value is not None or cell.has_style
                     or cell._comment is not None
                     or cell._hyperlink is not None):
            _refuse("append_row: sheet {0!r} has content at or below row "
                    "{1} under table {2!r}; appending would need to shift "
                    "it. Move that content, or restructure the "
                    "edit.".format(ws.title, below, table_name))

    # validate EVERY column before any mutation (a late
    # calc-column refusal left the totals row moved and half a data row
    # written — "Nothing was written" must be true)
    for i in range(n_cols):
        tc = tbl.tableColumns[i] if i < len(tbl.tableColumns) else None
        calc = getattr(tc, "calculatedColumnFormula", None) if tc else None
        if calc is not None and getattr(calc, "attr_text", None) \
                and row_values[i] is not None:
            _refuse("append_row: column {0!r} is a calculated column; "
                    "its value derives from the column formula "
                    "(={1}).".format(tc.name, calc.attr_text))

    planned_values = []
    try:
        from openpyxl.formula.translate import TranslatorError

        for i, col in enumerate(range(min_col, max_col + 1)):
            tc = tbl.tableColumns[i] if i < len(tbl.tableColumns) else None
            calc = getattr(tc, "calculatedColumnFormula", None) if tc else None
            given = row_values[i]
            planned = given
            if calc is not None and getattr(calc, "attr_text", None):
                planned = "=" + calc.attr_text
            elif calc is None and given is None \
                    and new_data_row - 1 > min_row:
                above = ws._cells.get((new_data_row - 1, col))
                if above is not None and above.data_type == "f" \
                        and isinstance(above.value, str):
                    planned = Translator(
                        above.value,
                        origin="{0}{1}".format(
                            get_column_letter(col), new_data_row - 1)
                    ).translate_formula(
                        "{0}{1}".format(
                            get_column_letter(col), new_data_row))
            planned_values.append(_normalize_append_value(
                ws, new_data_row, col, planned, table_name))
    except TranslatorError as exc:
        from openpyxl.errors import BoundaryViolationError

        raise BoundaryViolationError(
            "append_row cannot translate an inherited formula within "
            "Excel's grid for table {0!r}. Nothing was changed."
            .format(table_name)) from exc

    affected = [(new_data_row, col)
                for col in range(min_col, max_col + 1)]
    planned_totals = {}
    if totals:
        for col in range(min_col, max_col + 1):
            source_value = ws._cells.get((max_row, col))
            planned_totals[col] = _normalize_append_value(
                ws, max_row + 1, col,
                source_value.value if source_value is not None else None,
                table_name)
        affected.extend(
            (row, col)
            for row in (max_row, max_row + 1)
            for col in range(min_col, max_col + 1))
    _preflight_append_protection(ws, affected)

    # totals row moves down one: rewrite its cells at +1 first
    if totals:
        from copy import copy

        for col in range(min_col, max_col + 1):
            src = ws.cell(row=max_row, column=col)
            dst = ws.cell(row=max_row + 1, column=col)
            dst.value = planned_totals[col]
            dst._style = copy(src._style) if src.has_style else None
            dst.comment = src.comment
            if src.hyperlink is not None:
                dst.hyperlink = copy(src.hyperlink)
            else:
                dst.hyperlink = None
            src.value = None
            src.comment = None
            src.hyperlink = None
            # the freed slot becomes a DATA row: style it like the row
            # above, not like the totals row it used to be
            model = ws.cell(row=max_row - 1, column=col)
            src._style = model._style if model.has_style else None

    # write the new data row (calculated columns re-derive)
    for col, planned in zip(range(min_col, max_col + 1), planned_values):
        ws.cell(row=new_data_row, column=col).value = planned

    # extend the ref; keep the autoFilter over the header+data region
    new_max_row = max_row + 1
    new_ref = "{0}{1}:{2}{3}".format(
        get_column_letter(min_col), min_row,
        get_column_letter(max_col), new_max_row)
    tbl.ref = new_ref
    if tbl.autoFilter is not None:
        tbl.autoFilter.ref = "{0}{1}:{2}{3}".format(
            get_column_letter(min_col), min_row,
            get_column_letter(max_col), new_max_row - totals)
    return None


def _rels_path(part_name):
    folder, _, base = part_name.rpartition("/")
    return "{0}/_rels/{1}.rels".format(folder, base) if folder \
        else "_rels/{0}.rels".format(base)


def _resolve_target(from_part, target):
    """Resolve an OPC relative target against the source part's folder."""
    if target.startswith("/"):
        return target[1:]
    base = from_part.rpartition("/")[0].split("/") if "/" in from_part \
        else []
    for piece in target.split("/"):
        if piece == "..":
            base = base[:-1]
        elif piece != ".":
            base.append(piece)
    return "/".join(base)


TABLE_CONTENT_TYPE = ("application/vnd.openxmlformats-officedocument."
                      "spreadsheetml.table+xml")
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def plan_table_lifecycle(wb, ws, sheet_part, zin, armed_names, plan,
                         part_plan, names):
    """Plan table ADD/REMOVE for one sheet: parts via the engine, the
    sheet's tableParts element rebuilt as crafted bytes (returned; the
    caller rides them through the region splice). ``armed_names`` is the
    arm-time tuple of table names."""
    current = set(ws.tables.keys())
    armed = set(armed_names)
    added_names = sorted(current - armed)
    removed_names = sorted(armed - current)

    original_parts = sheet_table_parts(zin, sheet_part)
    rels_part = _rels_path(sheet_part)

    # rId bookkeeping for surviving originals: target -> rId
    rid_by_part = {}
    if rels_part in names:
        root = crosspart.scan_small(zin.read(rels_part), "Relationships",
                                    max_depth=1)
        for child in root.children:
            if child.local() != "Relationship":
                continue
            target = _resolve_target(sheet_part,
                                     child.attrs.get("Target", ""))
            rid_by_part[target] = child.attrs.get("Id", "")

    # hyperlink additions allocate rIds on the same rels part through a
    # separate planner: refusing the combination keeps both allocators
    # deterministic
    from .regions import hyperlink_signatures

    led = wb._paper_ledger
    armed_links = led.region_snapshots.get(ws, {}).get("hyperlinks", {})
    if added_names and hyperlink_signatures(ws) != armed_links:
        _refuse("sheet {0!r} adds tables AND changes hyperlinks in the "
                "same save; their relationship allocations would collide. "
                "Save between the two edits.".format(ws.title))

    # removals: engine drops the part + CT + the sheet rel
    for name in removed_names:
        if name not in original_parts:
            _refuse("table {0!r} was removed but its part cannot be "
                    "resolved in the original package.".format(name))
        part_name, _payload = original_parts[name]
        part_plan.remove_part(
            part_name,
            referencing_rels=[(rels_part, part_name)])

    # additions: engine creates the part + CT + rel (explicit rIds so the
    # tablePart elements can reference them now)
    rels_payload = zin.read(rels_part) if rels_part in names else None
    existing_ids = set()
    existing_numbers = []
    all_names = set(names) | set(part_plan.added)
    for n in all_names:
        m = re.match(r"xl/tables/table(\d+)\.xml$", n)
        if m:
            existing_numbers.append(int(m.group(1)))
    # table ids are WORKBOOK-unique (ECMA-376): scan every table part in
    # the package, not just this sheet's (duplicate id=1)
    for n in names:
        if n.startswith("xl/tables/") and n.endswith(".xml"):
            m = re.search(br'<table[^>]*\sid="(\d+)"', zin.read(n))
            if m:
                existing_ids.add(int(m.group(1)))
    for payload in part_plan.added.values():
        if isinstance(payload, bytes):
            m = re.search(br'<table[^>]*\sid="(\d+)"', payload)
            if m:
                existing_ids.add(int(m.group(1)))
    next_part_num = max(existing_numbers, default=0) + 1
    next_table_id = max(existing_ids, default=0) + 1

    new_entries = []      # (rid,) for the tableParts element
    for i, name in enumerate(added_names):
        tbl = ws.tables[name]
        validate_table(tbl, tbl.ref)
        _check_display_name(wb, ws, tbl, original_names=set())
        tbl.id = next_table_id + i
        part_name = "xl/tables/table{0}.xml".format(next_part_num + i)
        payload = tostring(tbl.to_tree())
        if not payload.startswith(b"<?xml"):
            payload = (b'<?xml version="1.0" encoding="UTF-8" '
                       b'standalone="yes"?>\n' + payload)
        rid = part_plan.reserve_rid(rels_part, rels_payload)
        part_plan.add_part(part_name, payload,
                           content_type=TABLE_CONTENT_TYPE,
                           relate_from=sheet_part,
                           rel_type=_REL_NS + "/table",
                           rel_id=rid)
        new_entries.append(rid)

    # the rebuilt tableParts element: surviving originals keep their rIds
    entries = []
    for name in sorted(current & armed):
        part_name, _payload = original_parts.get(name, (None, None))
        rid = rid_by_part.get(part_name)
        if rid is None:
            _refuse("table {0!r} has no relationship entry in the "
                    "original sheet rels.".format(name))
        entries.append(rid)
    entries.extend(new_entries)
    if not entries:
        return None                       # element drops entirely
    # per-element xmlns:r declaration: roots that never referenced rels
    # (no drawing/hyperlinks) do not declare the prefix (v0 lesson from
    # the <sheet> entries)
    blob = b"".join(
        b'<tablePart xmlns:r="%s" r:id="%s"/>' % (
            _REL_NS.encode("ascii"), rid.encode("ascii"))
        for rid in entries)
    return (b'<tableParts count="%d">' % len(entries)) + blob \
        + b"</tableParts>"
