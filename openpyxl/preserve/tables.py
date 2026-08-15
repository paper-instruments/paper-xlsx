# paper-xlsx: table support under preserve

"""Loaded-table mutation, and the table row discipline.

Table parts are fully modeled upstream (Table is Serialisable end to end),
so a mutated table re-renders whole from the model — located by
displayName in the ORIGINAL sheet rels, never by guessed part numbering.
Guards refuse geometry the discipline cannot keep coherent: the header row
moves, the data region vanishes, or the column count disagrees with
tableColumns.
"""

import io
import re
import zipfile
from copy import copy

from openpyxl.errors import UnsupportedStructureError
from openpyxl.utils.cell import get_column_letter, range_boundaries
from openpyxl.xml.functions import tostring

from . import crosspart

_TABLE_REL_TYPE_SUFFIX = "/table"
_DISPLAY_NAME_RE = re.compile(
    br'displayName=(?:"([^"]*)"|\'([^\']*)\')')
_REF_RE = re.compile(
    br'<(?:(?:\w+):)?table(?=[\s>])[^>]*\sref=(?:"([^"]*)"|\'([^\']*)\')')


def _refuse(msg):
    raise UnsupportedStructureError(msg + " Nothing was written.")


def _totals_row_count(tbl):
    """Return the physical totals-row count for the supported table shape.

    Excel producers use either ``totalsRowCount=1`` or
    ``totalsRowShown=1`` (and sometimes both) for the single totals row.
    More than one totals row has no closed append contract here.
    """
    count = tbl.totalsRowCount
    if count not in (None, 0, 1):
        _refuse("table {0!r} declares {1} totals rows; table-row append "
                "supports at most one".format(tbl.displayName, count))
    return int(count == 1 or bool(tbl.totalsRowShown))


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
    except (TypeError, ValueError):
        _refuse("table {0!r}: ref {1!r} is not a rectangular "
                "range.".format(tbl.displayName, tbl.ref))
    o_min_col, o_min_row, _oc, _or = range_boundaries(original_ref)
    header = tbl.headerRowCount if tbl.headerRowCount is not None else 1
    totals = _totals_row_count(tbl)
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
        except (TypeError, ValueError):
            _refuse("table {0!r}: autoFilter ref {1!r} is not a "
                    "range.".format(tbl.displayName, tbl.autoFilter.ref))
        if (a_min_col < min_col or a_max_col > max_col
                or a_min_row < min_row or a_max_row > max_row):
            _refuse("table {0!r}: autoFilter ref {1!r} lies outside the "
                    "table ref {2!r}; sync it (append_table_row() does "
                    "this automatically).".format(
                        tbl.displayName, tbl.autoFilter.ref, tbl.ref))


def plan_table_mutations(wb, ws, sheet_part, zin, changed_names, plan,
                         armed_tables=None):
    """Patch each changed loaded table into its ORIGINAL part."""
    from .lexical import patch_xml

    armed_tables = armed_tables or {}
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
        rendered = tostring(tbl.to_tree())
        payload = patch_xml(
            original, armed_tables.get(name), rendered, "table")
        if payload is None:
            if b"<extLst" in original or b"xr:uid" in original \
                    or b"xmlns:xr" in original:
                _refuse("table {0!r} on sheet {1!r} carries extension "
                        "content that cannot survive this structural table "
                        "edit. Recreate the table or edit without "
                        "preserve=True.".format(name, ws.title))
            payload = rendered
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
    from openpyxl.formula.tokenizer import TokenizerError
    from openpyxl.utils.cell import range_boundaries
    from openpyxl.errors import BoundaryViolationError

    try:
        tokens = Tokenizer(formula).items
    except (IndexError, TypeError, ValueError, TokenizerError):
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
                "append_table_row() would generate formula {0!r} outside Excel's "
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
    return value


def _append_refuse(table_name, message, kind="table-append-unsupported"):
    raise UnsupportedStructureError(
        "append_table_row(): table {0!r} {1}. Nothing was changed.".format(
            table_name, message),
        kind=kind,
        anchor=table_name,
    )


def _node_paths(root):
    paths = {}

    def visit(node, path):
        paths[path] = node
        counts = {}
        for child in node.children:
            tag = child.local()
            index = counts.get(tag, 0)
            counts[tag] = index + 1
            visit(child, path + ((tag, index),))

    visit(root, ((root.local(), 0),))
    return paths


def _attribute_local(name):
    return name.split(":", 1)[-1]


def _validate_closed_table_xml(table_name, original, baseline):
    """Reject table markup outside the model shape proven at load time."""
    try:
        original_nodes = _node_paths(
            crosspart.scan_small(original, "table", max_depth=64))
        baseline_nodes = _node_paths(
            crosspart.scan_small(baseline, "table", max_depth=64))
    except (TypeError, ValueError, UnsupportedStructureError) as exc:
        _append_refuse(
            table_name,
            "has table XML that cannot be inspected safely ({0})".format(exc),
            kind="uninspectable-table-xml",
        )
    if set(original_nodes) != set(baseline_nodes):
        _append_refuse(
            table_name,
            "contains unmodeled table elements or extensions",
            kind="extended-table-unsupported",
        )
    for path, original_node in original_nodes.items():
        baseline_node = baseline_nodes[path]
        if b":" in original_node.name:
            _append_refuse(
                table_name,
                "uses namespace-prefixed table markup outside the supported "
                "worksheet-table subset",
                kind="extended-table-unsupported",
            )
        modeled = {
            _attribute_local(key)
            for key in baseline_node.attrs
            if key != "xmlns" and not key.startswith("xmlns:")
        }
        for key in original_node.attrs:
            if key.startswith("xmlns:"):
                _append_refuse(
                    table_name,
                    "declares extension namespaces outside the supported "
                    "worksheet-table subset",
                    kind="extended-table-unsupported",
                )
            if key == "xmlns":
                continue
            if _attribute_local(key) not in modeled:
                _append_refuse(
                    table_name,
                    "contains the unmodeled table attribute {0!r}".format(key),
                    kind="extended-table-unsupported",
                )


def _source_append_context(ws, table_name):
    """Return retained source evidence needed by append preflight."""
    wb = ws.parent
    led = getattr(wb, "_paper_ledger", None)
    armed = bool(led is not None and led.armed)
    if getattr(wb, "_paper_loaded_from_package", False) and not armed:
        _append_refuse(
            table_name,
            "was loaded without preserve mode, so its original OOXML and "
            "relationships are unavailable; reopen with preserve=True",
            kind="preserve-source-required",
        )
    table_baselines = {}
    if armed:
        table_baselines = led.object_snapshots.get(ws, {}).get("table", {})
    loaded_table = table_name in table_baselines
    context = {
        "loaded_table": loaded_table,
        "table_xml": None,
        "table_baseline": table_baselines.get(table_name),
        "original_ref": None,
        "sheet_xml": None,
        "sheet_scan": None,
    }
    if not armed or ws in led.added_sheets:
        return context
    source = getattr(wb, "_paper_source", None)
    if source is None:
        _append_refuse(
            table_name,
            "has no retained package source for closed-world preflight",
            kind="missing-preserve-source",
        )
    from .saver import _package_info
    from .xmlscan import ScanRefusal, scan_sheet

    with zipfile.ZipFile(io.BytesIO(source)) as zin:
        names = set(zin.namelist())
        _workbook_part, sheet_parts = _package_info(zin)
        original_title = led.renames.get(ws, ws.title)
        sheet_part = sheet_parts.get(original_title)
        if sheet_part is None or sheet_part not in names:
            _append_refuse(
                table_name,
                "cannot resolve its worksheet through the retained package "
                "relationships",
                kind="missing-worksheet-part",
            )
        sheet_xml = zin.read(sheet_part)
        try:
            sheet_scan = scan_sheet(sheet_xml)
        except ScanRefusal as exc:
            _append_refuse(
                table_name,
                "is on a worksheet that cannot be scanned safely ({0})".format(
                    exc),
                kind="uninspectable-worksheet",
            )
        rels_part = _rels_path(sheet_part)
        if rels_part in names:
            rels = crosspart.scan_small(
                zin.read(rels_part), "Relationships", max_depth=1)
            for child in rels.children:
                if child.local() != "Relationship":
                    continue
                rel_type = child.attrs.get("Type", "")
                if rel_type.endswith("/queryTable") \
                        or "slicer" in rel_type.lower():
                    _append_refuse(
                        table_name,
                        "is on a worksheet with query-table or slicer "
                        "relationships",
                        kind="connected-table-unsupported",
                    )
        if loaded_table:
            parts = sheet_table_parts(zin, sheet_part)
            if table_name not in parts:
                _append_refuse(
                    table_name,
                    "cannot resolve its original table part",
                    kind="missing-table-part",
                )
            table_part, table_xml = parts[table_name]
            table_rels = _rels_path(table_part)
            if table_rels in names:
                root = crosspart.scan_small(
                    zin.read(table_rels), "Relationships", max_depth=1)
                if any(child.local() == "Relationship"
                       for child in root.children):
                    _append_refuse(
                        table_name,
                        "has table-part relationships outside the supported "
                        "worksheet-table subset",
                        kind="connected-table-unsupported",
                    )
            context["table_xml"] = table_xml
            match = _REF_RE.search(table_xml)
            if match is None:
                _append_refuse(
                    table_name,
                    "has no inspectable ref in its original table part",
                    kind="invalid-table-geometry",
                )
            raw_ref = match.group(1) if match.group(1) is not None \
                else match.group(2)
            context["original_ref"] = _unescape(raw_ref.decode("ascii"))
        context["sheet_xml"] = sheet_xml
        context["sheet_scan"] = sheet_scan
    return context


def _bounds_intersect(left, right):
    return not (
        left[2] < right[0] or right[2] < left[0]
        or left[3] < right[1] or right[3] < left[1]
    )


def _style_locked(wb, style):
    if style is None:
        return True
    index = style.protectionId
    if index < 0 or index >= len(wb._protections):
        return True
    return bool(wb._protections[index].locked)


def _validate_model_table_contract(tbl, table_name):
    if tbl.displayName != table_name or tbl.name not in (None, table_name):
        _append_refuse(
            table_name,
            "has inconsistent name/displayName metadata",
            kind="invalid-table-geometry",
        )
    if tbl.tableType not in (None, "worksheet") \
            or tbl.connectionId is not None:
        _append_refuse(
            table_name,
            "is a query, XML, or externally connected table",
            kind="connected-table-unsupported",
        )
    if not isinstance(tbl.id, int) or isinstance(tbl.id, bool) or tbl.id < 1:
        _append_refuse(
            table_name,
            "has an invalid table id",
            kind="invalid-table-geometry",
        )
    if bool(tbl.insertRow) or bool(tbl.insertRowShift):
        _append_refuse(
            table_name,
            "uses an active Excel insert-row state",
            kind="table-insert-state-unsupported",
        )
    if tbl.sortState is not None \
            or (tbl.autoFilter is not None
                and tbl.autoFilter.sortState is not None):
        _append_refuse(
            table_name,
            "has active sort metadata that cannot be expanded without "
            "re-sorting the table",
            kind="sorted-table-unsupported",
        )
    if vars(tbl).get("extLst") is not None:
        _append_refuse(
            table_name,
            "contains table extensions",
            kind="extended-table-unsupported",
        )
    if tbl.autoFilter is not None \
            and vars(tbl.autoFilter).get("extLst") is not None:
        _append_refuse(
            table_name,
            "contains auto-filter extensions",
            kind="extended-table-unsupported",
        )
    if tbl.autoFilter is not None and any(
            vars(column).get("extLst") is not None
            for column in tbl.autoFilter.filterColumn):
        _append_refuse(
            table_name,
            "contains filter-column extensions",
            kind="extended-table-unsupported",
        )
    for column in tbl.tableColumns:
        if column.queryTableFieldId is not None \
                or column.xmlColumnPr is not None \
                or column.extLst is not None:
            _append_refuse(
                table_name,
                "contains connected or extended table-column metadata",
                kind="connected-table-unsupported",
            )
        for formula in (column.calculatedColumnFormula,
                        column.totalsRowFormula):
            if formula is not None and bool(formula.array):
                _append_refuse(
                    table_name,
                    "contains an array table-column formula",
                    kind="formula-surface-unsupported",
                )
    rendered = tostring(tbl.to_tree())
    if re.search(br"<(?:(?:\w+):)?extLst(?=[\s>])", rendered) \
            or b"AlternateContent" in rendered:
        _append_refuse(
            table_name,
            "contains unsupported extension markup",
            kind="extended-table-unsupported",
        )


def _table_geometry(ws, tbl, table_name, loaded_table):
    try:
        min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
    except (TypeError, ValueError) as exc:
        _append_refuse(
            table_name,
            "has a non-rectangular ref {0!r} ({1})".format(tbl.ref, exc),
            kind="invalid-table-geometry",
        )
    if min_col < 1 or min_row < 1 or max_col > 16384 or max_row > 1048576:
        _append_refuse(
            table_name,
            "extends outside Excel's worksheet grid",
            kind="table-boundary-violation",
        )
    header = tbl.headerRowCount if tbl.headerRowCount is not None else 1
    if header != 1:
        _append_refuse(
            table_name,
            "does not have exactly one header row",
            kind="headerless-table-unsupported",
        )
    totals = _totals_row_count(tbl)
    if max_row - min_row + 1 < header + totals + 1:
        _append_refuse(
            table_name,
            "does not contain an existing data row",
            kind="invalid-table-geometry",
        )
    n_cols = max_col - min_col + 1
    headers = []
    for col in range(min_col, max_col + 1):
        cell = ws._cells.get((min_row, col))
        value = getattr(cell, "_value", None)
        if not isinstance(value, str) or not value:
            _append_refuse(
                table_name,
                "has a missing or non-text header at {0}{1}".format(
                    get_column_letter(col), min_row),
                kind="invalid-table-header",
            )
        headers.append(value)
    folded = [name.casefold() for name in headers]
    if len(set(folded)) != len(folded):
        _append_refuse(
            table_name,
            "has duplicate case-insensitive column headers",
            kind="ambiguous-table-header",
        )
    columns = list(tbl.tableColumns)
    initialize_columns = False
    if not columns:
        if loaded_table:
            _append_refuse(
                table_name,
                "has no tableColumns metadata in its loaded table part",
                kind="invalid-table-geometry",
            )
        initialize_columns = True
    elif len(columns) != n_cols:
        _append_refuse(
            table_name,
            "spans {0} columns but defines {1} tableColumns".format(
                n_cols, len(columns)),
            kind="invalid-table-geometry",
        )
    else:
        names = [column.name for column in columns]
        if any(not isinstance(name, str) or not name for name in names) \
                or [name.casefold() for name in names] != folded:
            _append_refuse(
                table_name,
                "has tableColumns that do not exactly match its headers",
                kind="invalid-table-header",
            )
        ids = [column.id for column in columns]
        if any(not isinstance(value, int) or isinstance(value, bool)
               or value < 1 for value in ids) or len(set(ids)) != len(ids):
            _append_refuse(
                table_name,
                "has missing, non-positive, or duplicate tableColumn ids",
                kind="invalid-table-geometry",
            )
    filter_includes_totals = False
    if tbl.autoFilter is not None:
        if not tbl.autoFilter.ref:
            _append_refuse(
                table_name,
                "has an autoFilter without a range",
                kind="invalid-table-filter",
            )
        data_only = (min_col, min_row, max_col, max_row - totals)
        full_table = (min_col, min_row, max_col, max_row)
        try:
            actual = range_boundaries(tbl.autoFilter.ref)
        except (TypeError, ValueError):
            actual = None
        accepted = {data_only}
        if totals:
            # Producers disagree on whether a totals row belongs to the
            # serialized table autoFilter range. Both shapes are common and
            # have unambiguous expansion semantics, so retain the producer's
            # convention instead of silently normalizing it.
            accepted.add(full_table)
        if actual not in accepted:
            _append_refuse(
                table_name,
                "has an autoFilter range that does not match its header and "
                "data rows",
                kind="invalid-table-filter",
            )
        for column in tbl.autoFilter.filterColumn:
            if not isinstance(column.colId, int) \
                    or isinstance(column.colId, bool) \
                    or column.colId < 0 or column.colId >= n_cols:
                _append_refuse(
                    table_name,
                    "has an autoFilter column outside the table",
                    kind="invalid-table-filter",
                )
        filter_includes_totals = bool(totals and actual == full_table)
    return (min_col, min_row, max_col, max_row, totals,
            tuple(headers), initialize_columns, filter_includes_totals)


def _normalize_row_values(ws, tbl, table_name, values, headers,
                          initialize_columns, row, min_col):
    from collections.abc import Mapping

    n_cols = len(headers)
    if isinstance(values, Mapping):
        by_name = dict(values)
        unknown = set(by_name) - set(headers)
        if unknown:
            _append_refuse(
                table_name,
                "received unknown column(s) {0}".format(sorted(unknown)),
                kind="unknown-table-column",
            )
        row_values = [by_name.get(name) for name in headers]
    else:
        if isinstance(values, (str, bytes, bytearray)):
            raise TypeError("append_table_row() values must be a mapping or "
                            "an iterable of cell values")
        row_values = list(values)
        if len(row_values) > n_cols:
            _append_refuse(
                table_name,
                "received {0} values for a {1}-column table".format(
                    len(row_values), n_cols),
                kind="table-value-count-mismatch",
            )
        row_values += [None] * (n_cols - len(row_values))
    columns = list(tbl.tableColumns)
    planned = []
    for index, col in enumerate(range(min_col, min_col + n_cols)):
        column = None if initialize_columns else columns[index]
        formula = getattr(column, "calculatedColumnFormula", None) \
            if column is not None else None
        formula_text = getattr(formula, "attr_text", None) \
            if formula is not None else None
        if formula is not None and bool(formula.array):
            _append_refuse(
                table_name,
                "has an array calculated-column formula for {0!r}".format(
                    headers[index]),
                kind="formula-surface-unsupported",
            )
        if formula is not None and not formula_text:
            _append_refuse(
                table_name,
                "has an empty calculated-column formula for {0!r}".format(
                    headers[index]),
                kind="invalid-calculated-column",
            )
        if formula_text and row_values[index] is not None:
            _append_refuse(
                table_name,
                "received a value for calculated column {0!r}".format(
                    headers[index]),
                kind="calculated-column-value-conflict",
            )
        value = formula_text if formula_text \
            and formula_text.startswith("=") else (
                "=" + formula_text if formula_text else row_values[index])
        planned.append(_normalize_append_value(
            ws, row, col, value, table_name))
    return tuple(planned)


def _preflight_append_surfaces(ws, table_name, bounds, totals, source):
    from openpyxl.cell import MergedCell
    from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

    min_col, min_row, max_col, max_row = bounds
    data_row = max_row - totals + 1
    destination_row = max_row + 1
    affected = {
        (data_row, col) for col in range(min_col, max_col + 1)
    }
    if totals:
        affected.update(
            (destination_row, col) for col in range(min_col, max_col + 1))
    expanded = (min_col, min_row, max_col, destination_row)
    for other in ws.tables.values():
        other_name = other.displayName
        if other is ws.tables[table_name]:
            continue
        try:
            other_bounds = range_boundaries(other.ref)
        except (TypeError, ValueError):
            _append_refuse(
                table_name,
                "shares a worksheet with malformed table {0!r}".format(
                    other_name),
                kind="overlapping-table-unsupported",
            )
        if _bounds_intersect(expanded, other_bounds):
            _append_refuse(
                table_name,
                "would overlap table {0!r}".format(other_name),
                kind="overlapping-table-unsupported",
            )
    for merged in ws.merged_cells.ranges:
        merged_bounds = (merged.min_col, merged.min_row,
                         merged.max_col, merged.max_row)
        if _bounds_intersect(expanded, merged_bounds):
            _append_refuse(
                table_name,
                "would write into merged range {0}".format(merged.coord),
                kind="merged-table-destination",
            )
    for coordinate in affected:
        cell = ws._cells.get(coordinate)
        if isinstance(cell, MergedCell):
            _append_refuse(
                table_name,
                "would write into merged cell {0}".format(cell.coordinate),
                kind="merged-table-destination",
            )
    for cell in ws._cells.values():
        value = getattr(cell, "_value", None)
        if not isinstance(value, (ArrayFormula, DataTableFormula)):
            continue
        try:
            formula_bounds = range_boundaries(value.ref)
        except (TypeError, ValueError):
            _append_refuse(
                table_name,
                "is on a worksheet with an uninspectable array/data-table "
                "formula",
                kind="formula-surface-unsupported",
            )
        if _bounds_intersect(expanded, formula_bounds):
            _append_refuse(
                table_name,
                "would intersect array, spill, or data-table formula range "
                "{0}".format(value.ref),
                kind="formula-surface-unsupported",
            )
    scan = source.get("sheet_scan")
    sheet_xml = source.get("sheet_xml")
    if scan is not None:
        for array_bounds in scan.array_bounds:
            if _bounds_intersect(expanded, array_bounds):
                _append_refuse(
                    table_name,
                    "would intersect an original array or spill formula",
                    kind="formula-surface-unsupported",
                )
        for row, col in affected:
            row_span = scan.rows.get(row)
            cell_span = row_span.cells.get(col) if row_span else None
            if cell_span is None:
                continue
            if cell_span.shared_si is not None or cell_span.array_ref is not None \
                    or cell_span.has_extlst \
                    or cell_span.has_unowned_children \
                    or "cm" in cell_span.attrs or "vm" in cell_span.attrs:
                _append_refuse(
                    table_name,
                    "would rewrite cell {0}{1} with formula or extension "
                    "metadata outside the append contract".format(
                        get_column_letter(col), row),
                    kind="formula-surface-unsupported",
                )
            raw_cell = sheet_xml[cell_span.start:cell_span.end]
            if re.search(
                    br"<(?:(?:\w+):)?f(?=[\s>])[^>]*\bt\s*=\s*"
                    br"(?:\"dataTable\"|'dataTable')", raw_cell):
                _append_refuse(
                    table_name,
                    "would intersect a what-if data-table formula",
                    kind="formula-surface-unsupported",
                )
    for col in range(min_col, max_col + 1):
        target = ws._cells.get((destination_row, col))
        if target is None:
            continue
        if getattr(target, "_value", None) is not None \
                or bool(getattr(target, "has_style", False)) \
                or getattr(target, "_comment", None) is not None \
                or getattr(target, "_hyperlink", None) is not None:
            _append_refuse(
                table_name,
                "would overwrite existing content or formatting at {0}{1}"
                .format(get_column_letter(col), destination_row),
                kind="table-destination-conflict",
            )
    return affected


def _preflight_append_protection(ws, table_name, data_styles, totals_cells):
    if not bool(ws.protection.sheet) or not getattr(
            ws.parent, "strict_protection", False):
        return
    for index, style in enumerate(data_styles):
        if _style_locked(ws.parent, style):
            _append_refuse(
                table_name,
                "would create a locked data cell in column {0} on strictly "
                "protected sheet {1!r}; unlock the existing table data "
                "column or unprotect the sheet".format(index + 1, ws.title),
                kind="protected-table-append",
            )
    for cell in totals_cells:
        if _style_locked(ws.parent, getattr(cell, "_style", None)):
            _append_refuse(
                table_name,
                "would move a locked totals cell {0} on strictly protected "
                "sheet {1!r}; unprotect the sheet before appending".format(
                    cell.coordinate, ws.title),
                kind="protected-table-append",
            )


class _TableAppendPlan:
    """Complete immutable-by-convention plan for one table-row append."""

    def __init__(self, ws, table, table_name, bounds, totals, headers,
                 initialize_columns, filter_includes_totals, values,
                 data_styles, totals_values, totals_styles, affected):
        self.ws = ws
        self.table = table
        self.table_name = table_name
        self.min_col, self.min_row, self.max_col, self.max_row = bounds
        self.totals = totals
        self.headers = headers
        self.initialize_columns = initialize_columns
        self.filter_includes_totals = filter_includes_totals
        self.values = values
        self.data_styles = data_styles
        self.totals_values = totals_values
        self.totals_styles = totals_styles
        self.affected = frozenset(affected)
        self.data_row = self.max_row - totals + 1
        self.destination_row = self.max_row + 1
        self.new_ref = "{0}{1}:{2}{3}".format(
            get_column_letter(self.min_col), self.min_row,
            get_column_letter(self.max_col), self.destination_row)
        self.filter_will_exist = table.autoFilter is not None \
            or initialize_columns
        self.new_filter_ref = "{0}{1}:{2}{3}".format(
            get_column_letter(self.min_col), self.min_row,
            get_column_letter(self.max_col),
            self.destination_row if filter_includes_totals
            else self.destination_row - totals)


def _plan_table_append(ws, table_name, values):
    tbl = ws.tables[table_name]
    source = _source_append_context(ws, table_name)
    _validate_model_table_contract(tbl, table_name)
    if source["loaded_table"]:
        _validate_closed_table_xml(
            table_name, source["table_xml"], source["table_baseline"])
        validate_table(tbl, source["original_ref"])
    geometry = _table_geometry(
        ws, tbl, table_name, source["loaded_table"])
    min_col, min_row, max_col, max_row, totals, headers, initialize, \
        filter_includes_totals = geometry
    if max_row >= 1048576:
        from openpyxl.errors import BoundaryViolationError

        raise BoundaryViolationError(
            "append_table_row() would extend table {0!r} past row 1048576, "
            "Excel's hard sheet limit. Nothing was changed.".format(
                table_name))
    affected = _preflight_append_surfaces(
        ws, table_name, (min_col, min_row, max_col, max_row), totals, source)
    data_row = max_row - totals + 1
    planned_values = _normalize_row_values(
        ws, tbl, table_name, values, headers, initialize,
        data_row, min_col)
    template_row = data_row - 1
    data_styles = []
    for col in range(min_col, max_col + 1):
        template = ws._cells.get((template_row, col))
        if template is None:
            data_styles.append(None)
        else:
            data_styles.append(
                copy(template._style) if template._style is not None else None)
    totals_values = []
    totals_styles = []
    totals_cells = []
    if totals:
        for col in range(min_col, max_col + 1):
            cell = ws._cells.get((max_row, col))
            if cell is None:
                totals_values.append(None)
                totals_styles.append(None)
                continue
            if cell._comment is not None or cell._hyperlink is not None:
                _append_refuse(
                    table_name,
                    "has a totals-row comment or hyperlink at {0}".format(
                        cell.coordinate),
                    kind="totals-relationship-unsupported",
                )
            value = cell._value
            formula = value if isinstance(value, str) \
                and value.startswith("=") else getattr(value, "text", None)
            if isinstance(formula, str) and formula.startswith("=") \
                    and "[" not in formula:
                _append_refuse(
                    table_name,
                    "has totals formula {0!r} whose row-expansion semantics "
                    "cannot be proven from a structured reference".format(
                        formula),
                    kind="totals-formula-unsupported",
                )
            totals_values.append(_normalize_append_value(
                ws, max_row + 1, col, value, table_name))
            totals_styles.append(
                copy(cell._style) if cell._style is not None else None)
            totals_cells.append(cell)
    _preflight_append_protection(
        ws, table_name, data_styles, totals_cells)
    return _TableAppendPlan(
        ws, tbl, table_name,
        (min_col, min_row, max_col, max_row), totals, headers, initialize,
        filter_includes_totals, planned_values, tuple(data_styles),
        tuple(totals_values), tuple(totals_styles), affected)


class _TableAppendTransaction:
    """Bounded rollback journal for one table-row append."""

    def __init__(self, plan):
        self.ws = plan.ws
        self.table = plan.table
        self.coordinates = {}
        for coordinate in plan.affected:
            cell = self.ws._cells.get(coordinate)
            if cell is None:
                self.coordinates[coordinate] = (False, None, None)
                continue
            hyperlink = getattr(cell, "_hyperlink", None)
            comment = getattr(cell, "_comment", None)
            state = (
                cell._value,
                cell._data_type,
                cell._style,
                hyperlink,
                getattr(hyperlink, "ref", None),
                comment,
                getattr(comment, "_parent", None),
            )
            self.coordinates[coordinate] = (True, cell, state)
        self.current_row = self.ws._current_row
        self.table_ref = self.table.ref
        self.table_columns = self.table.tableColumns
        self.table_column_items = list(self.table_columns)
        self.auto_filter = self.table.autoFilter
        self.filter_ref = self.table.autoFilter.ref \
            if self.table.autoFilter is not None else None
        self.ledger = getattr(self.ws.parent, "_paper_ledger", None)
        if self.ledger is not None and not self.ledger.armed:
            self.ledger = None
        self.ledger_states = {}
        if self.ledger is not None:
            for coordinate in self.coordinates:
                self.ledger_states[coordinate] = tuple(
                    self._capture(mapping, coordinate, bucket_type)
                    for mapping, bucket_type in (
                        (self.ledger.cells, set),
                        (self.ledger.value_overwrites, set),
                        (self.ledger.cache_writes, dict),
                    ))
            self.formulas_changed = self.ledger.formulas_changed
            self.was_warned = self.ws in self.ledger.protection_warned
        registry = self.ws.parent._number_formats
        self.number_formats = (
            registry, list(registry), registry.clean, dict(registry._dict))
        self.active = True

    def _capture(self, mapping, coordinate, bucket_type):
        bucket = mapping.get(self.ws)
        if bucket is None:
            return False, None, False, None, bucket_type
        present = coordinate in bucket
        value = bucket.get(coordinate) if bucket_type is dict else None
        return True, bucket, present, value, bucket_type

    def _restore(self, mapping, coordinate, state):
        existed, original, present, value, bucket_type = state
        current = mapping.get(self.ws)
        if existed:
            if current is not original:
                mapping[self.ws] = original
            bucket = original
        else:
            bucket = current
            if bucket is None:
                return
        if bucket_type is set:
            bucket.add(coordinate) if present else bucket.discard(coordinate)
        elif present:
            bucket[coordinate] = value
        else:
            bucket.pop(coordinate, None)
        if not existed and not bucket:
            mapping.pop(self.ws, None)

    def commit(self):
        self.active = False

    def rollback(self):
        if not self.active:
            return
        for coordinate, (existed, cell, state) in reversed(
                list(self.coordinates.items())):
            current = self.ws._cells.get(coordinate)
            if not existed:
                self.ws._cells.pop(coordinate, None)
                continue
            self.ws._cells[coordinate] = cell
            value, data_type, style, hyperlink, hyperlink_ref, comment, \
                comment_parent = state
            current_comment = getattr(current, "_comment", None)
            if current is not cell and current_comment is not None \
                    and getattr(current_comment, "_parent", None) is current:
                current_comment._parent = None
            cell._value = value
            cell._data_type = data_type
            cell._style = style
            cell._hyperlink = hyperlink
            if hyperlink is not None:
                hyperlink.ref = hyperlink_ref
            cell._comment = comment
            if comment is not None:
                comment._parent = comment_parent
        self.ws._current_row = self.current_row
        self.table.ref = self.table_ref
        self.table_columns[:] = self.table_column_items
        self.table.autoFilter = self.auto_filter
        if self.auto_filter is not None:
            self.auto_filter.ref = self.filter_ref
        if self.ledger is not None:
            for coordinate, states in reversed(list(self.ledger_states.items())):
                self._restore(self.ledger.cells, coordinate, states[0])
                self._restore(self.ledger.value_overwrites,
                              coordinate, states[1])
                self._restore(self.ledger.cache_writes, coordinate, states[2])
            self.ledger.formulas_changed = self.formulas_changed
            if self.was_warned:
                self.ledger.protection_warned.add(self.ws)
            else:
                self.ledger.protection_warned.discard(self.ws)
        registry, values, clean, index = self.number_formats
        registry[:] = values
        registry._dict.clear()
        registry._dict.update(index)
        registry.clean = clean
        self.active = False


def _append_commit_point(_name):
    """Test seam for fault injection between append commit stages."""


def _restore_planned_style(cell, planned):
    """Reapply inherited formatting without discarding date auto-formatting."""
    if planned is None:
        return
    bound = cell._style
    final = copy(planned)
    if final.numFmtId == 0 and bound is not None and bound.numFmtId != 0:
        final.numFmtId = bound.numFmtId
    cell._style = final


def _apply_table_append(plan):
    ws = plan.ws
    tbl = plan.table
    if plan.initialize_columns:
        tbl._initialise_columns()
        for column, name in zip(tbl.tableColumns, plan.headers):
            column.name = name
        _append_commit_point("columns-initialized")
    if plan.totals:
        for index, col in enumerate(
                range(plan.min_col, plan.max_col + 1)):
            destination = ws.cell(row=plan.destination_row, column=col)
            style = plan.totals_styles[index]
            destination._style = copy(style) if style is not None else None
            _append_commit_point(
                "totals-style-{0}".format(destination.coordinate))
            destination.value = plan.totals_values[index]
            _append_commit_point(
                "totals-value-{0}".format(destination.coordinate))
            if style is not None:
                _restore_planned_style(destination, style)
                _append_commit_point(
                    "totals-format-{0}".format(destination.coordinate))
    for index, col in enumerate(range(plan.min_col, plan.max_col + 1)):
        destination = ws.cell(row=plan.data_row, column=col)
        style = plan.data_styles[index]
        destination._style = copy(style) if style is not None else None
        _append_commit_point("data-style-{0}".format(destination.coordinate))
        destination.value = plan.values[index]
        _append_commit_point("data-value-{0}".format(destination.coordinate))
        if style is not None:
            _restore_planned_style(destination, style)
            _append_commit_point(
                "data-format-{0}".format(destination.coordinate))
    tbl.ref = plan.new_ref
    _append_commit_point("table-ref")
    if plan.filter_will_exist:
        tbl.autoFilter.ref = plan.new_filter_ref
        _append_commit_point("auto-filter-ref")


def append_table_row(ws, table_name, values):
    """Atomically append one row to a supported named worksheet table.

    The operation plans values, calculated columns, styles, totals movement,
    protection, table geometry, and retained OOXML before changing the model.
    Unsupported connected, extended, merged, spill, sorted, or ambiguous table
    states raise a typed refusal and leave the workbook unchanged.
    """
    if table_name not in ws.tables:
        from openpyxl.errors import TargetNotFoundError

        raise TargetNotFoundError(
            "no table named {0!r} on sheet {1!r}.".format(
                table_name, ws.title))
    plan = _plan_table_append(ws, table_name, values)
    transaction = _TableAppendTransaction(plan)
    try:
        _apply_table_append(plan)
    except BaseException:
        transaction.rollback()
        raise
    transaction.commit()
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
            m = re.search(
                br'<(?:(?:\w+):)?table(?=[\s>])[^>]*\sid="(\d+)"',
                zin.read(n))
            if m:
                existing_ids.add(int(m.group(1)))
    for payload in part_plan.added.values():
        if isinstance(payload, bytes):
            m = re.search(
                br'<(?:(?:\w+):)?table(?=[\s>])[^>]*\sid="(\d+)"',
                payload)
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
