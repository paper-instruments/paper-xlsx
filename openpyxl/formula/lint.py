# paper-xlsx: formula pre-flight linter (PLAN-v0.1 5.2, PR-1 §4.2)

"""Tokenizer-based formula checks — no evaluation, ever.

``lint_formula`` returns findings as ``{"code", "message", "anchor"}``
dicts (the pinned shape); an empty list means nothing to report. With a
workbook, references are checked against real sheets, defined names and
tables. Under preserve mode the value-bind chokepoint runs this per
``wb.formula_lint`` ("off" | "warn" | "refuse", default "warn").

Honest scope: formulas using LET/LAMBDA locals skip the unknown-name
check entirely (locals are indistinguishable from workbook names without
evaluation); R1C1 and external-workbook references are never judged.
"""

import re

from .catalog import EXCEL_FUNCTIONS
from .tokenizer import Token, Tokenizer, TokenizerError

LINT_MODES = ("off", "warn", "refuse")

_A1_RE = re.compile(
    r"^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$")
_COL_SPAN_RE = re.compile(r"^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$")
_ROW_SPAN_RE = re.compile(r"^\$?\d+:\$?\d+$")
_NAME_RE = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.?\\]*$")
_STRUCTURED_RE = re.compile(r"^([A-Za-z_\\][A-Za-z0-9_.?\\]*)?\[(.*)\]$",
                            re.S)
_COLUMN_PART_RE = re.compile(r"\[([^\[\]]*)\]")


def _finding(code, message, anchor):
    return {"code": code, "message": message, "anchor": anchor}


def _strip_prefix(name):
    for prefix in ("_xlfn.", "_xlws.", "_xll."):
        if name.startswith(prefix):
            name = name[len(prefix):]
    return name


def _split_sheet(operand):
    """(sheet_or_None, rest) honouring quoted sheet names; external
    references ([1]Sheet1!A1) return the marker sheet None ('skip')."""
    if operand.startswith("["):
        return ("<external>", None)
    if operand.startswith("'"):
        end = 1
        while end < len(operand):
            if operand[end] == "'":
                if operand[end + 1:end + 2] == "'":
                    end += 2
                    continue
                break
            end += 1
        if operand[end + 1:end + 2] == "!":
            sheet = operand[1:end].replace("''", "'")
            if "[" in sheet:
                # '[Budget.xlsx]Sheet One'!A1 / 'C:\path\[B.xlsx]S'!A1:
                # the quoted storage form of external-workbook references
                # — never judged (Batch-5 gate: flagged as unknown-sheet)
                return ("<external>", None)
            return (sheet, operand[end + 2:])
        return (None, operand)
    if "!" in operand:
        sheet, rest = operand.split("!", 1)
        if sheet.startswith("["):
            return ("<external>", None)
        return (sheet, rest)
    return (None, operand)


def _workbook_tables(workbook):
    tables = {}
    for ws in workbook.worksheets:
        ws_tables = getattr(ws, "tables", None) or {}
        for name in ws_tables:
            tables[name.casefold()] = ws_tables[name]
    return tables


def _known_names(workbook, sheet):
    names = {n.casefold() for n in workbook.defined_names}
    if sheet is not None:
        names |= {n.casefold() for n in sheet.defined_names}
    for ws in workbook.worksheets:
        names |= {n.casefold() for n in ws.defined_names}
    return names


def lint_formula(text, *, workbook=None, sheet=None):
    """Pre-flight findings for one formula string (leading '=' optional).

    Returns ``[{"code", "message", "anchor"}, ...]`` — the pinned finding
    shape. Codes: ``parse-error``, ``unbalanced-parens``,
    ``semicolon-separator``, ``unknown-function``, ``unknown-sheet``,
    ``unknown-name``, ``unknown-table``, ``unknown-column``.
    """
    if not isinstance(text, str):
        raise TypeError("lint_formula expects a formula string")
    if not text.startswith("="):
        text = "=" + text

    findings = []
    try:
        tokens = Tokenizer(text).items
    except (TokenizerError, IndexError) as exc:
        return [_finding("parse-error",
                         "the formula could not be tokenized: "
                         "{0}".format(exc), text)]

    depth = 0
    array_depth = 0
    has_locals = False
    for token in tokens:
        if token.type in (Token.FUNC, Token.PAREN):
            if token.subtype == Token.OPEN:
                depth += 1
            else:
                depth -= 1
                if depth < 0:
                    findings.append(_finding(
                        "unbalanced-parens",
                        "a closing parenthesis has no opening match",
                        token.value))
                    depth = 0
        elif token.type == Token.ARRAY:
            array_depth += 1 if token.subtype == Token.OPEN else -1
        elif token.type == Token.SEP and token.value == ";" \
                and array_depth == 0:
            findings.append(_finding(
                "semicolon-separator",
                "';' used as an argument separator: storage-canonical "
                "formulas ALWAYS use ',' (';' is a locale display "
                "convention; Excel will read this as one malformed "
                "argument)", token.value))
        if token.type == Token.FUNC and token.subtype == Token.OPEN:
            name = _strip_prefix(token.value[:-1].rstrip("("))
            if name.upper() in ("LET", "LAMBDA"):
                has_locals = True
            if name.upper() not in EXCEL_FUNCTIONS \
                    and not name.upper().startswith("ISO."):
                findings.append(_finding(
                    "unknown-function",
                    "{0!r} is not a known Excel function (a typo writes "
                    "#NAME? into the cell; user-defined/macro functions "
                    "are legal and can ignore this)".format(name),
                    token.value))
    if depth != 0:
        findings.append(_finding(
            "unbalanced-parens",
            "{0} unclosed parenthes{1}".format(
                depth, "is" if depth == 1 else "es"),
            text))

    if workbook is None:
        return findings

    sheetnames = {t.casefold() for t in workbook.sheetnames}
    tables = _workbook_tables(workbook)
    names = _known_names(workbook, sheet)

    for token in tokens:
        if token.type != Token.OPERAND or token.subtype != Token.RANGE:
            continue
        operand = token.value
        sheet_part, rest = _split_sheet(operand)
        if sheet_part == "<external>":
            continue                    # external workbook: never judged
        if sheet_part is not None:
            # 3-D spans name two endpoint sheets
            for endpoint in sheet_part.split(":"):
                if endpoint.casefold() not in sheetnames:
                    findings.append(_finding(
                        "unknown-sheet",
                        "sheet {0!r} does not exist in this "
                        "workbook".format(endpoint), operand))
            continue
        if _A1_RE.match(rest) or _COL_SPAN_RE.match(rest) \
                or _ROW_SPAN_RE.match(rest):
            continue                    # plain reference on the own sheet
        m = _STRUCTURED_RE.match(rest)
        if m:
            table_name = m.group(1)
            if table_name:
                table = tables.get(table_name.casefold())
                if table is None:
                    findings.append(_finding(
                        "unknown-table",
                        "table {0!r} does not exist in this "
                        "workbook".format(table_name), operand))
                    continue
                if not table.column_names:
                    # an in-session table has no tableColumns until save
                    # (openpyxl derives them from the header row at write
                    # time): its columns are unknowable here, never
                    # unknown (Batch-5 gate: false refusals on the normal
                    # add-table-then-write-formulas order)
                    continue
                if "'" in m.group(2):
                    # Excel's ' escape for [ ] # ' @ in column names:
                    # decoding needs a full structured-ref parser —
                    # unknowable, never unknown (Batch-5 gate)
                    continue
                columns = {c.casefold()
                           for c in (table.column_names or [])}
                specials = {"#all", "#data", "#headers", "#totals",
                            "#this row"}
                for part in _COLUMN_PART_RE.findall(m.group(2)) \
                        or [m.group(2)]:
                    part = part.strip()
                    if not part or part.startswith("@"):
                        part = part.lstrip("@")
                    if not part:
                        continue
                    if part.casefold() in specials:
                        continue
                    if part.casefold() not in columns:
                        findings.append(_finding(
                            "unknown-column",
                            "table {0!r} has no column {1!r}".format(
                                table_name, part), operand))
            continue
        if _NAME_RE.match(rest) and not has_locals:
            if _strip_prefix(rest).upper() in EXCEL_FUNCTIONS:
                continue      # eta-reduced function ref: REDUCE(0,A,SUM)
            if rest.casefold() not in names \
                    and rest.casefold() not in tables:
                findings.append(_finding(
                    "unknown-name",
                    "{0!r} is neither a defined name, a table, nor a "
                    "cell reference (Excel will show #NAME?)".format(
                        rest), operand))
    return findings


def lint_on_bind(cell, formula_text):
    """The value-bind chokepoint hook (preserve mode only; the caller
    already checked the ledger is armed). Consults wb.formula_lint."""
    ws = cell.parent
    wb = ws.parent
    mode = getattr(wb, "formula_lint", "warn")
    if mode not in LINT_MODES:
        raise ValueError(
            "wb.formula_lint must be one of {0!r} (got {1!r})".format(
                LINT_MODES, mode))
    if mode == "off":
        return
    findings = lint_formula(formula_text, workbook=wb, sheet=ws)
    if not findings:
        return
    summary = "; ".join(
        "[{0}] {1}".format(f["code"], f["message"]) for f in findings[:6])
    if mode == "refuse":
        from openpyxl.errors import UnsupportedStructureError

        raise UnsupportedStructureError(
            "formula for {0}!{1} failed the pre-flight lint and "
            "wb.formula_lint is 'refuse': {2}. Nothing was "
            "changed.".format(ws.title, cell.coordinate, summary))
    import warnings

    from openpyxl.errors import LintWarning

    warnings.warn(LintWarning(
        "formula for {0}!{1}: {2}".format(ws.title, cell.coordinate,
                                          summary)), stacklevel=4)
