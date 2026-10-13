# paper-xlsx: LibreOffice-free error evidence

"""Report cached error values and broken formula references."""

import codecs
import re

from openpyxl.formula.tokenizer import (
    EXCEL_ERROR_CODES,
    Token,
    Tokenizer,
    TokenizerError,
)

ERROR_TOKENS = frozenset(EXCEL_ERROR_CODES)

# t="e", or a t value holding a character reference, which could spell "e".
# An attribute value cannot hold "<", so the scan stops at the tag's end.
_ERROR_TYPE = re.compile(
    br"""\bt\s*=\s*(?:"(?:e"|[^"<]*&#)|'(?:e'|[^'<]*&#))""")


def _may_hold_error_cells(payload):
    """Whether a worksheet part could hold a cell whose ``t`` is ``e``.

    Most sheets hold none, and skipping their parse is most of the scan's
    time. The byte patterns hold only in an ASCII-compatible encoding, so a
    part that does not open with "<" (after a UTF-8 BOM), such as UTF-16, is
    always parsed. A DTD can declare an entity that spells ``e``, so a part
    with one is always parsed too; the guarded parser refuses it where it
    can.
    """
    start = len(codecs.BOM_UTF8) if payload.startswith(codecs.BOM_UTF8) else 0
    if payload[start:start + 1] != b"<" \
            or b"\x00" in payload[start:start + 4]:
        return True
    return _ERROR_TYPE.search(payload) is not None or b"<!DOCTYPE" in payload


def current_titles_by_part(wb, zin):
    """Map worksheet part names to current live titles."""
    from .saver import _package_info

    led = getattr(wb, "_paper_ledger", None)
    current_by_original = {}
    if led is not None:
        for ws_obj, original in getattr(led, "renames", {}).items():
            current_by_original[original] = ws_obj.title
    live_titles = {ws.title for ws in wb.worksheets}
    _wb_part, mapping = _package_info(zin)
    out = {}
    for title, part in mapping.items():
        current = current_by_original.get(title, title)
        if current in live_titles:
            out[part] = current
    return out


def _formula_errors(formula, address):
    from openpyxl.errors import UnsupportedStructureError

    try:
        tokens = Tokenizer(formula).items
    except (TokenizerError, IndexError) as exc:
        raise UnsupportedStructureError(
            "scan_errors() cannot tokenize formula {0}: {1}. No partial "
            "formula-error report was returned.".format(address, exc),
            kind="unscannable-formula",
            anchor=address,
        ) from exc
    found = []
    for token in tokens:
        if token.type != Token.OPERAND or token.subtype == Token.TEXT:
            continue
        if token.subtype == Token.ERROR and token.value in ERROR_TOKENS:
            found.append(token.value)
            continue
        # The upstream tokenizer classifies Sheet!#REF! as a RANGE operand.
        if token.subtype == Token.RANGE:
            for error in EXCEL_ERROR_CODES:
                if token.value.endswith(error) \
                        and token.value[:-len(error)].endswith("!"):
                    found.append(error)
                    break
    return tuple(dict.fromkeys(found))


def scan_errors(wb):
    """Return cached values and actual formula error operands."""
    from openpyxl.workbook.workbook import _require_materialized_cells

    _require_materialized_cells(wb, "scan_errors()")
    results = []
    seen = set()
    for ws in wb.worksheets:
        for cell in sorted(ws._cells.values(), key=lambda item: item.coordinate):
            value = cell._value
            if cell.data_type == "f" and not isinstance(value, str):
                value = getattr(value, "text", None)
            address = "{0}!{1}".format(ws.title, cell.coordinate)
            if not isinstance(value, str):
                continue
            if cell.data_type == "f":
                for error in _formula_errors(value, address):
                    results.append({"address": address, "value": error,
                                    "source": "formula"})
                    seen.add(address)
            elif cell.data_type != "f" and value.strip() in ERROR_TOKENS:
                results.append({"address": address, "value": value.strip(),
                                "source": "value"})
                seen.add(address)

    source = getattr(wb, "_paper_source", None)
    if not source:
        return results

    import io
    import zipfile

    from .crosscheck import _iter_sheet_cells

    with zipfile.ZipFile(io.BytesIO(source)) as zin:
        title_by_part = current_titles_by_part(wb, zin)
        names = set(zin.namelist())
        for part, title in sorted(title_by_part.items()):
            if part not in names:
                continue
            payload = zin.read(part)
            if not _may_hold_error_cells(payload):
                continue
            for ref, cell in _iter_sheet_cells(payload):
                if cell.get("t") != "e":
                    continue
                cached = cell.find(cell.tag[:-1] + "v")
                if cached is None:
                    continue
                address = "{0}!{1}".format(title, ref)
                if address in seen:
                    continue
                results.append({
                    "address": address,
                    "value": cached.text or "",
                    "source": "cache",
                })
                seen.add(address)
    return results
