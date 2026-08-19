# paper-xlsx: LibreOffice-free error evidence

"""Report cached error values and broken formula references."""

import re

from openpyxl.formula.tokenizer import (
    EXCEL_ERROR_CODES,
    Token,
    Tokenizer,
    TokenizerError,
)

ERROR_TOKENS = frozenset(EXCEL_ERROR_CODES)

_ERROR_CELL_RE = re.compile(
    br"<c\b([^>]*)\bt=(?:\"e\"|'e')([^>]*)>(.*?)</c>", re.S)
_R_ATTR_RE = re.compile(br"\br=(?:\"([A-Za-z]+\d+)\"|'([A-Za-z]+\d+)')")


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

    with zipfile.ZipFile(io.BytesIO(source)) as zin:
        title_by_part = current_titles_by_part(wb, zin)
        names = set(zin.namelist())
        for part, title in sorted(title_by_part.items()):
            if part not in names:
                continue
            for match in _ERROR_CELL_RE.finditer(zin.read(part)):
                ref_match = _R_ATTR_RE.search(match.group(1) + match.group(2))
                value_match = re.search(br"<v[^>]*>([^<]*)</v>", match.group(3))
                if ref_match is None or value_match is None:
                    continue
                ref = (ref_match.group(1) or ref_match.group(2)).decode("ascii")
                address = "{0}!{1}".format(title, ref)
                if address in seen:
                    continue
                results.append({
                    "address": address,
                    "value": value_match.group(1).decode("utf-8", "replace"),
                    "source": "cache",
                })
                seen.add(address)
    return results
