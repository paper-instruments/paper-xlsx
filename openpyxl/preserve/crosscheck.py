# paper-xlsx: the ledger cross-check (CONVENTIONS §3.3; debug mode)

"""Cross-check the splice output against the ledger's claims.

A cell the splice changed that the ledger never recorded is corruption
INSIDE the safety tooling — a release-blocking bug class — so this check
raises hard, never warns. Enabled via PAPER_LEDGER_CROSSCHECK=1 (the paper
test suite turns it on for every preserve-mode save it performs).
"""

import io
import zipfile
from xml.etree import ElementTree as ET

from openpyxl.xml.constants import SHEET_MAIN_NS

_ROW = "{%s}row" % SHEET_MAIN_NS
_CELL = "{%s}c" % SHEET_MAIN_NS
_SHEETDATA = "{%s}sheetData" % SHEET_MAIN_NS


class LedgerCrossCheckError(RuntimeError):
    """The splice changed cells the ledger never recorded."""


def _cell_signature(el):
    """Canonical per-cell signature; s missing is equivalent to s='0'."""
    attrs = dict(el.attrib)
    attrs.pop("r", None)
    attrs.setdefault("s", "0")
    body = ET.canonicalize(ET.tostring(el))
    return (tuple(sorted(attrs.items())), body)


def _sheet_cells(payload):
    root = ET.fromstring(payload)
    cells = {}
    for sheetdata in root.iter(_SHEETDATA):
        for row in sheetdata.findall(_ROW):
            for cell in row.findall(_CELL):
                ref = cell.get("r")
                if ref:
                    cells[ref] = _cell_signature(cell)
        break
    return cells


def _coord_ref(row, col):
    letters = ""
    c = col
    while c:
        c, rem = divmod(c - 1, 26)
        letters = chr(65 + rem) + letters
    return "{0}{1}".format(letters, row)


def verify_splice(source_bytes, output_bytes, dirty_by_part, baselines=None):
    """Assert that in every spliced part, the set of semantically changed
    cells is a subset of the ledger's dirty claims.

    ``baselines`` maps parts to their post-shift bytes (Phase 6b): those
    parts are checked against the renumbered baseline (the renumber pass is
    covered by its own tests and the oracle property tests)."""
    baselines = baselines or {}
    with zipfile.ZipFile(io.BytesIO(source_bytes)) as zin, \
            zipfile.ZipFile(io.BytesIO(output_bytes)) as zout:
        for part, dirty in dirty_by_part.items():
            before = _sheet_cells(baselines.get(part) or zin.read(part))
            after = _sheet_cells(zout.read(part))
            allowed = {_coord_ref(r, c) for (r, c) in dirty}
            changed = set()
            for ref in set(before) | set(after):
                if before.get(ref) != after.get(ref):
                    changed.add(ref)
            rogue = changed - allowed
            if rogue:
                raise LedgerCrossCheckError(
                    "ledger cross-check FAILED for {0}: the splice changed "
                    "cell(s) {1} that the ledger never recorded. This is "
                    "corruption inside the safety tooling; the save output "
                    "must not be trusted.".format(part, sorted(rogue)))
