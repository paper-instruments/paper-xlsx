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


def _region_signatures(payload):
    """Canonical signatures of every top-level element EXCEPT sheetData
    (whose cells and rows the sheetData checks cover, keyed to the
    ledger's claims). Known accepted limits, unreachable from the splice
    (edits are span-bounded): inter-tag document order, root attributes,
    and inter-element text are not compared."""
    root = ET.fromstring(payload)
    regions = {}
    for el in root:
        tag = el.tag.rsplit("}", 1)[-1]
        if tag == "sheetData":
            continue
        regions.setdefault(tag, []).append(
            ET.canonicalize(ET.tostring(el)))
    return regions


def _sheet_rows(payload):
    """Per-row signature: (attrs minus r, multiplicity) — catches row
    attribute drift, duplication, and deletion that the per-cell check
    (keyed by r only) is blind to."""
    root = ET.fromstring(payload)
    rows = {}
    for sheetdata in root.iter(_SHEETDATA):
        for row in sheetdata.findall(_ROW):
            ref = row.get("r")
            attrs = tuple(sorted((k, v) for k, v in row.attrib.items()
                                 if k != "r"))
            sig, count = rows.get(ref, (attrs, 0))
            rows[ref] = (attrs if count == 0 else sig, count + 1)
        break
    return rows


def verify_splice(source_bytes, output_bytes, dirty_by_part, baselines=None,
                  region_claims=None, row_claims=None):
    """Assert that in every spliced part, the set of semantically changed
    cells is a subset of the ledger's dirty claims — and (PLAN-v0.1 0.4)
    that no region the saver didn't claim differs, and no row's display
    attributes/multiplicity change outside the claimed rows.

    ``baselines`` maps parts to their post-shift bytes (Phase 6b): those
    parts are checked against the renumbered baseline (the renumber pass is
    covered by its own tests and the oracle property tests).
    ``region_claims`` maps parts to the region tags the saver knowingly
    rewrote; an unclaimed region that differs is corruption inside the
    safety tooling, exactly like an unclaimed cell.
    ``row_claims`` maps parts to row indices whose display attributes the
    saver knowingly rewrote; rows holding dirty cells are implicitly
    allowed (the splice recomputes their spans/attrs when rebuilding)."""
    baselines = baselines or {}
    region_claims = region_claims or {}
    row_claims = row_claims or {}
    with zipfile.ZipFile(io.BytesIO(source_bytes)) as zin, \
            zipfile.ZipFile(io.BytesIO(output_bytes)) as zout:
        for part, dirty in dirty_by_part.items():
            baseline = baselines.get(part) or zin.read(part)
            output = zout.read(part)
            before = _sheet_cells(baseline)
            after = _sheet_cells(output)
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

            regions_before = _region_signatures(baseline)
            regions_after = _region_signatures(output)
            claimed = region_claims.get(part, set())
            rogue_regions = {
                tag for tag in set(regions_before) | set(regions_after)
                if regions_before.get(tag) != regions_after.get(tag)
                and tag not in claimed}
            if rogue_regions:
                raise LedgerCrossCheckError(
                    "ledger cross-check FAILED for {0}: the splice changed "
                    "region(s) {1} the saver never claimed. This is "
                    "corruption inside the safety tooling; the save output "
                    "must not be trusted.".format(
                        part, sorted(rogue_regions)))

            rows_before = _sheet_rows(baseline)
            rows_after = _sheet_rows(output)
            allowed_rows = {str(r) for (r, _c) in dirty} \
                | {str(r) for r in row_claims.get(part, set())}
            rogue_rows = {
                ref for ref in set(rows_before) | set(rows_after)
                if rows_before.get(ref) != rows_after.get(ref)
                and ref not in allowed_rows}
            if rogue_rows:
                raise LedgerCrossCheckError(
                    "ledger cross-check FAILED for {0}: the splice changed "
                    "row(s) {1} (attributes or multiplicity) that the "
                    "ledger never claimed. This is corruption inside the "
                    "safety tooling; the save output must not be "
                    "trusted.".format(part, sorted(rogue_rows)))
