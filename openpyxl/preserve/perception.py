# paper-xlsx: perception — manifest and dependency sketch (PLAN Phase 4)

"""What an agent needs to know about a workbook before touching it:
what is in it, what of that is preservable under the active mode, and
which cells feed which.

The manifest's confession block is the honesty core: charts, pivots, VBA,
extensions and shapes are enumerated from the ACTUAL package (the retained
bytes under preserve, the load-time loss inventory otherwise) — never from
the model, which under-reports exactly the content that is at risk.
"""

import io
import re
import zipfile

from openpyxl.utils.cell import range_boundaries

# CONVENTIONS §3.7 (pinned): nondeterministic volatiles are excluded from
# certification; INDIRECT/OFFSET are volatile but deterministic given inputs
VOLATILE_NONDETERMINISTIC = ("NOW", "TODAY", "RAND", "RANDBETWEEN",
                             "RANDARRAY")
VOLATILE_DETERMINISTIC = ("INDIRECT", "OFFSET")

_VOLATILE_RE = re.compile(
    r"\b(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|INDIRECT|OFFSET)\s*\(",
    re.IGNORECASE)


class WorkbookManifest:

    SCHEMA = "workbook_manifest"
    VERSION = 1

    def __init__(self, doc):
        self._doc = doc

    def to_dict(self):
        return dict(self._doc)

    def __repr__(self):
        conf = self._doc["confession"]
        return ("WorkbookManifest({0} sheets, {1} charts, vba={2}, "
                "mode={3!r})".format(
                    len(self._doc["sheets"]), conf["chart_parts"],
                    conf["vba_present"], self._doc["preservation"]["mode"]))


def build_manifest(wb):
    """Build a :class:`WorkbookManifest` for any workbook (fresh, stock,
    or preserve-mode)."""
    sheets = []
    volatile = {}
    for ws in wb.worksheets:
        formulas = 0
        for (row, col), cell in ws._cells.items():
            if cell.data_type == "f" and isinstance(cell._value, str):
                formulas += 1
                for match in _VOLATILE_RE.finditer(cell._value):
                    name = match.group(1).upper()
                    volatile.setdefault(name, []).append(
                        "{0}!{1}".format(_quoted(ws.title), cell.coordinate))
        entry = {
            "title": ws.title,
            "state": ws.sheet_state,
            "dimensions": ws.calculate_dimension(),
            "cell_count": len(ws._cells),
            "formula_count": formulas,
            "merged_ranges": sorted(str(r) for r in ws.merged_cells.ranges),
            "tables": sorted(ws.tables.keys()),
            "charts": len(getattr(ws, "_charts", []) or []),
            "images": len(getattr(ws, "_images", []) or []),
            "data_validations": len(ws.data_validations.dataValidation)
            if ws.data_validations else 0,
            "conditional_formatting_blocks":
                len(list(ws.conditional_formatting)),
            "freeze_panes": ws.freeze_panes,
            "protection": bool(ws.protection.sheet),
            "defined_names": {name: dn.value for name, dn
                              in sorted(ws.defined_names.items())},
        }
        sheets.append(entry)
    for addresses in volatile.values():
        addresses.sort()

    doc = {
        "schema": WorkbookManifest.SCHEMA,
        "version": WorkbookManifest.VERSION,
        "sheets": sheets,
        "chartsheets": sorted(cs.title for cs in wb.chartsheets),
        "defined_names": {name: dn.value
                          for name, dn in sorted(wb.defined_names.items())},
        "external_links": len(wb._external_links),
        "volatile_functions": {
            "nondeterministic": {k: v for k, v in sorted(volatile.items())
                                 if k in VOLATILE_NONDETERMINISTIC},
            "deterministic": {k: v for k, v in sorted(volatile.items())
                              if k in VOLATILE_DETERMINISTIC},
        },
        "workbook_protection": bool(
            wb.security is not None
            and (wb.security.lockStructure or wb.security.lockWindows
                 or wb.security.workbookPassword
                 or wb.security.workbookPasswordCharacterSet
                 or wb.security.revisionsPassword)),
        "confession": _confession(wb),
        "preservation": _preservation(wb),
    }
    return WorkbookManifest(doc)


def _quoted(title):
    from openpyxl.utils.cell import quote_sheetname

    return quote_sheetname(title)


def _confession(wb):
    """What the package actually contains, from the package itself."""
    conf = {
        "chart_parts": 0,
        "drawing_parts": 0,
        "pivot_parts": 0,
        "vba_present": False,
        "custom_xml": False,
        "printer_settings": False,
        "worksheet_extensions": [],
        "at_risk_content": [],
    }
    source = getattr(wb, "_paper_source", None)
    if source:
        conf["source"] = "package"
        with zipfile.ZipFile(io.BytesIO(source)) as z:
            names = z.namelist()
        conf["chart_parts"] = sum(
            1 for n in names if n.startswith("xl/charts/")
            and n.endswith(".xml") and "/_rels/" not in n
            and not n.endswith(("colors.xml", "style.xml")))
        conf["drawing_parts"] = sum(
            1 for n in names
            if n.startswith("xl/drawings/") and n.endswith(".xml"))
        conf["pivot_parts"] = sum(
            1 for n in names if n.startswith(("xl/pivotTables/",
                                              "xl/pivotCache/")))
        conf["vba_present"] = "xl/vbaProject.bin" in names
        conf["custom_xml"] = any(n.startswith("customXml/") for n in names)
        conf["printer_settings"] = any(
            n.startswith("xl/printerSettings/") for n in names)
    else:
        # stock loads retain no archive: only model-visible facts remain,
        # and the model under-reports exactly the at-risk content — say so
        conf["source"] = "model (stock load retains no package; " \
            "part-level counts unavailable — open with preserve=True " \
            "for a package-accurate confession)"
        conf["chart_parts"] = sum(
            len(getattr(ws, "_charts", []) or []) for ws in wb.worksheets)
        conf["vba_present"] = wb.vba_archive is not None

    inventory = getattr(wb, "_paper_loss_inventory", None)
    if inventory:
        exts = sorted({
            loss["detail"].split(" extension")[0]
            for loss in inventory.losses
            if loss["kind"] == "worksheet-extension"})
        conf["worksheet_extensions"] = exts
        conf["at_risk_content"] = [
            {"kind": loss["kind"], "location": loss["location"],
             "detail": loss["detail"]}
            for loss in sorted(inventory.losses,
                               key=lambda l: (l["kind"], l["location"]))]
    return conf


def _preservation(wb):
    if getattr(wb, "_preserve", False):
        return {
            "mode": "preserve",
            "guarantee": (
                "untouched parts survive byte-identical (raw copy of the "
                "retained package); touched sheets are spliced, preserving "
                "unmodeled content; unsafe operations raise a typed "
                "PaperRefusal instead of proceeding lossily"),
            "at_risk": [],
        }
    inventory = getattr(wb, "_paper_loss_inventory", None)
    at_risk = sorted({loss["kind"] for loss in inventory.losses}) \
        if inventory else []
    return {
        "mode": "stock",
        "guarantee": (
            "the file is regenerated from the in-memory model on save: "
            "everything the model does not represent is rebuilt lossily or "
            "dropped (a LossySaveWarning enumerates it). Open with "
            "preserve=True for lossless custody."),
        "at_risk": at_risk,
    }


# ---------------------------------------------------------------------
# dependency sketch (feeds the Phase 6 guards)

class DependencySketch:
    """Coarse formula-dependency map: which cells feed which.

    ``references`` maps each formula cell (sheet-qualified A1) to the list
    of references its formula makes, as (sheet_title, bounds, raw) tuples —
    bounds may contain None for open-ended (whole-row/column) ranges.
    Table/structured references cannot be resolved to cells and are listed
    in ``unresolved`` (Phase 6 treats them as always-intersecting).
    """

    def __init__(self):
        self.references = {}      # "Model!B6" -> [(sheet, bounds, raw)]
        self.unresolved = {}      # "Model!B6" -> [raw operand]

    def cells_referencing(self, sheet_title, bounds):
        """Formula cells whose references intersect ``bounds`` on the given
        sheet — plus every cell with an unresolved (structured/table)
        reference, reported conservatively."""
        min_col, min_row, max_col, max_row = bounds
        title = sheet_title.casefold()   # Excel: sheet names case-insensitive
        hits = []
        for address, refs in self.references.items():
            for ref_sheet, ref_bounds, _raw in refs:
                if ref_sheet.casefold() != title:
                    continue
                if _intersects(ref_bounds, min_col, min_row, max_col, max_row):
                    hits.append(address)
                    break
        hits.extend(self.unresolved)
        return sorted(set(hits))

    def to_dict(self):
        return {
            "schema": "dependency_sketch",
            "version": 1,
            "references": {
                address: sorted(raw for (_s, _b, raw) in refs)
                for address, refs in sorted(self.references.items())
            },
            "unresolved": {address: sorted(raws) for address, raws
                           in sorted(self.unresolved.items())},
        }


def _intersects(bounds, min_col, min_row, max_col, max_row):
    b_min_col, b_min_row, b_max_col, b_max_row = bounds
    if b_min_col is None:
        b_min_col, b_max_col = 1, 1 << 20
    if b_min_row is None:
        b_min_row, b_max_row = 1, 1 << 22
    return not (b_max_col < min_col or b_min_col > max_col
                or b_max_row < min_row or b_min_row > max_row)


_SHEET_REF_RE = re.compile(r"^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$")


def dependency_sketch(wb):
    """Build a :class:`DependencySketch` from every formula in the model
    (tokenizer-based; PLAN Phase 4 — coarse is fine)."""
    from openpyxl.formula import Tokenizer

    sketch = DependencySketch()
    token_cache = {}
    for ws in wb.worksheets:
        for (row, col), cell in sorted(ws._cells.items()):
            if cell.data_type != "f" or not isinstance(cell._value, str):
                continue
            formula = cell._value
            address = "{0}!{1}".format(_quoted(ws.title), cell.coordinate)
            cached = token_cache.get(formula)
            if cached is None:
                try:
                    tokens = Tokenizer(formula).items
                except Exception:
                    sketch.unresolved.setdefault(address, []).append(formula)
                    continue
                operands = [t.value for t in tokens
                            if t.type == "OPERAND" and t.subtype == "RANGE"]
                # INDIRECT/OFFSET with computed-string targets leave no
                # RANGE operand at all (Batch-1 gate): the formula must
                # count as unresolved (always-intersecting), never as
                # invisible
                indirect = any(
                    t.type == "FUNC" and t.subtype == "OPEN"
                    and t.value.upper().lstrip("_XLFN.")
                    in ("INDIRECT(", "OFFSET(")
                    for t in tokens)
                cached = (operands, indirect)
                token_cache[formula] = cached
            operands, indirect = cached
            if indirect:
                sketch.unresolved.setdefault(address, []).append(formula)
            for raw in operands:
                _classify(sketch, wb, ws, address, raw)
    return sketch


def _classify(sketch, wb, ws, address, raw):
    ref = raw
    sheet_title = ws.title
    m = _SHEET_REF_RE.match(ref)
    if m:
        sheet_title = (m.group(1) or m.group(2))
        if m.group(1):
            sheet_title = sheet_title.replace("''", "'")
        ref = m.group(3)

    if "[" in raw or "]" in raw:
        # structured/table or external-workbook reference: not resolvable
        sketch.unresolved.setdefault(address, []).append(raw)
        return
    if ":" in sheet_title:
        # a 3-D span (Sheet1:Sheet3!A1) is not one sheet: classify it
        # conservatively as unresolved (always-intersecting) rather than
        # recording a phantom sheet name nothing can ever match
        # (Batch-1 gate: the phantom key silently defeated the recalc
        # guard and certification taint)
        sketch.unresolved.setdefault(address, []).append(raw)
        return

    plain = ref.replace("$", "")
    try:
        bounds = range_boundaries(plain)
    except Exception:
        # not A1-shaped: a defined name — expand via its destinations
        name = wb.defined_names.get(raw) or ws.defined_names.get(raw)
        if name is None:
            sketch.unresolved.setdefault(address, []).append(raw)
            return
        if name.value and "[" in name.value:
            # external-workbook reference hiding behind the name: the
            # expansion would drop the external marker (Batch-1 gate)
            sketch.unresolved.setdefault(address, []).append(raw)
            return
        try:
            for dest_sheet, dest_ref in name.destinations:
                dest_bounds = range_boundaries(dest_ref.replace("$", ""))
                sketch.references.setdefault(address, []).append(
                    (dest_sheet, dest_bounds, raw))
        except Exception:
            sketch.unresolved.setdefault(address, []).append(raw)
        return
    sketch.references.setdefault(address, []).append(
        (sheet_title, bounds, raw))
