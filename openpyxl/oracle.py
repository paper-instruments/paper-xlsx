# paper-xlsx: the LibreOffice oracle (PLAN Phase 5; PR-0 §7/D17)

"""Bounded recalculation and certification via headless LibreOffice.

This library never calculates — a partial engine is a silent-wrongness
machine (CONVENTIONS §5). Instead it routes to a real implementation of
Excel's semantics and reports MEASUREMENTS, never judgments:

- :func:`recalc` recomputes all cells on a TEMP COPY and scans for Excel
  error tokens, returning the skill-compatible JSON shape.
- :func:`certify` checks whether LibreOffice reproduces the file's own
  cached values (Excel's answer key for its current inputs) within the
  pinned tolerance, excluding cells downstream of nondeterministic volatile
  functions (CONVENTIONS §3.7).

Driver rules, all measured in Phase 0 (OPEN-QUESTIONS Q10):
the caller's file is NEVER handed to LibreOffice (temp copies only — tested
invariant); every invocation gets its own ``-env:UserInstallation`` profile
(shared profiles fail nondeterministically); success is ``returncode == 0
AND the output file exists`` (soffice exits 0 on unloadable input); stderr
is never parsed (successful runs emit noise); timeouts kill the whole
process group. Custody never depends on this module: everything
preservation-related works with no LibreOffice installed.
"""

import io
import os
import re
import shutil
import subprocess
import tempfile
import zipfile

from openpyxl.errors import OracleTimeoutError, OracleUnavailableError

_DARWIN_APP = "/Applications/LibreOffice.app/Contents/MacOS/soffice"

# Excel error tokens (the reference skill's set)
ERROR_TOKENS = frozenset((
    "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!",
))

# pinned numeric tolerance (CONVENTIONS §3.7)
REL_TOL = 1e-9
ABS_TOL = 1e-11

_RECALC_ALWAYS_XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" \
xmlns:xs="http://www.w3.org/2001/XMLSchema">
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load">\
<prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""


def find_soffice():
    """Locate the LibreOffice binary, or None."""
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    if os.path.exists(_DARWIN_APP):
        return _DARWIN_APP
    return None


def available():
    """True when a LibreOffice installation can be found."""
    return find_soffice() is not None


def _read_source(source):
    if isinstance(source, (bytes, bytearray)):
        return bytes(source)
    if hasattr(source, "read"):
        return source.read()
    with open(source, "rb") as f:
        return f.read()


def _with_forced_recalc(data):
    """Patch the TEMP COPY's calcPr so LibreOffice actually recalculates:
    without fullCalcOnLoad/forceFullCalc, LO keeps whatever cached values
    the file carries — the oracle would then 'compute' the very cache it is
    supposed to check (measured on a tampered fixture)."""
    import io as _io

    from openpyxl.preserve import crosspart, zipio

    with zipfile.ZipFile(_io.BytesIO(data)) as zin:
        wb_part = _find_workbook_part_name(zin)
        original = zin.read(wb_part)
        root = crosspart.scan_small(original, "workbook", max_depth=2)
        calc_nodes = [c for c in root.children if c.local() == "calcPr"]
        if calc_nodes:
            node = calc_nodes[0]
            edits = [crosspart._patch_attr(original, node, "fullCalcOnLoad", "1")]
            patched = crosspart.apply_edits(original, edits)
            root2 = crosspart.scan_small(patched, "workbook", max_depth=2)
            node2 = [c for c in root2.children if c.local() == "calcPr"][0]
            patched = crosspart.apply_edits(
                patched,
                [crosspart._patch_attr(patched, node2, "forceFullCalc", "1")])
        else:
            blob = b'<calcPr calcId="124519" fullCalcOnLoad="1" forceFullCalc="1"/>'
            patched = crosspart.apply_edits(
                original,
                [crosspart._wb_insert_edit(root, {}, "calcPr", blob)])

        def build(zout):
            for info in zin.infolist():
                if info.filename == wb_part:
                    zipio.write_entry(zout, wb_part, patched)
                else:
                    zipio.copy_entry(zin, info, zout)

        return zipio.build_archive_bytes(build)


def _find_workbook_part_name(zin):
    from openpyxl.packaging.manifest import Manifest
    from openpyxl.xml.constants import ARC_CONTENT_TYPES, XLSM, XLSX, XLTM, XLTX
    from openpyxl.xml.functions import fromstring

    package = Manifest.from_tree(fromstring(zin.read(ARC_CONTENT_TYPES)))
    for ct in (XLTM, XLTX, XLSM, XLSX):
        part = package.find(ct)
        if part:
            return part.PartName[1:]
    raise OracleUnavailableError("the package has no workbook part")


def _recalculate_bytes(data, timeout, suffix=".xlsx"):
    """Round-trip ``data`` through headless LibreOffice; returns the
    recalculated package bytes. Never touches any caller path."""
    soffice = find_soffice()
    if soffice is None:
        raise OracleUnavailableError(
            "no LibreOffice installation found (looked for 'soffice' and "
            "'libreoffice' on PATH, and {0}). Install LibreOffice — e.g. "
            "apt-get install libreoffice-calc — to use the oracle. "
            "Preservation does not depend on it.".format(_DARWIN_APP))

    workdir = tempfile.mkdtemp(prefix="paper_oracle_")
    try:
        profile = os.path.join(workdir, "profile")
        # pre-seed the fresh profile with "always recalculate on load":
        # without it LibreOffice keeps whatever cached values the file
        # carries and the oracle would "compute" the cache under test
        # (measured on a tampered fixture; calcPr flags alone are not
        # honored by the headless converter)
        userdir = os.path.join(profile, "user")
        os.makedirs(userdir)
        with open(os.path.join(userdir, "registrymodifications.xcu"),
                  "w") as f:
            f.write(_RECALC_ALWAYS_XCU)
        outdir = os.path.join(workdir, "out")
        os.makedirs(outdir)
        tmp_input = os.path.join(workdir, "input" + suffix)
        with open(tmp_input, "wb") as f:
            f.write(_with_forced_recalc(data))

        cmd = [
            soffice,
            "--headless",
            "-env:UserInstallation=file://{0}".format(profile),
            "--convert-to", "xlsx",
            "--outdir", outdir,
            tmp_input,
        ]
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE,
                                start_new_session=True)
        try:
            stdout, _stderr = proc.communicate(timeout=timeout)
        except subprocess.TimeoutExpired:
            # kill the whole process group: soffice spawns children
            # (oosplash -> soffice.bin) that a child-only kill orphans
            import signal

            try:
                os.killpg(os.getpgid(proc.pid), signal.SIGKILL)
            except (ProcessLookupError, PermissionError, OSError):
                proc.kill()
            proc.wait()
            raise OracleTimeoutError(
                "LibreOffice did not finish within {0:g}s. The process "
                "group was killed; no caller file was touched.".format(
                    timeout))
        proc = type("_Result", (), {"returncode": proc.returncode,
                                    "stdout": stdout})()
        out_path = os.path.join(outdir, "input.xlsx")
        # rc==0 alone lies: soffice exits 0 on unloadable input
        if proc.returncode != 0 or not os.path.exists(out_path):
            raise OracleUnavailableError(
                "LibreOffice could not recalculate the file (rc={0}, "
                "output {1}): stdout={2!r}".format(
                    proc.returncode,
                    "present" if os.path.exists(out_path) else "missing",
                    proc.stdout[-300:]))
        with open(out_path, "rb") as f:
            return f.read()
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


class RecalcResult:

    SCHEMA = "oracle_recalc"
    VERSION = 1

    def __init__(self, cells_scanned, formula_cells, errors):
        self.cells_scanned = cells_scanned
        self.formula_cells = formula_cells
        self.errors = errors          # [{"sheet", "cell", "value"}]

    @property
    def status(self):
        return "errors" if self.errors else "ok"

    def to_dict(self):
        return {
            "schema": self.SCHEMA,
            "version": self.VERSION,
            "status": self.status,
            "cells_scanned": self.cells_scanned,
            "formula_cells": self.formula_cells,
            "error_cells": len(self.errors),
            "errors": list(self.errors),
        }

    def __repr__(self):
        return "RecalcResult(status={0!r}, errors={1})".format(
            self.status, len(self.errors))


def recalc(source, *, output_path=None, in_place=False, timeout=120.0):
    """Recalculate every cell with LibreOffice and scan for error tokens.

    ``source``: path, bytes, or binary file-like. The source is NEVER
    handed to LibreOffice or modified — unless ``in_place=True``, in which
    case the recalculated bytes replace the source path atomically after a
    successful run. ``output_path`` writes them elsewhere instead. At most
    one of the two may be given.
    """
    if output_path is not None and in_place:
        raise ValueError("pass either output_path or in_place=True, not both")
    if in_place and not isinstance(source, (str, os.PathLike)):
        raise ValueError("in_place=True requires a filesystem path source")

    data = _read_source(source)

    if output_path is not None or in_place:
        # LibreOffice converts to plain xlsx: writing that output over a
        # macro workbook would strip the entire VBA project silently
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            if "xl/vbaProject.bin" in z.namelist():
                from openpyxl.errors import UnsupportedStructureError

                raise UnsupportedStructureError(
                    "recalc output for a macro-enabled workbook would strip "
                    "its VBA project (LibreOffice converts to plain .xlsx). "
                    "Run recalc() without output_path/in_place for the "
                    "error scan, or use certify() for value checking. "
                    "Nothing was written.")
    recalculated = _recalculate_bytes(data, timeout)

    from openpyxl.reader.excel import load_workbook

    wb_values = load_workbook(io.BytesIO(recalculated), data_only=True)
    wb_formulas = load_workbook(io.BytesIO(recalculated), data_only=False)

    cells_scanned = 0
    formula_cells = 0
    errors = []
    for ws in wb_values.worksheets:
        ws_f = wb_formulas[ws.title]
        for (row, col), cell in sorted(ws._cells.items()):
            cells_scanned += 1
            fcell = ws_f._cells.get((row, col))
            if fcell is not None and fcell.data_type == "f":
                formula_cells += 1
            value = cell._value
            if isinstance(value, str) and value.strip() in ERROR_TOKENS:
                errors.append({"sheet": ws.title, "cell": cell.coordinate,
                               "value": value.strip()})

    if output_path is not None:
        from openpyxl.preserve import zipio
        zipio.deliver(recalculated, output_path)
    elif in_place:
        from openpyxl.preserve import zipio
        zipio.deliver(recalculated, os.fspath(source))

    return RecalcResult(cells_scanned, formula_cells, errors)


class CertificationResult:

    SCHEMA = "oracle_certification"
    VERSION = 1
    CERTIFIED = "CERTIFIED"
    DIVERGED = "DIVERGED"
    BASELINE_UNVERIFIABLE = "BASELINE_UNVERIFIABLE"

    def __init__(self, status, checked, divergences, volatile_excluded,
                 unverifiable, external_excluded=None,
                 unsupported_excluded=None):
        self.status = status
        self.checked = checked
        self.divergences = divergences          # [{"address", "cached", "computed"}]
        self.volatile_excluded = volatile_excluded
        self.unverifiable = unverifiable        # formula cells without a cache
        # excluded-with-reason (PLAN-v0.1 1.7): DIVERGED keeps meaning
        # "genuine disagreement" because known-unverifiable classes are
        # named, never silently checked-and-wrong or silently skipped
        self.external_excluded = external_excluded or []
        self.unsupported_excluded = unsupported_excluded or []

    def to_dict(self):
        return {
            "schema": self.SCHEMA,
            "version": self.VERSION,
            "status": self.status,
            "cells_checked": self.checked,
            "divergences": list(self.divergences),
            "volatile_excluded": list(self.volatile_excluded),
            "external_excluded": list(self.external_excluded),
            "unsupported_excluded": list(self.unsupported_excluded),
            "unverifiable": list(self.unverifiable),
        }

    def __repr__(self):
        return "CertificationResult({0}, checked={1}, diverged={2})".format(
            self.status, self.checked, len(self.divergences))


def _values_match(cached, computed):
    if isinstance(cached, bool) or isinstance(computed, bool):
        # a boolean only ever matches a boolean: Python's True == 1 would
        # otherwise mask real divergences
        if not (isinstance(cached, bool) and isinstance(computed, bool)):
            return False
        return cached is computed
    if isinstance(cached, (int, float)) and isinstance(computed, (int, float)):
        diff = abs(float(cached) - float(computed))
        return diff <= max(ABS_TOL, REL_TOL * max(abs(float(cached)),
                                                  abs(float(computed))))
    # text and error values compare exactly (pinned)
    return cached == computed


def certify(source, *, timeout=120.0):
    """The divergence check: does LibreOffice reproduce the file's own
    cached values? Pre-flights on an untouched temp copy; the caller's file
    is never modified. Returns measurements, never judgments."""
    from openpyxl.reader.excel import load_workbook
    from openpyxl.preserve.perception import (
        VOLATILE_NONDETERMINISTIC,
        dependency_sketch,
    )

    data = _read_source(source)

    wb_formulas = load_workbook(io.BytesIO(data), data_only=False)
    wb_cached = load_workbook(io.BytesIO(data), data_only=True)

    # excluded-with-reason (§3.7 + PLAN-v0.1 1.7): nondeterministic
    # volatiles, oracle-unsupported functions, and external-workbook
    # references are all excluded from certification — so DIVERGED keeps
    # meaning "genuine disagreement" — with the reason recorded, never a
    # silent shrink of the check. Downstream cells inherit the taint.
    sketch = dependency_sketch(wb_formulas)
    reasons = _exclusion_seeds(wb_formulas)
    tainted = set(reasons)
    changed = True
    while changed:
        changed = False
        for address, refs in sorted(sketch.references.items()):
            key = _address_key(address, wb_formulas)
            if key in tainted:
                continue
            for ref_sheet, bounds, _raw in refs:
                hit = _bounds_hit_tainted(ref_sheet, bounds, tainted)
                if hit is not None:
                    tainted.add(key)
                    reasons[key] = reasons[hit]
                    changed = True
                    break

    formula_cells = []
    for ws in wb_formulas.worksheets:
        cached_ws = wb_cached[ws.title]
        for (row, col), cell in sorted(ws._cells.items()):
            if cell.data_type != "f":
                continue
            ccell = cached_ws._cells.get((row, col))
            cached = ccell._value if ccell is not None else None
            formula_cells.append((ws.title, row, col, cell.coordinate,
                                  cached))

    if not formula_cells:
        return CertificationResult(
            CertificationResult.BASELINE_UNVERIFIABLE, 0, [], [], [])
    if all(cached is None or cached == ""
           for (_s, _r, _c, _coord, cached) in formula_cells):
        # openpyxl-written files carry empty <v></v>: no answer key exists
        return CertificationResult(
            CertificationResult.BASELINE_UNVERIFIABLE, 0, [], [],
            ["{0}!{1}".format(s, coord)
             for (s, _r, _c, coord, _v) in formula_cells])

    recalculated = _recalculate_bytes(data, timeout)
    wb_computed = load_workbook(io.BytesIO(recalculated), data_only=True)

    checked = 0
    divergences = []
    volatile_excluded = []
    external_excluded = []
    unsupported_excluded = []
    unverifiable = []
    for (sheet, row, col, coord, cached) in formula_cells:
        address = "{0}!{1}".format(sheet, coord)
        reason = reasons.get((sheet, row, col))
        if reason is not None:
            if reason == "external-link":
                external_excluded.append(address)
            elif reason.startswith("unsupported:"):
                unsupported_excluded.append(
                    "{0} ({1})".format(address, reason[12:]))
            else:
                volatile_excluded.append(address)
            continue
        if cached is None or cached == "":
            unverifiable.append(address)
            continue
        computed_ws = wb_computed[sheet]
        ccell = computed_ws._cells.get((row, col))
        computed = ccell._value if ccell is not None else None
        checked += 1
        if not _values_match(cached, computed):
            divergences.append({"address": address, "cached": cached,
                                "computed": computed})

    status = (CertificationResult.DIVERGED if divergences
              else CertificationResult.CERTIFIED)
    return CertificationResult(status, checked, divergences,
                               sorted(volatile_excluded),
                               sorted(unverifiable),
                               external_excluded=sorted(external_excluded),
                               unsupported_excluded=sorted(
                                   unsupported_excluded))


# functions LibreOffice's recalc cannot be trusted to reproduce (version-
# dependent or environment-dependent semantics). Conservative by design:
# an excluded cell is NAMED with its reason; a mis-checked cell would be a
# false DIVERGED (PLAN-v0.1 1.7).
ORACLE_UNSUPPORTED_FUNCS = frozenset([
    "LAMBDA", "LET", "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL",
    "MAKEARRAY", "ISOMITTED",
    "STOCKHISTORY", "RTD", "WEBSERVICE", "FILTERXML", "IMAGE", "PY",
    "CUBEVALUE", "CUBEMEMBER", "CUBESET", "CUBESETCOUNT",
    "CUBERANKEDMEMBER", "CUBEMEMBERPROPERTY", "CUBEKPIMEMBER",
])

# an operand that opens with a bracketed workbook token is an external-
# workbook reference ([1]Sheet!A1, '[Budget.xlsx]Sheet One'!A1); table
# structured refs (Table1[Col]) never START with the bracket
_EXTERNAL_REF_RE = re.compile(r"^'?\[[^\]]+\]")


def _exclusion_seeds(wb_formulas):
    """{(sheet, row, col): reason} for every formula cell certification
    must exclude: volatile, unsupported function, external reference, or
    unparseable. Tokenizer-precise — string literals that merely contain
    a trigger name do not taint."""
    from openpyxl.formula import Tokenizer
    from openpyxl.preserve.perception import VOLATILE_NONDETERMINISTIC

    volatile_funcs = {name + "(" for name in VOLATILE_NONDETERMINISTIC}
    unsupported_funcs = {name + "(" for name in ORACLE_UNSUPPORTED_FUNCS}
    reasons = {}
    for ws in wb_formulas.worksheets:
        for (row, col), cell in ws._cells.items():
            if cell.data_type != "f" or not isinstance(cell._value, str):
                continue
            key = (ws.title, row, col)
            try:
                tokens = Tokenizer(cell._value).items
            except Exception:
                reasons[key] = "unparseable"
                continue
            for token in tokens:
                if token.type == "FUNC" and token.subtype == "OPEN":
                    up = token.value.upper()
                    if up in volatile_funcs:
                        reasons[key] = "volatile"
                        break                       # volatile outranks all
                    if up in unsupported_funcs and key not in reasons:
                        reasons[key] = "unsupported:" + up.rstrip("(")
                elif (token.type == "OPERAND"
                        and token.subtype == "RANGE"
                        and _EXTERNAL_REF_RE.match(token.value)
                        and key not in reasons):
                    reasons[key] = "external-link"
    return reasons


def _address_key(address, wb):
    """('Sheet', row, col) from a sketch address like "'Sheet'!B6"."""
    from openpyxl.utils.cell import coordinate_to_tuple

    title, _, coord = address.rpartition("!")
    if title.startswith("'") and title.endswith("'"):
        title = title[1:-1].replace("''", "'")
    row, col = coordinate_to_tuple(coord)
    return (title, row, col)


def _bounds_hit_tainted(sheet, bounds, tainted):
    min_col, min_row, max_col, max_row = bounds
    if min_col is None:
        min_col, max_col = 1, 1 << 20
    if min_row is None:
        min_row, max_row = 1, 1 << 22
    for (t_sheet, t_row, t_col) in tainted:
        if t_sheet != sheet:
            continue
        if min_row <= t_row <= max_row and min_col <= t_col <= max_col:
            return (t_sheet, t_row, t_col)
    return None
