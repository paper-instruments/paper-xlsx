# paper-xlsx: the LibreOffice oracle

"""Bounded recalculation and certification via headless LibreOffice.

This library never calculates — a partial engine is a silent-wrongness
machine. Instead it routes to a real implementation of
Excel's semantics and reports MEASUREMENTS, never judgments:

- :func:`recalc` recomputes all cells on a TEMP COPY and scans for Excel
  error tokens, returning the skill-compatible JSON shape.
- :func:`certify` checks whether LibreOffice reproduces the file's own
  cached values (Excel's answer key for its current inputs) within the
  pinned tolerance, excluding cells downstream of nondeterministic volatile
  functions.

Driver rules, all measured:
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

# pinned numeric tolerance
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


def _recalculate_bytes(data, timeout, suffix=".xlsx", profile_root=None):
    """Round-trip ``data`` through headless LibreOffice; returns the
    recalculated package bytes. Never touches any caller path.
    ``profile_root``: reuse (and lazily seed) a persistent profile
    directory — the evaluate_many warm pool; None keeps the fully
    isolated per-call profile."""
    soffice = find_soffice()
    if soffice is None:
        raise OracleUnavailableError(
            "no LibreOffice installation found (looked for 'soffice' and "
            "'libreoffice' on PATH, and {0}). Install LibreOffice — e.g. "
            "apt-get install libreoffice-calc — to use the oracle. "
            "Preservation does not depend on it.".format(_DARWIN_APP))

    workdir = tempfile.mkdtemp(prefix="paper_oracle_")
    try:
        if profile_root is not None:
            profile = os.path.join(profile_root, "profile")
        else:
            profile = os.path.join(workdir, "profile")
        # pre-seed the profile with "always recalculate on load":
        # without it LibreOffice keeps whatever cached values the file
        # carries and the oracle would "compute" the cache under test
        # (measured on a tampered fixture; calcPr flags alone are not
        # honored by the headless converter)
        userdir = os.path.join(profile, "user")
        if not os.path.isdir(userdir):
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
                 unsupported_excluded=None, input_excluded=None):
        self.status = status
        self.checked = checked
        self.divergences = divergences          # [{"address", "cached", "computed"}]
        self.volatile_excluded = volatile_excluded
        self.unverifiable = unverifiable        # formula cells without a cache
        # excluded-with-reason: DIVERGED keeps meaning
        # "genuine disagreement" because known-unverifiable classes are
        # named, never silently checked-and-wrong or silently skipped
        self.external_excluded = external_excluded or []
        self.unsupported_excluded = unsupported_excluded or []
        # scenario-runner inputs and their downstream cells: legitimately different from the file's caches, never a
        # divergence
        self.input_excluded = input_excluded or []

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
            "input_excluded": list(self.input_excluded),
            "unverifiable": list(self.unverifiable),
        }

    def __repr__(self):
        return "CertificationResult({0}, checked={1}, diverged={2})".format(
            self.status, self.checked, len(self.divergences))


def _values_match(cached, computed, epoch=None):
    import datetime as _dt

    def _serialize(v):
        if isinstance(v, (_dt.datetime, _dt.date, _dt.time, _dt.timedelta)):
            from openpyxl.utils.datetime import WINDOWS_EPOCH, to_excel

            return to_excel(v, epoch if epoch is not None
                            else WINDOWS_EPOCH)
        return v

    # a date SERIAL and its parsed datetime are the same value: compare
    # numerically (write_back's own serials were judged
    # DIVERGED by its own certification)
    if isinstance(cached, (_dt.datetime, _dt.date, _dt.time,
                           _dt.timedelta))             or isinstance(computed, (_dt.datetime, _dt.date, _dt.time,
                                     _dt.timedelta)):
        cached, computed = _serialize(cached), _serialize(computed)
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
    result, _recalculated = _certify_impl(_read_source(source), timeout)
    return result


def _certify_impl(data, timeout, recalculated=None, input_seeds=None):
    """certify()'s engine, reusable with an EXISTING recalc result (one
    LibreOffice run serves evaluation + certification) and with extra
    taint seeds for scenario inputs. Returns (result, recalculated_bytes)
    — the bytes are None when certification early-returned before any
    recalc was needed."""
    from openpyxl.reader.excel import load_workbook
    from openpyxl.preserve.perception import (
        VOLATILE_NONDETERMINISTIC,
        dependency_sketch,
    )

    wb_formulas = load_workbook(io.BytesIO(data), data_only=False)
    wb_cached = load_workbook(io.BytesIO(data), data_only=True)

    # excluded-with-reason: nondeterministic
    # volatiles, oracle-unsupported functions, and external-workbook
    # references are all excluded from certification — so DIVERGED keeps
    # meaning "genuine disagreement" — with the reason recorded, never a
    # silent shrink of the check. Downstream cells inherit the taint.
    sketch = dependency_sketch(wb_formulas)
    reasons = _exclusion_seeds(wb_formulas)
    for key in (input_seeds or ()):
        reasons[key] = "input"
    if input_seeds:
        # unresolved references (INDIRECT/OFFSET, structured refs, 3-D
        # spans) are always-intersecting: with scenario inputs in play
        # they may read ANY input, so they and their downstream inherit
        # the input taint (a cell fed only through
        # INDIRECT escaped, and the certification falsely DIVERGED)
        for address in sketch.unresolved:
            key = _address_key(address, wb_formulas)
            reasons.setdefault(key, "input")
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

    def _bucket_reasons():
        volatile, external, unsupported = [], [], []
        for (sheet, row, col, coord, _cached) in formula_cells:
            reason = reasons.get((sheet, row, col))
            if reason is None:
                continue
            address = "{0}!{1}".format(sheet, coord)
            if reason == "external-link":
                external.append(address)
            elif reason == "input":
                pass
            elif reason.startswith("unsupported:"):
                unsupported.append("{0} ({1})".format(address, reason[12:]))
            else:
                volatile.append(address)
        return sorted(volatile), sorted(external), sorted(unsupported)

    if not formula_cells:
        return CertificationResult(
            CertificationResult.BASELINE_UNVERIFIABLE, 0, [], [],
            []), recalculated
    if all(cached is None or cached == ""
           for (_s, _r, _c, _coord, cached) in formula_cells):
        # openpyxl-written files carry empty <v></v>: no answer key
        # exists — but the exclusion classes still ride along, so
        # write_back(allow_uncertified=True) never writes volatile/
        # external/unsupported cells
        vol, ext, uns = _bucket_reasons()
        return CertificationResult(
            CertificationResult.BASELINE_UNVERIFIABLE, 0, [], vol,
            ["{0}!{1}".format(s, coord)
             for (s, _r, _c, coord, _v) in formula_cells],
            external_excluded=ext,
            unsupported_excluded=uns), recalculated

    if recalculated is None:
        recalculated = _recalculate_bytes(data, timeout)
    wb_computed = load_workbook(io.BytesIO(recalculated), data_only=True)

    checked = 0
    divergences = []
    volatile_excluded = []
    external_excluded = []
    unsupported_excluded = []
    input_excluded = []
    unverifiable = []
    for (sheet, row, col, coord, cached) in formula_cells:
        address = "{0}!{1}".format(sheet, coord)
        reason = reasons.get((sheet, row, col))
        if reason is not None:
            if reason == "external-link":
                external_excluded.append(address)
            elif reason == "input":
                input_excluded.append(address)
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
        if not _values_match(cached, computed, epoch=wb_formulas.epoch):
            divergences.append({"address": address, "cached": cached,
                                "computed": computed})

    status = (CertificationResult.DIVERGED if divergences
              else CertificationResult.CERTIFIED)
    return CertificationResult(
        status, checked, divergences,
        sorted(volatile_excluded),
        sorted(unverifiable),
        external_excluded=sorted(external_excluded),
        unsupported_excluded=sorted(unsupported_excluded),
        input_excluded=sorted(input_excluded)), recalculated


# functions LibreOffice's recalc cannot be trusted to reproduce (version-
# dependent or environment-dependent semantics). Conservative by design:
# an excluded cell is NAMED with its reason; a mis-checked cell would be a
# false DIVERGED.
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
                    # Excel serializes post-2007 functions as _xlfn.NAME(
                    # — the catalog match must see the bare name
                    # (the prefixed form evaded exclusion)
                    up = token.value.upper()
                    if up.startswith("_XLFN."):
                        up = up[6:]
                    if up in volatile_funcs:
                        reasons[key] = "volatile"
                        break                       # volatile outranks all
                    if up in unsupported_funcs and key not in reasons:
                        reasons[key] = "unsupported:" + up.rstrip("(")
                elif (token.type == "OPERAND"
                        and token.subtype == "RANGE"
                        and key not in reasons):
                    if _EXTERNAL_REF_RE.match(token.value):
                        reasons[key] = "external-link"
                    else:
                        # an external ref hiding behind a defined name
                        # resolve the name, tag the cell
                        name = wb_formulas.defined_names.get(token.value)                             or ws.defined_names.get(token.value)
                        if name is not None and name.value                                 and "[" in name.value:
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
    folded = sheet.casefold()        # Excel: sheet names case-insensitive
    for (t_sheet, t_row, t_col) in tainted:
        if t_sheet.casefold() != folded:
            continue
        if min_row <= t_row <= max_row and min_col <= t_col <= max_col:
            return (t_sheet, t_row, t_col)
    return None


# ---------------------------------------------------------------------
# scenario runner

class Evaluation:
    """One what-if run: inputs applied to a TEMP COPY through the spine,
    LibreOffice recalculated, outputs harvested. Pinned surface."""

    SCHEMA = "evaluation"
    VERSION = 1

    def __init__(self, inputs, outputs, errors, certification):
        self.inputs = inputs            # {address: value} as given
        self.outputs = outputs          # {address: computed value}
        self.errors = errors            # [{"sheet", "cell", "value"}]
        self.certification = certification

    @property
    def status(self):
        return "errors" if self.errors else "ok"

    @property
    def error_cells(self):
        return ["{0}!{1}".format(e["sheet"], e["cell"])
                for e in self.errors]

    def to_dict(self):
        return {
            "schema": self.SCHEMA,
            "version": self.VERSION,
            "status": self.status,
            "inputs": dict(self.inputs),
            "outputs": dict(self.outputs),
            "error_cells": list(self.error_cells),
            "certification": self.certification.to_dict()
            if self.certification is not None else None,
        }

    def __repr__(self):
        return "Evaluation(status={0!r}, outputs={1})".format(
            self.status, len(self.outputs))


def _resolve_single_cell(wb, address):
    """(worksheet, row, col) for a sheet-qualified single-cell A1 address
    or a defined name resolving to one cell. Typed refusals otherwise."""
    from openpyxl.errors import TargetNotFoundError
    from openpyxl.utils.cell import range_boundaries

    def _fail(msg):
        raise TargetNotFoundError(
            "{0!r}: {1}".format(address, msg))

    ref = address
    if "!" not in ref:
        dn = wb.defined_names.get(ref)
        if dn is None:
            for ws in wb.worksheets:
                if ref in ws.defined_names:
                    dn = ws.defined_names[ref]
                    break
        if dn is None:
            _fail("not a sheet-qualified address and no defined name of "
                  "this name exists (defined names and 'Sheet1!B2' "
                  "addresses are accepted)")
        destinations = list(dn.destinations)
        if len(destinations) != 1:
            _fail("the defined name resolves to {0} areas; single cells "
                  "only".format(len(destinations)))
        title, coord = destinations[0]
        ref = "'{0}'!{1}".format(title.replace("'", "''"),
                                 coord.replace("$", ""))
    if ref.startswith("'"):
        end = ref.index("'!", 1)
        title = ref[1:end].replace("''", "'")
        coord = ref[end + 2:]
    else:
        title, coord = ref.split("!", 1)
    matches = [ws for ws in wb.worksheets
               if ws.title.casefold() == title.casefold()]
    if not matches:
        _fail("sheet {0!r} does not exist".format(title))
    coord = coord.replace("$", "")
    try:
        min_col, min_row, max_col, max_row = range_boundaries(coord)
    except ValueError as exc:
        _fail(str(exc))
    if (min_col, min_row) != (max_col, max_row) or min_col is None \
            or min_row is None:
        _fail("must resolve to a SINGLE cell (got {0})".format(coord))
    return matches[0], min_row, min_col


def _set_input_cell(ws, row, col, value, address):
    """Assign one scenario input, refusing merged-cell interiors typed
    (a raw AttributeError is not a legal outcome)."""
    from openpyxl.cell.cell import MergedCell
    from openpyxl.errors import TargetNotFoundError

    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        raise TargetNotFoundError(
            "{0!r} is inside a merged range; write the input to the "
            "merge's anchor cell instead.".format(address))
    cell.value = value


def _scan_errors(recalculated):
    from openpyxl.reader.excel import load_workbook

    wb_values = load_workbook(io.BytesIO(recalculated), data_only=True)
    errors = []
    for ws in wb_values.worksheets:
        for (row, col), cell in sorted(ws._cells.items()):
            value = cell._value
            if isinstance(value, str) and value.strip() in ERROR_TOKENS:
                errors.append({"sheet": ws.title, "cell": cell.coordinate,
                               "value": value.strip()})
    return errors


def evaluate(source, set, read, *, timeout=120.0):
    """Scenario run against ``source`` (path/bytes/file-like): apply
    ``set`` inputs to a temp copy through the preserve spine, recalculate
    with LibreOffice, harvest ``read`` outputs. The source and every
    caller file stay untouched. One LibreOffice run serves both the
    outputs and the certification (original caches vs computed, with
    inputs' downstream cells excluded as ``input_excluded``)."""
    from openpyxl.reader.excel import load_workbook

    data = _read_source(source)
    wb = load_workbook(io.BytesIO(data), preserve=True)
    input_seeds = []
    for address in sorted(set or {}):
        ws, row, col = _resolve_single_cell(wb, address)
        _set_input_cell(ws, row, col, set[address], address)
        input_seeds.append((ws.title, row, col))
    buf = io.BytesIO()
    wb.save(buf)
    spliced = buf.getvalue()

    recalculated = _recalculate_bytes(spliced, timeout)
    wb_values = load_workbook(io.BytesIO(recalculated), data_only=True)
    outputs = {}
    for address in (read or []):
        ws, row, col = _resolve_single_cell(wb_values, address)
        cell = ws._cells.get((row, col))
        outputs[address] = cell._value if cell is not None else None
    errors = _scan_errors(recalculated)
    certification, _ = _certify_impl(spliced, timeout,
                                     recalculated=recalculated,
                                     input_seeds=input_seeds)
    return Evaluation(dict(set or {}), outputs, errors, certification)


def evaluate_many(source, cases, read, *, pool_size=2, timeout=120.0):
    """``evaluate`` for a list of input dicts, sharing warm LibreOffice
    profiles across cases (the pool is an implementation
    detail — ``pool_size`` per-thread-isolated profiles, created lazily,
    crash-replaced once, destroyed before return)."""
    import shutil as _shutil
    import tempfile as _tempfile
    import threading
    from concurrent.futures import ThreadPoolExecutor

    from openpyxl.reader.excel import load_workbook

    data = _read_source(source)
    if not cases:
        return []
    pool_size = max(1, min(int(pool_size), len(cases)))

    # spine work is model-side and fast: build every spliced copy first
    prepared = []
    for case in cases:
        wb = load_workbook(io.BytesIO(data), preserve=True)
        seeds = []
        for address in sorted(case or {}):
            ws, row, col = _resolve_single_cell(wb, address)
            _set_input_cell(ws, row, col, case[address], address)
            seeds.append((ws.title, row, col))
        buf = io.BytesIO()
        wb.save(buf)
        prepared.append((case, buf.getvalue(), seeds))

    local = threading.local()
    roots = []
    roots_lock = threading.Lock()

    def _profile_root():
        root = getattr(local, "root", None)
        if root is None:
            root = _tempfile.mkdtemp(prefix="paper_oracle_pool_")
            local.root = root
            with roots_lock:
                roots.append(root)
        return root

    def _run(prepared_case):
        case, spliced, seeds = prepared_case
        try:
            recalculated = _recalculate_bytes(
                spliced, timeout, profile_root=_profile_root())
        except (OracleUnavailableError, OracleTimeoutError):
            # crash-replaced once: a poisoned profile must not sink every
            # following case on this worker
            root = local.root
            local.root = None
            _shutil.rmtree(root, ignore_errors=True)
            recalculated = _recalculate_bytes(
                spliced, timeout, profile_root=_profile_root())
        wb_values = load_workbook(io.BytesIO(recalculated), data_only=True)
        outputs = {}
        for address in (read or []):
            ws, row, col = _resolve_single_cell(wb_values, address)
            cell = ws._cells.get((row, col))
            outputs[address] = cell._value if cell is not None else None
        errors = _scan_errors(recalculated)
        certification, _ = _certify_impl(spliced, timeout,
                                         recalculated=recalculated,
                                         input_seeds=seeds)
        return Evaluation(dict(case or {}), outputs, errors, certification)

    try:
        with ThreadPoolExecutor(max_workers=pool_size) as pool:
            return list(pool.map(_run, prepared))
    finally:
        with roots_lock:
            for root in roots:
                _shutil.rmtree(root, ignore_errors=True)


# ---------------------------------------------------------------------
# certification-gated write-back

class WriteBackResult:

    SCHEMA = "oracle_write_back"
    VERSION = 1

    def __init__(self, cells_written, written, verified_unchanged,
                 excluded, uncertified, cleared_fullcalc, certification,
                 package_diff):
        self.cells_written = cells_written
        self.written = written                    # addresses updated
        self.verified_unchanged = verified_unchanged
        self.excluded = excluded                  # {address: reason}
        self.uncertified = uncertified
        self.cleared_fullcalc = cleared_fullcalc
        self.certification = certification
        self.package_diff = package_diff          # part names that changed

    def to_dict(self):
        return {
            "schema": self.SCHEMA,
            "version": self.VERSION,
            "cells_written": self.cells_written,
            "written": list(self.written),
            "verified_unchanged": list(self.verified_unchanged),
            "excluded": dict(self.excluded),
            "uncertified": self.uncertified,
            "cleared_fullcalc": self.cleared_fullcalc,
            "certification": self.certification.to_dict()
            if self.certification is not None else None,
            "package_diff": list(self.package_diff),
        }

    def __repr__(self):
        return ("WriteBackResult(written={0}, excluded={1}, "
                "uncertified={2})".format(
                    self.cells_written, len(self.excluded),
                    self.uncertified))


def _clear_fullcalc(package_bytes):
    """Remove fullCalcOnLoad/forceFullCalc from the package's calcPr.
    Returns (new_bytes, changed)."""
    from openpyxl.preserve import crosspart

    with zipfile.ZipFile(io.BytesIO(package_bytes)) as zin:
        wb_part = _find_workbook_part_name(zin)
        payload = zin.read(wb_part)
    changed = False
    for attr in ("fullCalcOnLoad", "forceFullCalc"):
        root = crosspart.scan_small(payload, "workbook", max_depth=1)
        for child in root.children:
            if child.local() == "calcPr" and attr in child.attrs:
                start, end, head = crosspart._patch_attr(
                    payload, child, attr, "1", drop_value="1")
                payload = payload[:start] + head + payload[end:]
                changed = True
                break
    if not changed:
        return package_bytes, False
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(package_bytes)) as zin, \
            zipfile.ZipFile(out, "w") as zout:
        from openpyxl.preserve import zipio
        for info in zin.infolist():
            if info.filename == wb_part:
                zipio.write_entry(zout, wb_part, payload)
            else:
                zipio.copy_entry(zin, info, zout)
    return out.getvalue(), True


def write_back(source, *, timeout=120.0, allow_uncertified=False):
    """Recalculate a temp copy with LibreOffice and splice the computed
    cached values into the ORIGINAL package at ``source`` (a filesystem
    path) — values only, formulas untouched, LibreOffice bytes never
    enter the output (macro-safe by construction).

    Certification-gated: on DIVERGED or
    BASELINE_UNVERIFIABLE the call refuses unless
    ``allow_uncertified=True``, and then the result carries a loud
    ``uncertified=True``. Cells excluded from certification
    (volatile/external/unsupported and their downstream) are never
    written. fullCalcOnLoad clears only when every formula cell ended
    verified or written."""
    from openpyxl.errors import UnsupportedStructureError
    from openpyxl.reader.excel import load_workbook

    if not isinstance(source, (str, os.PathLike)):
        raise ValueError(
            "write_back writes the recalculated values INTO the source; "
            "pass a filesystem path")
    data = _read_source(source)

    certification, recalculated = _certify_impl(data, timeout)
    uncertified = certification.status != CertificationResult.CERTIFIED
    if uncertified and not allow_uncertified:
        raise UnsupportedStructureError(
            "write-back is certification-gated and this workbook is {0} "
            "({1} divergences, {2} unverifiable). Nothing was written. "
            "Inspect certify(...).to_dict(), or pass "
            "allow_uncertified=True to write anyway with a loud "
            "uncertified stamp.".format(
                certification.status, len(certification.divergences),
                len(certification.unverifiable)))
    if recalculated is None:
        recalculated = _recalculate_bytes(data, timeout)

    wb = load_workbook(io.BytesIO(data), preserve=True)
    wb_computed = load_workbook(io.BytesIO(recalculated), data_only=True)

    excluded = {}
    for a in certification.volatile_excluded:
        excluded[a] = "volatile"
    for a in certification.external_excluded:
        excluded[a] = "external-link"
    for a in certification.unsupported_excluded:
        addr = a.rsplit(" (", 1)[0]
        excluded[addr] = "oracle-unsupported"
    diverged = {d["address"] for d in certification.divergences}

    led = wb._paper_ledger
    written = []
    verified_unchanged = []
    covered = True
    for ws in wb.worksheets:
        computed_ws = wb_computed[ws.title]
        for (row, col), cell in sorted(ws._cells.items()):
            if cell.data_type != "f":
                continue
            address = "{0}!{1}".format(ws.title, cell.coordinate)
            if address in excluded:
                covered = False
                continue
            ccell = computed_ws._cells.get((row, col))
            computed = ccell._value if ccell is not None else None
            if computed is None:
                excluded[address] = "no-computed-value"
                covered = False
                continue
            if address in diverged:
                # reachable only under allow_uncertified
                led.cache_writes.setdefault(ws, {})[(row, col)] = computed
                written.append(address)
                continue
            if address in set(certification.unverifiable):
                # previously cache-less: the whole point
                led.cache_writes.setdefault(ws, {})[(row, col)] = computed
                written.append(address)
                continue
            # verified: the cache already equals the computed value
            verified_unchanged.append(address)

    buf = io.BytesIO()
    wb.save(buf)
    out = buf.getvalue()
    cleared = False
    if covered and not uncertified:
        # an UNCERTIFIED write must leave the recalc-on-load flag alone:
        # clearing it would make Excel trust caches nobody verified
        out, cleared = _clear_fullcalc(out)

    package_diff = []
    with zipfile.ZipFile(io.BytesIO(data)) as za, \
            zipfile.ZipFile(io.BytesIO(out)) as zb:
        names_a, names_b = set(za.namelist()), set(zb.namelist())
        for name in sorted(names_a | names_b):
            if name not in names_a or name not in names_b \
                    or za.read(name) != zb.read(name):
                package_diff.append(name)

    from openpyxl.preserve import zipio
    zipio.deliver(out, os.fspath(source))
    return WriteBackResult(len(written), written, verified_unchanged,
                           excluded, uncertified, cleared, certification,
                           package_diff)
