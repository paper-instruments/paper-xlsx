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
import hashlib
import math
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import zipfile

from openpyxl.errors import OracleTimeoutError, OracleUnavailableError

_DARWIN_APP = "/Applications/LibreOffice.app/Contents/MacOS/soffice"

# Excel error tokens shared by oracle scans and preserve-mode cache write-back.
ERROR_TOKENS = frozenset((
    "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!",
    "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#UNKNOWN!",
    "#CONNECT!", "#BUSY!", "#PYTHON!", "#GETTING_DATA",
))

_MAX_MULTI_CELL_FORMULA_RESULTS = 1000000


def _artifact_sha256(data):
    return hashlib.sha256(data).hexdigest()

# Pinned comparison budget: ordinary numerics get four binary64 steps.
NUMERIC_ULPS = 4
DATE_SERIAL_ABS_FLOOR = 1e-11
_MIN_NORMAL_DOUBLE = float.fromhex("0x1.0p-1022")
_MIN_SUBNORMAL_DOUBLE = float.fromhex("0x0.0000000000001p-1022")


def _finite_double_ulp(value):
    value = abs(float(value))
    if value < _MIN_NORMAL_DOUBLE:
        return _MIN_SUBNORMAL_DOUBLE
    _mantissa, exponent = math.frexp(value)
    return math.ldexp(1.0, exponent - 53)

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
    from openpyxl.preserve.limits import read_bounded

    try:
        return read_bounded(source, context="oracle workbook")
    except TypeError as exc:
        raise ValueError(
            "file-like oracle sources must be seekable so the complete "
            "package can be read and the caller's cursor restored"
        ) from exc


def _read_source_with_custody(source):
    """Read ``source`` and retain path identity for any later delivery."""
    if isinstance(source, (str, os.PathLike)):
        from openpyxl.preserve.zipio import read_path_snapshot

        return read_path_snapshot(source, context="oracle workbook")
    return _read_source(source), None


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


def _workbook_content_type(data):
    from openpyxl.packaging.manifest import Manifest
    from openpyxl.xml.constants import ARC_CONTENT_TYPES, XLSM, XLSX, XLTM, XLTX
    from openpyxl.xml.functions import fromstring

    with zipfile.ZipFile(io.BytesIO(data)) as zin:
        package = Manifest.from_tree(fromstring(zin.read(ARC_CONTENT_TYPES)))
    for content_type in (XLTM, XLTX, XLSM, XLSX):
        if package.find(content_type):
            return content_type
    raise OracleUnavailableError("the package has no workbook part")


def _refuse_template_conversion(data):
    from openpyxl.errors import UnsupportedStructureError
    from openpyxl.xml.constants import XLTM, XLTX

    if _workbook_content_type(data) in (XLTM, XLTX):
        raise UnsupportedStructureError(
            "LibreOffice oracle recalculation does not have a proven "
            "format-preserving route for Excel templates (.xltx/.xltm); "
            "conversion would produce plain .xlsx content. Nothing was "
            "written.")


def _profile_uri(profile):
    return Path(profile).resolve().as_uri()


def _popen_session_kwargs():
    if os.name == "nt":
        flag = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        return {"creationflags": flag} if flag else {}
    return {"start_new_session": True}


def _terminate_process_tree(proc):
    if os.name == "nt":
        try:
            subprocess.run(
                ["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                check=False)
        except OSError:
            proc.kill()
        else:
            if proc.poll() is None:
                proc.kill()
        return

    import signal

    try:
        os.killpg(os.getpgid(proc.pid), signal.SIGKILL)
    except (ProcessLookupError, PermissionError, OSError):
        proc.kill()


def _recalculate_bytes(data, timeout, suffix=".xlsx", profile_root=None):
    """Round-trip ``data`` through headless LibreOffice; returns the
    recalculated package bytes. Never touches any caller path.
    ``profile_root``: reuse (and lazily seed) a persistent profile
    directory — the evaluate_many warm pool; None keeps the fully
    isolated per-call profile."""
    from openpyxl.preserve.zipguard import validate_package_bytes

    validate_package_bytes(data, context="oracle workbook")
    _refuse_template_conversion(data)
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
            "-env:UserInstallation={0}".format(_profile_uri(profile)),
            "--convert-to", "xlsx",
            "--outdir", outdir,
            tmp_input,
        ]
        popen_kwargs = _popen_session_kwargs()
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE, **popen_kwargs)
        try:
            stdout, _stderr = proc.communicate(timeout=timeout)
        except subprocess.TimeoutExpired:
            # soffice spawns children (oosplash -> soffice.bin), so terminate
            # the platform's process tree rather than only the direct child.
            _terminate_process_tree(proc)
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

    def __init__(self, cells_scanned, formula_cells, errors,
                 artifact_sha256=None):
        self.cells_scanned = cells_scanned
        self.formula_cells = formula_cells
        self.errors = errors          # [{"sheet", "cell", "value"}]
        self.artifact_sha256 = artifact_sha256

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
            "artifact_sha256": self.artifact_sha256,
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

    data, source_identity = _read_source_with_custody(source)

    destination = source if in_place else output_path
    if destination is not None and os.path.splitext(
            os.fspath(destination))[1].lower() in (".xltx", ".xltm"):
        from openpyxl.errors import UnsupportedStructureError

        raise UnsupportedStructureError(
            "recalc cannot write plain .xlsx content to an Excel template "
            "(.xltx/.xltm) destination. Nothing was written.")

    expected_identity = None
    if destination is not None:
        from openpyxl.preserve import zipio

        expected_identity = source_identity if in_place else \
            zipio.path_identity(output_path, allow_missing=True)

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
            if cell.data_type == "e" and isinstance(value, str) \
                    and value.strip() in ERROR_TOKENS:
                errors.append({"sheet": ws.title, "cell": cell.coordinate,
                               "value": value.strip()})

    def validate_source():
        if source_identity is not None:
            zipio._assert_path_identity(source_identity)

    if output_path is not None:
        from openpyxl.preserve import zipio
        zipio.deliver(
            recalculated, output_path, expected_identity=expected_identity,
            precommit=validate_source, postcommit=validate_source)
    elif in_place:
        from openpyxl.preserve import zipio
        zipio.deliver(
            recalculated, os.fspath(source),
            expected_identity=expected_identity,
            precommit=validate_source)

    return RecalcResult(cells_scanned, formula_cells, errors,
                        _artifact_sha256(recalculated))


class CertificationResult:

    SCHEMA = "oracle_certification"
    VERSION = 1
    CERTIFIED = "CERTIFIED"
    DIVERGED = "DIVERGED"
    BASELINE_UNVERIFIABLE = "BASELINE_UNVERIFIABLE"

    def __init__(self, status, checked, divergences, volatile_excluded,
                 unverifiable, external_excluded=None,
                 unsupported_excluded=None, input_excluded=None,
                 artifact_sha256=None):
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
        self.artifact_sha256 = artifact_sha256

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
            "artifact_sha256": self.artifact_sha256,
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
    temporal = (
        isinstance(cached, (_dt.datetime, _dt.date, _dt.time,
                            _dt.timedelta))
        or isinstance(computed, (_dt.datetime, _dt.date, _dt.time,
                                 _dt.timedelta))
    )
    if temporal:
        cached, computed = _serialize(cached), _serialize(computed)
    if isinstance(cached, bool) or isinstance(computed, bool):
        # a boolean only ever matches a boolean: Python's True == 1 would
        # otherwise mask real divergences
        if not (isinstance(cached, bool) and isinstance(computed, bool)):
            return False
        return cached is computed
    if isinstance(cached, (int, float)) and isinstance(computed, (int, float)):
        if isinstance(cached, int) and isinstance(computed, int):
            return cached == computed
        cached_float = float(cached)
        computed_float = float(computed)
        if not (math.isfinite(cached_float)
                and math.isfinite(computed_float)):
            return False
        if cached == computed:
            return True
        if isinstance(cached, int) and cached_float == computed_float:
            return False
        if isinstance(computed, int) and computed_float == cached_float:
            return False
        budget = NUMERIC_ULPS * max(
            _finite_double_ulp(cached_float),
            _finite_double_ulp(computed_float))
        if temporal:
            budget = max(DATE_SERIAL_ABS_FLOOR, budget)
        return abs(cached_float - computed_float) <= budget
    # text and error values compare exactly (pinned)
    return cached == computed


def _formula_results_match(cached, cached_type, computed, computed_type,
                           epoch=None):
    return cached_type == computed_type and _values_match(
        cached, computed, epoch=epoch)


def _formula_result_cells(wb_formulas, wb_cached):
    """Formula result cells, including every member of an array range.

    Each tuple ends with the formula anchor key used for exclusion reasons.
    Ordinary formulas are their own anchor.
    """
    from openpyxl.utils.cell import get_column_letter, range_boundaries
    from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

    result = []
    seen = set()
    for ws in wb_formulas.worksheets:
        cached_ws = wb_cached[ws.title]
        for (row, col), cell in sorted(ws._cells.items()):
            if cell.data_type != "f":
                continue
            coordinates = ((row, col),)
            if isinstance(cell._value, (ArrayFormula, DataTableFormula)) \
                    and cell._value.ref:
                min_col, min_row, max_col, max_row = range_boundaries(
                    cell._value.ref)
                count = ((max_row - min_row + 1)
                         * (max_col - min_col + 1))
                if count > _MAX_MULTI_CELL_FORMULA_RESULTS:
                    from openpyxl.errors import UnsupportedStructureError

                    raise UnsupportedStructureError(
                        "multi-cell formula at {0}!{1} declares {2} result "
                        "cells, past the {3}-cell oracle safety cap. Narrow "
                        "the formula range before certification. Nothing "
                        "was written.".format(
                            ws.title, cell.coordinate, count,
                            _MAX_MULTI_CELL_FORMULA_RESULTS))
                coordinates = (
                    (result_row, result_col)
                    for result_row in range(min_row, max_row + 1)
                    for result_col in range(min_col, max_col + 1)
                )
            anchor = (ws.title, row, col)
            for result_row, result_col in coordinates:
                key = (ws.title, result_row, result_col)
                if key in seen:
                    continue
                seen.add(key)
                cached_cell = cached_ws._cells.get((result_row, result_col))
                cached = cached_cell._value if cached_cell is not None else None
                cached_type = cached_cell.data_type \
                    if cached_cell is not None else None
                coordinate = "{0}{1}".format(
                    get_column_letter(result_col), result_row)
                result.append((ws.title, result_row, result_col, coordinate,
                               cached, cached_type, anchor))
    return result


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
    from openpyxl.preserve.perception import dependency_sketch
    from openpyxl.preserve.zipguard import validate_package_bytes

    validate_package_bytes(data, context="oracle certification input")

    wb_formulas = load_workbook(io.BytesIO(data), data_only=False)
    wb_cached = load_workbook(io.BytesIO(data), data_only=True)
    source_digest = _artifact_sha256(data)

    # excluded-with-reason: nondeterministic
    # volatiles, oracle-unsupported functions, unresolved dependencies, and
    # external-workbook references are all excluded from certification — so
    # DIVERGED keeps meaning "genuine disagreement" — with the reason
    # recorded, never a silent shrink of the check. Downstream cells inherit
    # the taint.
    sketch = dependency_sketch(wb_formulas)
    reasons = _exclusion_seeds(wb_formulas)
    for key in (input_seeds or ()):
        reasons[key] = "input"
    if input_seeds:
        reasons.update(_formula_input_dependencies(
            wb_formulas, set(input_seeds)))
    # An unresolved dependency may read any cell. Ordinary certification
    # cannot prove its value, while scenario certification must assume it may
    # read an input. Seed it before expanding multi-cell formula results so
    # every result cell and downstream formula inherits the same reason.
    unresolved_reason = (
        "input" if input_seeds else "unsupported:unresolved-reference")
    for address in sorted(sketch.unresolved):
        key = _address_key(address, wb_formulas)
        reasons.setdefault(key, unresolved_reason)
    formula_cells = _formula_result_cells(wb_formulas, wb_cached)
    for sheet, row, col, _coord, _cached, _cached_type, anchor \
            in formula_cells:
        if anchor in reasons:
            reasons.setdefault((sheet, row, col), reasons[anchor])
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

    def _bucket_reasons():
        volatile, external, unsupported, inputs = [], [], [], []
        for (sheet, row, col, coord, _cached, _cached_type, _anchor) \
                in formula_cells:
            reason = reasons.get((sheet, row, col))
            if reason is None:
                continue
            address = "{0}!{1}".format(sheet, coord)
            if reason == "external-link":
                external.append(address)
            elif reason == "input":
                inputs.append(address)
            elif reason.startswith("unsupported:"):
                unsupported.append("{0} ({1})".format(address, reason[12:]))
            else:
                volatile.append(address)
        return (sorted(volatile), sorted(external), sorted(unsupported),
                sorted(inputs))

    if not formula_cells:
        return CertificationResult(
            CertificationResult.BASELINE_UNVERIFIABLE, 0, [], [],
            [], artifact_sha256=source_digest), recalculated
    if all(cached is None or cached == ""
           for (_s, _r, _c, _coord, cached, _cached_type, _anchor)
           in formula_cells):
        # openpyxl-written files carry empty <v></v>: no answer key
        # exists — but the exclusion classes still ride along, so
        # write_back(allow_uncertified=True) never writes volatile/
        # external/unsupported cells
        vol, ext, uns, inp = _bucket_reasons()
        return CertificationResult(
            CertificationResult.BASELINE_UNVERIFIABLE, 0, [], vol,
            ["{0}!{1}".format(s, coord)
             for (s, _r, _c, coord, _v, _type, _anchor) in formula_cells],
            external_excluded=ext,
            unsupported_excluded=uns,
            input_excluded=inp,
            artifact_sha256=source_digest), recalculated

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
    for (sheet, row, col, coord, cached, cached_type, _anchor) \
            in formula_cells:
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
        computed_type = ccell.data_type if ccell is not None else None
        checked += 1
        if cached_type == "e" or computed_type == "e":
            divergences.append({
                "address": address, "cached": cached,
                "computed": computed, "reason": "formula-error"})
        elif not _formula_results_match(
                cached, cached_type, computed, computed_type,
                epoch=wb_formulas.epoch):
            divergences.append({"address": address, "cached": cached,
                                "computed": computed})

    complete_coverage = (
        checked > 0
        and not unverifiable
        and not volatile_excluded
        and not external_excluded
        and not unsupported_excluded
        and not input_excluded
    )
    status = (CertificationResult.DIVERGED if divergences
              else CertificationResult.CERTIFIED
              if complete_coverage
              else CertificationResult.BASELINE_UNVERIFIABLE)
    return CertificationResult(
        status, checked, divergences,
        sorted(volatile_excluded),
        sorted(unverifiable),
        external_excluded=sorted(external_excluded),
        unsupported_excluded=sorted(unsupported_excluded),
        input_excluded=sorted(input_excluded),
        artifact_sha256=source_digest), recalculated


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

_EXTERNAL_BOOK_RE = re.compile(r"\[[^\]]+\]")


def _is_external_reference(value):
    """True when a range operand's sheet qualifier names a workbook.

    The workbook token may follow a local/UNC/URL path. Structured table
    references put their brackets after any sheet separator, so they do not
    match.
    """
    if not isinstance(value, str) or "!" not in value:
        return False
    qualifier = value.lstrip("=").split("!", 1)[0]
    return _EXTERNAL_BOOK_RE.search(qualifier) is not None


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
    from openpyxl.worksheet.formula import DataTableFormula

    def formula_reason(formula, ws, seen=frozenset()):
        if not isinstance(formula, str):
            return None
        if not formula.startswith("="):
            formula = "=" + formula
        try:
            tokens = Tokenizer(formula).items
        except Exception:
            return "unparseable"
        for token in tokens:
            if token.type == "FUNC" and token.subtype == "OPEN":
                up = token.value.upper()
                if up.startswith("_XLFN."):
                    up = up[6:]
                if up in volatile_funcs:
                    return "volatile"
                if up in unsupported_funcs:
                    return "unsupported:" + up.rstrip("(")
            elif token.type == "OPERAND" and token.subtype == "RANGE":
                if _is_external_reference(token.value):
                    return "external-link"
                from openpyxl.preserve.perception import _defined_name

                name = _defined_name(wb_formulas, ws, token.value)
                if name is None or id(name) in seen:
                    continue
                reason = formula_reason(
                    name.value, ws, seen | {id(name)})
                if reason is not None:
                    return reason
        return None

    for ws in wb_formulas.worksheets:
        for (row, col), cell in ws._cells.items():
            if cell.data_type != "f":
                continue
            formula = cell._value
            if isinstance(formula, DataTableFormula):
                reasons[(ws.title, row, col)] = "unsupported:data-table"
                continue
            if not isinstance(formula, str):
                formula = getattr(formula, "text", None)
            if not isinstance(formula, str):
                continue
            key = (ws.title, row, col)
            reason = formula_reason(formula, ws)
            if reason is not None:
                reasons[key] = reason
    return reasons


def _formula_input_dependencies(wb, input_seeds):
    """Formula cells that read scenario inputs, including through names."""
    from openpyxl.formula import Tokenizer
    from openpyxl.preserve.perception import _defined_name
    from openpyxl.utils.cell import range_boundaries, range_to_tuple

    def reads_input(formula, ws, seen=frozenset()):
        if not isinstance(formula, str):
            return False
        formula = formula if formula.startswith("=") else "=" + formula
        try:
            tokens = Tokenizer(formula).items
        except Exception:
            return True
        for token in tokens:
            if token.type != "OPERAND" or token.subtype != "RANGE":
                continue
            name = _defined_name(wb, ws, token.value)
            if name is not None and id(name) not in seen:
                if reads_input(name.value, ws, seen | {id(name)}):
                    return True
                continue
            try:
                if "!" in token.value:
                    title, bounds = range_to_tuple(token.value)
                else:
                    title, bounds = ws.title, range_boundaries(token.value)
            except (TypeError, ValueError):
                continue
            min_col, min_row, max_col, max_row = bounds
            if any(seed_title.casefold() == title.casefold()
                   and min_row <= row <= max_row
                   and min_col <= col <= max_col
                   for seed_title, row, col in input_seeds):
                return True
        return False

    reasons = {}
    for ws in wb.worksheets:
        for (row, col), cell in ws._cells.items():
            if cell.data_type != "f":
                continue
            formula = cell._value if isinstance(cell._value, str) else \
                getattr(cell._value, "text", None)
            if reads_input(formula, ws):
                reasons[(ws.title, row, col)] = "input"
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
    ordered = sorted(
        tainted,
        key=lambda item: (item[0].casefold(), item[1], item[2], item[0]),
    )
    for (t_sheet, t_row, t_col) in ordered:
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

    def __init__(self, inputs, outputs, errors, certification,
                 artifact_sha256=None):
        self.inputs = inputs            # {address: value} as given
        self.outputs = outputs          # {address: computed value}
        self.errors = errors            # [{"sheet", "cell", "value"}]
        self.certification = certification
        self.artifact_sha256 = artifact_sha256

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
            "artifact_sha256": self.artifact_sha256,
        }

    def __repr__(self):
        return "Evaluation(status={0!r}, outputs={1})".format(
            self.status, len(self.outputs))


def _resolve_single_cell(wb, address):
    """(worksheet, row, col) for a sheet-qualified single-cell A1 address
    or a defined name resolving to one cell. Typed refusals otherwise."""
    from openpyxl.errors import AmbiguousTargetError, TargetNotFoundError
    from openpyxl.utils.cell import range_boundaries

    def _fail(msg):
        raise TargetNotFoundError(
            "{0!r}: {1}".format(address, msg))

    ref = address
    if "!" not in ref:
        dn = wb.defined_names.get(ref)
        if dn is None:
            local = [(ws, ws.defined_names[ref]) for ws in wb.worksheets
                     if ref in ws.defined_names]
            if len(local) > 1:
                options = [ws.title for ws, _dn in local]
                raise AmbiguousTargetError(
                    "{0!r}: sheet-scoped defined name exists on {1} sheets: "
                    "{2}. Use a sheet-qualified address.".format(
                        address, len(options), ", ".join(options)),
                    kind="ambiguous-name",
                    options=options,
                )
            if local:
                dn = local[0][1]
        if dn is None:
            _fail("not a sheet-qualified address and no defined name of "
                  "this name exists (defined names and 'Sheet1!B2' "
                  "addresses are accepted)")
        try:
            destinations = list(dn.destinations)
        except Exception as exc:
            _fail("the defined name cannot be resolved: {0}".format(exc))
        if len(destinations) != 1:
            _fail("the defined name resolves to {0} areas; single cells "
                  "only".format(len(destinations)))
        title, coord = destinations[0]
        ref = "'{0}'!{1}".format(title.replace("'", "''"),
                                 coord.replace("$", ""))
    if ref.startswith("'"):
        try:
            end = ref.index("'!", 1)
        except ValueError:
            _fail("has an unterminated quoted sheet name")
        title = ref[1:end].replace("''", "'")
        coord = ref[end + 2:]
    else:
        if "!" not in ref:
            _fail("does not resolve to a sheet-qualified address")
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
    from openpyxl.errors import TargetNotFoundError, UnsupportedStructureError

    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        raise TargetNotFoundError(
            "{0!r} is inside a merged range; write the input to the "
            "merge's anchor cell instead.".format(address))
    if cell.data_type == "f":
        raise UnsupportedStructureError(
            "{0!r} holds a formula; scenario inputs never overwrite "
            "calculations. Nothing was changed.".format(address),
            kind="input-is-calculation",
            anchor=address,
        )
    cell.value = value


def _scan_errors(recalculated):
    from openpyxl.reader.excel import load_workbook

    wb_values = load_workbook(io.BytesIO(recalculated), data_only=True)
    errors = []
    for ws in wb_values.worksheets:
        for (row, col), cell in sorted(ws._cells.items()):
            value = cell._value
            if cell.data_type == "e" and isinstance(value, str) \
                    and value.strip() in ERROR_TOKENS:
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
    return Evaluation(dict(set or {}), outputs, errors, certification,
                      _artifact_sha256(spliced))


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
        return Evaluation(dict(case or {}), outputs, errors, certification,
                          _artifact_sha256(spliced))

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
                 package_diff, artifact_sha256=None):
        self.cells_written = cells_written
        self.written = written                    # addresses updated
        self.verified_unchanged = verified_unchanged
        self.excluded = excluded                  # {address: reason}
        self.uncertified = uncertified
        self.cleared_fullcalc = cleared_fullcalc
        self.certification = certification
        self.package_diff = package_diff          # part names that changed
        self.artifact_sha256 = artifact_sha256

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
            "artifact_sha256": self.artifact_sha256,
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


def _cache_write_preserves_type(value, data_type, epoch):
    """Whether the value-only cache splicer preserves the OOXML type."""
    import datetime

    if isinstance(value, (datetime.datetime, datetime.date, datetime.time,
                          datetime.timedelta)):
        # The cache splicer emits temporal values as untyped Excel serials.
        # Without also proving the target style, reloading can expose a
        # number instead of the computed temporal value.
        return False
    from openpyxl.preserve.splice import _serialize_cached_value

    type_attr, _payload = _serialize_cached_value(value, epoch)
    if type_attr is None:
        serialized_type = "n"
    else:
        serialized_type = {b"b": "b", b"e": "e", b"str": "s"}.get(
            type_attr)
    return serialized_type == data_type


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
    data, source_identity = _read_source_with_custody(source)

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
    for a in certification.input_excluded:
        excluded[a] = "input-dependent"
    diverged = {d["address"] for d in certification.divergences}

    led = wb._paper_ledger
    written = []
    verified_unchanged = []
    covered = True
    unverifiable = set(certification.unverifiable)
    planned_writes = []
    for sheet, row, col, coord, _computed, computed_type, _anchor in \
            _formula_result_cells(wb, wb_computed):
        ws = wb[sheet]
        computed_ws = wb_computed[ws.title]
        address = "{0}!{1}".format(ws.title, coord)
        if address in excluded:
            covered = False
            continue
        ccell = computed_ws._cells.get((row, col))
        computed = ccell._value if ccell is not None else None
        computed_type = ccell.data_type if ccell is not None else None
        if computed is None:
            excluded[address] = "no-computed-value"
            covered = False
            continue
        if address in diverged:
            if computed_type == "e":
                excluded[address] = "formula-error"
                covered = False
                continue
            if not _cache_write_preserves_type(
                    computed, computed_type, wb_computed.epoch):
                excluded[address] = "computed-cache-type-not-writable"
                covered = False
                continue
            planned_writes.append((ws, row, col, computed, address))
            continue
        if address in unverifiable:
            if computed_type == "e":
                excluded[address] = "formula-error"
                covered = False
                continue
            if not _cache_write_preserves_type(
                    computed, computed_type, wb_computed.epoch):
                excluded[address] = "computed-cache-type-not-writable"
                covered = False
                continue
            planned_writes.append((ws, row, col, computed, address))
            continue
        # verified: the cache already equals the computed value
        verified_unchanged.append(address)

    for ws, row, col, computed, address in planned_writes:
        led.cache_writes.setdefault(ws, {})[(row, col)] = computed
        written.append(address)

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

    def validate_source():
        zipio._assert_path_identity(source_identity)

    zipio.deliver(
        out, os.fspath(source), expected_identity=source_identity,
        precommit=validate_source)
    return WriteBackResult(len(written), written, verified_unchanged,
                           excluded, uncertified, cleared, certification,
                           package_diff, _artifact_sha256(out))
