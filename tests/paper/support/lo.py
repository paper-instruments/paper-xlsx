"""LibreOffice test driver (lo_smoke tier).

Operational rules distilled from the Phase-0 oracle probes
(agent_docs/OPEN-QUESTIONS.md Q10), which this module must never violate:

- The caller's file is NEVER handed to LibreOffice: every conversion copies
  the input into a fresh temp dir first (asserted by tests).
- Every invocation gets its own ``-env:UserInstallation`` profile — shared
  profiles fail nondeterministically (hard DeploymentException aborts or
  silent IPC delegation).
- Success predicate is ``returncode == 0 AND output file exists``: soffice
  exits 0 on unloadable input, and successful runs may emit stderr noise.
- Timeouts kill the whole process group (macOS/Debian both spawn children).

This is a test helper; the production oracle (Phase 5) re-implements the same
rules as package code with typed errors.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile

import pytest

_DARWIN_FALLBACK = "/Applications/LibreOffice.app/Contents/MacOS/soffice"


def find_soffice():
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    if os.path.exists(_DARWIN_FALLBACK):
        return _DARWIN_FALLBACK
    return None


def lo_available():
    return find_soffice() is not None


def require_lo():
    """Skip loudly when LibreOffice is absent — unless PAPER_REQUIRE_LO=1,
    in which case absence is a hard failure (the LO-equipped CI leg)."""
    if lo_available():
        return
    if os.environ.get("PAPER_REQUIRE_LO") == "1":
        pytest.fail(
            "PAPER_REQUIRE_LO=1 but no LibreOffice found (looked for 'soffice', "
            "'libreoffice' on PATH and {0}). Install libreoffice-calc.".format(_DARWIN_FALLBACK)
        )
    pytest.skip(
        "LibreOffice not installed — lo_smoke assertion SKIPPED, not verified. "
        "Install LibreOffice (apt: libreoffice-calc) or run on the LO-equipped CI leg."
    )


class LOConversionError(Exception):
    pass


def lo_convert(src_path, fmt="xlsx", timeout=120.0):
    """Convert a TEMP COPY of ``src_path`` with headless LibreOffice.

    Returns the converted file's bytes. Never touches ``src_path`` beyond the
    initial read. Raises LOConversionError on failure or missing output.
    """
    soffice = find_soffice()
    if soffice is None:
        raise LOConversionError("LibreOffice not available")
    workdir = tempfile.mkdtemp(prefix="paper_lo_")
    try:
        profile = os.path.join(workdir, "profile")
        os.makedirs(profile)
        indir = os.path.join(workdir, "in")
        outdir = os.path.join(workdir, "out")
        os.makedirs(indir)
        os.makedirs(outdir)
        base = os.path.basename(src_path)
        tmp_input = os.path.join(indir, base)
        shutil.copyfile(src_path, tmp_input)

        cmd = [
            soffice,
            "--headless",
            "-env:UserInstallation=file://{0}".format(profile),
            "--convert-to",
            fmt,
            "--outdir",
            outdir,
            tmp_input,
        ]
        try:
            proc = subprocess.run(
                cmd,
                capture_output=True,
                timeout=timeout,
                start_new_session=True,
            )
        except subprocess.TimeoutExpired:
            raise LOConversionError(
                "LibreOffice timed out after {0:g}s converting {1}".format(timeout, base)
            )
        stem = os.path.splitext(base)[0]
        out_path = os.path.join(outdir, stem + "." + fmt.split(":")[0])
        # rc==0 alone is a lie (soffice exits 0 on unloadable input); the
        # output-exists check is the load-bearing half of the predicate.
        if proc.returncode != 0 or not os.path.exists(out_path):
            raise LOConversionError(
                "LibreOffice conversion failed (rc={0}, output {1}): stdout={2!r} stderr={3!r}".format(
                    proc.returncode,
                    "present" if os.path.exists(out_path) else "MISSING",
                    proc.stdout[-300:],
                    proc.stderr[-300:],
                )
            )
        with open(out_path, "rb") as f:
            return f.read()
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


def lo_recalc_bytes(src_path, timeout=120.0):
    """Round-trip through LibreOffice (recalculates cached values)."""
    return lo_convert(src_path, fmt="xlsx", timeout=timeout)


def lo_loads(src_path, timeout=120.0):
    """True if LibreOffice can load/convert the file — the independent-loader
    smoke check (contract harness assertion 4)."""
    try:
        lo_convert(src_path, fmt="xlsx", timeout=timeout)
        return True
    except LOConversionError:
        return False
