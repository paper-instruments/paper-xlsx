"""Verify that the frozen ``openpyxl`` import belongs to ``paper-xlsx``."""

from __future__ import annotations

import base64
import csv
import hashlib
import hmac
import importlib
import json
import sys
from importlib.metadata import Distribution, PackageNotFoundError, distribution
from io import StringIO
from pathlib import Path, PurePosixPath
from typing import Iterable, Optional, Tuple


class DoctorError(RuntimeError):
    """The installed ``openpyxl`` package cannot be trusted."""


_REMEDY = (
    "python -m pip uninstall -y openpyxl paper-xlsx && "
    "python -m pip install --force-reinstall paper-xlsx"
)


def verify_install() -> str:
    """Verify ownership, installed bytes, and the fork sentinel."""
    paper = _installed_distribution("paper-xlsx")
    upstream = _installed_distribution("openpyxl")
    if paper is None:
        raise DoctorError("paper-xlsx distribution metadata is missing")
    if upstream is not None:
        raise DoctorError(
            "paper-xlsx and openpyxl are both installed and own the same "
            "openpyxl package"
        )

    _verify_openpyxl_record(paper)

    try:
        package = importlib.import_module("openpyxl")
    except Exception as exc:
        raise DoctorError("openpyxl cannot be imported: {0}".format(exc)) \
            from exc
    sentinel = getattr(package, "__paper_version__", None)
    if sentinel is None:
        raise DoctorError("openpyxl.__paper_version__ is missing")
    if sentinel != paper.version:
        raise DoctorError(
            "openpyxl.__paper_version__ does not match the installed "
            "paper-xlsx version ({0!r} != {1!r})".format(
                sentinel, paper.version)
        )
    return paper.version


def main() -> int:
    """Console entry point for ``paper-xlsx-doctor``."""
    try:
        version = verify_install()
    except DoctorError as exc:
        print("paper-xlsx-doctor: FAIL: {0}".format(exc), file=sys.stderr)
        print("Remedy: {0}".format(_REMEDY), file=sys.stderr)
        return 1
    print("paper-xlsx-doctor: OK (paper-xlsx {0})".format(version))
    return 0


def _installed_distribution(name: str) -> Optional[Distribution]:
    try:
        return distribution(name)
    except PackageNotFoundError:
        return None


def _is_editable_install(dist: Distribution) -> bool:
    """PEP 660 editable installs record their state in direct_url.json."""
    payload = dist.read_text("direct_url.json")
    if payload is None:
        return False
    try:
        direct_url = json.loads(payload)
    except ValueError:
        return False
    dir_info = direct_url.get("dir_info")
    return isinstance(dir_info, dict) and dir_info.get("editable") is True


def _verify_openpyxl_record(dist: Distribution) -> None:
    record = dist.read_text("RECORD")
    if record is None:
        raise DoctorError(
            "paper-xlsx has no RECORD to verify against. Wheel installs "
            "always carry one; legacy setup.py/egg installs cannot be "
            "verified — reinstall from a wheel (or, for development, use a "
            "modern editable install: pip install -e .)")
    entries = tuple(
        (relative_path, hash_spec)
        for relative_path, hash_spec in _openpyxl_record_entries(record)
        if hash_spec
    )
    if not entries:
        if _is_editable_install(dist):
            # An editable install materializes import shims, not hashed
            # openpyxl files; the source tree is the live code, so there is
            # nothing meaningful for file hashes to attest. The
            # dual-distribution and sentinel checks still apply.
            return
        raise DoctorError(
            "paper-xlsx RECORD has no hashed openpyxl package files")
    for relative_path, hash_spec in entries:
        path = Path(dist.locate_file(relative_path))
        if not path.is_file():
            raise DoctorError(
                "paper-xlsx file is missing: {0}".format(relative_path))
        algorithm, expected = _parse_hash(hash_spec, relative_path)
        actual = _file_digest(path, algorithm)
        if not hmac.compare_digest(actual, expected):
            raise DoctorError(
                "paper-xlsx file hash mismatch: {0}".format(relative_path))


def _openpyxl_record_entries(
        record: str) -> Iterable[Tuple[PurePosixPath, str]]:
    for row in csv.reader(StringIO(record)):
        if len(row) != 3:
            raise DoctorError("paper-xlsx RECORD contains a malformed row")
        raw_path, hash_spec, _size = row
        path = PurePosixPath(raw_path)
        if not path.parts or path.parts[0] != "openpyxl":
            continue
        if path.is_absolute() or ".." in path.parts:
            raise DoctorError(
                "paper-xlsx RECORD contains an unsafe path: {0}"
                .format(raw_path))
        yield path, hash_spec


def _parse_hash(hash_spec: str,
                relative_path: PurePosixPath) -> Tuple[str, str]:
    try:
        algorithm, expected = hash_spec.split("=", 1)
        hashlib.new(algorithm)
    except (TypeError, ValueError):
        raise DoctorError(
            "paper-xlsx RECORD has an invalid hash for {0}"
            .format(relative_path)) from None
    if not expected:
        raise DoctorError(
            "paper-xlsx RECORD has an invalid hash for {0}"
            .format(relative_path))
    return algorithm, expected.rstrip("=")


def _file_digest(path: Path, algorithm: str) -> str:
    digest = hashlib.new(algorithm)
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return base64.urlsafe_b64encode(digest.digest()).rstrip(b"=").decode(
        "ascii")


__all__ = ["DoctorError", "main", "verify_install"]
