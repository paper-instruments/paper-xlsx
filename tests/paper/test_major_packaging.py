from __future__ import annotations

import zipfile
from pathlib import Path, PurePosixPath

import pytest

import openpyxl
from openpyxl._distribution import assert_single_openpyxl_distribution
from paper_xlsx_doctor import DoctorError, _openpyxl_record_entries


class _ArchiveMetadata:

    def __init__(self, infos):
        self._infos = infos

    def infolist(self):
        return self._infos

    def namelist(self):
        return [info.filename for info in self._infos]


def _zip_info(name, file_size, compress_size=None):
    info = zipfile.ZipInfo(name)
    info.file_size = file_size
    info.compress_size = file_size if compress_size is None else compress_size
    return info


def test_distribution_guard_accepts_paper_xlsx_alone():
    assert_single_openpyxl_distribution(["paper-xlsx"])


def test_distribution_guard_rejects_shared_import_ownership():
    with pytest.raises(ImportError, match="both provide.*openpyxl"):
        assert_single_openpyxl_distribution(["paper_xlsx", "openpyxl"])


def test_runtime_version_comes_from_packaging_source():
    from openpyxl._paper_version import __paper_version__

    assert openpyxl.__paper_version__ == __paper_version__


def test_release_candidate_version_pin():
    # Deliberate pin: bump alongside openpyxl/_paper_version.py at release.
    assert openpyxl.__paper_version__ == "0.2.0"


class _StubDistribution:
    """Just enough of importlib.metadata.Distribution for record checks."""

    def __init__(self, files):
        self._files = files

    def read_text(self, name):
        return self._files.get(name)

    def locate_file(self, path):
        return Path(str(path))


def test_doctor_accepts_editable_install_without_hashed_files():
    from paper_xlsx_doctor import _verify_openpyxl_record

    record = "__editable__.paper_xlsx_finder.py,sha256=abc,1\n"
    editable = _StubDistribution({
        "RECORD": record,
        "direct_url.json":
            '{"url": "file:///src", "dir_info": {"editable": true}}',
    })
    _verify_openpyxl_record(editable)

    wheel_like = _StubDistribution({"RECORD": record})
    with pytest.raises(DoctorError, match="no hashed openpyxl"):
        _verify_openpyxl_record(wheel_like)


def test_doctor_record_filter_only_accepts_safe_openpyxl_paths():
    record = (
        "openpyxl/__init__.py,sha256=abc,1\n"
        "paper_xlsx-0.1.1.dist-info/METADATA,sha256=def,2\n"
    )
    assert list(_openpyxl_record_entries(record)) == [
        (PurePosixPath("openpyxl/__init__.py"), "sha256=abc")]

    unsafe = "openpyxl/../outside.py,sha256=abc,1\n"
    with pytest.raises(DoctorError, match="unsafe path"):
        list(_openpyxl_record_entries(unsafe))


def test_fixture_request_document_exists():
    root = Path(__file__).resolve().parents[2]
    requests = root / "FIXTURE-REQUESTS.md"
    assert requests.is_file()
    text = requests.read_text(encoding="utf-8")
    assert "Google Sheets" in text
    assert "pivot cache" in text


def test_paper_does_not_expose_package_eligibility_caps():
    from openpyxl.preserve import zipguard
    from openpyxl.reader import excel

    for name in (
        "MAX_PART_BYTES", "MAX_TOTAL_BYTES", "MAX_ENTRIES",
        "_PRESERVE_DECOMPRESSION_MAX_PART", "_DECOMPRESSION_MAX_PART",
        "_DECOMPRESSION_MAX_TOTAL", "_DECOMPRESSION_MAX_ENTRIES",
        "_DECOMPRESSION_RATIO_FLOOR",
    ):
        assert not hasattr(zipguard, name)
        assert not hasattr(excel, name)
