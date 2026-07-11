from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

import openpyxl
from openpyxl._distribution import assert_single_openpyxl_distribution
from openpyxl.errors import UnsupportedStructureError
from openpyxl.reader import excel
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


def test_release_candidate_version_is_newer_than_published_release():
    assert openpyxl.__paper_version__ == "0.1.1"


def test_doctor_record_filter_only_accepts_safe_openpyxl_paths():
    record = (
        "openpyxl/__init__.py,sha256=abc,1\n"
        "paper_xlsx-0.1.1.dist-info/METADATA,sha256=def,2\n"
    )
    assert list(_openpyxl_record_entries(record)) == [
        (Path("openpyxl/__init__.py"), "sha256=abc")]

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


@pytest.mark.parametrize(
    ("infos", "message"),
    [
        (
            [_zip_info(
                "large-part.xml",
                excel._PRESERVE_DECOMPRESSION_MAX_PART + 1,
            )],
            "part .* cap",
        ),
        (
            [
                _zip_info("aggregate-{0}.xml".format(index),
                          excel._DECOMPRESSION_MAX_TOTAL // 3 + 1)
                for index in range(3)
            ],
            "aggregate uncompressed",
        ),
        (
            [
                _zip_info("entry-{0}.xml".format(index), 1)
                for index in range(
                    excel._DECOMPRESSION_MAX_ENTRIES + 1)
            ],
            "entries",
        ),
    ],
)
def test_tighter_archive_limits_apply_only_in_preserve_mode(infos, message):
    archive = _ArchiveMetadata(infos)

    excel._check_decompression_caps(archive)
    with pytest.raises(UnsupportedStructureError, match=message):
        excel._check_decompression_caps(archive, preserve=True)


def test_default_archive_check_retains_stock_part_limit():
    archive = _ArchiveMetadata([
        _zip_info("too-large.xml", excel._DECOMPRESSION_MAX_PART + 1),
    ])

    with pytest.raises(UnsupportedStructureError, match="part .* cap"):
        excel._check_decompression_caps(archive)


def test_default_archive_check_retains_compression_ratio_limit():
    archive = _ArchiveMetadata([
        _zip_info(
            "bomb.xml",
            excel._DECOMPRESSION_RATIO_FLOOR + 1,
            compress_size=1,
        ),
    ])

    with pytest.raises(UnsupportedStructureError, match="inflates"):
        excel._check_decompression_caps(archive)


def test_excel_reader_forwards_preserve_mode_to_archive_preflight(monkeypatch):
    archive = _ArchiveMetadata([
        _zip_info(
            "large-part.xml",
            excel._PRESERVE_DECOMPRESSION_MAX_PART + 1,
        ),
    ])
    monkeypatch.setattr(excel, "ZipFile", lambda *args, **kwargs: archive)

    reader = excel.ExcelReader(io.BytesIO(b"stub"), preserve=False)
    assert reader.archive is archive
    with pytest.raises(UnsupportedStructureError, match="part .* cap"):
        excel.ExcelReader(io.BytesIO(b"stub"), preserve=True)
