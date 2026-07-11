from __future__ import annotations

import io
import hashlib
import zipfile

import pytest

from openpyxl import Workbook
from openpyxl.errors import UnsupportedStructureError
from openpyxl.package import diff_cells, diff_package, xml_equivalent
from openpyxl.preserve.diffreport import diff_workbooks
from openpyxl.preserve.receipts import receipt


def _book_bytes(value, data_type=None):
    wb = Workbook()
    cell = wb.active["A1"]
    cell.value = value
    if data_type is not None:
        cell.data_type = data_type
    stream = io.BytesIO()
    wb.save(stream)
    return stream.getvalue()


def _duplicate_archive():
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w") as archive:
        archive.writestr("same.xml", b"<x>first</x>")
        archive.writestr("same.xml", b"<x>second</x>")
    return stream.getvalue()


def test_duplicate_entries_refuse_diff_and_receipt():
    duplicate = _duplicate_archive()
    with pytest.raises(UnsupportedStructureError, match="duplicate ZIP"):
        diff_package(duplicate, duplicate)
    with pytest.raises(UnsupportedStructureError, match="duplicate ZIP"):
        receipt(duplicate, duplicate)


def test_package_bounds_are_checked_before_member_reads(monkeypatch):
    import openpyxl.package.diff as package_diff
    import openpyxl.preserve.receipts as receipts

    payload = _book_bytes("bounded")
    monkeypatch.setattr(package_diff, "_MAX_ZIP_UNCOMPRESSED", 1)
    monkeypatch.setattr(receipts, "_MAX_ZIP_UNCOMPRESSED", 1)
    with pytest.raises(UnsupportedStructureError, match="aggregate uncompressed"):
        diff_package(payload, payload)
    with pytest.raises(UnsupportedStructureError, match="aggregate uncompressed"):
        receipt(payload, payload)


def test_package_part_bounds_are_checked_before_member_reads(monkeypatch):
    import openpyxl.package.diff as package_diff
    import openpyxl.preserve.receipts as receipts

    payload = _book_bytes("bounded")
    monkeypatch.setattr(package_diff, "_MAX_ZIP_PART", 1)
    monkeypatch.setattr(receipts, "_MAX_ZIP_PART", 1)
    with pytest.raises(UnsupportedStructureError, match="part .* diff cap"):
        diff_package(payload, payload)
    with pytest.raises(UnsupportedStructureError, match="part .* receipt cap"):
        receipt(payload, payload)


def test_preserve_loader_enforces_aggregate_bound(monkeypatch):
    import openpyxl.reader.excel as excel

    monkeypatch.setattr(excel, "_DECOMPRESSION_MAX_TOTAL", 1)
    with pytest.raises(UnsupportedStructureError, match="aggregate uncompressed"):
        excel.load_workbook(
            io.BytesIO(_book_bytes("bounded")), preserve=True)


def test_xml_leaf_and_xml_space_whitespace_are_significant():
    assert not xml_equivalent(b"<t> </t>", b"<t></t>")
    assert not xml_equivalent(
        b'<r xml:space="preserve"><a/> </r>',
        b'<r xml:space="preserve"><a/></r>')
    assert xml_equivalent(b"<r><a/>\n  <b/></r>",
                          b"<r><a></a><b/></r>")


def test_semantic_xml_diff_refuses_dtd_entity_expansion():
    entity = (b'<!DOCTYPE root [<!ENTITY value "same">]>'
              b'<root>&value;</root>')
    literal = b"<root>same</root>"
    with pytest.raises(UnsupportedStructureError, match="DTD-bearing"):
        xml_equivalent(entity, literal)


def test_agent_facing_package_reads_never_use_unbounded_read():
    payload = _book_bytes("bounded")

    class NoUnboundedRead(io.BytesIO):
        def read(self, size=-1):
            assert size >= 0
            return super().read(size)

    for operation in (diff_package, diff_cells, diff_workbooks, receipt):
        left = NoUnboundedRead(payload)
        right = NoUnboundedRead(payload)
        left.seek(3)
        right.seek(5)
        operation(left, right)
        assert left.tell() == 3
        assert right.tell() == 5


def test_bytes_and_file_like_sources_are_full_and_cursor_safe():
    before = _book_bytes("before")
    after = _book_bytes("after")
    assert not diff_cells(before, after).clean
    left = io.BytesIO(before)
    right = io.BytesIO(after)
    left.seek(17)
    right.seek(23)
    assert not diff_cells(left, right).clean
    assert left.tell() == 17
    assert right.tell() == 23
    left.seek(31)
    right.seek(37)
    assert not diff_package(left, right).clean
    assert left.tell() == 31
    assert right.tell() == 37
    receipt(left, right)
    assert left.tell() == 31
    assert right.tell() == 37


def test_cell_diffs_distinguish_same_python_value_with_different_type():
    numeric = _book_bytes(1)
    boolean = _book_bytes(True)
    change = diff_cells(numeric, boolean).changes[0]
    assert change["old_type"] == "n"
    assert change["new_type"] == "b"
    report = diff_workbooks(numeric, boolean)
    assert report.changed[0]["before_type"] == "n"
    assert report.changed[0]["after_type"] == "b"


def test_cell_diff_distinguishes_formula_text_from_literal_text():
    formula = _book_bytes("=A1")
    literal = _book_bytes("=A1", data_type="s")
    change = diff_cells(formula, literal).changes[0]
    assert change["old_formula"] == "=A1"
    assert change["new_value"] == "=A1"


def test_multi_cell_formula_diff_is_deterministic(fixture_copy):
    source = fixture_copy("features/shared_formulas.xlsx")
    assert diff_cells(source, source).clean
    assert diff_workbooks(source, source).changed == []


def test_receipt_refuses_unbound_or_cross_workbook_verification():
    before = _book_bytes("before")
    after = _book_bytes("after")

    class Result:
        def __init__(self, digest=None):
            self.artifact_sha256 = digest

        def to_dict(self):
            return {"status": "CERTIFIED",
                    "artifact_sha256": self.artifact_sha256}

    with pytest.raises(UnsupportedStructureError, match="not bound"):
        receipt(before, after, recalc=Result())
    with pytest.raises(UnsupportedStructureError, match="different workbook"):
        receipt(before, after, recalc=Result(hashlib.sha256(before).hexdigest()))

    bound = receipt(
        before, after,
        recalc=Result(hashlib.sha256(after).hexdigest()))
    assert bound.recalc["status"] == "CERTIFIED"
