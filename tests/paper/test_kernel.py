"""The package kernel: openpyxl.package (CONVENTIONS §7, PR-0 §5).

The kernel is verified against the harness's own independent diff helpers —
the safety tooling and the shipped kernel implement the same contract twice,
so a bug in one is caught by the other.
"""
from __future__ import annotations

import shutil
import zipfile

from openpyxl.package import diff_package, xml_equivalent, xml_semantic_diff


class TestXmlEquivalent:

    def test_prefix_and_attr_order_insignificant(self):
        a = b'<w xmlns="urn:m"><c r="A1" t="n"/></w>'
        b = b'<x:w xmlns:x="urn:m"><x:c t="n" r="A1"/></x:w>'
        assert xml_equivalent(a, b)

    def test_cell_text_never_normalized(self):
        assert not xml_equivalent(b"<t>0.1</t>", b"<t>0.10</t>")
        assert not xml_equivalent(b"<t> pad </t>", b"<t>pad</t>")

    def test_structural_differences_detected(self):
        assert not xml_equivalent(b"<r><a/></r>", b"<r><a/><a/></r>")
        assert not xml_equivalent(b'<c r="A1"/>', b'<c r="A2"/>')

    def test_self_closing_vs_expanded_equivalent(self):
        assert xml_equivalent(b"<v></v>", b"<v/>")

    def test_diff_reports_are_human_readable(self):
        diffs = xml_semantic_diff(b'<c r="A1"><v>1</v></c>', b'<c r="A1"><v>2</v></c>')
        assert len(diffs) == 1 and "text" in diffs[0]


class TestDiffPackage:

    def test_identical_package_is_clean(self, fixture_copy):
        a = fixture_copy("gauntlet/gauntlet.xlsx", "a.xlsx")
        b = fixture_copy("gauntlet/gauntlet.xlsx", "b.xlsx")
        d = diff_package(a, b)
        assert d.clean
        assert not d.added and not d.removed and not d.changed
        assert len(d.identical) > 10

    def test_detects_changed_added_removed(self, fixture_copy, tmp_path):
        a = fixture_copy("minimal/minimal_clean.xlsx", "a.xlsx")
        b = str(tmp_path / "b.xlsx")
        with zipfile.ZipFile(a) as zin, zipfile.ZipFile(b, "w") as zout:
            for name in zin.namelist():
                payload = zin.read(name)
                if name.startswith("xl/worksheets/sheet"):
                    payload = payload.replace(b"apples", b"oranges")
                if name == "docProps/app.xml":
                    continue  # removed part
                zout.writestr(name, payload)
            zout.writestr("xl/media/new.bin", b"\x00\x01")  # added part
        d = diff_package(a, b)
        assert not d.clean
        assert d.added == ["xl/media/new.bin"]
        assert d.removed == ["docProps/app.xml"]
        assert [c.part for c in d.changed] == ["xl/worksheets/sheet1.xml"]
        assert d.changed[0].kind == "xml"

    def test_byte_different_but_semantically_equivalent_xml(self, fixture_copy, tmp_path):
        a = fixture_copy("minimal/minimal_clean.xlsx", "a.xlsx")
        b = str(tmp_path / "b.xlsx")
        with zipfile.ZipFile(a) as zin, zipfile.ZipFile(b, "w") as zout:
            for name in zin.namelist():
                payload = zin.read(name)
                if name == "xl/workbook.xml":
                    # insignificant whitespace change only
                    payload = payload.replace(b"><", b">\n<")
                zout.writestr(name, payload)
        d = diff_package(a, b)
        assert d.clean
        assert d.equivalent == ["xl/workbook.xml"]

    def test_binary_parts_compared_by_hash(self, fixture_copy, tmp_path):
        a = fixture_copy("features/macro_stub.xlsm", "a.xlsm")
        b = str(tmp_path / "b.xlsm")
        with zipfile.ZipFile(a) as zin, zipfile.ZipFile(b, "w") as zout:
            for name in zin.namelist():
                payload = zin.read(name)
                if name == "xl/vbaProject.bin":
                    payload = payload + b"\x00"
                zout.writestr(name, payload)
        d = diff_package(a, b)
        assert [c.part for c in d.changed] == ["xl/vbaProject.bin"]
        assert d.changed[0].kind == "binary"

    def test_to_dict_schema(self, fixture_copy):
        a = fixture_copy("minimal/minimal_clean.xlsx", "a.xlsx")
        doc = diff_package(a, a).to_dict()
        assert doc["schema"] == "package_diff"
        assert doc["version"] == 1
        for key in ("added", "removed", "changed", "byte_identical",
                    "semantically_equivalent"):
            assert key in doc

    def test_kernel_agrees_with_harness_partdiff(self, fixture_copy):
        from .support.partdiff import diff_parts

        a = fixture_copy("gauntlet/gauntlet.xlsx", "a.xlsx")
        b = fixture_copy("features/schedule.xlsx", "b.xlsx")
        kernel = diff_package(a, b)
        harness = diff_parts(a, b)
        assert set(kernel.added) == harness.added
        assert set(kernel.removed) == harness.removed
