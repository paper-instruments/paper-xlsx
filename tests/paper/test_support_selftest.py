"""Self-tests for the contract-harness support helpers: the safety tooling
must itself be tested, or a bug in it silently weakens every other test."""
from __future__ import annotations

import zipfile

import pytest

from .support.partdiff import diff_parts, part_payloads, xml_equivalent, xml_semantic_diff
from .support.harness import assert_part_budget, assert_refusal_atomic


def _make_pkg(path, entries):
    with zipfile.ZipFile(path, "w") as z:
        for name, payload in entries.items():
            z.writestr(name, payload)
    return str(path)


class TestPartDiff:

    def test_identical_packages_diff_clean(self, tmp_path):
        entries = {"a.xml": b"<a/>", "b/c.bin": b"\x00\x01"}
        p1 = _make_pkg(tmp_path / "one.zip", entries)
        p2 = _make_pkg(tmp_path / "two.zip", entries)
        d = diff_parts(p1, p2)
        assert d.clean
        assert d.identical == {"a.xml", "b/c.bin"}

    def test_changed_added_removed_detected(self, tmp_path):
        p1 = _make_pkg(tmp_path / "one.zip", {"a.xml": b"<a/>", "gone.xml": b"<g/>"})
        p2 = _make_pkg(tmp_path / "two.zip", {"a.xml": b"<a x='1'/>", "new.xml": b"<n/>"})
        d = diff_parts(p1, p2)
        assert d.changed == {"a.xml"}
        assert d.added == {"new.xml"}
        assert d.removed == {"gone.xml"}
        assert not d.clean

    def test_payload_identity_ignores_zip_metadata(self, tmp_path):
        p1 = tmp_path / "one.zip"
        p2 = tmp_path / "two.zip"
        with zipfile.ZipFile(p1, "w") as z:
            z.writestr(zipfile.ZipInfo("a.xml", (2020, 1, 1, 0, 0, 0)), b"<a/>")
        with zipfile.ZipFile(p2, "w") as z:
            z.writestr(zipfile.ZipInfo("a.xml", (2024, 6, 6, 6, 6, 6)), b"<a/>")
        assert diff_parts(str(p1), str(p2)).clean

    def test_part_payloads_accepts_bytes(self, tmp_path):
        p1 = _make_pkg(tmp_path / "one.zip", {"a.xml": b"<a/>"})
        with open(p1, "rb") as f:
            data = f.read()
        assert part_payloads(data) == {"a.xml": b"<a/>"}


class TestXmlSemanticDiff:

    def test_attribute_order_is_insignificant(self):
        assert xml_equivalent(b'<c r="A1" s="2" t="n"/>', b'<c t="n" r="A1" s="2"/>')

    def test_namespace_prefix_is_insignificant_clark_names_compared(self):
        a = b'<w xmlns="urn:x"><c/></w>'
        b = b'<x:w xmlns:x="urn:x"><x:c/></x:w>'
        assert xml_equivalent(a, b)

    def test_inter_element_whitespace_is_insignificant(self):
        assert xml_equivalent(b"<r><a/><b/></r>", b"<r>\n  <a/>\n  <b/>\n</r>")

    def test_cell_text_never_normalized(self):
        assert not xml_equivalent(b"<t>  padded  </t>", b"<t>padded</t>")
        assert not xml_equivalent(b"<t>0.1</t>", b"<t>0.10</t>")

    def test_attribute_value_change_detected(self):
        diffs = xml_semantic_diff(b'<c r="A1"/>', b'<c r="A2"/>')
        assert diffs and "attrib" in diffs[0]

    def test_child_count_change_detected(self):
        diffs = xml_semantic_diff(b"<r><a/></r>", b"<r><a/><a/></r>")
        assert diffs and "child count" in diffs[0]

    def test_self_closing_vs_expanded_is_equivalent(self):
        assert xml_equivalent(b"<v></v>", b"<v/>")


class TestHarnessHelpers:

    def test_part_budget_pass_and_fail(self, tmp_path):
        p1 = _make_pkg(tmp_path / "one.zip", {"a.xml": b"<a/>", "b.xml": b"<b/>"})
        p2 = _make_pkg(tmp_path / "two.zip", {"a.xml": b"<a x='1'/>", "b.xml": b"<b/>"})
        assert_part_budget(p1, p2, expect_changed={"a.xml"})
        with pytest.raises(AssertionError, match="part budget violated"):
            assert_part_budget(p1, p2, expect_changed=set())
        with pytest.raises(AssertionError, match="part budget violated"):
            # expected-but-unchanged is also a violation: the budget is literal
            assert_part_budget(p1, p2, expect_changed={"a.xml", "b.xml"})

    def test_refusal_atomicity_helper(self, tmp_path, fixture_copy):
        src = fixture_copy("minimal/minimal_clean.xlsx")

        class FakeRefusal(Exception):
            pass

        def refusing_mutation(wb, path):
            raise FakeRefusal("refused before touching disk")

        exc = assert_refusal_atomic(src, tmp_path, refusing_mutation, FakeRefusal)
        assert "refused" in str(exc)

    def test_refusal_atomicity_helper_catches_dirty_writes(self, tmp_path, fixture_copy):
        src = fixture_copy("minimal/minimal_clean.xlsx")

        class FakeRefusal(Exception):
            pass

        def dirty_mutation(wb, path):
            with open(path, "ab") as f:
                f.write(b"corruption")
            raise FakeRefusal("refused AFTER touching disk — not atomic")

        with pytest.raises(AssertionError, match="not atomic"):
            assert_refusal_atomic(src, tmp_path, dirty_mutation, FakeRefusal)


@pytest.mark.lo_smoke
class TestLoDriver:

    def test_lo_convert_never_touches_the_original(self, lo, fixture_copy, tmp_path):
        src = fixture_copy("minimal/minimal_clean.xlsx")
        with open(src, "rb") as f:
            before = f.read()
        data = lo.lo_convert(src, fmt="xlsx")
        assert data[:2] == b"PK"
        with open(src, "rb") as f:
            assert f.read() == before, "lo_convert mutated its input"

    def test_lo_loads_smoke(self, lo, fixture_copy):
        assert lo.lo_loads(fixture_copy("gauntlet/gauntlet.xlsx"))

    def test_lo_convert_fails_loudly_on_corrupt_input(self, lo, fixture_copy):
        src = fixture_copy("corrupt/truncated.xlsx")
        with pytest.raises(lo.LOConversionError):
            lo.lo_convert(src, fmt="xlsx")
