"""Permanent property invariants.

Two properties that hold across the WHOLE corpus, forever:

- a zero-edit preserve save is byte-identical on every part, for every
  loadable fixture — the invariant the fork is named for. An early
  false-dirty bug slipped through because the no-op test enumerated
  fixtures by hand and skipped the one that failed; this one enumerates
  the corpus.
- the ledger cross-check covers REGIONS, not just cells: a region the
  saver never claimed may never differ in the output.
"""
from __future__ import annotations

import pathlib

import pytest

from openpyxl import load_workbook

from .support.partdiff import part_payloads

FIXTURE_ROOT = pathlib.Path(__file__).parent / "fixtures"

# every workbook in the corpus that is supposed to load (corrupt/ exists
# to be unloadable; generators/ holds the corpus builder, not fixtures)
ALL_LOADABLE = sorted(
    str(p.relative_to(FIXTURE_ROOT)).replace("\\", "/")
    for pattern in ("*.xlsx", "*.xlsm")
    for p in FIXTURE_ROOT.rglob(pattern)
    if "corrupt" not in p.parts and "generators" not in p.parts)


def test_corpus_enumeration_is_alive():
    # if the corpus moves, this file must fail loudly, not skip silently
    assert len(ALL_LOADABLE) >= 12


class TestNoOpByteIdentityEverywhere:

    @pytest.mark.parametrize("fixture", ALL_LOADABLE)
    def test_noop_save_is_byte_identical(self, fixture_copy, tmp_path,
                                         fixture):
        src = fixture_copy(fixture)
        wb = load_workbook(src, preserve=True)
        out = str(tmp_path / ("noop_" + pathlib.Path(fixture).name))
        wb.save(out)
        assert part_payloads(src) == part_payloads(out)


class TestRegionClaimsCrossCheck:
    """The extended cross-check itself (unit level): unclaimed region
    drift raises; claimed drift passes; the suite-wide env then applies
    it to every preserve save the suite performs."""

    _NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'

    def _pkg(self, sheet_xml):
        import io
        import zipfile

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as z:
            z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        return buf.getvalue()

    def test_unclaimed_region_change_raises(self):
        from openpyxl.preserve.crosscheck import (
            LedgerCrossCheckError,
            verify_splice,
        )

        before = self._pkg(
            '<worksheet {0}><sheetData/><pageMargins left="0.75"/>'
            '</worksheet>'.format(self._NS))
        after = self._pkg(
            '<worksheet {0}><sheetData/><pageMargins left="1.25"/>'
            '</worksheet>'.format(self._NS))
        with pytest.raises(LedgerCrossCheckError, match="pageMargins"):
            verify_splice(before, after,
                          {"xl/worksheets/sheet1.xml": set()},
                          region_claims={"xl/worksheets/sheet1.xml": set()})

    def test_claimed_region_change_passes(self):
        from openpyxl.preserve.crosscheck import verify_splice

        before = self._pkg(
            '<worksheet {0}><sheetData/><pageMargins left="0.75"/>'
            '</worksheet>'.format(self._NS))
        after = self._pkg(
            '<worksheet {0}><sheetData/><pageMargins left="1.25"/>'
            '</worksheet>'.format(self._NS))
        verify_splice(before, after,
                      {"xl/worksheets/sheet1.xml": set()},
                      region_claims={
                          "xl/worksheets/sheet1.xml": {"pageMargins"}})

    def test_unclaimed_row_attr_drift_raises(self):
        # a blind spot: row display attributes rewritten
        # on rows the ledger never claimed
        from openpyxl.preserve.crosscheck import (
            LedgerCrossCheckError,
            verify_splice,
        )

        before = self._pkg(
            '<worksheet {0}><sheetData><row r="1"><c r="A1"><v>1</v></c>'
            '</row></sheetData></worksheet>'.format(self._NS))
        after = self._pkg(
            '<worksheet {0}><sheetData><row r="1" ht="99" customHeight="1">'
            '<c r="A1"><v>1</v></c></row></sheetData>'
            '</worksheet>'.format(self._NS))
        with pytest.raises(LedgerCrossCheckError, match="row"):
            verify_splice(before, after,
                          {"xl/worksheets/sheet1.xml": set()})
        # claimed: same drift passes
        verify_splice(before, after, {"xl/worksheets/sheet1.xml": set()},
                      row_claims={"xl/worksheets/sheet1.xml": {1}})

    def test_unclaimed_row_duplication_raises(self):
        from openpyxl.preserve.crosscheck import (
            LedgerCrossCheckError,
            verify_splice,
        )

        before = self._pkg(
            '<worksheet {0}><sheetData><row r="3"><c r="A3"><v>1</v></c>'
            '</row></sheetData></worksheet>'.format(self._NS))
        after = self._pkg(
            '<worksheet {0}><sheetData><row r="3"><c r="A3"><v>1</v></c>'
            '</row><row r="3"><c r="A3"><v>1</v></c></row></sheetData>'
            '</worksheet>'.format(self._NS))
        with pytest.raises(LedgerCrossCheckError, match="row"):
            verify_splice(before, after,
                          {"xl/worksheets/sheet1.xml": set()})


class TestCellSignatureNeverLaxerThanC14N:
    """Receipts and the cross-check call two cells equal by their signature.
    The signature reads the parsed tree instead of canonicalizing each cell.
    It may call more cells changed than the C14N comparison it replaced, and
    only where documented, but never fewer: a change it calls "same" would be
    missing from a receipt."""

    _OPEN = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/'
        '2006/main" xmlns:x14ac="http://schemas.microsoft.com/office/'
        'spreadsheetml/2009/9/ac" xmlns:a="urn:a" xmlns:b="urn:b">'
        '<sheetData><row r="1">')
    _CLOSE = '</row></sheetData></worksheet>'

    CELLS = [
        '<c r="A1"><v>1</v></c>',
        '<c r="A1"><v>2</v></c>',
        '<c r="A1" s="0"><v>1</v></c>',
        '<c r="A1" s="3" t="n"><v>1</v></c>',
        '<c r="A1" t="n" s="3"><v>1</v></c>',
        '<c r="A1" x14ac:dyDescent="0.25"><v>1</v></c>',
        '<c r="A1" a:x="1" b:y="2"><v>1</v></c>',
        '<c r="A1" b:y="2" a:x="1"><v>1</v></c>',
        '<c r="A1"><f>B1+1</f><v>2</v></c>',
        '<c r="A1"><f>B1+1</f>\n  <v>2</v></c>',
        '<c r="A1"><v>2</v><f>B1+1</f></c>',
        '<c r="A1"><f>B1+1</f></c>',
        '<c r="A1"><f>B1+1</f><v></v></c>',
        '<c r="A1"><f>B1+1</f><v/></c>',
        '<c r="A1"><f t="shared" si="0" ref="A1:A9">B1+1</f><v>2</v></c>',
        '<c r="A1"><f ref="A1:A9" si="0" t="shared">B1+1</f><v>2</v></c>',
        '<c r="A1" t="str"><f>"a&amp;b"</f><v>a&amp;b</v></c>',
        '<c r="A1" t="str"><f>"a&#38;b"</f><v>a&#38;b</v></c>',
        '<c r="A1" t="inlineStr"><is><t>a</t></is></c>',
        '<c r="A1" t="inlineStr"><is><t xml:space="preserve"> a </t></is></c>',
        '<c r="A1" t="inlineStr"><is><t xml:space="preserve">a </t></is></c>',
        '<c r="A1" t="inlineStr"><is><t>x&#13;y</t></is></c>',
        '<c r="A1" t="inlineStr"><is><t>x&#10;y</t></is></c>',
        '<c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>a</t></r></is></c>',
        '<c r="A1" t="inlineStr"><is><r><rPr><i/></rPr><t>a</t></r></is></c>',
    ]

    # Pairs C14N called equal that the signature calls changed: attributes
    # in another order, and a carriage return C14N read as a line feed.
    STRICTER = {
        frozenset({3, 4}), frozenset({14, 15}), frozenset({21, 22}),
    }

    def _cell(self, xml):
        from xml.etree import ElementTree as ET

        root = ET.fromstring(self._OPEN + xml + self._CLOSE)
        return next(root.iter(
            "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"))

    def test_never_calls_a_canonical_change_the_same(self):
        from xml.etree import ElementTree as ET

        from openpyxl.preserve.crosscheck import _cell_signature

        def c14n_signature(el):
            attrs = dict(el.attrib)
            attrs.pop("r", None)
            attrs.setdefault("s", "0")
            return (tuple(sorted(attrs.items())),
                    ET.canonicalize(ET.tostring(el)))

        cells = [self._cell(xml) for xml in self.CELLS]
        expected = [c14n_signature(el) for el in cells]
        actual = [_cell_signature(el) for el in cells]
        stricter = set()
        for i in range(len(cells)):
            for j in range(i + 1, len(cells)):
                if actual[i] == actual[j]:
                    assert expected[i] == expected[j], \
                        (self.CELLS[i], self.CELLS[j])
                elif expected[i] == expected[j]:
                    stricter.add(frozenset({i, j}))
        assert stricter == self.STRICTER

    def test_deep_nesting_does_not_hit_the_recursion_limit(self):
        import sys

        from openpyxl.preserve.crosscheck import _cell_signature

        depth = sys.getrecursionlimit() * 2
        inner = "<a:n>" * depth + "{0}" + "</a:n>" * depth
        cell = '<c r="A1"><v>1</v><extLst>' + inner + '</extLst></c>'
        same = _cell_signature(self._cell(cell.format("x")))
        assert same == _cell_signature(self._cell(cell.format("x")))
        assert same != _cell_signature(self._cell(cell.format("y")))

    def test_text_after_a_cell_is_a_change(self):
        # C14N dropped whitespace after a cell and raised on other text, so
        # junk spliced in after an untouched cell must not compare equal.
        from openpyxl.preserve.crosscheck import _cell_signature

        plain = _cell_signature(self._cell('<c r="A1"><v>1</v></c>'))
        assert plain == _cell_signature(
            self._cell('<c r="A1"><v>1</v></c>\n  '))
        assert plain != _cell_signature(
            self._cell('<c r="A1"><v>1</v></c>junk'))
