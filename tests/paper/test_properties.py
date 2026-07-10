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
