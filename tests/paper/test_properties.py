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


class TestTextPrefiltersMatchTheTokenizer:
    """The rename and shift helpers skip the tokenizer when a name cannot
    appear in a formula's text. For every formula and title here, each
    helper must answer exactly as tokenizing gives."""

    FORMULAS = [
        "=Data!A1*2", "=data!a1+'Other Sheet'!B2", "='O''Brien'!C3",
        "='o''brien'!C3+1", "=SUM(First:Last!A2)", "=SUM('Q1 Data:Q4'!A2)",
        '=INDIRECT("Data!A1")', '=indirect("\'O\'\'Brien\'!A1")',
        "=OFFSET(Data!A1,1,0)", "=_xlfn.EVALUATE(A1)", "=offset (A1,1,1)",
        '="Data"&A1', '="indirect("&A1', "=SUM(A1:B2)*Inputs!B3",
        "=rate*2", "=Straße!A1", "=STRASSE!A1", "=İnputs!A1", "=ınputs!A1",
        "=1+(2", '="unterminated', "=A1+", "=Data!A1:Data!B2",
        "=@Data!A1", "=Data!#REF!", "=[1]Data!A1", "Data!A1",
        '=İNDİRECT("Data!A1")', '=İNDİRECT("Data!A1)', "=oﬀset(A1,1,0)",
        "=oﬀset(A1,1,0", "=ındırect(A1)", "=ındırect(A1",
    ]
    TITLES = ["Data", "DATA", "O'Brien", "First", "Q1 Data", "Straße",
              "STRASSE", "Inputs", "İnputs", "Missing"]

    @staticmethod
    def _tokens(formula):
        from openpyxl.formula import Tokenizer

        return Tokenizer(formula).items

    def test_dynamic_reference_checks(self):
        from openpyxl.preserve.references import (
            _DYNAMIC_FUNCTIONS,
            _DYNAMIC_REFERENCE,
            _dynamic_reference_affects_shift,
            _has_dynamic_reference,
        )

        for value in self.FORMULAS:
            probe = value[1:] if value.startswith("#") else value
            formula = probe if probe.startswith("=") else "=" + probe
            try:
                expected = any(
                    t.type == "FUNC" and t.subtype == "OPEN"
                    and t.value.upper() in _DYNAMIC_FUNCTIONS
                    for t in self._tokens(formula))
            except Exception:
                expected = bool(_DYNAMIC_REFERENCE.search(probe))
            assert _has_dynamic_reference(value) is expected, value
            if not expected:
                assert _dynamic_reference_affects_shift(
                    value, "Data", "Data", "rows", 1, 1, False) is False

    def test_rename_rewrites_and_literal_checks(self):
        from openpyxl.preserve.rewrite import (
            _RENAME_PREFIX_RE,
            rename_sheets_in_formula,
            title_in_string_literals,
        )

        for formula in self.FORMULAS:
            for title in self.TITLES:
                folded = title.casefold()
                try:
                    tokens = self._tokens(formula)
                except Exception:
                    tokens = None
                in_literal = (folded in formula.casefold() if tokens is None
                              else any(t.type == "OPERAND"
                                       and t.subtype == "TEXT"
                                       and folded in t.value.casefold()
                                       for t in tokens))
                assert title_in_string_literals(formula, title) is \
                    in_literal, (formula, title)

                names_title = False
                if formula.startswith("=") and tokens is not None:
                    for t in tokens:
                        if t.type != "OPERAND" or t.subtype != "RANGE" \
                                or "[" in t.value:
                            continue
                        m = _RENAME_PREFIX_RE.match(t.value.lstrip("@"))
                        if m is None:
                            continue
                        sheet = (m.group(1).replace("''", "'")
                                 if m.group(1) else m.group(2))
                        if folded in (p.casefold()
                                      for p in sheet.split(":")):
                            names_title = True
                _rewritten, changed = rename_sheets_in_formula(
                    formula, {title: "Renamed"})
                assert changed is names_title, (formula, title)
