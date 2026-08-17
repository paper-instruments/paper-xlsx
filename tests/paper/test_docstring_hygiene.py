"""Docstrings must stay MDX-safe.

The published API reference is generated from these docstrings, so a docstring is not just
prose any more - it is input to a build. MDX parses a raw `<xdr:col>` as a JSX component and
fails outright, and reST leftovers (`|Name|`, `:meth:`target``) render as literal noise on the
published page.

The failure is silent in the worst way: nothing here imports a docstring, no test exercises one,
and a paragraph can lose its meaning without anything going red. A check is the only thing that
catches it.

The rewriter is `tools/docstring_hazards.py`; `--fix` repairs what this reports.
"""
from __future__ import annotations

import pathlib
import sys

import pytest

TOOLS = pathlib.Path(__file__).resolve().parents[2] / "tools"
sys.path.insert(0, str(TOOLS))

from docstring_hazards import HAZARDS, SKIP, SRC, _rewrite, scan  # noqa: E402


def _offenders(hazard):
    """Return `path (count)` for every file still carrying `hazard` in a docstring."""
    found = []
    for path in sorted(SRC.rglob("*.py")):
        if any(part in SKIP for part in path.relative_to(SRC).parts):
            continue
        _, counts = scan(path)
        if counts[hazard]:
            found.append(f"{path.relative_to(SRC.parent)} ({counts[hazard]})")
    return found


class TestDocstringHygiene:

    @pytest.mark.parametrize("hazard", HAZARDS)
    def test_no_hazard_survives(self, hazard):
        offenders = _offenders(hazard)
        assert not offenders, (
            f"{len(offenders)} file(s) still carry a '{hazard}' hazard in a docstring, which "
            f"the generated API reference cannot render. Run "
            f"`uv run python tools/docstring_hazards.py --list` to see every site, then `--fix`. "
            f"Offenders: {', '.join(offenders[:10])}"
        )


class TestTheRewriter:

    @pytest.mark.parametrize(("before", "after"), [
        ("the |Workbook| object", "the `Workbook` object"),
        ("see :meth:`Workbook.save`", "see `Workbook.save`"),
        ("see :class:`.PaperRefusal`", "see `PaperRefusal`"),
        ("emits <xdr:col> markers", "emits `<xdr:col>` markers"),
        ("keys are {sheet: rows}", "keys are `{sheet: rows}`"),
    ])
    def test_it_converts_each_hazard_class(self, before, after):
        assert _rewrite(before)[0] == after

    @pytest.mark.parametrize("text", [
        "already `<xdr:col>` backticked",
        "``a reST literal with {braces} inside``",
        "plain prose with no markup at all",
    ])
    def test_it_leaves_safe_text_alone(self, text):
        assert _rewrite(text)[0] == text

    def test_it_does_not_touch_ordinary_string_literals(self, tmp_path):
        """A hazard outside a docstring is code, not documentation."""
        module = tmp_path / "sample.py"
        module.write_text(
            '"""Docstring naming |Workbook|."""\n'
            'SEPARATOR = "|Workbook|"\n'
            'TAG = "<xdr:col>"\n',
            encoding="utf-8",
        )
        rewritten, counts = scan(module)

        assert counts["substitution"] == 1, "only the docstring occurrence converts"
        assert '"""Docstring naming `Workbook`."""' in rewritten
        assert 'SEPARATOR = "|Workbook|"' in rewritten
        assert 'TAG = "<xdr:col>"' in rewritten
