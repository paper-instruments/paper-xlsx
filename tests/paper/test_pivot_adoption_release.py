"""PR 12: adoption release-gate honesty. No new adoption verbs."""
from __future__ import annotations

import json
import os
import re
import sys

import pytest

from openpyxl import load_workbook
from openpyxl.errors import UnsupportedStructureError
from openpyxl.pivot.adopt_evidence import excel_equivalence_proved
from .test_pivot_adoption_qualification import _paper_then_foreign
from .test_pivot_adopt_dedicated import _fingerprint


_MATRIX = os.path.join(
    os.path.dirname(__file__), "fixtures", "pivots", "RELEASE_MATRIX.json")
_FORBIDDEN_CLAIMS = (
    "edit any existing PivotTable",
    "full Excel PivotTable parity",
)


def test_release_matrix_does_not_claim_excel_adoption():
    with open(_MATRIX, encoding="utf-8") as handle:
        matrix = json.load(handle)
    assert matrix["excel"]["windows_build"] is None
    assert matrix["excel"]["macos_build"] is None
    assert matrix["excel"]["marker_survives"] is None
    assert matrix["excel"]["foreign_adoption_proved"] is False
    assert matrix["adoption"]["eligible_without_excel_evidence"] is False
    assert matrix["adoption"]["excel_authored_fixtures"] == "pending"
    assert "adopt" in matrix["operations"]
    assert "qualify_adoption" in matrix["operations"]
    assert "shared-isolation" in matrix["adoption"]["strategies"]
    assert excel_equivalence_proved() is False


def test_excel_runner_stays_a_stub_for_every_adoption_kind(monkeypatch):
    try:
        monkeypatch.delenv("PAPER_EXCEL_PIVOT", raising=False)
        from .support.excel_pivot import (
            TRANSCRIPT_KINDS,
            excel_available,
            run_transcript,
        )

        assert excel_available() is False
        for kind in TRANSCRIPT_KINDS:
            with pytest.raises(RuntimeError):
                run_transcript("unused.xlsx", {}, kind=kind)
        monkeypatch.setenv("PAPER_EXCEL_PIVOT", "1")
        for kind in TRANSCRIPT_KINDS:
            with pytest.raises(NotImplementedError):
                run_transcript("unused.xlsx", {}, kind=kind)
    finally:
        for name in list(sys.modules):
            if "excel_pivot" in name:
                sys.modules.pop(name, None)


def test_public_docs_do_not_overclaim_adoption():
    roots = (
        os.path.join(os.path.dirname(__file__), "..", "..", "README.md"),
        os.path.join(os.path.dirname(__file__), "..", "..", "doc", "paper.rst"),
        os.path.join(os.path.dirname(__file__), "..", "..", "doc", "changes.rst"),
    )
    for path in roots:
        text = open(path, encoding="utf-8").read()
        lowered = text.lower()
        assert "adopt" in lowered
        for claim in _FORBIDDEN_CLAIMS:
            for match in re.finditer(re.escape(claim.lower()), lowered):
                window = lowered[max(0, match.start() - 48):match.end() + 16]
                assert "not" in window, path


def test_adopt_remains_gated_without_excel_evidence(fixture_copy, tmp_path):
    wb, path = _paper_then_foreign(fixture_copy, tmp_path)
    before = _fingerprint(wb)
    source = open(path, "rb").read()
    with pytest.raises(UnsupportedStructureError) as exc:
        wb["Data"].pivots["ByRegion"].adopt()
    assert exc.value.kind == "unsupported-pivot-operation"
    assert "foreign-managed-equivalence-unproved" in (exc.value.options or [])
    assert _fingerprint(wb) == before
    assert open(path, "rb").read() == source


def test_package_does_not_import_the_excel_runner():
    import openpyxl.pivot.adopt as adopt_mod
    import openpyxl.pivot.adopt_evidence as evidence_mod

    leaked = [
        name for name in sys.modules
        if name.startswith("openpyxl") and "excel_pivot" in name
    ]
    assert leaked == []
    assert "support.excel_pivot" not in adopt_mod.__name__
    assert evidence_mod.excel_equivalence_proved() is False
