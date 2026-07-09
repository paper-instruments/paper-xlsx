"""The pinned-surface CI check (PLAN-v0.1 process amendment 1).

Every exception class, result state, and return type pinned in CONVENTIONS
or an approved API proposal must be raised/produced by at least one test,
or carry an explicit entry in PAPER.md's pinned-surface debt ledger. This
mechanizes away the v0 breach class: AddressRemap was pinned and returned
None for two phases with nothing failing; three refusal classes were
defined and never raised.

A debt entry is a line in PAPER.md of the form:
    - `Name` — owed to Batch N (...)
Paying the debt (implementing + testing the surface) makes the ledger
entry stale; this check then REQUIRES its removal (a paid debt may not
linger and mask a future regression).
"""
from __future__ import annotations

import pathlib
import re

import pytest

REPO = pathlib.Path(__file__).resolve().parents[2]

# exception classes pinned by CONVENTIONS §2 / PR-0 §2
PINNED_EXCEPTIONS = [
    "PaperRefusal",
    "AmbiguousTargetError",
    "TargetNotFoundError",
    "UnsupportedStructureError",
    "BoundaryViolationError",
    "RelationshipPolicyError",
    "OracleUnavailableError",
    "OracleTimeoutError",
    "LossySaveWarning",
    "StructuralShiftWarning",
    "ProtectedWriteWarning",
    "LintWarning",
]

# result states pinned by PR-0 §7 (oracle) — each must be produced somewhere
PINNED_RESULT_STATES = ["CERTIFIED", "DIVERGED", "BASELINE_UNVERIFIABLE"]

# return types pinned by CONVENTIONS §2 (structural edits return a remap)
# + PR-1 §4 (computation layer results)
PINNED_RETURN_TYPES = ["AddressRemap", "Evaluation", "WriteBackResult"]


def _source_files(root):
    for p in root.rglob("*.py"):
        if "tests" in p.parts and root.name == "openpyxl":
            continue                       # upstream's embedded test dirs
        if p.name == "test_pinned_surface.py":
            continue                       # the checker cannot vouch for itself
        yield p


_COMMENT = re.compile(r"#[^\n]*")


def _grep(root, pattern, strip_comments=False):
    rx = re.compile(pattern)
    for p in _source_files(root):
        try:
            text = p.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            continue
        if strip_comments:
            # a TODO comment must never satisfy the produced-arm (gate
            # finding: comment mentions flipped the check)
            text = _COMMENT.sub("", text)
        if rx.search(text):
            return True
    return False


def _debt_ledger():
    paper = (REPO / "PAPER.md").read_text(encoding="utf-8")
    section = re.search(
        r"## Pinned-surface debt ledger\n(.*?)(?:\n## |\Z)", paper, re.S)
    if not section:
        return set()
    return set(re.findall(r"^- `(\w+)` — owed to",
                          section.group(1), re.M))


@pytest.mark.parametrize("name", PINNED_EXCEPTIONS)
def test_pinned_exception_is_raised_and_tested_or_ledgered(name):
    if name.endswith("Warning"):
        produced = _grep(REPO / "openpyxl",
                         r"warn.*{0}|{0}\(".format(name),
                         strip_comments=True)
        tested = _grep(REPO / "tests" / "paper",
                       r"warns\(\s*{0}|with_warning.*{0}|"
                       r"simplefilter.*\n.*{0}|{0}\)".format(name),
                       strip_comments=True)
    else:
        produced = (_grep(REPO / "openpyxl", r"raise {0}\b".format(name),
                          strip_comments=True)
                    or _grep(REPO / "openpyxl",
                             r"class \w+\({0}\)".format(name),
                             strip_comments=True))    # raised via subclass
        # the tested-arm demands the exception be ASSERTED, not mentioned
        # (gate finding: bare imports/comments satisfied a \b-grep)
        tested = _grep(REPO / "tests" / "paper",
                       r"raises\(\s*{0}\b|except {0}\b".format(name),
                       strip_comments=True)
    ledgered = name in _debt_ledger()
    if produced and tested:
        assert not ledgered, (
            "{0} is implemented and tested but still carries a debt entry "
            "in PAPER.md — remove the paid debt.".format(name))
        return
    assert ledgered, (
        "{0} is pinned but not raised-and-tested (produced={1}, "
        "tested={2}) and carries no PAPER.md debt entry. Either raise it "
        "with a test, or ledger the debt with its owning batch.".format(
            name, produced, tested))


@pytest.mark.parametrize("state", PINNED_RESULT_STATES)
def test_pinned_result_state_is_produced_and_tested(state):
    assert _grep(REPO / "openpyxl", r'"{0}"'.format(state))
    assert _grep(REPO / "tests" / "paper", r"\b{0}\b".format(state))


@pytest.mark.parametrize("name", PINNED_RETURN_TYPES)
def test_pinned_return_type_exists_and_tested_or_ledgered(name):
    produced = _grep(REPO / "openpyxl", r"class {0}\b".format(name))
    tested = _grep(REPO / "tests" / "paper", r"\b{0}\b(?!`)".format(name))
    ledgered = name in _debt_ledger()
    if produced and tested:
        assert not ledgered, (
            "{0} is implemented and tested but still carries a debt entry "
            "in PAPER.md — remove the paid debt.".format(name))
        return
    assert ledgered, (
        "{0} is pinned but absent (produced={1}, tested={2}) and carries "
        "no PAPER.md debt entry.".format(name, produced, tested))
