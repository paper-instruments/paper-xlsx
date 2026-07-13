"""The pinned-surface CI check.

Every exception class, result state, and return type in the public
contract must be raised/produced by at least one test, or carry an
explicit entry in the ``KNOWN_DEBTS`` set below. This mechanizes away a
whole breach class: AddressRemap was pinned and returned None with
nothing failing; three refusal classes were defined and never raised.

Paying a debt (implementing + testing the surface) makes its entry
stale; this check then REQUIRES its removal (a paid debt may not linger
and mask a future regression).
"""
from __future__ import annotations

import pathlib
import re

import pytest

REPO = pathlib.Path(__file__).resolve().parents[2]

# exception classes in the public contract
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

# oracle result states — each must be produced somewhere
PINNED_RESULT_STATES = ["CERTIFIED", "DIVERGED", "BASELINE_UNVERIFIABLE"]

# return types in the public contract (structural edits return a remap;
# computation-layer results)
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
            # a TODO comment must never satisfy the produced-arm
            # (comment mentions would otherwise flip the check)
            text = _COMMENT.sub("", text)
        if rx.search(text):
            return True
    return False


# Pinned surface that is intentionally not yet produced-and-tested.
# Empty = every pinned name below is implemented and asserted somewhere.
# To record a future debt, add the exact pinned name here with a comment
# explaining why it isn't yet raised-and-tested.
KNOWN_DEBTS: set[str] = set()


def _debt_ledger():
    return KNOWN_DEBTS


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
        # (bare imports/comments would otherwise satisfy a \b-grep)
        tested = _grep(REPO / "tests" / "paper",
                       r"raises\(\s*{0}\b|except {0}\b".format(name),
                       strip_comments=True)
    ledgered = name in _debt_ledger()
    if produced and tested:
        assert not ledgered, (
            "{0} is implemented and tested but still carries a debt entry "
            "in KNOWN_DEBTS — remove the paid debt.".format(name))
        return
    assert ledgered, (
        "{0} is pinned but not raised-and-tested (produced={1}, "
        "tested={2}) and carries no KNOWN_DEBTS entry. Either raise it "
        "with a test, or add it to KNOWN_DEBTS.".format(
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
            "in KNOWN_DEBTS — remove the paid debt.".format(name))
        return
    assert ledgered, (
        "{0} is pinned but absent (produced={1}, tested={2}) and carries "
        "no KNOWN_DEBTS entry.".format(name, produced, tested))
