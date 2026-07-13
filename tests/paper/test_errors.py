"""The pinned exception taxonomy."""
from __future__ import annotations

import pytest

from openpyxl.errors import (
    AmbiguousTargetError,
    BoundaryViolationError,
    LossySaveWarning,
    OracleTimeoutError,
    OracleUnavailableError,
    PaperRefusal,
    RelationshipPolicyError,
    TargetNotFoundError,
    UnsupportedStructureError,
)

REFUSALS = [
    AmbiguousTargetError,
    TargetNotFoundError,
    UnsupportedStructureError,
    BoundaryViolationError,
    RelationshipPolicyError,
    OracleUnavailableError,
    OracleTimeoutError,
]


@pytest.mark.parametrize("exc", REFUSALS)
def test_every_refusal_is_a_paper_refusal(exc):
    assert issubclass(exc, PaperRefusal)
    assert issubclass(exc, Exception)


def test_paper_refusal_is_not_a_builtin_error_subclass():
    # callers must be able to catch PaperRefusal without catching
    # TypeError/ValueError programmer errors, and vice versa
    assert not issubclass(PaperRefusal, (TypeError, ValueError))


def test_lossy_save_warning_carries_structured_losses():
    losses = [{"kind": "vba", "location": "xl/vbaProject.bin", "detail": "d"}]
    w = LossySaveWarning("msg", losses)
    assert issubclass(LossySaveWarning, UserWarning)
    assert w.losses == losses
    assert LossySaveWarning("msg").losses == []
