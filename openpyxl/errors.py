# paper-xlsx: typed refusals and structured warnings (CONVENTIONS §2, PR-0 §2)

"""Typed exceptions for paper-xlsx safety refusals.

Every operation in preserve mode has exactly three legal outcomes: done
correctly; refused with a :class:`PaperRefusal` subclass saying what was found
and why it was unsafe; or done with a loud warning enumerating exactly what
could not be preserved. A refused operation leaves the in-memory model, the
dirty ledger, and any file on disk exactly as they were.

Programmer errors (invalid argument combinations, wrong types) remain
``TypeError``/``ValueError`` and are deliberately NOT part of this hierarchy.
"""


class PaperRefusal(Exception):
    """Base class for all safe refusals.

    Refusals are atomic: when one is raised, the workbook model, the dirty
    ledger, and every file on disk are exactly as they were before the
    refused operation began.
    """


class AmbiguousTargetError(PaperRefusal):
    """The addressed target matches more than one candidate."""


class TargetNotFoundError(PaperRefusal):
    """The addressed target does not exist in the workbook or package."""


class UnsupportedStructureError(PaperRefusal):
    """The operation would require understanding or rewriting structure this
    library cannot handle safely; performing it would risk silent damage."""


class BoundaryViolationError(PaperRefusal):
    """The operation would cross a declared boundary (range, sheet, or
    package region) it is not allowed to cross."""


class RelationshipPolicyError(PaperRefusal):
    """The operation would rewrite or renumber package relationships in a
    way that could detach preserved content."""


class OracleUnavailableError(PaperRefusal):
    """No LibreOffice installation could be found to act as the oracle."""


class OracleTimeoutError(PaperRefusal):
    """The LibreOffice oracle did not finish within the allowed time."""


class StructuralShiftWarning(UserWarning):
    """A row/column shift on a loaded workbook: the cells move but nothing
    that references them is updated — formulas, defined names and chart
    ranges keep pointing at the old cells."""


class LossySaveWarning(UserWarning):
    """Loud warning on a save path that is about to rebuild or drop content
    it cannot preserve.

    ``losses`` is a list of dicts, each ``{"kind": ..., "location": ...,
    "detail": ...}``; the rendered message enumerates them.
    """

    def __init__(self, message, losses=None):
        super().__init__(message)
        self.losses = list(losses) if losses else []
