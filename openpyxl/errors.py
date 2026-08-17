# paper-xlsx: typed refusals and protection warnings

"""Typed exceptions for paper-xlsx safety refusals.

Every preserve-mode operation completes correctly or refuses with a
`PaperRefusal` subclass saying what was found and why it was unsafe.
Writes to protected cells can also emit an explicit advisory warning. A
refused operation leaves the in-memory model, dirty ledger, and destination
exactly as they were.

Programmer errors (invalid argument combinations, wrong types) remain
``TypeError``/``ValueError`` and are deliberately NOT part of this hierarchy.
"""

__all__ = [
    "AmbiguousTargetError",
    "BoundaryViolationError",
    "HandleRebindWarning",
    "OracleTimeoutError",
    "OracleUnavailableError",
    "PaperRefusal",
    "ProtectedWriteWarning",
    "RelationshipPolicyError",
    "TargetNotFoundError",
    "UnsupportedStructureError",
]


class PaperRefusal(Exception):
    """Base class for all safe refusals.

    Refusals are atomic: when one is raised, the workbook model, the dirty
    ledger, and every file on disk are exactly as they were before the
    refused operation began.

    Structured fields (populated progressively — message
    text is always the source of truth):

    - ``kind``: stable machine-readable string ("ambiguous-label", ...)
    - ``anchor``: sheet-qualified address or part name the refusal is
      about, or None
    - ``options``: suggested remedies / candidate addresses (list)
    """

    def __init__(self, *args, kind=None, anchor=None, options=None):
        super().__init__(*args)
        self.kind = kind
        self.anchor = anchor
        self.options = list(options) if options else []


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


class ProtectedWriteWarning(UserWarning):
    """A write landed on a locked cell of a protected sheet. The write
    proceeds — openpyxl-level protection is advisory, and this library
    reports it rather than enforcing or bypassing it — but
    the human who protected the sheet expected the cell to be read-only.
    Set ``wb.strict_protection = True`` to turn these writes into typed
    refusals."""


class HandleRebindWarning(UserWarning):
    """A path-backed save committed correctly, but the caller's open file
    handle could not be rebound to the replacement file."""
