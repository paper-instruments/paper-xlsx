# paper-xlsx: the preserve-mode save entry point (PR-0 §3)

"""Save dispatch target for preserve-mode workbooks.

Build stage: Phase 2a. Retention, the loss inventory, and the zip machinery
are live; the splice writer (Phase 2c) and the dirty ledger (Phase 2b) are
not. Until the splice lands there is no safe way to apply model mutations to
the retained bytes, and no ledger to prove there were none — so save refuses,
atomically, with the alternative spelled out. This refusal narrows as the
spine grows; it never silently widens what it writes.
"""

from openpyxl.errors import UnsupportedStructureError


def save_preserved(workbook, target):
    """Save a preserve-mode workbook to ``target`` (path or binary file-like)."""
    raise UnsupportedStructureError(
        "preserve-mode save is not available yet at this build stage: the "
        "splice writer (Phase 2c) has not landed, so edits cannot be applied "
        "to the retained package without risking silent loss. Nothing was "
        "written. To save with stock (lossy) behavior, reopen the file "
        "without preserve=True — the save will warn about anything it cannot "
        "preserve."
    )
