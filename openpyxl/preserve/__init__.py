# paper-xlsx: the spine

"""Preserve mode: the original package is the source of truth; the object
model is a source of edits to it.

Enabled by default for editable OOXML workbooks loaded with
``load_workbook(path)``. Untouched parts survive byte-identical by
construction (raw compressed-stream copy where possible); touched worksheet
parts are spliced, never re-serialized.
"""

from .inventory import LossInventory, scan_archive
from .saver import save_preserved

__all__ = ["AddressRemap", "DirtyLedger", "LossInventory", "scan_archive",
           "save_preserved", "scan_errors", "findings", "receipt",
           "diff_workbooks", "copy_format", "apply_profile"]


def __getattr__(name):
    # DirtyLedger lives in .ledger, which several early-imported modules
    # (cell, styleable, worksheet) pull helpers from; exposing it lazily
    # here keeps this package importable from anywhere without cycles
    if name == "DirtyLedger":
        from .ledger import DirtyLedger
        return DirtyLedger
    if name == "AddressRemap":
        from .structural import AddressRemap
        return AddressRemap
    if name == "scan_errors":
        from .hygiene import scan_errors
        return scan_errors
    if name == "findings":
        from .hygiene import findings
        return findings
    if name == "receipt":
        # the module is named receiptS so this attribute can only ever
        # resolve to the FUNCTION (a same-named submodule
        # import shadowed the function with the module object)
        from .receipts import receipt
        return receipt
    if name == "diff_workbooks":
        from .diffreport import diff_workbooks
        return diff_workbooks
    if name == "copy_format":
        from .styleverbs import copy_format
        return copy_format
    if name == "apply_profile":
        from .styleverbs import apply_profile
        return apply_profile
    raise AttributeError(name)
