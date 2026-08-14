# paper-xlsx: the spine

"""Preserve mode: the original package is the source of truth; the object
model is a source of edits to it.

Enabled by default for editable OOXML workbooks loaded with
``load_workbook(path)``. Untouched parts survive byte-identical by
construction (raw compressed-stream copy where possible); touched worksheet
parts are spliced, never re-serialized.
"""

__all__ = ["AddressRemap", "scan_errors", "receipt", "diff_workbooks",
           "copy_format"]


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
    raise AttributeError(name)
