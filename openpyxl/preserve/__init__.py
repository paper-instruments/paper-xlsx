# paper-xlsx: the spine (CONVENTIONS §3, PR-0 §§3-6)

"""Preserve mode: the original package is the source of truth; the object
model is a source of edits to it.

Loaded via ``load_workbook(path, preserve=True)``. Untouched parts survive
byte-identical by construction (raw compressed-stream copy where possible);
touched worksheet parts are spliced, never re-serialized (Phase 2c).
"""

from .inventory import LossInventory, scan_archive
from .saver import save_preserved

__all__ = ["DirtyLedger", "LossInventory", "scan_archive", "save_preserved"]


def __getattr__(name):
    # DirtyLedger lives in .ledger, which several early-imported modules
    # (cell, styleable, worksheet) pull helpers from; exposing it lazily
    # here keeps this package importable from anywhere without cycles
    if name == "DirtyLedger":
        from .ledger import DirtyLedger
        return DirtyLedger
    raise AttributeError(name)
