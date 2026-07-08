# paper-xlsx: the package kernel (CONVENTIONS §7, PR-0 §5)

"""Package-level perception: semantic XML comparison and part-by-part
package diffing.

Under preserve mode, patch-writing *is* the save path; this module exists so
tests and agents can verify what a save did. The byte-identity invariant is
defined on part payloads, never whole-archive bytes: zip entry metadata
(timestamps, permissions) is out of scope.
"""

from .cells import CellsDiff, diff_cells
from .diff import (
    PackageDiff,
    PartChange,
    diff_package,
    xml_equivalent,
    xml_semantic_diff,
)

__all__ = [
    "CellsDiff",
    "PackageDiff",
    "PartChange",
    "diff_cells",
    "diff_package",
    "xml_equivalent",
    "xml_semantic_diff",
]
