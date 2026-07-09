# paper-xlsx: the edit receipt (PLAN-v0.1 Batch 6.6, PR-1 §5)

"""One artifact answering "what did that save actually do?": cells-diff
+ package-diff + confession + optional recalc/certification status."""

import io
import zipfile


class EditReceipt:

    SCHEMA = "edit_receipt"
    VERSION = 1

    def __init__(self, cells_changed, parts_changed, parts_added,
                 parts_removed, confession, recalc):
        self.cells_changed = cells_changed    # {part: {ref: kind}}
        self.parts_changed = parts_changed
        self.parts_added = parts_added
        self.parts_removed = parts_removed
        self.confession = confession          # loss-inventory style dicts
        self.recalc = recalc                  # dict or None

    def to_dict(self):
        return {
            "schema": self.SCHEMA,
            "version": self.VERSION,
            "cells_changed": {part: dict(refs)
                              for part, refs in self.cells_changed.items()},
            "parts_changed": list(self.parts_changed),
            "parts_added": list(self.parts_added),
            "parts_removed": list(self.parts_removed),
            "confession": list(self.confession),
            "recalc": self.recalc,
        }

    def __repr__(self):
        cells = sum(len(refs) for refs in self.cells_changed.values())
        return ("EditReceipt({0} cells, {1} parts changed, +{2}/-{3} "
                "parts)".format(cells, len(self.parts_changed),
                                len(self.parts_added),
                                len(self.parts_removed)))


def _read(source):
    if isinstance(source, bytes):
        return source
    if hasattr(source, "read"):
        pos = source.tell() if hasattr(source, "tell") else None
        if hasattr(source, "seek"):
            source.seek(0)
        data = source.read()
        if pos is not None and hasattr(source, "seek"):
            source.seek(pos)
        return data
    with open(source, "rb") as f:
        return f.read()


def receipt(before, after, *, recalc=None):
    """Build an :class:`EditReceipt` from two package states (paths,
    bytes, or binary file-likes). ``recalc``: an oracle result
    (RecalcResult/CertificationResult/Evaluation/WriteBackResult) whose
    ``to_dict()`` rides along."""
    from .crosscheck import _sheet_cells

    data_a, data_b = _read(before), _read(after)
    with zipfile.ZipFile(io.BytesIO(data_a)) as za, \
            zipfile.ZipFile(io.BytesIO(data_b)) as zb:
        names_a, names_b = set(za.namelist()), set(zb.namelist())
        parts_added = sorted(names_b - names_a)
        parts_removed = sorted(names_a - names_b)
        parts_changed = []
        cells_changed = {}
        for name in sorted(names_a & names_b):
            payload_a = za.read(name)
            payload_b = zb.read(name)
            if payload_a == payload_b:
                continue
            parts_changed.append(name)
            if name.startswith("xl/worksheets/") \
                    and name.endswith(".xml"):
                before_cells = _sheet_cells(payload_a)
                after_cells = _sheet_cells(payload_b)
                refs = {}
                for ref in sorted(set(before_cells) | set(after_cells)):
                    if before_cells.get(ref) != after_cells.get(ref):
                        if ref not in before_cells:
                            refs[ref] = "added"
                        elif ref not in after_cells:
                            refs[ref] = "removed"
                        else:
                            refs[ref] = "changed"
                if refs:
                    cells_changed[name] = refs

    from .inventory import scan_archive

    with zipfile.ZipFile(io.BytesIO(data_b)) as zb2:
        inventory = scan_archive(zb2, zb2.namelist())
    confession = [dict(loss) for loss in inventory.losses]

    recalc_dict = None
    if recalc is not None:
        recalc_dict = recalc.to_dict() if hasattr(recalc, "to_dict") \
            else dict(recalc)
    return EditReceipt(cells_changed, parts_changed, parts_added,
                       parts_removed, confession, recalc_dict)
