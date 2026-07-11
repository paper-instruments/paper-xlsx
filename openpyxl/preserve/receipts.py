# paper-xlsx: the edit receipt

"""One artifact answering "what did that save actually do?": cells-diff
+ package-diff + confession + optional recalc/certification status."""

import io
import hashlib
import zipfile
from collections import Counter

from openpyxl.errors import UnsupportedStructureError

from .zipguard import (
    MAX_ENTRIES as _MAX_ZIP_ENTRIES,
    MAX_PART_BYTES as _MAX_ZIP_PART,
    MAX_TOTAL_BYTES as _MAX_ZIP_UNCOMPRESSED,
)


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
    from .limits import read_bounded

    return read_bounded(source, context="receipt workbook")


def _validated_names(archive):
    infos = archive.infolist()
    names = [info.filename for info in infos]
    duplicates = sorted(name for name, count in Counter(names).items()
                        if count > 1)
    if duplicates:
        raise UnsupportedStructureError(
            "archive contains duplicate ZIP entry names ({0}); receipt "
            "generation refuses because choosing one copy could produce a "
            "false-clean receipt.".format(", ".join(duplicates)))
    if len(infos) > _MAX_ZIP_ENTRIES:
        raise UnsupportedStructureError(
            "archive declares {0} entries, past the {1}-entry cap; refusing "
            "before inflation.".format(len(infos), _MAX_ZIP_ENTRIES))
    oversized = next(
        (info for info in infos if info.file_size > _MAX_ZIP_PART), None)
    if oversized is not None:
        raise UnsupportedStructureError(
            "archive part {0!r} declares {1} uncompressed bytes, past the "
            "{2}-byte receipt cap; refusing before inflation.".format(
                oversized.filename, oversized.file_size, _MAX_ZIP_PART))
    total = sum(info.file_size for info in infos)
    if total > _MAX_ZIP_UNCOMPRESSED:
        raise UnsupportedStructureError(
            "archive declares {0} aggregate uncompressed bytes, past the "
            "{1}-byte cap; refusing before inflation.".format(
                total, _MAX_ZIP_UNCOMPRESSED))
    return names


def receipt(before, after, *, recalc=None):
    """Build an :class:`EditReceipt` from two package states (paths,
    bytes, or binary file-likes). ``recalc``: an oracle result
    (RecalcResult/CertificationResult/Evaluation/WriteBackResult) whose
    ``to_dict()`` rides along. The result must carry ``artifact_sha256``
    matching ``after``; unbound or cross-workbook verification refuses."""
    from .crosscheck import _sheet_cells

    data_a, data_b = _read(before), _read(after)
    from .zipguard import validate_package_bytes

    validate_package_bytes(data_a, context="receipt before-package")
    validate_package_bytes(data_b, context="receipt after-package")
    with zipfile.ZipFile(io.BytesIO(data_a)) as za, \
            zipfile.ZipFile(io.BytesIO(data_b)) as zb:
        names_a_list = _validated_names(za)
        names_b_list = _validated_names(zb)
        names_a, names_b = set(names_a_list), set(names_b_list)
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

    with zipfile.ZipFile(io.BytesIO(data_a)) as za2, \
            zipfile.ZipFile(io.BytesIO(data_b)) as zb2:
        before_inventory = scan_archive(za2, names_a_list)
        after_inventory = scan_archive(zb2, names_b_list)
    retained = {(loss["kind"], loss["location"], loss["detail"])
                for loss in after_inventory.losses}
    confession = []
    for loss in before_inventory.losses:
        key = (loss["kind"], loss["location"], loss["detail"])
        if key in retained:
            continue
        actual = dict(loss)
        actual["detail"] = "content present before save is absent from output"
        confession.append(actual)

    recalc_dict = None
    if recalc is not None:
        recalc_dict = recalc.to_dict() if hasattr(recalc, "to_dict") \
            else dict(recalc)
        claimed_digest = getattr(recalc, "artifact_sha256", None) \
            or recalc_dict.get("artifact_sha256")
        actual_digest = hashlib.sha256(data_b).hexdigest()
        if not claimed_digest:
            raise UnsupportedStructureError(
                "the supplied recalc/certification result is not bound to "
                "an artifact digest, so it cannot verify this receipt")
        if claimed_digest != actual_digest:
            raise UnsupportedStructureError(
                "the supplied recalc/certification result describes a "
                "different workbook (artifact SHA-256 does not match the "
                "receipt output)")
    return EditReceipt(cells_changed, parts_changed, parts_added,
                       parts_removed, confession, recalc_dict)
