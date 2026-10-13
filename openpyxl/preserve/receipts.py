# paper-xlsx: the edit receipt

"""One artifact answering "what did that save actually do?": cells-diff
+ package-diff + confession + optional recalc/certification status."""

import io
import hashlib
import re
import zipfile
from collections import Counter
from xml.etree.ElementTree import ParseError, fromstring

from openpyxl.errors import UnsupportedStructureError
from openpyxl.xml.constants import SHEET_MAIN_NS

class EditReceipt:

    SCHEMA = "edit_receipt"
    VERSION = 2

    def __init__(self, cells_changed, parts_changed, parts_added,
                 parts_removed, confession, recalc, derived_effects=()):
        self.cells_changed = cells_changed    # {part: {ref: kind}}
        self.parts_changed = parts_changed
        self.parts_added = parts_added
        self.parts_removed = parts_removed
        self.confession = confession          # loss-inventory style dicts
        self.recalc = recalc                  # dict or None
        self.derived_effects = list(derived_effects)

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
            "derived_effects_version": 1,
            "derived_effects": list(self.derived_effects),
        }

    def __repr__(self):
        cells = sum(len(refs) for refs in self.cells_changed.values())
        return ("EditReceipt({0} cells, {1} parts changed, +{2}/-{3} "
                "parts)".format(cells, len(self.parts_changed),
                                len(self.parts_added),
                                len(self.parts_removed)))


def _read(source):
    from .sourceio import read_source_bytes

    return read_source_bytes(source, context="receipt workbook")


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
    return names


_MAX_ROW = 1048576
_MAX_COL = 16384


def _cell_formula_cache_state(payload):
    """{ref: (has formula, has cached value)} for each cell, read as XML
    like the rest of the receipt (a regex over "<c ...>body</c>" ran a
    self-closing cell on into the next cell), and each array formula as
    (anchor ref, (min_col, min_row, max_col, max_row)). An open ref such as
    "D:D" runs to the sheet's edge."""
    from openpyxl.utils.cell import range_boundaries

    from .crosscheck import _iter_sheet_cells

    states = {}
    arrays = []
    for ref, cell in _iter_sheet_cells(payload):
        ns = cell.tag[:-1]
        formula = cell.find(ns + "f")
        states[ref] = (formula is not None, cell.find(ns + "v") is not None)
        if formula is not None and formula.get("t") == "array":
            try:
                min_col, min_row, max_col, max_row = range_boundaries(
                    formula.get("ref") or ref)
            except (TypeError, ValueError):
                continue
            arrays.append((ref, (min_col or 1, min_row or 1,
                                 max_col or _MAX_COL, max_row or _MAX_ROW)))
    return states, arrays


def _coordinate(ref):
    """(row, col) of an A1 reference, or None when it is not one."""
    from openpyxl.utils.cell import coordinate_to_tuple

    try:
        return coordinate_to_tuple(ref)
    except (TypeError, ValueError):
        return None


# An array at most this many rows tall is found by its first row; taller
# ones (a whole-column array) are few and checked directly.
_SHORT_ARRAY_ROWS = 1024


def _cleared_array_member(arrays, old, new):
    """A test for whether a reference lies in an array formula's range whose
    anchor lost its cached value in the same comparison. A save strips the
    whole range at once; a value cleared under an anchor that kept its cache
    was an edit, not a derived effect. Short arrays are kept sorted by first
    row, so a reference checks only those starting within _SHORT_ARRAY_ROWS
    rows above it, and memory stays one entry per array."""
    from bisect import bisect_left, bisect_right

    short = []
    tall = []
    for anchor, (min_col, min_row, max_col, max_row) in arrays:
        if old.get(anchor) != (True, True) \
                or new.get(anchor) != (True, False):
            continue
        bounds = (min_row, min_col, max_row, max_col)
        if max_row - min_row < _SHORT_ARRAY_ROWS:
            short.append(bounds)
        else:
            tall.append(bounds)
    short.sort()
    starts = [bounds[0] for bounds in short]

    def member(ref):
        coordinate = _coordinate(ref)
        if coordinate is None:
            return False
        row, col = coordinate
        first = bisect_left(starts, row - _SHORT_ARRAY_ROWS + 1)
        last = bisect_right(starts, row)
        for index in range(first, last):
            _min_row, min_col, max_row, max_col = short[index]
            if row <= max_row and min_col <= col <= max_col:
                return True
        return any(min_row <= row <= max_row and min_col <= col <= max_col
                   for min_row, min_col, max_row, max_col in tall)

    return member


def _pivot_refresh_enabled(payload):
    """Return whether a main-namespace pivot cache requests refresh."""
    try:
        root = fromstring(payload)
    except (ParseError, ValueError, TypeError):
        return False
    expected = "{{{0}}}pivotCacheDefinition".format(SHEET_MAIN_NS)
    refresh = root.attrib.get("refreshOnLoad", "false").casefold()
    enabled = root.attrib.get("enableRefresh")
    return root.tag == expected and refresh in ("1", "true") \
        and (enabled is None or enabled.casefold() in ("1", "true"))


def _derived_effects(za, zb, names_a, names_b, *, ledger=None,
                     workbook=None, invalidated=None):
    effects = []
    image_rels = {
        request["rels_part"]
        for request in getattr(ledger, "image_replacements", {}).values()
    }
    pivot_parts = set(getattr(ledger, "pivot_refresh_requests", ()))
    if ledger is not None and pivot_parts:
        from .pivots import source_impacts

        if workbook is None:
            sheet = next(iter(ledger.value_overwrites), None)
            workbook = getattr(sheet, "parent", None)
        if workbook is not None:
            for impact in source_impacts(workbook, ledger):
                if impact["part"] not in pivot_parts:
                    continue
                effects.append({
                    "kind": "pivot_source_changed_requires_refresh",
                    "part": impact["part"],
                    "cause": impact["cause"],
                    "pivots": impact["pivots"],
                    "source": impact["source"],
                    "requirement": "excel_refresh_on_open",
                })
    cause = "formula_changed" if getattr(ledger, "formulas_changed", False) \
        else "input_changed"
    if "xl/calcChain.xml" in names_a and "xl/calcChain.xml" not in names_b:
        effects.append({"kind": "calc_chain_removed",
                        "part": "xl/calcChain.xml", "cause": cause})
    for name in sorted(names_a & names_b):
        before = za.read(name)
        after = zb.read(name)
        if before == after:
            continue
        if name.startswith("xl/worksheets/") and name.endswith(".xml"):
            old, arrays = _cell_formula_cache_state(before)
            new, _arrays = _cell_formula_cache_state(after)
            # (row, col) the save stripped; compared as coordinates, since
            # a part may spell a reference d3 or D03
            if invalidated is not None:
                stripped = set(invalidated.get(name, ()))

                def array_cache_cleared(ref):
                    return _coordinate(ref) in stripped
            else:
                array_cache_cleared = _cleared_array_member(arrays, old, new)
            for ref in sorted(set(old) & set(new)):
                # An array formula's result spans its range, and the cells
                # besides its anchor hold no formula. A save that strips
                # caches names the cells it stripped; comparing two packages
                # infers them from the anchor losing its cache too.
                if (old[ref] == (True, True) and new[ref] == (True, False)) \
                        or (old[ref] == (False, True)
                            and new[ref] == (False, False)
                            and array_cache_cleared(ref)):
                    effects.append({
                        "kind": "formula_cache_removed", "part": name,
                        "cell": ref, "cause": cause,
                    })
        elif name.startswith("xl/charts/") and name.endswith(".xml"):
            old_count = len(re.findall(
                br"<(?:(?:\w+):)?(?:numCache|strCache)\b", before))
            new_count = len(re.findall(
                br"<(?:(?:\w+):)?(?:numCache|strCache)\b", after))
            for _ in range(max(0, old_count - new_count)):
                effects.append({
                    "kind": "chart_cache_removed", "part": name,
                    "cause": "chart_repointed",
                })
        elif name in pivot_parts \
                or name.startswith("xl/pivotCache/pivotCacheDefinition"):
            old_enabled = _pivot_refresh_enabled(before)
            new_enabled = _pivot_refresh_enabled(after)
            if not old_enabled and new_enabled:
                effects.append({
                    "kind": "pivot_refresh_on_load_enabled", "part": name,
                    "cause": "explicit_request",
                })
        if name.endswith(".rels"):
            rel_cause = "image_replaced" if name in image_rels \
                else "supported_lifecycle_edit"
            effects.append({"kind": "relationship_changed", "part": name,
                            "cause": rel_cause})
        elif name == "[Content_Types].xml":
            effects.append({"kind": "content_type_changed", "part": name,
                            "cause": "supported_lifecycle_edit"})
    for name in sorted(names_b - names_a):
        if name.endswith(".rels"):
            effects.append({"kind": "relationship_added", "part": name,
                            "cause": "supported_lifecycle_edit"})
    # calcPr is workbook-level recalculation metadata. Resolve the workbook
    # part rather than assuming xl/workbook.xml.
    try:
        from .saver import _package_info

        workbook_part, _sheets = _package_info(za)
        if workbook_part in names_b \
                and za.read(workbook_part) != zb.read(workbook_part):
            before_match = re.search(
                br"<(?:(?:\w+):)?calcPr\b[^>]*/?>", za.read(workbook_part))
            after_match = re.search(
                br"<(?:(?:\w+):)?calcPr\b[^>]*/?>", zb.read(workbook_part))
            before_calc = before_match.group(0) if before_match else None
            after_calc = after_match.group(0) if after_match else None
            if before_calc != after_calc:
                effects.append({
                    "kind": "recalculation_metadata_changed",
                    "part": workbook_part, "cause": cause,
                })
    except (KeyError, ValueError):
        pass
    return effects


def receipt(before, after, *, recalc=None, _ledger=None, _workbook=None,
            _invalidated=None):
    """Build an `EditReceipt` from two package states (paths,
    bytes, or binary file-likes). ``recalc``: an oracle result
    (RecalcResult/CertificationResult/Evaluation) whose
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
        derived_effects = _derived_effects(
            za, zb, names_a, names_b, ledger=_ledger,
            workbook=_workbook, invalidated=_invalidated)

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
                       parts_removed, confession, recalc_dict,
                       derived_effects)
