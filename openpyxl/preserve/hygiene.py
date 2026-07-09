# paper-xlsx: LibreOffice-free hygiene checks (PLAN-v0.1 Batch 6)

"""Cheap first-pass checks for blind environments — measurements, never
judgments (the advisory fence stands: findings flag, they never decide).
"""

import re

ERROR_TOKENS = frozenset([
    "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
    "#SPILL!", "#CALC!", "#GETTING_DATA",
])

# both quote styles, ALWAYS (the recurring gate lesson)
_ERROR_CELL_RE = re.compile(
    br"<c\b([^>]*)\bt=(?:\"e\"|'e')([^>]*)>(.*?)</c>", re.S)
_R_ATTR_RE = re.compile(br"\br=(?:\"([A-Za-z]+\d+)\"|'([A-Za-z]+\d+)')")


def current_titles_by_part(wb, zin):
    """{part_name: CURRENT sheet title} for a preserve workbook — the
    package still uses ORIGINAL titles after an in-session rename, and
    sheets removed this session map to nothing (shared by scan_errors
    and the manifest — Batch-6 gate: stale/None attributions)."""
    from .saver import _package_info

    led = getattr(wb, "_paper_ledger", None)
    current_by_original = {}
    if led is not None:
        for ws_obj, original in getattr(led, "renames", {}).items():
            current_by_original[original] = ws_obj.title
    live_titles = {ws.title for ws in wb.worksheets}
    _wb_part, mapping = _package_info(zin)
    out = {}
    for title, part in mapping.items():
        current = current_by_original.get(title, title)
        if current in live_titles:
            out[part] = current
    return out


def _require_materialized_cells(wb, api):
    from openpyxl.workbook.workbook import _require_materialized_cells

    _require_materialized_cells(wb, api)


def scan_errors(wb):
    """Error evidence without LibreOffice (PR-1 §5): cached error tokens
    (from live cell values AND, under preserve, the original bytes — both
    load views from one workbook) plus #REF! markers inside formulas.

    Returns ``[{"address", "value", "source"}, ...]`` where source is
    "value" (a live cell holds the error), "cache" (the preserved bytes
    hold it), or "formula" (a formula's text contains #REF!).
    """
    _require_materialized_cells(wb, "scan_errors()")
    results = []
    seen = set()
    for ws in wb.worksheets:
        for (row, col), cell in sorted(ws._cells.items()):
            value = cell._value
            if cell.data_type == "f" and not isinstance(value, str):
                # ArrayFormula/DataTableFormula: scan the TEXT
                value = getattr(value, "text", None)
            address = "{0}!{1}".format(ws.title, cell.coordinate)
            if isinstance(value, str):
                if cell.data_type == "f":
                    if "#REF!" in value:
                        results.append({"address": address,
                                        "value": "#REF!",
                                        "source": "formula"})
                        seen.add(address)
                elif value.strip() in ERROR_TOKENS:
                    results.append({"address": address,
                                    "value": value.strip(),
                                    "source": "value"})
                    seen.add(address)

    source = getattr(wb, "_paper_source", None)
    if source:
        # the OTHER load view: cached <v> error tokens under t="e" in the
        # original bytes (invisible on a data_only=False load)
        import io
        import zipfile

        with zipfile.ZipFile(io.BytesIO(source)) as zin:
            title_by_part = current_titles_by_part(wb, zin)
            for part, title in sorted(title_by_part.items()):
                if part not in zin.namelist():
                    continue
                payload = zin.read(part)
                for m in _ERROR_CELL_RE.finditer(payload):
                    ref_m = _R_ATTR_RE.search(m.group(1) + m.group(2))
                    v_m = re.search(br"<v[^>]*>([^<]*)</v>", m.group(3))
                    if ref_m is None or v_m is None:
                        continue
                    ref = (ref_m.group(1) or ref_m.group(2)).decode("ascii")
                    address = "{0}!{1}".format(title, ref)
                    if address in seen:
                        continue
                    results.append({
                        "address": address,
                        "value": v_m.group(1).decode("utf-8", "replace"),
                        "source": "cache"})
                    seen.add(address)
    return results


# ---------------------------------------------------------------------
# hygiene findings (PLAN-v0.1 Batch 6.8) — ADVISORY lint only: findings
# flag, they never decide; every finding carries evidence addresses

class Finding:

    def __init__(self, kind, evidence, detail):
        self.kind = kind
        self.evidence = list(evidence)
        self.detail = detail

    def to_dict(self):
        return {"kind": self.kind, "evidence": list(self.evidence),
                "detail": self.detail}

    def __repr__(self):
        return "Finding({0!r}, {1} cells)".format(self.kind,
                                                  len(self.evidence))


FINDING_KINDS = (
    "hardcode-in-formula", "inconsistent-row-formula", "error-cell",
    "orphaned-name", "external-link", "hidden-sheet", "hidden-rows",
    "merged-hazard", "volatile", "magnitude-outlier",
)

# calendar/percent/counting constants that are not "hardcodes" worth
# flagging (advisory heuristic, documented)
_BLESSED_LITERALS = frozenset([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 24,
                               60, 100, 365, 366, 52, 1000])

# neutralize ROW NUMBERS inside cell references (digits following a
# column letter or $) so =A1*1.17 and =A2*1.17 share a shape — literals
# stay, so a changed literal breaks the pattern
_ROWNUM_RE = re.compile(r"(?<=[A-Za-z$])\d+")


def findings(wb):
    """Advisory hygiene findings for any workbook (PR-1 §5, taxonomy
    pinned). Measurements only — nothing here refuses or rewrites."""
    from openpyxl.formula import Tokenizer

    _require_materialized_cells(wb, "findings()")
    out = []

    # error-cell (reuses the LibreOffice-free scan)
    errors = scan_errors(wb)
    if errors:
        out.append(Finding("error-cell",
                           [e["address"] for e in errors],
                           "cells holding error values or #REF! markers"))

    # external-link
    links = getattr(wb, "_external_links", None) or []
    if links:
        targets = []
        for link in links:
            book = getattr(getattr(link, "file_link", None), "Target",
                           None)
            targets.append(str(book) if book else "external workbook")
        out.append(Finding("external-link", targets,
                           "values depend on files outside this package"))

    # hidden-sheet
    hidden = ["{0} ({1})".format(s.title, s.sheet_state)
              for s in wb._sheets
              if getattr(s, "sheet_state", "visible") != "visible"]
    if hidden:
        out.append(Finding("hidden-sheet", hidden,
                           "hidden content ships with the file"))

    volatile_cells = []
    hardcodes = []
    for ws in wb.worksheets:
        # hidden-rows
        hidden_rows = sorted(
            idx for idx, dim in ws.row_dimensions.items()
            if getattr(dim, "hidden", False))
        if hidden_rows:
            out.append(Finding(
                "hidden-rows",
                ["{0}!row {1}".format(ws.title, r) for r in hidden_rows],
                "hidden rows on sheet {0!r}".format(ws.title)))

        shapes = {}
        for (row, col), cell in sorted(ws._cells.items()):
            if cell.data_type != "f" or not isinstance(cell._value, str):
                continue
            address = "{0}!{1}".format(ws.title, cell.coordinate)
            formula = cell._value
            from .perception import _VOLATILE_RE

            if _VOLATILE_RE.search(formula):
                volatile_cells.append(address)
            try:
                tokens = Tokenizer(formula).items
            except Exception:
                tokens = []
            for token in tokens:
                if token.type == "OPERAND" and token.subtype == "NUMBER":
                    try:
                        num = float(token.value)
                    except ValueError:
                        continue
                    if num in _BLESSED_LITERALS and num == int(num):
                        continue
                    hardcodes.append("{0} ({1})".format(address,
                                                        token.value))
                    break
            # inconsistent-row-formula: group column runs by shape
            shape = _ROWNUM_RE.sub("#", formula)
            shapes.setdefault(col, []).append((row, shape, address))

        for col, entries in sorted(shapes.items()):
            entries.sort()
            run = []
            for row, shape, address in entries + [(None, None, None)]:
                if run and (row is None or row != run[-1][0] + 1):
                    _check_run(run, out, ws)
                    run = []
                if row is not None:
                    run.append((row, shape, address))

        # merged-hazard: merged interiors holding (shadowed) content
        hazards = []
        for rng in ws.merged_cells.ranges:
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    if (r, c) == (rng.min_row, rng.min_col):
                        continue
                    cell = ws._cells.get((r, c))
                    if cell is not None and cell._value is not None:
                        hazards.append("{0}!{1}".format(
                            ws.title, cell.coordinate))
        if hazards:
            out.append(Finding(
                "merged-hazard", hazards,
                "cells inside merged ranges hold shadowed values on "
                "sheet {0!r}".format(ws.title)))

        # magnitude-outlier: contiguous numeric column runs
        _magnitude_outliers(ws, out)

    _byte_level_merged_hazards(wb, out)

    if volatile_cells:
        out.append(Finding("volatile", sorted(volatile_cells),
                           "volatile functions recompute on every edit"))
    if hardcodes:
        out.append(Finding("hardcode-in-formula", hardcodes,
                           "numeric literals buried in formulas "
                           "(advisory heuristic: calendar/percent "
                           "constants are not flagged)"))

    # orphaned-name
    orphans = []
    sheetnames = {t.casefold() for t in wb.sheetnames}
    for name in list(wb.defined_names):
        dn = wb.defined_names[name]
        value = dn.value or ""
        if "#REF!" in value:
            orphans.append("{0} = {1}".format(name, value))
            continue
        m = re.match(r"^(?:'((?:[^']|'')+)'|([^'!\[]+))!", value)
        if m:
            title = (m.group(1) or m.group(2)).replace("''", "'")
            if title.casefold() not in sheetnames:
                orphans.append("{0} = {1}".format(name, value))
    if orphans:
        out.append(Finding("orphaned-name", orphans,
                           "defined names pointing at nothing"))
    return out


def _check_run(run, out, ws):
    if len(run) < 3:
        return
    from collections import Counter

    counts = Counter(shape for (_r, shape, _a) in run)
    dominant, dom_count = counts.most_common(1)[0]
    if dom_count < len(run) - dom_count or dom_count == len(run):
        return
    odd = [address for (_r, shape, address) in run if shape != dominant]
    out.append(Finding(
        "inconsistent-row-formula", odd,
        "formula run of {0} in column {1} on sheet {2!r} breaks pattern "
        "at {3} cell(s)".format(len(run), run[0][2].split("!")[1].rstrip(
            "0123456789"), ws.title, len(odd))))


def _magnitude_outliers(ws, out):
    import math

    by_col = {}
    for (row, col), cell in sorted(ws._cells.items()):
        if cell.data_type == "f":
            continue
        value = cell._value
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            continue
        if value == 0:
            continue
        by_col.setdefault(col, []).append((row, abs(float(value)),
                                           cell.coordinate))
    for col, entries in sorted(by_col.items()):
        entries.sort()
        run = []
        for row, mag, coord in entries + [(None, None, None)]:
            if run and (row is None or row != run[-1][0] + 1):
                if len(run) >= 5:
                    logs = sorted(math.log10(m) for (_r, m, _c) in run)
                    median = logs[len(logs) // 2]
                    odd = [c for (_r, m, c) in run
                           if abs(math.log10(m) - median) >= 3]
                    if odd:
                        out.append(Finding(
                            "magnitude-outlier",
                            ["{0}!{1}".format(ws.title, c) for c in odd],
                            "values orders of magnitude off their "
                            "column neighbours"))
                run = []
            if row is not None:
                run.append((row, mag, coord))


def _byte_level_merged_hazards(wb, out):
    """merged-hazard from the PRESERVED bytes: the model DISCARDS
    shadowed interior values at load (MergedCell._value is always None),
    so the real evidence lives only in the package (Batch-6 gate: the
    model-side detector could never fire on a loaded file)."""
    source = getattr(wb, "_paper_source", None)
    if not source:
        return
    import io
    import zipfile

    from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries

    with zipfile.ZipFile(io.BytesIO(source)) as zin:
        title_by_part = current_titles_by_part(wb, zin)
        byte_hazards = []
        for part, title in sorted(title_by_part.items()):
            payload = zin.read(part)
            merges = re.findall(
                br"<mergeCell\b[^>]*\bref=(?:\"([^\"]+)\"|'([^']+)')",
                payload)
            interiors = set()
            for g1, g2 in merges:
                ref = (g1 or g2).decode("ascii", "replace")
                try:
                    c1, r1, c2, r2 = range_boundaries(ref)
                except Exception:
                    continue
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        if (r, c) != (r1, c1):
                            interiors.add((r, c))
            if not interiors:
                continue
            for m in re.finditer(
                    br"<c\b[^>]*\br=(?:\"([A-Z]+\d+)\"|'([A-Z]+\d+)')"
                    br"[^>]*>(.*?)</c>", payload, re.S):
                ref = (m.group(1) or m.group(2)).decode("ascii")
                if b"<v" not in m.group(3) and b"<is" not in m.group(3):
                    continue
                if coordinate_to_tuple(ref) in interiors:
                    byte_hazards.append("{0}!{1}".format(title, ref))
        if byte_hazards:
            out.append(Finding(
                "merged-hazard", sorted(byte_hazards),
                "the PRESERVED bytes carry shadowed values inside "
                "merged ranges (invisible in the loaded model)"))
