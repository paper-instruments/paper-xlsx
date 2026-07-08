# paper-xlsx: the structural-edit guard (PLAN Phase 6a; PR-0 §8)

"""Analyze what a row/column shift would strand.

The scariest damage in the model is here: ``insert_rows`` moves cells while
updating NOTHING — not formulas, not defined names, not chart ranges — so
one inserted row silently corrupts every SUM below it with numbers that
look plausible (measured: LibreOffice computes 1100/6399/5400 where the
correct answers are 7499/6500 — OPEN-QUESTIONS Q11). Under preserve mode
the shift refuses with the precise list of what would break; the stock path
keeps stock behavior plus a loud warning.
"""

import io
import re
import zipfile

MAX_COL = 1 << 20
MAX_ROW = 1 << 22


def shift_bounds(kind, index):
    """The cell region a shift at ``index`` moves or destroys: everything
    at or after the index (deletes also destroy the range itself)."""
    if kind in ("insert_rows", "delete_rows"):
        return (1, index, MAX_COL, MAX_ROW)
    return (index, 1, MAX_COL, MAX_ROW)


def _intersects(bounds, min_col, min_row, max_col, max_row):
    b_min_col, b_min_row, b_max_col, b_max_row = bounds
    return not (b_max_col < min_col or b_min_col > max_col
                or b_max_row < min_row or b_min_row > max_row)


def analyze_shift(ws, kind, index):
    """Everything a shift would strand, as human-readable impact lines."""
    from openpyxl.utils.cell import range_boundaries

    from .perception import dependency_sketch

    wb = ws.parent
    bounds = shift_bounds(kind, index)
    min_col, min_row, max_col, max_row = bounds
    impacts = []

    # formulas (cross-sheet included) whose references intersect the region
    sketch = dependency_sketch(wb)
    formulas = sketch.cells_referencing(ws.title, bounds)
    if formulas:
        shown = ", ".join(formulas[:8])
        if len(formulas) > 8:
            shown += ", ... ({0} total)".format(len(formulas))
        impacts.append(
            "formulas referencing the shifted cells would keep their old "
            "ranges and silently compute wrong numbers: {0}".format(shown))

    # defined names pointing into the region
    stranded_names = []
    name_sources = [(None, wb.defined_names)]
    name_sources += [(sheet.title, sheet.defined_names)
                     for sheet in wb.worksheets]
    for scope, names in name_sources:
        for name, dn in names.items():
            try:
                for dest_sheet, dest_ref in dn.destinations:
                    if dest_sheet != ws.title:
                        continue
                    dest_bounds = range_boundaries(dest_ref.replace("$", ""))
                    if _intersects(dest_bounds, *_norm(bounds)):
                        stranded_names.append(name)
                        break
            except Exception:
                continue
    if stranded_names:
        impacts.append(
            "defined name(s) {0} would keep pointing at the old "
            "cells".format(", ".join(sorted(set(stranded_names)))))

    # merged ranges
    merged = [str(r) for r in ws.merged_cells.ranges
              if _ranges_hit(r, bounds)]
    if merged:
        impacts.append("merged range(s) {0} would not move with their "
                       "content".format(", ".join(sorted(merged))))

    # conditional formatting and data validation ranges
    cf_hit = []
    for cf in ws.conditional_formatting:
        for rng in getattr(cf.sqref, "ranges", []):
            if _ranges_hit(rng, bounds):
                cf_hit.append(str(cf.sqref))
                break
    if cf_hit:
        impacts.append("conditional formatting on {0} would keep the old "
                       "ranges".format(", ".join(sorted(set(cf_hit)))))
    dv_hit = []
    if ws.data_validations:
        for dv in ws.data_validations.dataValidation:
            for rng in getattr(dv.sqref, "ranges", []):
                if _ranges_hit(rng, bounds):
                    dv_hit.append(str(dv.sqref))
                    break
    if dv_hit:
        impacts.append("data validation on {0} would keep the old "
                       "ranges".format(", ".join(sorted(set(dv_hit)))))

    # tables
    tables_hit = [name for name, ref in ws.tables.items()
                  if _ref_hit(ref, bounds)]
    if tables_hit:
        impacts.append("table(s) {0} would keep their old extents".format(
            ", ".join(sorted(tables_hit))))

    # preserved charts: their XML is raw retained bytes — a shift makes them
    # point at wrong rows with no error anywhere (PR-0 §8: refusal is the
    # only honest v0 answer on chart-referenced sheets)
    charts = _charts_referencing(wb, ws.title)
    if charts:
        impacts.append(
            "preserved chart(s) ({0}) reference this sheet; their series "
            "ranges live in preserved bytes and would point at the wrong "
            "rows".format(", ".join(sorted(charts))))

    return impacts


def _norm(bounds):
    return bounds


def _ranges_hit(cell_range, bounds):
    return _intersects((cell_range.min_col, cell_range.min_row,
                        cell_range.max_col, cell_range.max_row), *bounds)


def _ref_hit(ref, bounds):
    from openpyxl.utils.cell import range_boundaries

    try:
        return _intersects(range_boundaries(ref.replace("$", "")), *bounds)
    except Exception:
        return True   # unparseable: assume affected (conservative)


def _charts_referencing(wb, sheet_title):
    """Names of retained chart parts whose XML mentions the sheet."""
    source = getattr(wb, "_paper_source", None)
    if not source:
        return []
    needles = [sheet_title.encode("utf-8")]
    quoted = "'{0}'".format(sheet_title.replace("'", "''")).encode("utf-8")
    needles.append(quoted)
    hits = []
    with zipfile.ZipFile(io.BytesIO(source)) as z:
        for name in z.namelist():
            if (name.startswith("xl/charts/") and name.endswith(".xml")
                    and "/_rels/" not in name):
                payload = z.read(name)
                if any(needle in payload for needle in needles):
                    hits.append(name)
    return hits


STOCK_WARNING = (
    "{0}() moves cells but updates NOTHING that points at them: formulas "
    "keep their old ranges, defined names and chart series keep their old "
    "cells — the numbers that come out will look plausible and be wrong. "
    "Open the file with preserve=True to get a safety analysis instead."
)
