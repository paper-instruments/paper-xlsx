# paper-xlsx: the structural-edit guard

"""Analyze what a row/column shift would strand.

The scariest damage in the model is here: ``insert_rows`` moves cells while
updating NOTHING — not formulas, not defined names, not chart ranges — so
one inserted row silently corrupts every SUM below it with numbers that
look plausible (measured: LibreOffice computes 1100/6399/5400 where the
correct answers are 7499/6500). Under preserve mode
the shift refuses with the precise list of what would break; the stock path
keeps stock behavior plus a loud warning.
"""

import io
import re
import zipfile

from openpyxl.errors import UnsupportedStructureError

MAX_COL = 1 << 20
MAX_ROW = 1 << 22

# hard sheet limits (ECMA-376): shifting occupied cells past these is a
# BoundaryViolationError, not a silent wrap or drop
EXCEL_MAX_ROW = 1048576
EXCEL_MAX_COL = 16384


class StructuralEditTransaction:
    """In-memory rollback boundary for a preserve-mode row/column shift."""

    def __init__(self, ws, operation, warn_protection=False,
                 reference_rewrites=()):
        self.ws = ws
        self.operation = operation
        self.warn_protection = warn_protection
        self.reference_rewrites = tuple(reference_rewrites)
        self._snapshot = _capture_structural_state(ws.parent)
        self._active = True

    def commit(self):
        if not self._active:
            return
        if not self.warn_protection:
            self._active = False
            return
        led = self.ws.parent._paper_ledger
        led.protection_warned.add(self.ws)
        import warnings

        from openpyxl.errors import ProtectedWriteWarning

        warnings.warn(ProtectedWriteWarning(
            "{0}() on protected sheet {1!r}: the edit proceeds — "
            "protection is reported, never enforced — but Excel itself "
            "would block it. Set wb.strict_protection = True to refuse "
            "instead.".format(self.operation, self.ws.title)), stacklevel=4)
        self._active = False

    def rollback(self):
        if not self._active:
            return
        _restore_structural_state(self.ws.parent, self._snapshot)
        self._active = False


def _unique_by_identity(values):
    seen = set()
    unique = []
    for value in values:
        marker = id(value)
        if marker in seen:
            continue
        seen.add(marker)
        unique.append(value)
    return unique


def _capture_structural_state(wb):
    """Capture every model/ledger surface mutated by ``apply_model_shift``.

    Object identities are retained. Rollback restores existing Cell,
    DefinedName, rule, validation, dimension, and chart-reference objects in
    place, so callers holding those objects do not observe a half-shift.
    """
    sheet_states = []
    for sheet in wb.worksheets:
        cells = dict(sheet._cells)
        cell_states = []
        for cell in _unique_by_identity(cells.values()):
            link = getattr(cell, "_hyperlink", None)
            cell_states.append((
                cell,
                getattr(cell, "row", None),
                getattr(cell, "column", None),
                getattr(cell, "_value", None),
                getattr(cell, "_data_type", None),
                getattr(cell, "_style", None),
                getattr(link, "ref", None) if link is not None else None,
                getattr(link, "location", None) if link is not None else None,
            ))

        cf_rules = sheet.conditional_formatting._cf_rules
        rule_formulas = [
            (rule, rule.formula)
            for rules in cf_rules.values()
            for rule in rules
        ]
        validations = sheet.data_validations
        dv_list = validations.dataValidation if validations is not None else []
        dv_states = [
            (dv, dv.formula1, dv.formula2, dv.sqref)
            for dv in dv_list
        ]
        filter_columns = list(sheet.auto_filter.filterColumn)
        filter_column_states = [
            (column, column.colId) for column in filter_columns]
        sort_state = sheet.auto_filter.sortState
        sort_conditions = list(sort_state.sortCondition) \
            if sort_state is not None else []
        scenario_list = getattr(sheet, "scenarios", None)
        scenario_inputs = [
            input_cell
            for scenario in getattr(scenario_list, "scenario", ()) or ()
            for input_cell in scenario.inputCells
        ]
        row_items = dict(sheet.row_dimensions)
        col_items = dict(sheet.column_dimensions)
        row_dim_states = [(dim, dim.index)
                          for dim in _unique_by_identity(row_items.values())]
        col_dim_states = [(dim, dim.index, dim.min, dim.max)
                          for dim in _unique_by_identity(col_items.values())]
        sheet_states.append({
            "sheet": sheet,
            "cells": cells,
            "cell_states": cell_states,
            "current_row": sheet._current_row,
            "merged_cells": sheet.merged_cells,
            "cf_rules": cf_rules,
            "rule_formulas": rule_formulas,
            "dv_list": dv_list,
            "dv_states": dv_states,
            "auto_filter_ref": sheet.auto_filter.ref,
            "filter_columns": filter_columns,
            "filter_column_states": filter_column_states,
            "sort_state": sort_state,
            "sort_state_ref": getattr(sort_state, "ref", None),
            "sort_conditions": sort_conditions,
            "sort_condition_refs": [
                condition.ref for condition in sort_conditions],
            "scenario_list": scenario_list,
            "scenario_sqref": getattr(scenario_list, "sqref", None),
            "scenario_inputs": scenario_inputs,
            "scenario_input_refs": [item.r for item in scenario_inputs],
            "row_items": row_items,
            "col_items": col_items,
            "row_dim_states": row_dim_states,
            "col_dim_states": col_dim_states,
        })

    name_states = []
    for names in [wb.defined_names] + [s.defined_names for s in wb.worksheets]:
        name_states.extend((names[name], names[name].attr_text)
                           for name in names)

    chart_ref_states = []
    for sheet in wb.worksheets:
        for chart in getattr(sheet, "_charts", ()):
            chart_ref_states.extend(
                (ref, ref.f) for ref in _chart_source_ref_objects(chart))

    from .references import formula_surfaces

    formula_surface_states = [
        (surface, surface.value) for surface in formula_surfaces(wb)
        if not surface.cell
    ]

    led = wb._paper_ledger
    registries = []
    for name in ("_fonts", "_fills", "_borders", "_alignments",
                 "_protections", "_number_formats", "_cell_styles"):
        registry = getattr(wb, name)
        registries.append((registry, list(registry), registry.clean,
                           dict(registry._dict)))
    ledger_state = None
    if led is not None:
        def clone(value):
            if isinstance(value, dict):
                return {key: clone(item) for key, item in value.items()}
            if isinstance(value, set):
                return set(value)
            if isinstance(value, list):
                return list(value)
            return value

        mutable = tuple(
            name for name in led.__slots__
            if isinstance(getattr(led, name), (dict, list, set)))
        ledger_state = {
            "mutable": {
                name: (getattr(led, name), clone(getattr(led, name)))
                for name in mutable
            },
            "scalars": {
                name: clone(getattr(led, name)) for name in led.__slots__
                if name not in mutable
            },
        }
    return {
        "sheets": sheet_states,
        "names": name_states,
        "chart_refs": chart_ref_states,
        "formula_surfaces": formula_surface_states,
        "registries": registries,
        "dxfs": list(wb._differential_styles.styles),
        "ledger": ledger_state,
    }


def _restore_structural_state(wb, snapshot):
    for state in snapshot["sheets"]:
        sheet = state["sheet"]
        for (cell, row, column, value, data_type, style, link_ref,
             link_location) in \
                state["cell_states"]:
            cell.row = row
            cell.column = column
            if data_type is not None:
                cell._value = value
                cell._data_type = data_type
            cell._style = style
            link = getattr(cell, "_hyperlink", None)
            if link is not None:
                link.ref = link_ref
                link.location = link_location
        sheet._cells.clear()
        sheet._cells.update(state["cells"])
        sheet._current_row = state["current_row"]
        sheet.merged_cells = state["merged_cells"]
        sheet.conditional_formatting._cf_rules = state["cf_rules"]
        for rule, formula in state["rule_formulas"]:
            rule.formula = formula
        if sheet.data_validations is not None:
            sheet.data_validations.dataValidation = state["dv_list"]
        for dv, formula1, formula2, sqref in state["dv_states"]:
            dv.formula1 = formula1
            dv.formula2 = formula2
            dv.sqref = sqref
        sheet.auto_filter.ref = state["auto_filter_ref"]
        sheet.auto_filter.filterColumn = state["filter_columns"]
        for column, col_id in state["filter_column_states"]:
            column.colId = col_id
        sheet.auto_filter.sortState = state["sort_state"]
        if state["sort_state"] is not None:
            state["sort_state"].ref = state["sort_state_ref"]
            state["sort_state"].sortCondition = state["sort_conditions"]
        for condition, ref in zip(
                state["sort_conditions"], state["sort_condition_refs"]):
            condition.ref = ref
        sheet.scenarios = state["scenario_list"]
        if state["scenario_list"] is not None:
            state["scenario_list"].sqref = state["scenario_sqref"]
        for item, ref in zip(
                state["scenario_inputs"], state["scenario_input_refs"]):
            item.r = ref
        for dim, index in state["row_dim_states"]:
            dim.index = index
        for dim, index, min_value, max_value in state["col_dim_states"]:
            dim.index = index
            dim.min = min_value
            dim.max = max_value
        sheet.row_dimensions.clear()
        sheet.row_dimensions.update(state["row_items"])
        sheet.column_dimensions.clear()
        sheet.column_dimensions.update(state["col_items"])

    for name, attr_text in snapshot["names"]:
        name.attr_text = attr_text
    for ref, formula in snapshot["chart_refs"]:
        ref.f = formula
    for surface, value in snapshot["formula_surfaces"]:
        surface.replace(value)
    for registry, values, clean, index in snapshot["registries"]:
        registry[:] = values
        registry.clean = clean
        registry._dict = index
    wb._differential_styles.styles[:] = snapshot["dxfs"]

    led_state = snapshot["ledger"]
    led = wb._paper_ledger
    if led is not None and led_state is not None:
        for name, (container, value) in led_state["mutable"].items():
            if isinstance(container, dict):
                container.clear()
                container.update(value)
            elif isinstance(container, set):
                container.clear()
                container.update(value)
            elif isinstance(container, list):
                container[:] = value
            setattr(led, name, container)
        for name, value in led_state["scalars"].items():
            setattr(led, name, value)


class AddressRemap:
    """How one structural edit moved addresses:
    every pre-edit address must be remapped through this, never reused.

    ``map('Model!B12') -> 'Model!B13'``; addresses whose cells the edit
    deleted map to ``None``; addresses on other sheets (or untouched by
    the shift) come back unchanged. Accepts bare cells, ranges, and
    sheet-qualified forms; ``$`` markers are kept positionally, matching
    the rewriter's Excel semantics."""

    def __init__(self, sheet_title, operation, index, amount):
        self.sheet_title = sheet_title
        self.operation = operation
        self.index = index
        self.amount = amount

    def map(self, address):
        from .rewrite import REF_ERROR, shift_ref

        sheet, ref = None, address
        m = _SHEET_PREFIX_RE.match(address)
        if m:
            sheet = (m.group(1) or m.group(2)).replace("''", "'")
            ref = m.group(3)
        if sheet is not None \
                and sheet.casefold() != self.sheet_title.casefold():
            return address
        axis = "rows" if self.operation.endswith("_rows") else "cols"
        is_delete = self.operation.startswith("delete")
        shifted = shift_ref(ref, axis, self.index, self.amount, is_delete)
        if shifted == REF_ERROR:
            return None
        if sheet is None:
            return shifted
        return address[:m.end(0) - len(m.group(3))] + shifted

    def __repr__(self):
        return "AddressRemap({0!r}, {1}, index={2}, amount={3})".format(
            self.sheet_title, self.operation, self.index, self.amount)


_SHEET_PREFIX_RE = re.compile(r"^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$")


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
    # point at wrong rows with no error anywhere (refusal is the
    # only honest v0 answer on chart-referenced sheets)
    led_ref = getattr(wb, "_paper_ledger", None)
    chart_title = led_ref.renames.get(ws, ws.title) if led_ref else ws.title
    charts = _charts_referencing(wb, chart_title)
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


def _title_needles(sheet_title):
    """Byte needles for a sheet title inside preserved XML: raw and
    XML-escaped forms, quoted variants, all lower-cased for case-insensitive
    search (Excel resolves sheet names case-insensitively)."""
    escaped = (sheet_title.replace("&", "&amp;").replace("<", "&lt;")
               .replace(">", "&gt;"))
    variants = {sheet_title, escaped,
                "'{0}'".format(sheet_title.replace("'", "''")),
                "'{0}'".format(escaped.replace("'", "''")),
                "'{0}'".format(escaped.replace("'", "&apos;&apos;"))}
    return [v.encode("utf-8").lower() for v in variants]


def _charts_referencing(wb, sheet_title):
    """Names of retained chart parts whose XML mentions the sheet."""
    source = getattr(wb, "_paper_source", None)
    if not source:
        return []
    needles = _title_needles(sheet_title)
    hits = []
    with zipfile.ZipFile(io.BytesIO(source)) as z:
        for name in z.namelist():
            if (name.startswith("xl/charts/") and name.endswith(".xml")
                    and "/_rels/" not in name):
                payload = z.read(name).lower()
                if any(needle in payload for needle in needles):
                    hits.append(name)
    return hits


STOCK_WARNING = (
    "{0}() moves cells but updates NOTHING that points at them: formulas "
    "keep their old ranges, defined names and chart series keep their old "
    "cells — the numbers that come out will look plausible and be wrong. "
    "Open the file with preserve=True to get a safety analysis instead."
)


# ---------------------------------------------------------------------
# performing the shift (model fixups + byte renumber)

def shift_blockers(ws, operation, index, amount=1):
    """Content that makes a shift unsafe to REWRITE in v0 — anything whose
    references live outside the fully-modeled set. A non-empty result means
    refusal (with analyze_shift providing the victim list)."""
    wb = ws.parent
    blockers = []
    # multiple shifts per session compose: model fixups and snapshot
    # rebases run at edit time in order, the byte renumber replays the
    # recorded ops in order at save (retired the
    # one-shift-per-session refusal)
    led_ref = getattr(wb, "_paper_ledger", None)
    lookup_title = led_ref.renames.get(ws, ws.title) if led_ref else ws.title
    numeric_chart_parts = _charts_with_numeric_formula_entities(
        wb, lookup_title)
    if numeric_chart_parts:
        blockers.append(
            "chart part(s) {0} contain numeric character references in "
            "formula text; exact reference rewriting cannot be guaranteed"
            .format(", ".join(numeric_chart_parts)))
    from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula
    from .rewrite import shift_formula, title_in_string_literals

    axis = "rows" if "rows" in operation else "cols"
    is_delete = operation.startswith("delete")
    for other in wb.worksheets:
        for cell in other._cells.values():
            formula = cell._value
            formula_text = (formula if isinstance(formula, str)
                            else getattr(formula, "text", None))
            if isinstance(formula_text, str) and \
                    _three_d_formula_references_sheet(
                        wb, formula_text, ws.title):
                blockers.append(
                    "3-D formula {0}!{1} spans the shifted sheet; one "
                    "3-D reference cannot represent a shift applied to "
                    "only one member sheet".format(
                        other.title, cell.coordinate))
                continue
            if not isinstance(formula, ArrayFormula) \
                    or not isinstance(formula_text, str):
                if isinstance(formula, DataTableFormula):
                    for ref in (formula.r1, formula.r2):
                        if not isinstance(ref, str) or "!" not in ref:
                            continue
                        _rewritten, changed = shift_formula(
                            "=" + ref, other.title, ws.title, axis, index,
                            amount, is_delete)
                        if changed:
                            blockers.append(
                                "what-if data table {0}!{1} references the "
                                "shifted sheet through {2!r}; rewriting "
                                "data-table metadata is not supported"
                                .format(other.title, cell.coordinate, ref))
                continue
            _rewritten, changed = shift_formula(
                formula_text, other.title, ws.title, axis, index, amount,
                is_delete)
            if changed or title_in_string_literals(
                    formula_text, ws.title):
                blockers.append(
                    "array formula {0}!{1} references the shifted sheet; "
                    "rewriting loaded multi-cell formulas is not supported"
                    .format(other.title, cell.coordinate))
    # in-session charts are model-rendered at save: a delete that removes
    # their charted cells has no honest rewrite — block BEFORE any cell
    # moves
    if led_ref is not None and operation.startswith("delete"):
        from .rewrite import shift_name_value

        axis_ = "rows" if "rows" in operation else "cols"
        for sheet in wb.worksheets:
            armed_charts = (led_ref.object_snapshots.get(sheet) or {}).get(
                "chart", {})
            for i, chart in enumerate(getattr(sheet, "_charts", []) or []):
                if i in armed_charts:
                    continue
                for f in _chart_source_refs(chart):
                    new_f, chg = shift_name_value(f, ws.title, axis_,
                                                  index, amount, True)
                    if chg and "#REF" in new_f and "#REF" not in f:
                        blockers.append(
                            "an in-session chart charts {0!r}, which this "
                            "delete removes".format(f))
    part_payload = _sheet_payload(wb, lookup_title)
    if part_payload is None:
        blockers.append("the sheet's package part could not be located")
        return blockers
    if b"extLst" in part_payload:
        blockers.append(
            "the sheet carries extension content (extLst: sparklines, x14 "
            "rules, ...) whose cell ranges live in unmodeled bytes")
    if b"t=\"array\"" in part_payload or b"t='array'" in part_payload:
        blockers.append("the sheet carries array formulas (ref rewriting "
                        "for spill ranges is not supported in v0)")
    if b"t=\"dataTable\"" in part_payload \
            or b"t='dataTable'" in part_payload:
        blockers.append("the sheet carries what-if data tables; their "
                        "ref/r1/r2 inputs live in unmodeled bytes and "
                        "would silently mis-shift")
    if any(cell._comment is not None for cell in ws._cells.values()):
        blockers.append("the sheet carries comments; their anchors live in "
                        "comment/VML parts the shift cannot rewrite")
    if b"<legacyDrawing" in part_payload:
        blockers.append("the sheet references a legacy (VML) drawing")
    charts = _charts_referencing(wb, lookup_title)
    if charts:
        # chart series references and drawing anchors are
        # patchable in place — dry-run the planner; only its blockers
        # (extension machinery, would-be #REF!) refuse now
        from .chartpatch import plan_chart_updates

        overrides = {}
        if led_ref is not None:
            for prior_operation, prior_index, prior_amount in \
                    led_ref.shifts.get(ws, ()):
                prior_plans, prior_blockers = plan_chart_updates(
                    wb, lookup_title, prior_operation, prior_index,
                    prior_amount, overrides=overrides)
                blockers.extend(prior_blockers)
                overrides.update(prior_plans)
        _plans, chart_blockers = plan_chart_updates(
            wb, lookup_title, operation, index, amount,
            overrides=overrides)
        blockers.extend(chart_blockers)
    pivots = _pivots_referencing(wb, lookup_title)
    if pivots:
        blockers.append(
            "preserved pivot part(s) {0} reference this sheet".format(
                ", ".join(sorted(pivots))))
    if ws.row_breaks or ws.col_breaks:
        blockers.append("the sheet carries manual page breaks, which "
                        "anchor to row/column numbers (not rewritten in v0)")
    for other in wb.worksheets:
        for (_r, _c), cell in other._cells.items():
            if cell.data_type == "f" and isinstance(cell._value, str) \
                    and "[" in cell._value and ws.title in cell._value:
                blockers.append(
                    "formula {0}!{1} uses a structured/external reference "
                    "mentioning this sheet".format(other.title,
                                                   cell.coordinate))
                break
    return blockers


def _three_d_formula_references_sheet(wb, formula, target_title):
    """Whether a 3-D operand's sheet interval contains ``target_title``."""
    from openpyxl.formula import Tokenizer
    from .rewrite import _RENAME_PREFIX_RE

    try:
        tokens = Tokenizer(formula).items
    except Exception:
        return False
    order = {sheet.title.casefold(): index
             for index, sheet in enumerate(wb.worksheets)}
    target = order.get(target_title.casefold())
    if target is None:
        return False
    for token in tokens:
        if token.type != "OPERAND" or token.subtype != "RANGE" \
                or "[" in token.value:
            continue
        match = _RENAME_PREFIX_RE.match(token.value)
        if match is None:
            continue
        sheet_text = (match.group(1).replace("''", "'")
                      if match.group(1) else match.group(2))
        if ":" not in sheet_text:
            continue
        first, last = sheet_text.split(":", 1)
        first_index = order.get(first.casefold())
        last_index = order.get(last.casefold())
        if first_index is None or last_index is None:
            if target_title.casefold() in {
                    first.casefold(), last.casefold()}:
                return True
            continue
        if min(first_index, last_index) <= target <= max(
                first_index, last_index):
            return True
    return False


def _charts_with_numeric_formula_entities(wb, sheet_title=None):
    """Find relevant chart formulas that use numeric XML entities."""
    source = getattr(wb, "_paper_source", None)
    if not source:
        return []
    from html import unescape

    from .chartpatch import CHART_NS, _walk_leaf_texts
    from .rewrite import rename_sheet_in_formula

    hits = []
    with zipfile.ZipFile(io.BytesIO(source)) as z:
        for name in z.namelist():
            if not (name.startswith("xl/charts/") and name.endswith(".xml")
                    and "/_rels/" not in name):
                continue
            payload = z.read(name)
            try:
                formulas = [
                    payload[start:end]
                    for ns, local, _parent, start, end, _path
                    in _walk_leaf_texts(payload)
                    if ns == CHART_NS and local == b"f"
                ]
                if not any(b"&#" in formula for formula in formulas):
                    continue
                if sheet_title is None:
                    hits.append(name)
                    continue
                for formula in formulas:
                    decoded = unescape(formula.decode("utf-8"))
                    _rewritten, changed = rename_sheet_in_formula(
                        "=" + decoded, sheet_title,
                        sheet_title + "__paper_probe")
                    if changed:
                        hits.append(name)
                        break
            except Exception:
                if b"&#" in payload:
                    hits.append(name)
    return sorted(hits)


def _pivots_referencing(wb, sheet_title):
    source = getattr(wb, "_paper_source", None)
    if not source:
        return []
    needles = _title_needles(sheet_title)
    hits = []
    with zipfile.ZipFile(io.BytesIO(source)) as z:
        for name in z.namelist():
            if name.startswith(("xl/pivotTables/", "xl/pivotCache/")) \
                    and name.endswith(".xml"):
                payload = z.read(name).lower()
                if any(needle in payload for needle in needles):
                    hits.append(name)
    return hits


def _sheet_payload(wb, title):
    source = getattr(wb, "_paper_source", None)
    if not source:
        return None
    import zipfile as _zf

    from .saver import _package_info

    with _zf.ZipFile(io.BytesIO(source)) as z:
        _wb_part, mapping = _package_info(z)
        part = mapping.get(title)
        if part is None:
            return None
        return z.read(part)


def _chart_source_ref_objects(chart):
    """The live reference OBJECTS of a model chart's data sources (each
    carries a string ``.f``)."""
    from .references import chart_source_ref_objects

    yield from chart_source_ref_objects(chart)


def _chart_source_refs(chart):
    for ref in _chart_source_ref_objects(chart):
        yield ref.f


def _shift_added_chart_refs(chart, target_title, axis, index, amount,
                            is_delete, shift_name_value):
    """Rewrite one model chart's data-source references for a shift on
    ``target_title`` (deletes that would strand a chart were already
    blocked pre-move by shift_blockers)."""
    for ref in _chart_source_ref_objects(chart):
        new_f, changed = shift_name_value(
            ref.f, target_title, axis, index, amount, is_delete)
        if not changed:
            continue
        if "#REF" in new_f and "#REF" not in ref.f:
            raise UnsupportedStructureError(
                "internal: a delete stranding an in-session chart "
                "({0!r}) escaped the pre-move blocker; the model may be "
                "partially shifted — do not save.".format(ref.f))
        ref.f = new_f


def apply_model_shift(ws, operation, index, amount, reference_rewrites=None):
    """All the reference updates stock openpyxl skips, applied to the MODEL
    after the cells moved (Excel insert/delete semantics via rewrite.py).
    Also rebases the positional arm snapshots so pure moves are not
    mis-detected as user changes."""
    from .ledger import _armed_ledger_for_ws
    from .rewrite import row_mapping, shift_cell_range

    wb = ws.parent
    led = _armed_ledger_for_ws(ws)
    axis = "rows" if "rows" in operation else "cols"
    is_delete = operation.startswith("delete")
    mapper = row_mapping(operation, index, amount)

    # 0. rebase positional ledger/arm state FIRST: pre-shift dirty marks and
    # snapshots move with the cells; everything the fixups below record is
    # already in post-shift coordinates and must NOT be remapped again
    if led is not None:
        _rebase_snapshots(led, ws, mapper, axis)

    # 1. every modeled formula-like surface, planned before cell mutation.
    from .references import apply_rewrites, plan_shift

    rewrites = (plan_shift(wb, ws, operation, index, amount)
                if reference_rewrites is None else reference_rewrites)
    _saved_lint = getattr(wb, "formula_lint", "warn")
    wb.formula_lint = "off"
    try:
        apply_rewrites(rewrites)
    finally:
        wb.formula_lint = _saved_lint

    # 3. sheet-internal regions (fully modeled; the splice re-renders them)
    from collections import OrderedDict
    from openpyxl.formatting.formatting import ConditionalFormatting

    shifted_merges = _shifted_multi_range(
        ws.merged_cells, axis, index, amount, is_delete, shift_cell_range)
    _rebuild_merged_cells(ws, shifted_merges)

    # ConditionalFormatting objects are dict keys and CellRanges are set
    # members. Mutating either in place corrupts their hash collections.
    rebuilt_cf = OrderedDict()
    for cf, rules in ws.conditional_formatting._cf_rules.items():
        sqref = _shifted_multi_range(
            cf.sqref, axis, index, amount, is_delete, shift_cell_range)
        if not sqref:
            continue
        new_cf = ConditionalFormatting(sqref=sqref, pivot=cf.pivot)
        rebuilt_cf.setdefault(new_cf, []).extend(rules)
    ws.conditional_formatting._cf_rules = rebuilt_cf

    if ws.data_validations:
        for dv in list(ws.data_validations.dataValidation):
            dv.sqref = _shifted_multi_range(
                dv.sqref, axis, index, amount, is_delete, shift_cell_range)
        ws.data_validations.dataValidation = [
            dv for dv in ws.data_validations.dataValidation if dv.sqref]
    if ws.auto_filter and ws.auto_filter.ref:
        from openpyxl.worksheet.cell_range import CellRange

        original_filter = CellRange(ws.auto_filter.ref)
        mapped_filter_columns = None
        if axis == "cols":
            from .rewrite import row_mapping

            map_column = row_mapping(operation, index, amount)
            mapped_filter_columns = [
                (column, map_column(original_filter.min_col + column.colId))
                for column in ws.auto_filter.filterColumn
            ]
        cr = CellRange(ws.auto_filter.ref)
        filter_state = shift_cell_range(cr, axis, index, amount, is_delete)
        if filter_state == "deleted":
            ws.auto_filter.ref = None
            ws.auto_filter.filterColumn = []
            ws.auto_filter.sortState = None
        elif filter_state == "changed":
            ws.auto_filter.ref = cr.coord
        if mapped_filter_columns is not None and ws.auto_filter.ref:
            shifted_filter = CellRange(ws.auto_filter.ref)
            ws.auto_filter.filterColumn = [
                column for column, absolute in mapped_filter_columns
                if absolute is not None
                and shifted_filter.min_col <= absolute <=
                shifted_filter.max_col
            ]
            for column, absolute in mapped_filter_columns:
                if column in ws.auto_filter.filterColumn:
                    column.colId = absolute - shifted_filter.min_col
        sort_state = ws.auto_filter.sortState
        if sort_state is not None:
            if sort_state.ref:
                sort_range = CellRange(sort_state.ref)
                if shift_cell_range(
                        sort_range, axis, index, amount, is_delete) == \
                        "deleted":
                    sort_state.ref = None
                else:
                    sort_state.ref = sort_range.coord
            for condition in sort_state.sortCondition:
                if not condition.ref:
                    continue
                condition_range = CellRange(condition.ref)
                if shift_cell_range(
                        condition_range, axis, index, amount, is_delete) == \
                        "deleted":
                    condition.ref = None
                else:
                    condition.ref = condition_range.coord

    scenarios = getattr(ws, "scenarios", None)
    if scenarios is not None:
        if scenarios.sqref:
            scenarios.sqref = _shifted_multi_range(
                scenarios.sqref, axis, index, amount, is_delete,
                shift_cell_range)
        from openpyxl.worksheet.cell_range import CellRange

        for scenario in scenarios.scenario:
            for input_cell in scenario.inputCells:
                if not input_cell.r:
                    continue
                input_range = CellRange(input_cell.r)
                state = shift_cell_range(
                    input_range, axis, index, amount, is_delete)
                if state == "deleted":
                    raise UnsupportedStructureError(
                        "the structural edit would delete scenario input "
                        "{0}. Nothing was changed.".format(input_cell.r))
                input_cell.r = input_range.coord

    if axis == "rows":
        # 4. row display attributes move with their rows
        new_dims = {}
        for r, dim in list(ws.row_dimensions.items()):
            new_row = mapper(r)
            if new_row is None:
                continue
            dim.index = new_row
            new_dims[new_row] = dim
        ws.row_dimensions.clear()
        ws.row_dimensions.update(new_dims)
    else:
        from openpyxl.utils import column_index_from_string, get_column_letter

        new_dims = {}
        dims = sorted(ws.column_dimensions.values(), key=lambda dim:
                      dim.min or column_index_from_string(dim.index))
        for dim in dims:
            start = dim.min or column_index_from_string(dim.index)
            end = dim.max or start
            span = _shift_span_for_dimension(
                start, end, index, amount, is_delete)
            if span is None:
                continue
            dim.min, dim.max = span
            dim.index = get_column_letter(span[0])
            new_dims[dim.index] = dim
        ws.column_dimensions.clear()
        ws.column_dimensions.update(new_dims)

    # 5. hyperlink anchors track their cells (both axes)
    for (_r, _c), cell in ws._cells.items():
        link = getattr(cell, "_hyperlink", None)
        if link is not None:
            link.ref = cell.coordinate

    if led is not None:
        if ws not in led.added_sheets:
            led.shifts.setdefault(ws, []).append((operation, index, amount))
        led.formulas_changed = True


def _shifted_multi_range(multi_range, axis, index, amount, is_delete,
                         shift_cell_range):
    """Return a fresh MultiCellRange; never mutate hashed CellRanges."""
    from openpyxl.worksheet.cell_range import CellRange, MultiCellRange

    shifted = []
    for original in getattr(multi_range, "ranges", []):
        # MergedCellRange.__copy__ constructs against the live worksheet and
        # can recreate a stale anchor as a side effect. A plain CellRange is
        # the side-effect-free value object needed for coordinate planning.
        rng = CellRange(str(original))
        if shift_cell_range(rng, axis, index, amount, is_delete) != "deleted":
            shifted.append(rng)
    return MultiCellRange(shifted)


def _rebuild_merged_cells(ws, shifted_ranges):
    """Recreate merge anchors and placeholders from shifted coordinates."""
    from copy import copy

    from openpyxl.cell import Cell, MergedCell
    from openpyxl.worksheet.cell_range import MultiCellRange
    from openpyxl.worksheet.merge import MergedCellRange

    placeholder_styles = {}
    for key, cell in list(ws._cells.items()):
        if not isinstance(cell, MergedCell):
            continue
        placeholder_styles[key] = copy(cell._style)
        del ws._cells[key]

    rebuilt = []
    ranges = sorted(
        shifted_ranges.ranges,
        key=lambda rng: (rng.min_row, rng.min_col, rng.max_row, rng.max_col),
    )
    for rng in ranges:
        anchor_key = (rng.min_row, rng.min_col)
        if anchor_key not in ws._cells:
            anchor = Cell(ws, row=rng.min_row, column=rng.min_col)
            if anchor_key in placeholder_styles:
                anchor._style = placeholder_styles[anchor_key]
            ws._cells[anchor_key] = anchor
        merged = MergedCellRange(ws, rng.coord)
        rebuilt.append(merged)
        ws._clean_merge_range(merged)
    ws.merged_cells = MultiCellRange(rebuilt)


def _shift_span_for_dimension(start, end, index, amount, is_delete):
    from .rewrite import _shift_span

    return _shift_span(start, end, index, amount, is_delete)


def validate_model_shift(ws, operation, index, amount):
    """Dry-run every modeled formula reference before cells are moved."""
    wb = ws.parent
    from .references import plan_shift

    rewrites = plan_shift(wb, ws, operation, index, amount)
    strict_protection = getattr(wb, "strict_protection", False)
    if strict_protection:
        from .ledger import check_protection
    if strict_protection:
        for surface, _rewritten in rewrites:
            if surface.cell:
                check_protection(surface.owner)
    return rewrites


def _rebase_snapshots(led, ws, mapper, axis):
    """Every POSITIONAL piece of arm/ledger state (dirty cells, row attrs,
    hyperlink anchors, comments) must follow a pure move, or the save would
    lose pre-shift edits and mis-read the move as user changes."""
    if axis == "rows":
        def map_key(row, col):
            new_row = mapper(row)
            return None if new_row is None else (new_row, col)
    else:
        def map_key(row, col):
            new_col = mapper(col)   # the caller passes the column mapper
            return None if new_col is None else (row, new_col)

    dirty = led.cells.get(ws)
    if dirty:
        led.cells[ws] = {mapped for (r, c) in dirty
                         for mapped in (map_key(r, c),) if mapped is not None}
    overwrites = led.value_overwrites.get(ws)
    if overwrites:
        led.value_overwrites[ws] = {
            mapped for (r, c) in overwrites
            for mapped in (map_key(r, c),) if mapped is not None}
    links = led.region_snapshots.get(ws, {}).get("hyperlinks")
    if links:
        led.region_snapshots[ws]["hyperlinks"] = {
            mapped: sig for (r, c), sig in links.items()
            for mapped in (map_key(r, c),) if mapped is not None}
    comments = led.comment_snapshots.get(ws)
    if comments:
        led.comment_snapshots[ws] = {
            mapped: sig for (r, c), sig in comments.items()
            for mapped in (map_key(r, c),) if mapped is not None}
    if axis == "rows":
        rows = led.row_attr_snapshots.get(ws)
        if rows:
            led.row_attr_snapshots[ws] = {
                mapper(r): attrs for r, attrs in rows.items()
                if mapper(r) is not None}


def apply_shift_to_bytes(original, operation, index, amount):
    """The byte-level renumber pre-transform: deleted rows are cut; shifted
    rows get their r attributes (row and cells) rewritten; every other byte
    — cell contents, unmodeled attributes, spans — is copied verbatim. The
    result becomes the baseline the standard splice runs against."""
    from .rewrite import row_mapping
    from .xmlscan import scan_sheet

    axis = "rows" if "rows" in operation else "cols"
    scan = scan_sheet(original)
    edits = []

    if axis == "rows":
        mapper = row_mapping(operation, index, amount)
        for row_index, row_span in scan.rows.items():
            new_row = mapper(row_index)
            if new_row is None:
                edits.append((row_span.start, row_span.end, b""))
                continue
            if new_row == row_index:
                continue
            head_end = (row_span.content_start
                        if not row_span.self_closing else row_span.end)
            head = original[row_span.start:head_end]
            new_head = re.sub(
                br'(<row[^>]*?\sr=")%d(")' % row_index,
                br"\g<1>%d\g<2>" % new_row, head, 1)
            edits.append((row_span.start, head_end, new_head))
            for col, cell_span in row_span.cells.items():
                cell_head_end = original.index(b">", cell_span.start) + 1
                cell_head = original[cell_span.start:cell_head_end]
                old_ref = cell_span.attrs["r"].encode("ascii")
                letters = old_ref.rstrip(b"0123456789")
                new_ref = letters + str(new_row).encode("ascii")
                new_cell_head = cell_head.replace(
                    b'r="%s"' % old_ref, b'r="%s"' % new_ref, 1)
                edits.append((cell_span.start, cell_head_end, new_cell_head))
    else:
        from openpyxl.utils import get_column_letter

        from .rewrite import _shift_span
        is_delete = operation.startswith("delete")
        for row_index, row_span in scan.rows.items():
            for col, cell_span in sorted(row_span.cells.items()):
                span = _shift_span(col, col, index, amount, is_delete)
                if span is None:
                    edits.append((cell_span.start, cell_span.end, b""))
                    continue
                if span[0] == col:
                    continue
                cell_head_end = original.index(b">", cell_span.start) + 1
                cell_head = original[cell_span.start:cell_head_end]
                old_ref = cell_span.attrs["r"].encode("ascii")
                digits = old_ref.lstrip(
                    b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz")
                new_ref = get_column_letter(span[0]).encode("ascii") + digits
                edits.append((cell_span.start, cell_head_end,
                              cell_head.replace(b'r="%s"' % old_ref,
                                                b'r="%s"' % new_ref, 1)))

    if not edits:
        return original
    from .crosspart import apply_edits

    return apply_edits(original, edits)
