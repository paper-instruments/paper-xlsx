"""Modeled reference surfaces shared by structural edits and renames."""

import re


_DYNAMIC_REFERENCE = re.compile(
    r"(?i)(?<![A-Z0-9_.])(?:INDIRECT|EVALUATE|OFFSET|INDEX)\s*\(")
_DYNAMIC_FUNCTIONS = {"INDIRECT(", "EVALUATE(", "OFFSET(", "INDEX("}


def _has_dynamic_reference(value):
    """Return whether ``value`` calls a dynamic-reference function."""
    from openpyxl.formula import Tokenizer

    probe = value[1:] if value.startswith("#") else value
    formula = probe if probe.startswith("=") else "=" + probe
    try:
        tokens = Tokenizer(formula).items
    except Exception:
        return bool(_DYNAMIC_REFERENCE.search(probe))
    return any(
        token.type == "FUNC" and token.subtype == "OPEN"
        and token.value.upper() in _DYNAMIC_FUNCTIONS
        for token in tokens)


class FormulaSurface:
    """A formula-like string with enough context to rewrite it safely."""

    def __init__(self, sheet, owner, attribute, label, *, index=None,
                 name=False, prefix=False, cell=False, range_ref=False):
        self.sheet = sheet
        self.owner = owner
        self.attribute = attribute
        self.label = label
        self.index = index
        self.name = name
        self.prefix = prefix
        self.cell = cell
        self.range_ref = range_ref

    @property
    def value(self):
        value = getattr(self.owner, self.attribute)
        return value[self.index] if self.index is not None else value

    def replace(self, value):
        if self.cell:
            self.owner.value = value
        elif self.index is not None:
            values = getattr(self.owner, self.attribute)
            values[self.index] = value
        else:
            setattr(self.owner, self.attribute, value)


def chart_source_ref_objects(chart):
    """Yield every modeled chart reference, including titles and axes."""
    from openpyxl.chart.data_source import MultiLevelStrRef, NumRef, StrRef
    from openpyxl.descriptors.serialisable import Serialisable

    wanted = (NumRef, StrRef, MultiLevelStrRef)
    stack = [chart]
    seen = set()
    while stack:
        value = stack.pop()
        marker = id(value)
        if marker in seen:
            continue
        seen.add(marker)
        if isinstance(value, wanted):
            if isinstance(getattr(value, "f", None), str):
                yield value
            continue
        if isinstance(value, dict):
            stack.extend(value.values())
        elif isinstance(value, (list, tuple, set)):
            stack.extend(value)
        elif isinstance(value, Serialisable):
            stack.extend(value.__dict__.values())


def formula_surfaces(wb):
    """Yield all formula-like references represented by the live model."""
    ledger = getattr(wb, "_paper_ledger", None)
    for sheet in wb.worksheets:
        for attribute in ("print_area", "print_title_rows",
                          "print_title_cols"):
            if isinstance(getattr(sheet, attribute), str):
                yield FormulaSurface(
                    sheet, sheet, attribute,
                    "{0} on {1!r}".format(attribute, sheet.title))
        for cell in sheet._cells.values():
            if cell.data_type == "f" and isinstance(cell._value, str):
                yield FormulaSurface(
                    sheet, cell, "_value",
                    "formula at {0}!{1}".format(sheet.title, cell.coordinate),
                    cell=True)
            hyperlink = getattr(cell, "_hyperlink", None)
            if hyperlink is not None and isinstance(hyperlink.location, str):
                yield FormulaSurface(
                    sheet, hyperlink, "location",
                    "hyperlink at {0}!{1}".format(
                        sheet.title, cell.coordinate), prefix=True)

        for conditional in sheet.conditional_formatting:
            for rule in conditional.rules:
                for index, formula in enumerate(rule.formula or []):
                    if isinstance(formula, str):
                        yield FormulaSurface(
                            sheet, rule, "formula",
                            "conditional formatting on {0}!{1}".format(
                                sheet.title, conditional.sqref), index=index)
                for value_object in _conditional_formula_values(rule):
                    yield FormulaSurface(
                        sheet, value_object, "val",
                        "conditional formatting value on {0}!{1}".format(
                            sheet.title, conditional.sqref))

        if sheet.data_validations:
            for validation in sheet.data_validations.dataValidation:
                for attribute in ("formula1", "formula2"):
                    if isinstance(getattr(validation, attribute), str):
                        yield FormulaSurface(
                            sheet, validation, attribute,
                            "data validation on {0}!{1}".format(
                                sheet.title, validation.sqref))

        for table in sheet.tables.values():
            yield FormulaSurface(
                sheet, table, "ref",
                "range of table {0!r}".format(table.displayName),
                range_ref=True)
            table_filter = getattr(table, "autoFilter", None)
            if table_filter is not None and isinstance(table_filter.ref, str):
                yield FormulaSurface(
                    sheet, table_filter, "ref",
                    "filter range of table {0!r}".format(
                        table.displayName), range_ref=True)
            for column in table.tableColumns:
                for attribute in (
                        "calculatedColumnFormula", "totalsRowFormula"):
                    formula = getattr(column, attribute)
                    if formula is not None and isinstance(
                            formula.attr_text, str):
                        yield FormulaSurface(
                            sheet, formula, "attr_text",
                            "{0} in table {1!r}".format(
                                attribute, table.displayName))

        for definition in sheet.defined_names.values():
            if isinstance(definition.attr_text, str):
                yield FormulaSurface(
                    sheet, definition, "attr_text",
                    "sheet-scoped name {0!r}".format(definition.name),
                    name=True)

        armed_charts = ((ledger.object_snapshots.get(sheet) or {}).get(
            "chart", {}) if ledger is not None else {})
        for index, chart in enumerate(getattr(sheet, "_charts", ()) or ()):
            if index in armed_charts:
                continue
            for reference in chart_source_ref_objects(chart):
                yield FormulaSurface(
                    sheet, reference, "f",
                    "chart reference on {0!r}".format(sheet.title),
                    name=True)

    for definition in wb.defined_names.values():
        if isinstance(definition.attr_text, str):
            yield FormulaSurface(
                None, definition, "attr_text",
                "workbook name {0!r}".format(definition.name), name=True)


def _conditional_formula_values(rule):
    for container_name in ("colorScale", "dataBar", "iconSet"):
        container = getattr(rule, container_name, None)
        for value in getattr(container, "cfvo", ()) or ():
            if value.type == "formula" and isinstance(value.val, str):
                yield value


def plan_shift(wb, target_sheet, operation, index, amount):
    """Validate and return every modeled formula rewrite for one shift."""
    from .rewrite import shift_formula_fragment, shift_name_value

    axis = "rows" if "rows" in operation else "cols"
    is_delete = operation.startswith("delete")
    _validate_table_shift(
        target_sheet, operation, axis, index, amount, is_delete)
    rewrites = []
    for surface in formula_surfaces(wb):
        value = surface.value
        from .structural import _three_d_formula_references_sheet

        probe = value if value.startswith("=") else "=" + value
        if _three_d_formula_references_sheet(
                wb, probe, target_sheet.title):
            from openpyxl.errors import UnsupportedStructureError

            raise UnsupportedStructureError(
                "{0}() cannot rewrite the 3-D formula reference in {1}. "
                "Nothing was changed.".format(operation, surface.label),
                kind="three-dimensional-structural-reference",
                anchor=surface.label)
        if _has_dynamic_reference(value):
            from openpyxl.errors import UnsupportedStructureError

            raise UnsupportedStructureError(
                "{0}() cannot prove the target of a dynamic reference in "
                "{1}. Nothing was changed.".format(
                    operation, surface.label),
                kind="dynamic-structural-reference",
                anchor=surface.label)
        prefix = "#" if surface.prefix and value.startswith("#") else ""
        core = value[1:] if prefix else value
        if surface.name and surface.sheet is None:
            if _has_unqualified_cell_reference(core):
                from openpyxl.errors import UnsupportedStructureError

                raise UnsupportedStructureError(
                    "{0}() cannot assign a sheet context to the "
                    "unqualified reference in {1}. Nothing was changed."
                    .format(operation, surface.label),
                    kind="ambiguous-structural-reference",
                    anchor=surface.label)
            rewritten, changed = shift_name_value(
                core, target_sheet.title, axis, index, amount, is_delete)
        else:
            rewritten, changed = shift_formula_fragment(
                core,
                surface.sheet.title if surface.sheet is not None else None,
                target_sheet.title, axis, index, amount, is_delete)
        if changed:
            if surface.range_ref and "#REF" in rewritten \
                    and "#REF" not in core:
                from openpyxl.errors import UnsupportedStructureError

                raise UnsupportedStructureError(
                    "{0}() would remove {1}. Nothing was changed.".format(
                        operation, surface.label),
                    kind="structural-reference-deleted",
                    anchor=surface.label)
            rewrites.append((surface, prefix + rewritten))
    return rewrites


def _validate_table_shift(sheet, operation, axis, index, amount, is_delete):
    """Refuse shifts that require changing table column/header metadata."""
    from openpyxl.errors import UnsupportedStructureError
    from openpyxl.utils.cell import range_boundaries

    delete_end = index + amount - 1
    for table in sheet.tables.values():
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        unsafe = False
        if axis == "cols":
            if is_delete:
                unsafe = not (delete_end < min_col or index > max_col)
            else:
                unsafe = min_col < index <= max_col
        elif is_delete:
            header_hit = index <= min_row <= delete_end
            totals = bool(
                getattr(table, "totalsRowShown", False)
                or getattr(table, "totalsRowCount", 0))
            totals_hit = totals and index <= max_row <= delete_end
            overlap = max(
                0, min(max_row, delete_end) - max(min_row, index) + 1)
            minimum_rows = 2 + int(totals)
            too_short = overlap and max_row - min_row + 1 - overlap < \
                minimum_rows
            unsafe = header_hit or totals_hit or too_short
        if unsafe:
            raise UnsupportedStructureError(
                "{0}() intersects table {1!r} where table column, header, "
                "or totals metadata would also need changing. Nothing was "
                "changed.".format(operation, table.displayName),
                kind="table-structure-edit-unsupported",
                anchor=table.displayName)


def _has_unqualified_cell_reference(value):
    """Whether a workbook-scoped name contains a context-free A1 ref."""
    from openpyxl.formula import Tokenizer

    formula = value if value.startswith("=") else "=" + value
    try:
        tokens = Tokenizer(formula).items
    except Exception:
        return True
    for token in tokens:
        if token.type != "OPERAND" or token.subtype != "RANGE" \
                or "!" in token.value:
            continue
        candidate = token.value.replace("$", "")
        if re.match(r"^[A-Z]{1,3}[0-9]+(?::[A-Z]{1,3}[0-9]+)?$",
                    candidate, re.I):
            return True
    return False


def apply_rewrites(rewrites):
    for surface, value in rewrites:
        surface.replace(value)


def plan_rename(wb, old_title, new_title):
    """Return all modeled reference rewrites for a worksheet rename."""
    from .rewrite import rename_sheet_in_formula_fragment

    rewrites = []
    for surface in formula_surfaces(wb):
        value = surface.value
        if _has_dynamic_reference(value):
            from openpyxl.errors import UnsupportedStructureError

            raise UnsupportedStructureError(
                "renaming sheet {0!r} cannot prove the target of a dynamic "
                "reference in {1}. Nothing was changed.".format(
                    old_title, surface.label),
                kind="dynamic-structural-reference",
                anchor=surface.label)
        prefix = "#" if surface.prefix and value.startswith("#") else ""
        core = value[1:] if prefix else value
        rewritten, changed = rename_sheet_in_formula_fragment(
            core, old_title, new_title)
        if changed:
            rewrites.append((surface, prefix + rewritten))
    return rewrites
