# paper-xlsx: style verbs (PLAN-v0.1 Batch 7, PR-1 §6)

"""Format work an agent actually asks for, preserve-safe by construction
(cell style edits ride the splice; styles.xml stays append-only via the
D2 translator at save)."""

from copy import copy

from openpyxl.utils.cell import range_boundaries


def copy_format(ws, src_cell, dst_range):
    """Copy one cell's complete format onto every cell of a range —
    "make these look like B2". ``src_cell``/``dst_range`` are A1 strings
    (a single-cell dst is fine). Returns the number of cells formatted."""
    src = ws[src_cell.replace("$", "")] if isinstance(src_cell, str) \
        else src_cell
    style_array = getattr(src, "_style", None)
    min_col, min_row, max_col, max_row = range_boundaries(
        str(dst_range).replace("$", ""))
    count = 0
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            if (row, col) == (src.row, src.column):
                continue
            cell = ws.cell(row=row, column=col)
            # one public-setter write keeps every ledger chokepoint honest
            # (a bare _style assignment would bypass the dirty mark); the
            # remaining components share the SAME interned array, so this
            # is the style-array reuse the contract asks for
            cell._style = copy(style_array) if style_array is not None \
                else None
            from openpyxl.preserve.ledger import mark_styleable_dirty

            mark_styleable_dirty(cell)
            count += 1
    return count


# a small number-format library profiles can reference by name
NUMBER_FORMATS = {
    "comma": "#,##0.00",
    "comma0": "#,##0",
    "percent": "0.0%",
    "percent0": "0%",
    "currency": '"$"#,##0.00',
    "date": "yyyy-mm-dd",
    "general": "General",
}


def apply_profile(ws, profile):
    """Apply a formatting PROFILE (data, not code) to a sheet's cells by
    their model-map role. ``profile`` maps role names ("inputs",
    "calculations", "outputs", "constants") to format specs:

        {"inputs": {"fill": "FFF2CC", "number_format": "comma",
                    "bold": False, "font_color": "1F4E79",
                    "locked": False}, ...}

    number_format accepts a library name (NUMBER_FORMATS) or a literal
    format string. Returns {role: cells_formatted}. Measurements drive
    it (the model map); nothing here decides what a cell IS."""
    from openpyxl.styles import Font, PatternFill, Protection

    wb = ws.parent
    from .modelmap import build_model_map

    mm = build_model_map(wb)
    roles = mm.sheets.get(ws.title, {})
    counts = {}
    for role, spec in profile.items():
        if not isinstance(spec, dict):
            raise TypeError(
                "profile[{0!r}] must be a dict of format settings".format(
                    role))
        applied = 0
        for address in roles.get(role, []):
            cell = ws[address]
            if "number_format" in spec:
                fmt = spec["number_format"]
                cell.number_format = NUMBER_FORMATS.get(fmt, fmt)
            if "fill" in spec and spec["fill"]:
                cell.fill = PatternFill(start_color=spec["fill"],
                                        end_color=spec["fill"],
                                        fill_type="solid")
            if "bold" in spec or "font_color" in spec or "italic" in spec:
                font = copy(cell.font)
                if "bold" in spec:
                    font.bold = bool(spec["bold"])
                if "italic" in spec:
                    font.italic = bool(spec["italic"])
                if spec.get("font_color"):
                    from openpyxl.styles import Color

                    font.color = Color(rgb=spec["font_color"])
                cell.font = font
            if "locked" in spec:
                cell.protection = Protection(locked=bool(spec["locked"]))
            applied += 1
        counts[role] = applied
    return counts
