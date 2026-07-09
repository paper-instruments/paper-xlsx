# paper-xlsx: model-to-bytes emission for the splice (CONVENTIONS §3.6; PR-0 D2)

"""Serialize single cells and rows from the model as bare fragments ready to
splice into a default-namespace worksheet stream.

Per PR-0 D2 this is a THIN VARIANT of upstream's cell writer
(openpyxl/cell/_writer.py), not a direct call: the style index is an
explicit parameter — the file's xf numbering, from the StyleTranslator —
never ``cell.style_id`` (which interns into the MODEL's numbering; model and
file numbering drift on non-openpyxl producers, measured), and the
hyperlink-registration side effect is owned by the hyperlinks planner
instead. The value/formula/date/rich-text emission logic mirrors
``_set_attributes``/``etree_write_cell`` line for line.
"""

from datetime import timedelta

from openpyxl.cell.rich_text import CellRichText
from openpyxl.utils.datetime import to_excel, to_ISO8601
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula
from openpyxl.xml.functions import Element, SubElement, tostring, whitespace
from openpyxl.compat import safe_string


def _cell_attributes(cell, style_index):
    """Mirror of cell/_writer._set_attributes with the style index explicit
    and no hyperlink side effect."""
    attrs = {"r": cell.coordinate}
    if style_index:
        attrs["s"] = "{0}".format(style_index)

    if cell.data_type == "s":
        attrs["t"] = "inlineStr"
    elif cell.data_type != "f":
        attrs["t"] = cell.data_type

    value = cell._value

    if cell.data_type == "d":
        if hasattr(value, "tzinfo") and value.tzinfo is not None:
            raise TypeError("Excel does not support timezones in datetimes. "
                            "The tzinfo in the datetime/time object must be "
                            "set to None.")
        if cell.parent.parent.iso_dates and not isinstance(value, timedelta):
            value = to_ISO8601(value)
        else:
            attrs["t"] = "n"
            value = to_excel(value, cell.parent.parent.epoch)

    return value, attrs


def emit_cell(ws, cell, style_index):
    """Bytes of one ``<c>`` element, or ``None`` when the model holds
    nothing worth serializing (mirrors the stock writer's skip rule).

    ``style_index`` is the FILE xf index (StyleTranslator.resolve), or None.
    """
    if cell._value is None and style_index is None:
        return None
    value, attributes = _cell_attributes(cell, style_index)

    el = Element("c", attributes)
    if value is None or value == "":
        return tostring(el)

    if cell.data_type == "f":
        attrib = {}
        if isinstance(value, ArrayFormula):
            attrib = dict(value)
            value = value.text
        elif isinstance(value, DataTableFormula):
            attrib = dict(value)
            value = None
        formula = SubElement(el, "f", attrib)
        if value is not None and not attrib.get("t") == "dataTable":
            formula.text = value[1:]
            value = None

    if cell.data_type == "s":
        if isinstance(value, CellRichText):
            el.append(value.to_tree())
        else:
            inline_string = Element("is")
            text = Element("t")
            text.text = value
            whitespace(text)
            inline_string.append(text)
            el.append(inline_string)
    else:
        cell_content = SubElement(el, "v")
        if value is not None:
            cell_content.text = safe_string(value)

    return tostring(el)


_ENTITIES = (("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"),
             ("&quot;", '"'), ("&apos;", "'"))


def _unescape_value(value):
    """Scanner attribute values keep their raw entity escapes; expand them
    before re-escaping so carried values stay verbatim (PR-0 D6)."""
    for entity, char in _ENTITIES:
        value = value.replace(entity, char)
    return value


def carry_attributes(new_cell_bytes, original_attrs, drop_metadata=False):
    """PR-0 D6 attribute-carry rule: re-attach every original cell attribute
    the replacement does not intentionally rewrite (everything except r, s,
    t). cm/vm rich-value metadata drops ONLY when the cell's VALUE was
    overwritten (the cell stops being a rich value) — style-only edits,
    move re-emissions and dissolution re-emits must carry it
    (PLAN-v0.1 3.4, corrected by the Batch-3 gate)."""
    skip = ("r", "s", "t", "cm", "vm") if drop_metadata else ("r", "s", "t")
    carried = {}
    for k, v in original_attrs.items():
        if k in skip:
            continue
        if "&#" in v:
            from openpyxl.errors import UnsupportedStructureError

            raise UnsupportedStructureError(
                "cannot carry cell attribute {0!r}: it uses numeric "
                "character references the splice cannot round-trip. "
                "Nothing was written.".format(k))
        carried[k] = _unescape_value(v)
    if not carried:
        return new_cell_bytes
    head_end = new_cell_bytes.index(b">")
    self_closing = new_cell_bytes[head_end - 1:head_end] == b"/"
    insert_at = head_end - 1 if self_closing else head_end
    blob = b"".join(
        b' %s="%s"' % (k.encode("latin-1"), _escape_attr(v))
        for k, v in sorted(carried.items()))
    return (new_cell_bytes[:insert_at] + blob + new_cell_bytes[insert_at:])


def _escape_attr(value):
    return (value.encode("utf-8")
            .replace(b"&", b"&amp;")
            .replace(b"<", b"&lt;")
            .replace(b'"', b"&quot;"))


def row_start_tag(index, attrs, self_closing=False):
    """Bytes of a ``<row>`` start tag with the given display attributes."""
    parts = [b'<row r="%d"' % index]
    for key, value in sorted(attrs.items()):
        if key == "r":
            continue
        parts.append(b' %s="%s"' % (key.encode("latin-1"),
                                    _escape_attr(str(value))))
    parts.append(b"/>" if self_closing else b">")
    return b"".join(parts)


def emit_new_row(ws, index, cells_bytes, attrs):
    """Bytes of a whole new ``<row>`` element (sorted cells already emitted)."""
    if not cells_bytes:
        return row_start_tag(index, attrs, self_closing=True)
    return (row_start_tag(index, attrs)
            + b"".join(cells_bytes)
            + b"</row>")
