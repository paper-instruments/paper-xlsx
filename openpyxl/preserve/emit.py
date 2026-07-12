# paper-xlsx: model-to-bytes emission for the splice

"""Serialize single cells and rows from the model as bare fragments ready to
splice into a default-namespace worksheet stream.

This is a THIN VARIANT of upstream's cell writer
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


_XML_WHITESPACE = b" \t\r\n"
_ATTRIBUTE_NAME_END = b" \t\r\n=/>"


class _AttributeSpan:
    __slots__ = ("name", "start", "end", "value_start", "value_end")

    def __init__(self, name, start, end, value_start, value_end):
        self.name = name
        self.start = start
        self.end = end
        self.value_start = value_start
        self.value_end = value_end


def _start_tag_end(fragment):
    quote = None
    for index, byte in enumerate(fragment):
        if quote is not None:
            if byte == quote:
                quote = None
            continue
        if byte in (34, 39):
            quote = byte
        elif byte == 62:
            return index
    raise ValueError("fragment has no complete start tag")


def _start_tag_attributes(fragment):
    """Scan one start tag without confusing quoted text for markup."""
    tag_end = _start_tag_end(fragment)
    if not fragment.startswith(b"<") or fragment.startswith(b"</"):
        raise ValueError("fragment does not begin with a start tag")
    index = 1
    while index < tag_end and fragment[index] not in _XML_WHITESPACE + b"/":
        index += 1
    spans = []
    while index < tag_end:
        attribute_start = index
        while index < tag_end and fragment[index] in _XML_WHITESPACE:
            index += 1
        if index >= tag_end:
            break
        if fragment[index] == 47:
            break
        name_start = index
        while index < tag_end and fragment[index] not in _ATTRIBUTE_NAME_END:
            index += 1
        name = fragment[name_start:index]
        while index < tag_end and fragment[index] in _XML_WHITESPACE:
            index += 1
        if not name or index >= tag_end or fragment[index] != 61:
            raise ValueError("malformed start-tag attribute")
        index += 1
        while index < tag_end and fragment[index] in _XML_WHITESPACE:
            index += 1
        if index >= tag_end or fragment[index] not in (34, 39):
            raise ValueError("attribute value is not quoted")
        quote = fragment[index]
        index += 1
        value_start = index
        while index < tag_end and fragment[index] != quote:
            index += 1
        if index >= tag_end:
            raise ValueError("unterminated attribute value")
        value_end = index
        index += 1
        spans.append(_AttributeSpan(
            name, attribute_start, index, value_start, value_end))
    return tag_end, fragment[tag_end - 1:tag_end] == b"/", spans


def patch_start_tag_attribute(fragment, name, value):
    tag_end, self_closing, spans = _start_tag_attributes(fragment)
    matches = [span for span in spans if span.name == name]
    if len(matches) > 1:
        raise ValueError("duplicate attribute {0!r}".format(name))
    if matches:
        span = matches[0]
        if value is None:
            return fragment[:span.start] + fragment[span.end:]
        return fragment[:span.value_start] + value + fragment[span.value_end:]
    if value is None:
        return fragment
    insert_at = tag_end - 1 if self_closing else tag_end
    return (fragment[:insert_at] + b" " + name + b'="' + value + b'"'
            + fragment[insert_at:])


def carry_start_tag_attributes(new_fragment, original_fragment, skip=()):
    _old_end, _old_closing, old_spans = _start_tag_attributes(
        original_fragment)
    new_end, new_closing, new_spans = _start_tag_attributes(new_fragment)
    present = {span.name for span in new_spans}
    skipped = set(skip)
    carried = [
        original_fragment[span.start:span.end] for span in old_spans
        if span.name not in skipped and span.name not in present]
    if not carried:
        return new_fragment
    insert_at = new_end - 1 if new_closing else new_end
    return (new_fragment[:insert_at] + b"".join(carried)
            + new_fragment[insert_at:])


def patch_cell_style(original_cell_bytes, style_index):
    normalized = int(style_index or 0)
    value = None if normalized == 0 else str(normalized).encode("ascii")
    return patch_start_tag_attribute(original_cell_bytes, b"s", value)


def carry_attributes(new_cell_bytes, original_cell_bytes,
                     drop_metadata=False):
    """Attribute-carry rule: re-attach every original cell attribute
    the replacement does not intentionally rewrite (everything except r, s,
    t). cm/vm rich-value metadata drops ONLY when the cell's VALUE was
    overwritten (the cell stops being a rich value) — style-only edits,
    move re-emissions and dissolution re-emits must carry it."""
    skip = ((b"r", b"s", b"t", b"cm", b"vm") if drop_metadata
            else (b"r", b"s", b"t"))
    return carry_start_tag_attributes(
        new_cell_bytes, original_cell_bytes, skip=skip)


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
