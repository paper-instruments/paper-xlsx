# paper-xlsx: model-to-bytes emission for the splice (CONVENTIONS §3.6)

"""Serialize single cells and rows from the model, through upstream's own
cell-writing machinery (never string-formatted XML), as bare fragments ready
to splice into a default-namespace worksheet stream.

The upstream ``write_cell`` path has two side effects the splice must own
(PR-0 D2): it appends the cell's hyperlink to ``ws._hyperlinks`` (guarded
here by save/restore) and ``cell.style_id`` interns new StyleArrays into
``wb._cell_styles`` (wanted: that allocation is exactly how new xf indices
are assigned for the styles.xml append in Phase 2d).
"""

from openpyxl.cell._writer import etree_write_cell
from openpyxl.xml.functions import tostring


class _ElementCapture:
    """Minimal xf stand-in: captures the Element write_cell produces."""

    __slots__ = ("element",)

    def __init__(self):
        self.element = None

    def write(self, element):
        self.element = element


def emit_cell(ws, cell):
    """Bytes of one ``<c>`` element for ``cell``, or ``None`` when the model
    holds nothing worth serializing (mirrors the stock writer's skip rule)."""
    if cell._value is None and not cell.has_style:
        return None
    capture = _ElementCapture()
    saved_links = getattr(ws, "_hyperlinks", None)
    ws._hyperlinks = []
    try:
        etree_write_cell(capture, ws, cell, cell.has_style)
    finally:
        if saved_links is None:
            del ws._hyperlinks
        else:
            ws._hyperlinks = saved_links
    return tostring(capture.element)


def carry_attributes(new_cell_bytes, original_attrs):
    """PR-0 D6 attribute-carry rule: re-attach every original cell attribute
    the replacement does not intentionally rewrite (everything except r, s,
    t). ``cm``/``vm`` never reach here — the splice refuses them earlier."""
    carried = {k: v for k, v in original_attrs.items()
               if k not in ("r", "s", "t")}
    if not carried:
        return new_cell_bytes
    # inject into the start tag, before its terminating '>' or '/>'
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
