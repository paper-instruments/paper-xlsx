# paper-xlsx: chart-range rewriting in preserved bytes (PLAN Phase 6c)

"""Targeted patches of series references inside PRESERVED chart parts, and
of anchor positions inside preserved drawing parts, so row/column shifts can
proceed on chart-referenced sheets.

Scope is honest (PLAN: "if it slips, the refusal stands — never the silent
third option"): only ``<c:f>`` reference texts and ``<xdr:from>/<xdr:to>``
anchor markers are rewritten, located namespace-aware (openpyxl writes both
families with DEFAULT namespaces; Excel and LibreOffice write ``c:``/``xdr:``
prefixes — both forms are handled by resolving names, never by matching
prefixes). Charts carrying modern extension machinery (c15 filtered series,
extLst, AlternateContent) refuse; a shift that would delete charted data
refuses rather than write ``#REF!`` into a chart.
"""

import io
import re
import zipfile

from .xmlscan import ScanRefusal, _find_tag_end, _scan_name_end, _ATTR_RE

CHART_NS = b"http://schemas.openxmlformats.org/drawingml/2006/chart"
XDR_NS = b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


def _walk_leaf_texts(data):
    """Yield (clark_ns, local, parent_local, text_start, text_end) for every
    element with pure text content, tracking namespaces exactly like the
    sheet scanner. Raises ScanRefusal on constructions we must not touch."""
    pos = 0
    if data[:3] == b"\xef\xbb\xbf":
        pos = 3
    n = len(data)
    stack = []   # [local, default_ns, prefixes, ns, content_start]
    while pos < n:
        lt = data.find(b"<", pos)
        if lt == -1:
            break
        nxt = data[lt + 1]
        if nxt == 0x3F:
            end = data.find(b"?>", lt)
            if end == -1:
                raise ScanRefusal("unterminated processing instruction")
            pos = end + 2
            continue
        if nxt == 0x21:
            if data.startswith(b"<!--", lt):
                end = data.find(b"-->", lt)
                if end == -1:
                    raise ScanRefusal("unterminated comment")
                pos = end + 3
                continue
            raise ScanRefusal("CDATA/DOCTYPE in a chart or drawing part")
        if nxt == 0x2F:
            gt = data.find(b">", lt)
            if gt == -1 or not stack:
                raise ScanRefusal("unbalanced end tag")
            entry = stack.pop()
            if entry[4] is not None and entry[4] <= lt:
                text = data[entry[4]:lt]
                if b"<" not in text:
                    parent_local = stack[-1][0] if stack else b""
                    yield (entry[3], entry[0], parent_local, entry[4], lt)
            pos = gt + 1
            continue

        gt = _find_tag_end(data, lt)
        self_closing = data[gt - 1:gt] == b"/"
        head = data[lt + 1: gt - 1 if self_closing else gt]
        name_end = _scan_name_end(head)
        raw_name = head[:name_end]
        attr_blob = head[name_end:]

        parent_default = stack[-1][1] if stack else None
        parent_prefixes = stack[-1][2] if stack else {}
        default_ns, prefixes = parent_default, parent_prefixes
        if b"xmlns" in attr_blob:
            prefixes = dict(parent_prefixes)
            for m in _ATTR_RE.finditer(attr_blob):
                key = m.group(1)
                value = m.group(3) if m.group(3) is not None else m.group(4)
                if key == b"xmlns":
                    default_ns = value
                elif key.startswith(b"xmlns:"):
                    prefixes[key[6:]] = value

        if b":" in raw_name:
            prefix, local = raw_name.split(b":", 1)
            ns = prefixes.get(prefix)
        else:
            local, ns = raw_name, default_ns

        if not self_closing:
            stack.append([local, default_ns, prefixes, ns, gt + 1])
        pos = gt + 1


_ENTITY_MAP = ((b"&amp;", b"&"), (b"&lt;", b"<"), (b"&gt;", b">"),
               (b"&quot;", b'"'), (b"&apos;", b"'"))


def _unescape(text):
    if b"&#" in text:
        raise ScanRefusal("numeric character references in a chart formula")
    for entity, char in _ENTITY_MAP:
        text = text.replace(entity, char)
    return text


def _escape(text):
    text = text.replace(b"&", b"&amp;")
    return text.replace(b"<", b"&lt;").replace(b">", b"&gt;")


def patch_chart(payload, sheet_title, operation, index, amount):
    """(new_payload, changed, blockers): rewrite every chart formula text
    referencing ``sheet_title`` per the shift."""
    from .rewrite import shift_name_value

    blockers = []
    for marker, label in ((b"c15:", "c15 filtered-series machinery"),
                          (b"AlternateContent", "alternate-content blocks"),
                          (b"extLst", "chart extension lists")):
        if marker in payload:
            blockers.append("the chart carries {0}, whose references the "
                            "patch cannot see".format(label))
    if blockers:
        return payload, False, blockers

    axis = "rows" if "rows" in operation else "cols"
    is_delete = operation.startswith("delete")
    edits = []
    try:
        for ns, local, _parent, t_start, t_end in _walk_leaf_texts(payload):
            if local != b"f" or ns != CHART_NS:
                continue
            raw = payload[t_start:t_end]
            text = _unescape(raw).decode("utf-8")
            new_text, changed = shift_name_value(
                text, sheet_title, axis, index, amount, is_delete)
            if not changed:
                continue
            if "#REF!" in new_text:
                blockers.append(
                    "the shift would delete data charted by reference "
                    "{0!r}".format(text))
                continue
            edits.append((t_start, t_end,
                          _escape(new_text.encode("utf-8"))))
    except ScanRefusal as exc:
        return payload, False, [str(exc)]
    if blockers:
        return payload, False, blockers
    if not edits:
        return payload, False, []
    from .crosspart import apply_edits

    return apply_edits(payload, edits), True, []


def patch_drawing_anchors(payload, operation, index, amount):
    """(new_payload, changed): move drawing anchor markers with the cells
    (from/to row and col elements in the spreadsheetDrawing namespace)."""
    axis = "rows" if "rows" in operation else "cols"
    is_delete = operation.startswith("delete")
    wanted = b"row" if axis == "rows" else b"col"
    edits = []
    for ns, local, parent, t_start, t_end in _walk_leaf_texts(payload):
        if ns != XDR_NS or local != wanted or parent not in (b"from", b"to"):
            continue
        try:
            value = int(payload[t_start:t_end])
        except ValueError:
            continue
        anchor = value + 1     # anchors are 0-based; shifts are 1-based
        if is_delete:
            if anchor >= index + amount:
                new_anchor = anchor - amount
            elif anchor >= index:
                new_anchor = index
            else:
                continue
        else:
            if anchor < index:
                continue
            new_anchor = anchor + amount
        if new_anchor == anchor:
            continue
        edits.append((t_start, t_end, str(new_anchor - 1).encode("ascii")))
    if not edits:
        return payload, False
    from .crosspart import apply_edits

    return apply_edits(payload, edits), True


def plan_chart_updates(wb, sheet_title, operation, index, amount):
    """Plan every chart/drawing part patch for a shift on ``sheet_title``.

    Returns ({part_name: new_payload}, blockers). Used twice: as a dry run
    when the edit is attempted (blockers refuse before any mutation) and for
    real at save time.
    """
    from .structural import _charts_referencing

    source = getattr(wb, "_paper_source", None)
    plans = {}
    blockers = []
    if not source:
        return plans, blockers
    chart_parts = _charts_referencing(wb, sheet_title)
    with zipfile.ZipFile(io.BytesIO(source)) as z:
        for part in chart_parts:
            payload = z.read(part)
            new_payload, changed, part_blockers = patch_chart(
                payload, sheet_title, operation, index, amount)
            for blocker in part_blockers:
                blockers.append("{0}: {1}".format(part, blocker))
            if changed and not part_blockers:
                plans[part] = new_payload
        drawing_part = _sheet_drawing_part(z, sheet_title)
        if drawing_part is not None:
            payload = z.read(drawing_part)
            new_payload, changed = patch_drawing_anchors(
                payload, operation, index, amount)
            if changed:
                plans[drawing_part] = new_payload
    return plans, blockers


def _sheet_drawing_part(zin, sheet_title):
    from openpyxl.packaging.relationship import get_dependents, get_rels_path

    from .saver import _package_info

    _wb_part, mapping = _package_info(zin)
    sheet_part = mapping.get(sheet_title)
    if sheet_part is None:
        return None
    rels_path = get_rels_path(sheet_part)
    if rels_path not in zin.namelist():
        return None
    rels = get_dependents(zin, rels_path)
    for rel in rels:
        if rel.Type.endswith("/drawing"):
            return rel.target
    return None
