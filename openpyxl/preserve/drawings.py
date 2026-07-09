# paper-xlsx: drawing-part creation under preserve (PLAN-v0.1 Batch 4)

"""Charts and images added in-session become NEW drawing/chart/media parts
through the lifecycle engine — the stock writer's own serialization, never
a re-render of preserved bytes.

Three shapes, in increasing contact with original bytes:

- ADDED sheet (4.1): the stock worksheet writer already emitted
  ``<drawing r:id>``; this module supplies the parts and fills the rel.
- LOADED sheet, no drawing machinery (4.2): a fresh drawing part plus ONE
  ``<drawing r:id>`` element spliced into the sheet at its schema
  position.
- LOADED sheet, existing drawing (4.2): new anchors appended INTO the
  original drawing part — only when that part is anchor-only; anything
  else refuses.

Preserved drawing parts are never re-serialized: appends are byte
insertions before the closing tag, with namespaces declared on the
inserted anchors themselves.
"""

import re

from openpyxl.errors import UnsupportedStructureError

from . import crosspart
from .lifecycle import _rels_path, _resolve_target
from .xmlscan import ScanRefusal, _find_tag_end

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAWING_REL_TYPE = REL_NS + "/drawing"
CHART_REL_TYPE = REL_NS + "/chart"
IMAGE_REL_TYPE = REL_NS + "/image"
DRAWING_CT = "application/vnd.openxmlformats-officedocument.drawing+xml"
CHART_CT = ("application/vnd.openxmlformats-officedocument.drawingml."
            "chart+xml")
SHEET_DRAWING_NS = ("http://schemas.openxmlformats.org/drawingml/2006/"
                    "spreadsheetDrawing")

_IMAGE_MIME = {"png": "image/png", "jpeg": "image/jpeg", "gif": "image/gif",
               "bmp": "image/bmp", "tiff": "image/tiff"}

_ANCHOR_OPEN_RE = re.compile(
    br"<(oneCellAnchor|twoCellAnchor|absoluteAnchor)(?=[ >])")
_CNVPR_ID_RE = re.compile(br'\bid=(?:"(\d+)"|\'(\d+)\')')


def _refuse(msg):
    raise UnsupportedStructureError(msg + " Nothing was written.")


def _next_number(taken_names, pattern):
    """Next free 1-based number for a numbered part family, considering
    both original and engine-added parts."""
    rx = re.compile(pattern)
    highest = 0
    for name in taken_names:
        m = rx.match(name)
        if m:
            highest = max(highest, int(m.group(1)))
    return highest + 1


def _taken(names, part_plan):
    return set(names) | set(part_plan.added)


def _image_payload(image):
    """The image bytes, read non-destructively where possible. Refuses
    formats PIL cannot round-trip instead of crashing mid-save."""
    fmt = getattr(image, "format", "png") or "png"
    fmt = fmt.lower()
    if fmt not in _IMAGE_MIME:
        _refuse("image format {0!r} cannot be embedded (supported: "
                "{1}).".format(fmt, ", ".join(sorted(_IMAGE_MIME))))
    ref = getattr(image, "ref", None)
    try:
        if hasattr(ref, "getvalue"):
            data = ref.getvalue()
        elif hasattr(ref, "seek") and hasattr(ref, "read"):
            # PIL leaves the stream position wherever parsing stopped:
            # reading from the CURRENT position saves garbage media bytes
            # (Batch-4 gate) — read the whole stream, restore the position
            pos = ref.tell()
            ref.seek(0)
            data = ref.read()
            ref.seek(pos)
        elif isinstance(ref, str):
            with open(ref, "rb") as f:
                data = f.read()
        else:
            data = image._data()
    except UnsupportedStructureError:
        raise
    except Exception as exc:
        _refuse("image data could not be read ({0}).".format(exc))
    if not data:
        _refuse("image data is empty.")
    _SIGNATURES = {"png": b"\x89PNG", "gif": b"GIF8",
                   "jpeg": b"\xff\xd8\xff", "bmp": b"BM"}
    sig = _SIGNATURES.get(fmt)
    if sig is not None and not data.startswith(sig):
        _refuse("image data does not look like {0} (bad signature); the "
                "backing stream may have been consumed.".format(fmt))
    return data, fmt


def _render_drawing(charts, images):
    """(payload, rel_entries) for a fresh drawing holding exactly these
    objects — the stock SpreadsheetDrawing serialization. rel_entries are
    (rid, type, target, mode) with the stock rId1..rIdN local numbering
    matching the frame references inside the payload."""
    from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
    from openpyxl.xml.functions import tostring

    drawing = SpreadsheetDrawing()
    drawing.charts = list(charts)
    drawing.images = list(images)
    try:
        payload = tostring(drawing._write())
    except UnsupportedStructureError:
        raise
    except Exception as exc:
        _refuse("a chart/image anchor could not be serialized "
                "({0}).".format(exc))
    entries = []
    for i, rel in enumerate(drawing._rels, 1):
        entries.append(("rId{0}".format(i), rel.Type, rel.Target, None))
    return payload, entries


def _register_objects(workbook, ws, part_plan, names, charts, images):
    """Allocate part numbers, register chart/media parts with the engine,
    and stamp each object's _id so its ``path`` (the rel target inside the
    drawing) is final. Returns nothing; raises before ANY registration on
    invalid input (validate-then-mutate)."""
    taken = _taken(names, part_plan)
    # validate everything first: image data readable, charts single-use.
    # the seen-set lives on the PART PLAN (one save), not the workbook —
    # a second save of the same workbook replans the same additions and
    # must not false-refuse (Batch-4 gate suspicion, confirmed)
    image_data = [_image_payload(img) for img in images]
    seen = getattr(part_plan, "_drawn_charts", None)
    if seen is None:
        seen = set()
        part_plan._drawn_charts = seen
    for chart in charts:
        if id(chart) in seen:
            _refuse("the same chart object was added more than once; "
                    "charts are single-use (sheet {0!r}).".format(
                        ws.title))
        seen.add(id(chart))

    next_chart = _next_number(taken, r"xl/charts/chart(\d+)\.xml$")
    for i, chart in enumerate(charts):
        chart._id = next_chart + i
        part_name = "xl/charts/chart{0}.xml".format(chart._id)
        from openpyxl.xml.functions import tostring
        part_plan.add_part(part_name, tostring(chart._write()),
                           content_type=CHART_CT)
    next_media = _next_number(taken, r"xl/media/image(\d+)\.\w+$")
    for i, (image, (data, fmt)) in enumerate(zip(images, image_data)):
        image._id = next_media + i
        image.format = fmt
        part_name = "xl/media/image{0}.{1}".format(image._id, fmt)
        part_plan.add_part(part_name, data)
        part_plan.add_default(fmt, _IMAGE_MIME[fmt])


def plan_added_sheet_drawing(workbook, ws, part_plan, names, rel_entries):
    """Charts/images on an ADDED sheet (4.1): the stock writer emitted
    ``<drawing r:id>`` into the fresh sheet payload with an empty rel
    Target; supply the drawing/chart/media parts and fill the target.
    Returns the updated sheet rel entries."""
    charts = list(getattr(ws, "_charts", []) or [])
    images = list(getattr(ws, "_images", []) or [])
    if not charts and not images:
        return rel_entries
    _register_objects(workbook, ws, part_plan, names, charts, images)
    payload, drawing_rels = _render_drawing(charts, images)
    taken = _taken(names, part_plan)
    number = _next_number(taken, r"xl/drawings/drawing(\d+)\.xml$")
    drawing_part = "xl/drawings/drawing{0}.xml".format(number)
    part_plan.add_part(drawing_part, payload, content_type=DRAWING_CT)
    part_plan.add_part(_rels_path(drawing_part),
                       crosspart.render_rels_document(drawing_rels))
    out = []
    filled = False
    for (rid, rtype, target, mode) in rel_entries:
        if rtype == DRAWING_REL_TYPE and not target:
            target = "/" + drawing_part
            filled = True
        out.append((rid, rtype, target, mode))
    if not filled:
        _refuse("internal: the added sheet {0!r} has charts/images but "
                "its generated payload carries no drawing "
                "relationship.".format(ws.title))
    return out


def plan_fresh_drawing(workbook, ws, part_plan, names, sheet_part,
                       original_sheet_rels, charts, images):
    """Charts/images on a LOADED sheet with no drawing machinery (4.2):
    a fresh drawing part via the engine plus ONE spliced element. Returns
    the ``<drawing r:id>`` bytes for the region splice."""
    _register_objects(workbook, ws, part_plan, names, charts, images)
    payload, drawing_rels = _render_drawing(charts, images)
    taken = _taken(names, part_plan)
    number = _next_number(taken, r"xl/drawings/drawing(\d+)\.xml$")
    drawing_part = "xl/drawings/drawing{0}.xml".format(number)
    sheet_rels_part = _rels_path(sheet_part)
    rid = part_plan.reserve_rid(sheet_rels_part, original_sheet_rels)
    part_plan.add_part(drawing_part, payload, content_type=DRAWING_CT,
                       relate_from=sheet_part, rel_type=DRAWING_REL_TYPE,
                       rel_id=rid)
    part_plan.add_part(_rels_path(drawing_part),
                       crosspart.render_rels_document(drawing_rels))
    # xmlns:r declared inline: producers do not reliably declare it on
    # sheets that never used relationships (the tableParts lesson)
    return (b'<drawing xmlns:r="%s" r:id="%s"/>'
            % (REL_NS.encode("ascii"), rid.encode("ascii")))


def _existing_drawing_part(zin, names, sheet_part):
    """(drawing_part_name, rel_id) for the sheet's drawing, resolved
    through its ORIGINAL rels, or (None, None)."""
    rels_part = _rels_path(sheet_part)
    if rels_part not in names:
        return None, None
    payload = zin.read(rels_part)
    for m in re.finditer(br"<Relationship\b[^>]*>", payload):
        tag = m.group(0)
        type_m = re.search(br'Type=(?:"([^"]*)"|\'([^\']*)\')', tag)
        rtype = (type_m.group(1) or type_m.group(2)) if type_m else b""
        if not rtype.endswith(b"/drawing"):
            continue
        target_m = re.search(br'Target=(?:"([^"]*)"|\'([^\']*)\')', tag)
        if not target_m:
            continue
        id_m = re.search(br'Id=(?:"([^"]*)"|\'([^\']*)\')', tag)
        rid = ((id_m.group(1) or id_m.group(2)).decode("ascii")
               if id_m else None)
        target = (target_m.group(1) or target_m.group(2)).decode("utf-8")
        return _resolve_target(sheet_part, target), rid
    return None, None


def _iter_tags(body):
    """Yield (is_closing, local_name, self_closing) for every tag in
    ``body``, quote-aware (a '>' inside an attribute value never ends a
    tag — Batch-4 gate: false refusals). Raises ScanRefusal on
    unterminated tags."""
    pos = 0
    n = len(body)
    while pos < n:
        lt = body.find(b"<", pos)
        if lt == -1:
            return
        gt = _find_tag_end(body, lt)
        head = body[lt + 1:gt]
        closing = head.startswith(b"/")
        if closing:
            head = head[1:]
        self_closing = head.endswith(b"/")
        m = re.match(br"[\w:]+", head)
        name = m.group(0) if m else b""
        yield (closing, name.split(b":")[-1], self_closing)
        pos = gt + 1


def _cnvpr_max_id(payload):
    """The highest numeric cNvPr id in a drawing part, both quote styles,
    quote-aware tag boundaries."""
    max_id = 0
    pos = 0
    n = len(payload)
    while pos < n:
        lt = payload.find(b"<", pos)
        if lt == -1:
            break
        gt = _find_tag_end(payload, lt)
        head = payload[lt + 1:gt]
        if re.match(br"(?:\w+:)?cNvPr\b", head):
            for m in _CNVPR_ID_RE.finditer(head):
                max_id = max(max_id,
                             int(m.group(1) if m.group(1) is not None
                                 else m.group(2)))
        pos = gt + 1
    return max_id


def _wsdr_close(payload):
    """(insert_offset, root_uses_default_ns) for appending before the
    drawing root's close tag."""
    m = re.search(br"</(?:\w+:)?wsDr\s*>\s*$", payload)
    if m is None:
        _refuse("the drawing part has no recognizable closing tag.")
    root_m = re.search(br"<(\w+:)?wsDr\b", payload)
    return m.start(), root_m.group(1) is None if root_m else True


def _expand_self_closing_root(payload):
    """A schema-valid empty drawing may be a self-closing <wsDr .../>;
    expand it so the append has a body to land in. Returns the payload
    unchanged when the root is not self-closing."""
    m = re.search(br"<((?:\w+:)?wsDr)\b", payload)
    if m is None:
        return payload
    try:
        gt = _find_tag_end(payload, m.start())
    except ScanRefusal:
        return payload
    if payload[gt - 1:gt] != b"/":
        return payload
    if payload[gt + 1:].strip():
        return payload            # trailing content: not a bare root
    return (payload[:gt - 1] + b"></" + m.group(1) + b">"
            + payload[gt + 1:])


def _anchor_only(payload):
    """True when every top-level child of wsDr is an anchor element —
    the only drawing shape this module may append into."""
    payload = _expand_self_closing_root(payload)
    root_m = re.search(br"<(?:\w+:)?wsDr\b", payload)
    if root_m is None:
        return False
    try:
        root_gt = _find_tag_end(payload, root_m.start())
    except ScanRefusal:
        return False
    close_m = re.search(br"</(?:\w+:)?wsDr\s*>\s*$", payload)
    if close_m is None:
        return False
    body = payload[root_gt + 1:close_m.start()]
    if b"<!--" in body or b"<![" in body or b"<?" in body:
        return False              # comments/CDATA/PIs: hands off
    depth = 0
    try:
        for closing, local, self_closing in _iter_tags(body):
            if closing:
                depth -= 1
                if depth < 0:
                    return False
            else:
                if depth == 0:
                    if local not in (b"oneCellAnchor", b"twoCellAnchor",
                                     b"absoluteAnchor"):
                        return False
                if not self_closing:
                    depth += 1
    except ScanRefusal:
        return False
    return depth == 0


def plan_drawing_append(workbook, ws, part_plan, names, drawing_part,
                        original, existing_rels, charts, images):
    """Charts/images on a LOADED sheet whose drawing already exists (4.2):
    new anchors appended INTO the original drawing bytes (``original`` may
    already carry this save's shift patches) — only when that drawing is
    anchor-only. Returns the new drawing payload."""
    if not _anchor_only(original):
        _refuse("sheet {0!r} already has a drawing carrying content other "
                "than plain chart/image anchors (shapes, alternate "
                "content, ...); appending into it is not supported — "
                "the {1} would otherwise risk the existing drawing.".format(
                    ws.title, "chart" if charts else "image"))
    original = _expand_self_closing_root(original)
    _register_objects(workbook, ws, part_plan, names, charts, images)
    rendered, local_rels = _render_drawing(charts, images)
    root_m = re.search(br"<wsDr\b[^>]*>", rendered)
    close_m = re.search(br"</wsDr>\s*$", rendered)
    body = rendered[root_m.end():close_m.start()]

    # remap the render's local rId1..rIdN to ids reserved on the ORIGINAL
    # drawing's rels — through collision-proof placeholders: a sequential
    # in-place replace cross-wires anchors whenever a reserved id equals a
    # still-unreplaced local id (Batch-4 gate: chart anchor pointed at the
    # PNG, output unreadable)
    drawing_rels_part = _rels_path(drawing_part)
    rel_appends = []
    reserved = []
    for i, (local_rid, rtype, target, mode) in enumerate(local_rels, 1):
        rid = part_plan.reserve_rid(drawing_rels_part, existing_rels)
        body = body.replace(b'"rId%d"' % i, b'"__paperRid%d__"' % i, 1)
        reserved.append((i, rid))
        rel_appends.append((rid, rtype, target, mode))
    for i, rid in reserved:
        body = body.replace(b'"__paperRid%d__"' % i,
                            b'"%s"' % rid.encode("ascii"), 1)
    part_plan.rel_appends.setdefault(drawing_rels_part,
                                     []).extend(rel_appends)

    # shape ids bump past the original's maximum (quote-aware scan; the
    # render body itself is double-quoted by construction)
    max_id = _cnvpr_max_id(original)
    body = re.sub(
        br'(<cNvPr\b[^>]*?\bid=")(\d+)(")',
        lambda m: m.group(1) + b"%d" % (int(m.group(2)) + max_id)
        + m.group(3),
        body)

    insert_at, uses_default = _wsdr_close(original)
    if not uses_default:
        # the host document is prefix-namespaced: the appended anchors
        # carry their own default declaration (namespace-correct in any
        # host; the a:/c:/r: children already self-declare)
        body = _ANCHOR_OPEN_RE.sub(
            br'<\1 xmlns="' + SHEET_DRAWING_NS.encode("ascii") + b'"',
            body)
    return original[:insert_at] + body + original[insert_at:]
