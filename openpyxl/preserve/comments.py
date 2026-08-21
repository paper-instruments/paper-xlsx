# paper-xlsx: comment creation under preserve

"""Comments on sheets whose original package carries NO comment machinery —
the 80% case. The comments part and its legacy-VML anchor part are created
whole via the lifecycle engine (nothing to splice: both parts are new), and
one <legacyDrawing r:id> element rides the region splice.

Sheets that ALREADY carry comment parts keep refusing: editing preserved
VML is out of scope.
"""

from openpyxl.errors import UnsupportedStructureError
from openpyxl.xml.functions import tostring

from . import crosspart

_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
COMMENTS_CONTENT_TYPE = ("application/vnd.openxmlformats-officedocument."
                         "spreadsheetml.comments+xml")
VML_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.vmlDrawing"


def _refuse(msg):
    raise UnsupportedStructureError(msg + " Nothing was written.")


def sheet_has_comment_machinery(zin, sheet_part, names):
    """True only for comment relationships or VML note shapes."""
    return comment_machinery_kind(zin, sheet_part, names) == "comments"


def comment_machinery_kind(zin, sheet_part, names):
    """Classify a sheet's comment and VML relationships.

    :param zin: Open workbook package.
    :type zin: zipfile.ZipFile
    :param sheet_part: Worksheet part name in the package.
    :type sheet_part: str
    :param names: Part names present in the package.
    :type names: builtins.set[str]
    :return: ``"comments"``, ``"other-vml"``, or ``None``.
    :rtype: str or None
    """
    rels_part = _rels_path(sheet_part)
    if rels_part not in names:
        return None
    root = crosspart.scan_small(zin.read(rels_part), "Relationships",
                                max_depth=1)
    saw_comments = False
    saw_other_vml = False
    for child in root.children:
        rel_type = child.attrs.get("Type", "")
        if rel_type.endswith("/comments"):
            saw_comments = True
            continue
        if not rel_type.endswith("/vmlDrawing"):
            continue
        target = _resolve_target(sheet_part, child.attrs.get("Target", ""))
        payload = zin.read(target) if target in names else b""
        if _vml_is_comment_only(payload):
            saw_comments = True
        else:
            saw_other_vml = True
    if saw_other_vml:
        return "other-vml"
    return "comments" if saw_comments else None


def _vml_is_comment_only(payload):
    """Whether every drawing object in one VML part is a note shape."""
    try:
        from openpyxl.xml.functions import fromstring

        root = fromstring(payload)
    except Exception:
        return False

    vml_namespace = "urn:schemas-microsoft-com:vml"
    office_namespace = "urn:schemas-microsoft-com:office:office"
    excel_namespace = "urn:schemas-microsoft-com:office:excel"
    drawing_objects = {
        "arc", "curve", "group", "image", "line", "oval", "polyline",
        "rect", "roundrect", "shape",
    }
    shapes = []
    for element in root:
        namespace, local = _tag_parts(element)
        if namespace == vml_namespace and local == "shape":
            shapes.append(element)
        elif namespace == vml_namespace and local == "shapetype":
            continue
        elif namespace == office_namespace and local == "shapelayout":
            continue
        else:
            return False
    if not shapes:
        return False
    all_shapes = []
    for element in root.iter():
        namespace, local = _tag_parts(element)
        if namespace == vml_namespace and local in drawing_objects:
            if local != "shape":
                return False
            all_shapes.append(element)
    if len(all_shapes) != len(shapes):
        return False
    for shape in shapes:
        client_data = [element for element in shape.iter()
                       if _tag_parts(element) ==
                       (excel_namespace, "ClientData")]
        if len(client_data) != 1 \
                or client_data[0].get("ObjectType") != "Note":
            return False
    return True


def _tag_parts(element):
    tag = getattr(element, "tag", "")
    if isinstance(tag, str) and tag.startswith("{") and "}" in tag:
        return tuple(tag[1:].split("}", 1))
    return None, tag


def comment_anchors_precede_shift(zin, sheet_part, names, axis, index):
    """Whether every source note anchor is wholly before one shift."""
    rels_part = _rels_path(sheet_part)
    if rels_part not in names:
        return False
    root = crosspart.scan_small(zin.read(rels_part), "Relationships",
                                max_depth=1)
    found = False
    for child in root.children:
        if not child.attrs.get("Type", "").endswith("/vmlDrawing"):
            continue
        target = _resolve_target(sheet_part, child.attrs.get("Target", ""))
        if target not in names:
            return False
        try:
            from openpyxl.xml.functions import fromstring

            vml = fromstring(zin.read(target))
        except Exception:
            return False
        for client_data in vml.iter():
            if _tag_parts(client_data)[1] != "ClientData" \
                    or client_data.get("ObjectType") != "Note":
                continue
            found = True
            coordinates = {}
            for element in client_data:
                local = _tag_parts(element)[1]
                if local in ("Row", "Column", "Anchor"):
                    coordinates[local] = element.text
            local = "Row" if axis == "rows" else "Column"
            try:
                position = int(coordinates[local])
                if position < 0 or position + 1 >= index:
                    return False
                anchor = coordinates.get("Anchor")
                if anchor is not None:
                    values = [int(value.strip())
                              for value in anchor.split(",")]
                    if len(values) != 8:
                        return False
                    positions = values[2:7:4] if axis == "rows" \
                        else values[0:5:4]
                    if any(position < 0 or position + 1 >= index
                           for position in positions):
                        return False
            except (KeyError, TypeError, ValueError):
                return False
    return found


def plan_comment_creation(wb, ws, sheet_part, zin, part_plan, names):
    """Create the comments + VML parts for one comment-free sheet; returns
    the crafted <legacyDrawing r:id> bytes for the region splice."""
    from openpyxl.comments.comment_sheet import CommentRecord, CommentSheet
    from .regions import hyperlink_signatures

    led = wb._paper_ledger
    armed_links = led.region_snapshots.get(ws, {}).get("hyperlinks", {})
    if hyperlink_signatures(ws) != armed_links:
        _refuse("sheet {0!r} adds comments AND changes hyperlinks in the "
                "same save; their relationship allocations would collide. "
                "Save between the two edits.".format(ws.title))

    from openpyxl.utils.exceptions import IllegalCharacterError
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

    records = []
    for (_row, _col), cell in sorted(ws._cells.items()):
        if cell._comment is not None:
            comment = cell._comment
            for text in (comment.content or "", comment.author or ""):
                if ILLEGAL_CHARACTERS_RE.search(text):
                    # cell values get this guard in check_string; comments
                    # must too, or the save writes an unparseable part
                    # under the stdlib serializer
                    _refuse("comment on {0}!{1} contains characters that "
                            "cannot be written to XML (control "
                            "bytes).".format(ws.title, cell.coordinate))
            records.append(CommentRecord.from_cell(cell))
    if not records:
        _refuse("internal: comment creation planned with no comments on "
                "sheet {0!r}.".format(ws.title))

    cs = CommentSheet.from_comments(records)
    payload = tostring(cs.to_tree())
    if not payload.startswith(b"<?xml"):
        payload = (b'<?xml version="1.0" encoding="UTF-8" '
                   b'standalone="yes"?>\n' + payload)
    vml = cs.write_shapes(None)

    all_names = set(names) | set(part_plan.added)
    number = _next_number(all_names, r"xl/comments/comment(\d+)\.xml$")
    number = max(number, _next_number(all_names, r"xl/comments(\d+)\.xml$"))
    comments_part = "xl/comments/comment{0}.xml".format(number)
    vml_part = "xl/drawings/commentsDrawing{0}.vml".format(
        _next_number(all_names,
                     r"xl/drawings/commentsDrawing(\d+)\.vml$"))

    rels_part = _rels_path(sheet_part)
    rels_payload = zin.read(rels_part) if rels_part in names else None
    comments_rid = part_plan.reserve_rid(rels_part, rels_payload)
    vml_rid = part_plan.reserve_rid(rels_part, rels_payload)

    part_plan.add_part(comments_part, payload,
                       content_type=COMMENTS_CONTENT_TYPE,
                       relate_from=sheet_part,
                       rel_type=_REL_NS + "/comments",
                       rel_id=comments_rid)
    part_plan.add_part(vml_part, vml,
                       relate_from=sheet_part,
                       rel_type=_REL_NS + "/vmlDrawing",
                       rel_id=vml_rid)
    part_plan.add_default("vml", VML_CONTENT_TYPE)

    return (b'<legacyDrawing xmlns:r="%s" r:id="%s"/>' % (
        _REL_NS.encode("ascii"), vml_rid.encode("ascii")))


def _next_number(names, pattern):
    import re

    rx = re.compile(pattern)
    highest = 0
    for name in names:
        m = rx.match(name)
        if m:
            highest = max(highest, int(m.group(1)))
    return highest + 1


def _rels_path(part_name):
    folder, _, base = part_name.rpartition("/")
    return "{0}/_rels/{1}.rels".format(folder, base) if folder \
        else "_rels/{0}.rels".format(base)


def _resolve_target(from_part, target):
    if target.startswith("/"):
        return target[1:]
    base = from_part.rpartition("/")[0].split("/")
    for piece in target.split("/"):
        if piece == "..":
            base = base[:-1]
        elif piece != ".":
            base.append(piece)
    return "/".join(base)
