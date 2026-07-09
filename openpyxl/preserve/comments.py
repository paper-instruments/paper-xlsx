# paper-xlsx: comment creation under preserve (PLAN-v0.1 Batch 2; PR-1 §1.3)

"""Comments on sheets whose original package carries NO comment machinery —
the 80% case. The comments part and its legacy-VML anchor part are created
whole via the lifecycle engine (nothing to splice: both parts are new), and
one <legacyDrawing r:id> element rides the region splice.

Sheets that ALREADY carry comment parts keep refusing: editing preserved
VML is Batch-4-class work that earns its own probe.
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
    """True when the ORIGINAL sheet already references comment/VML parts."""
    rels_part = _rels_path(sheet_part)
    if rels_part not in names:
        return False
    root = crosspart.scan_small(zin.read(rels_part), "Relationships",
                                max_depth=1)
    for child in root.children:
        rel_type = child.attrs.get("Type", "")
        if rel_type.endswith("/comments") \
                or rel_type.endswith("/vmlDrawing"):
            return True
    return False


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
                    # under the stdlib serializer (Batch-2 gate)
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
