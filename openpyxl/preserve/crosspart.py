# paper-xlsx: cross-part edits (CONVENTIONS §3.5; PR-0 D2/D9/D11/D12/D13)

"""Targeted edits to non-worksheet parts under preserve mode.

Every part here is edited by byte splice against its ORIGINAL payload:
appended entries land before the container's end tag, removed entries are
cut at their scanned spans, and everything unmodeled (extLst,
mc:AlternateContent, fileVersion, unknown attributes) passes through
untouched. Rels are append-only with ``rId = max numeric existing + 1``
(PR-0 D11); [Content_Types].xml is edited by targeted append/remove, never
regenerated (D12); calcChain removal cascades to its content-type override
and workbook relationship (D13).
"""

import re

from openpyxl.errors import UnsupportedStructureError
from openpyxl.xml.functions import tostring

from .xmlscan import ScanRefusal, _find_tag_end, _scan_name_end, _ATTR_RE


class Node:
    __slots__ = ("name", "start", "end", "content_start", "content_end",
                 "self_closing", "attrs", "children")

    def __init__(self, name, start):
        self.name = name              # raw tag name (with prefix if any)
        self.start = start
        self.end = None
        self.content_start = None
        self.content_end = None
        self.self_closing = False
        self.attrs = {}
        self.children = []

    def local(self):
        name = self.name.decode("latin-1")
        return name.split(":", 1)[1] if ":" in name else name


def scan_small(data, expected_root, max_depth=3):
    """Compact span scanner for small XML parts (workbook.xml, styles.xml,
    [Content_Types].xml, rels). Records children down to ``max_depth``.
    Refuses prefixed roots — byte edits would land outside the schema."""
    pos = 0
    if data[:3] == b"\xef\xbb\xbf":
        pos = 3
    n = len(data)
    root = None
    stack = []

    while pos < n:
        lt = data.find(b"<", pos)
        if lt == -1:
            break
        nxt = data[lt + 1]
        if nxt == 0x3F:                       # <? ... ?>
            end = data.find(b"?>", lt)
            if end == -1:
                raise ScanRefusal("unterminated processing instruction")
            pos = end + 2
            continue
        if nxt == 0x21:                       # <!-- / <![CDATA[ / DOCTYPE
            if data.startswith(b"<!--", lt):
                end = data.find(b"-->", lt)
                if end == -1:
                    raise ScanRefusal("unterminated comment")
                pos = end + 3
                continue
            if data.startswith(b"<![CDATA[", lt):
                end = data.find(b"]]>", lt)
                if end == -1:
                    raise ScanRefusal("unterminated CDATA")
                pos = end + 3
                continue
            raise ScanRefusal("cannot edit this part: DOCTYPE or unknown "
                              "markup present")
        if nxt == 0x2F:                       # end tag
            gt = data.find(b">", lt)
            if gt == -1 or not stack:
                raise ScanRefusal("unbalanced end tag")
            node = stack.pop()
            node.end = gt + 1
            node.content_end = lt
            pos = gt + 1
            continue

        gt = _find_tag_end(data, lt)
        self_closing = data[gt - 1:gt] == b"/"
        tag_end = gt + 1
        head = data[lt + 1: gt - 1 if self_closing else gt]
        name_end = _scan_name_end(head)
        raw_name = head[:name_end]
        node = Node(raw_name, lt)
        node.self_closing = self_closing
        if self_closing:
            node.end = tag_end
        else:
            node.content_start = tag_end
        if len(stack) <= max_depth:
            for m in _ATTR_RE.finditer(head[name_end:]):
                key = m.group(1)
                value = m.group(3) if m.group(3) is not None else m.group(4)
                node.attrs[key.decode("latin-1")] = _unescape_attr(
                    value.decode("utf-8"))

        if root is None:
            if b":" in raw_name:
                raise ScanRefusal(
                    "cannot edit this part: its root element is namespace-"
                    "prefixed ({0!r})".format(raw_name.decode("latin-1")))
            if node.local() != expected_root:
                raise ScanRefusal(
                    "unexpected root element {0!r} (wanted {1!r})".format(
                        node.local(), expected_root))
            root = node
            if self_closing:
                return root
            stack.append(node)
            pos = tag_end
            continue

        if stack:
            stack[-1].children.append(node)
        if not self_closing:
            stack.append(node)
        pos = tag_end

    if stack or root is None:
        raise ScanRefusal("malformed XML part")
    return root


def _escape(value):
    return (value.replace("&", "&amp;").replace("<", "&lt;")
            .replace('"', "&quot;").encode("utf-8"))


_CHARREF_RE = re.compile(r"&#(?:x([0-9a-fA-F]+)|([0-9]+));")


def _unescape_attr(value):
    """Attribute values compare against MODEL strings (sheet titles, part
    names): entity escapes must be expanded or names like 'P&L' never match
    (measured: sheet-state changes silently dropped)."""
    if "&" not in value:
        return value
    value = _CHARREF_RE.sub(
        lambda m: chr(int(m.group(1), 16) if m.group(1) else int(m.group(2))),
        value)
    for entity, char in (("&lt;", "<"), ("&gt;", ">"), ("&quot;", '"'),
                         ("&apos;", "'"), ("&amp;", "&")):
        value = value.replace(entity, char)
    return value


def apply_edits(data, edits):
    """Apply sorted, non-overlapping (start, end, replacement) edits."""
    edits = sorted(edits, key=lambda e: (e[0], e[1]))
    for (s1, e1, _), (s2, e2, _) in zip(edits, edits[1:]):
        if e1 > s2:
            raise UnsupportedStructureError(
                "internal: overlapping cross-part edits. Nothing was written.")
    out = []
    pos = 0
    for start, end, replacement in edits:
        out.append(data[pos:start])
        out.append(replacement)
        pos = end
    out.append(data[pos:])
    return b"".join(out)


def _insert_into(node, data, payload):
    """(start, end, replacement) edit appending ``payload`` inside ``node``
    (expanding a self-closing container when necessary)."""
    if node.self_closing:
        head = data[node.start:node.end]
        assert head.endswith(b"/>")
        expanded = head[:-2] + b">" + payload + b"</" + node.name + b">"
        return (node.start, node.end, expanded)
    return (node.content_end, node.content_end, payload)


# ---------------------------------------------------------------------
# [Content_Types].xml

def ct_append_overrides(data, overrides):
    """Append Override entries; ``overrides`` = [(part_name, content_type)]."""
    root = scan_small(data, "Types", max_depth=1)
    payload = b"".join(
        b'<Override PartName="/%s" ContentType="%s"/>' % (
            _escape(part), _escape(ctype))
        for part, ctype in overrides)
    return apply_edits(data, [_insert_into(root, data, payload)])


def ct_append_defaults(data, defaults):
    """Append Default entries; ``defaults`` = [(extension, content_type)].
    Callers check for existing extensions first (duplicates are illegal)."""
    root = scan_small(data, "Types", max_depth=1)
    payload = b"".join(
        b'<Default Extension="%s" ContentType="%s"/>' % (
            _escape(ext), _escape(ctype))
        for ext, ctype in defaults)
    return apply_edits(data, [_insert_into(root, data, payload)])


def ct_remove_override(data, part_name):
    """Remove the Override for ``part_name``; no-op when absent."""
    root = scan_small(data, "Types", max_depth=1)
    target = "/" + part_name
    edits = []
    for child in root.children:
        if child.local() == "Override" and child.attrs.get("PartName") == target:
            edits.append((child.start, child.end, b""))
    if not edits:
        return data
    return apply_edits(data, edits)


# ---------------------------------------------------------------------
# relationship parts

def rels_next_rid(data):
    root = scan_small(data, "Relationships", max_depth=1)
    highest = 0
    for child in root.children:
        rid = child.attrs.get("Id", "")
        m = re.match(r"rId(\d+)$", rid)
        if m:
            highest = max(highest, int(m.group(1)))
    return highest + 1


def rels_append(data, entries):
    """Append relationship entries: [(rId, type, target, target_mode|None)]."""
    root = scan_small(data, "Relationships", max_depth=1)
    parts = []
    for rid, rtype, target, mode in entries:
        attrs = b'<Relationship Id="%s" Type="%s" Target="%s"' % (
            _escape(rid), _escape(rtype), _escape(target))
        if mode:
            attrs += b' TargetMode="%s"' % _escape(mode)
        parts.append(attrs + b"/>")
    return apply_edits(data, [_insert_into(root, data, b"".join(parts))])


def rels_remove_by_target_suffix(data, suffix):
    """Remove relationships whose Target ends with ``suffix``."""
    root = scan_small(data, "Relationships", max_depth=1)
    edits = [(c.start, c.end, b"") for c in root.children
             if c.local() == "Relationship"
             and c.attrs.get("Target", "").endswith(suffix)]
    if not edits:
        return data
    return apply_edits(data, edits)


def render_rels_document(entries):
    """A whole new .rels part (for sheets that had none)."""
    body = rels_append(
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/'
        b'2006/relationships"></Relationships>', entries)
    return body


# ---------------------------------------------------------------------
# workbook.xml

# CT_Workbook child sequence (ECMA-376 §18.2.27)
CT_WORKBOOK_ORDER = [
    "fileVersion", "fileSharing", "workbookPr", "workbookProtection",
    "bookViews", "sheets", "functionGroups", "externalReferences",
    "definedNames", "calcPr", "oleSize", "customWorkbookViews",
    "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing",
    "fileRecoveryPr", "webPublishObjects", "extLst",
]
CT_WORKBOOK_INDEX = {tag: i for i, tag in enumerate(CT_WORKBOOK_ORDER)}

# elements the model can faithfully re-render and the splice may replace
WB_SPLICEABLE = {"definedNames", "calcPr", "bookViews"}


def render_workbook_elements(wb):
    """Per-element model renders of workbook.xml (self-consistent between
    arm and save; 'sheets' is compared structurally, not by render)."""
    from openpyxl.workbook._writer import WorkbookWriter
    from openpyxl.xml.functions import fromstring

    rendered = WorkbookWriter(wb).write()
    tree = fromstring(rendered)
    out = {}
    for child in tree:
        tag = child.tag
        local = tag.split("}")[-1] if "}" in tag else tag
        if local == "sheets":
            continue
        out.setdefault(local, []).append(_serialize_plain(child))
    return {k: tuple(v) for k, v in out.items()}


_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _serialize_plain(el):
    """Serialize an element for splicing into a default-namespace document:
    main-namespace tags become plain, relationship-namespace attributes
    become r:... with a local declaration; anything else refuses."""
    parts = [b"<"]
    local = el.tag.split("}")[-1] if "}" in el.tag else el.tag
    parts.append(local.encode("ascii"))
    needs_r = False
    attrs = []
    for key, value in el.attrib.items():
        if key.startswith("{"):
            ns, attr_local = key[1:].split("}", 1)
            if ns == _REL_NS:
                attrs.append((b"r:" + attr_local.encode("ascii"), value))
                needs_r = True
            elif ns == "http://www.w3.org/XML/1998/namespace":
                attrs.append((b"xml:" + attr_local.encode("ascii"), value))
            else:
                raise UnsupportedStructureError(
                    "internal: cannot serialize attribute in namespace "
                    "{0!r}. Nothing was written.".format(ns))
        else:
            attrs.append((key.encode("ascii"), value))
    if needs_r:
        parts.append(b' xmlns:r="' + _REL_NS.encode("ascii") + b'"')
    for key, value in attrs:
        parts.append(b' %s="%s"' % (key, _escape(value)))
    children = list(el)
    text = el.text if el.text and el.text.strip() else None
    if not children and text is None:
        parts.append(b"/>")
        return b"".join(parts)
    parts.append(b">")
    if text is not None:
        parts.append(_escape_text(el.text))
    for child in children:
        parts.append(_serialize_plain(child))
    parts.append(b"</" + local.encode("ascii") + b">")
    return b"".join(parts)


def _escape_text(value):
    return (value.replace("&", "&amp;").replace("<", "&lt;")
            .encode("utf-8"))


def plan_workbook_xml(wb, led, original, new_sheet_entries, force_tags=()):
    """New workbook.xml bytes, or None if nothing changes.

    ``new_sheet_entries``: [(title, sheet_id, rid, state)] for added sheets.
    ``force_tags``: spliceable elements re-rendered even when the arm-vs-save
    diff sees no change (the recalc-on-load flag: the model defaults it, so
    only a forced re-render can write it into files that lack it).
    Raises for changes outside the v0-spliceable set.
    """
    current = render_workbook_elements(wb)
    snapshot = led.workbook_snapshot or {}
    changed = {}
    for tag in set(current) | set(snapshot):
        if current.get(tag) != snapshot.get(tag):
            changed[tag] = current.get(tag)
    for tag in force_tags:
        if tag in WB_SPLICEABLE and tag not in changed:
            changed[tag] = current.get(tag)

    unsupported = set(changed) - WB_SPLICEABLE
    if unsupported:
        raise UnsupportedStructureError(
            "workbook-level change(s) to {0} cannot be written under "
            "preserve mode in v0. Nothing was written.".format(
                sorted(unsupported)))

    state_changes = {}
    for title, arm_state in led.sheet_states.items():
        sheet = wb[title] if title in wb.sheetnames else None
        if sheet is not None and sheet.sheet_state != arm_state:
            state_changes[title] = sheet.sheet_state

    if not (changed or new_sheet_entries or state_changes):
        return None

    root = scan_small(original, "workbook", max_depth=2)
    by_local = {}
    for child in root.children:
        by_local.setdefault(child.local(), []).append(child)

    edits = []

    # spliceable element replacements/insertions/removals
    for tag, renders in changed.items():
        payload = b"".join(renders) if renders else b""
        spans = by_local.get(tag, [])
        if spans:
            first = spans[0]
            edits.append((first.start, first.end, payload))
            for extra in spans[1:]:
                edits.append((extra.start, extra.end, b""))
        elif payload:
            edits.append(_wb_insert_edit(root, by_local, tag, payload))

    # sheets element: state patches + appended entries only
    sheets_nodes = by_local.get("sheets", [])
    if not sheets_nodes:
        raise UnsupportedStructureError(
            "the workbook part has no sheets element. Nothing was written.")
    sheets_node = sheets_nodes[0]
    renamed = {}
    for ws_obj, original_title in getattr(led, "renames", {}).items():
        if ws_obj.title != original_title:
            renamed[original_title] = ws_obj.title

    # removals/reorder rebuild the sheets children from ORIGINAL entry
    # bytes (PLAN-v0.1 3.2); the per-entry patch path handles the rest
    removed = set(getattr(led, "removed_sheets", ()))
    current_originals = []
    for sheet in wb._sheets:
        title = sheet.title
        orig = getattr(led, "renames", {}).get(sheet, title)
        if orig in led.sheet_order or title in led.sheet_order:
            current_originals.append(orig)
    armed_order = [t for t in getattr(led, "sheet_order", ())
                   if t not in removed]
    rebuild = bool(removed) or current_originals != armed_order
    if rebuild:
        by_name = {}
        for entry in sheets_node.children:
            if entry.local() == "sheet":
                by_name[entry.attrs.get("name")] =                     original[entry.start:entry.end]
        pieces = []
        for orig_title in current_originals:
            blob = by_name.get(orig_title)
            if blob is None:
                raise UnsupportedStructureError(
                    "cannot rebuild the sheets element: no original entry "
                    "for {0!r}. Nothing was written.".format(orig_title))
            if orig_title in renamed:
                blob = _rename_entry_name(blob, renamed[orig_title])
            new_state = state_changes.get(
                renamed.get(orig_title, orig_title))
            if new_state is not None:
                blob = _repatch_entry_state(blob, new_state)
            pieces.append(blob)
        if new_sheet_entries:
            pieces.extend(
                _render_sheet_entry(title, sheet_id, rid, state)
                for title, sheet_id, rid, state in new_sheet_entries)
        inner_start = min(c.start for c in sheets_node.children)             if sheets_node.children else None
        inner_end = max(c.end for c in sheets_node.children)             if sheets_node.children else None
        if inner_start is None:
            raise UnsupportedStructureError(
                "cannot rebuild an empty sheets element. Nothing was "
                "written.")
        edits.append((inner_start, inner_end, b"".join(pieces)))
        new_sheet_entries = ()          # consumed by the rebuild
        state_changes = {}
        renamed = {}

    if state_changes or renamed:
        for entry in sheets_node.children:
            if entry.local() != "sheet":
                continue
            name = entry.attrs.get("name")
            if name in renamed:
                # the entry still carries the ORIGINAL name bytes; the
                # state key (led re-keyed at rename time) is the NEW one
                edits.append(_patch_attr(original, entry, "name",
                                         renamed[name]))
            effective = renamed.get(name, name)
            if effective in state_changes:
                edits.append(_patch_attr(original, entry, "state",
                                         state_changes[effective],
                                         drop_value="visible"))
    if new_sheet_entries:
        payload = b"".join(
            _render_sheet_entry(title, sheet_id, rid, state)
            for title, sheet_id, rid, state in new_sheet_entries)
        edits.append(_insert_into(sheets_node, original, payload))

    if not edits:
        return None
    return apply_edits(original, edits)


def _rename_entry_name(entry_bytes, new_title):
    import re as _re

    return _re.sub(
        br'(\sname=")[^"]*(")',
        lambda m: m.group(1) + _escape(new_title) + m.group(2),
        entry_bytes, count=1)


def _repatch_entry_state(entry_bytes, new_state):
    import re as _re

    if b' state="' in entry_bytes:
        if new_state == "visible":
            return _re.sub(br'\sstate="[^"]*"', b"", entry_bytes, count=1)
        return _re.sub(br'(\sstate=")[^"]*(")',
                       lambda m: m.group(1) + _escape(new_state) + m.group(2),
                       entry_bytes, count=1)
    if new_state == "visible":
        return entry_bytes
    return entry_bytes.replace(
        b"<sheet ", b'<sheet state="%s" ' % _escape(new_state), 1)


def _render_sheet_entry(title, sheet_id, rid, state):
    # xmlns:r is declared on the element itself: producers (openpyxl
    # included) do not reliably declare it on the workbook root
    out = b'<sheet name="%s" sheetId="%d"' % (_escape(title), sheet_id)
    if state and state != "visible":
        out += b' state="%s"' % _escape(state)
    out += b' xmlns:r="%s" r:id="%s"/>' % (_REL_NS.encode("ascii"),
                                           _escape(rid))
    return out


def _wb_insert_edit(root, by_local, tag, payload):
    order = CT_WORKBOOK_INDEX[tag]
    for child in root.children:
        idx = CT_WORKBOOK_INDEX.get(child.local(), len(CT_WORKBOOK_INDEX))
        if idx > order:
            return (child.start, child.start, payload)
    return (root.content_end, root.content_end, payload)


def _patch_attr(data, node, attr, value, drop_value=None):
    """Rewrite one attribute inside a node's start tag."""
    head_end = node.content_start if not node.self_closing else node.end
    head = data[node.start:head_end]
    pattern = re.compile(
        br'\s%s\s*=\s*("[^"]*"|\'[^\']*\')' % re.escape(attr.encode("ascii")))
    if drop_value is not None and value == drop_value:
        new_head, _count = pattern.subn(b"", head)
    elif pattern.search(head):
        new_head = pattern.sub(
            b' %s="%s"' % (attr.encode("ascii"), _escape(value)), head, 1)
    else:
        cut = 2 if head.endswith(b"/>") else 1
        new_head = (head[:-cut]
                    + b' %s="%s"' % (attr.encode("ascii"), _escape(value))
                    + head[-cut:])
    return (node.start, head_end, new_head)


# ---------------------------------------------------------------------
# styles.xml append (PR-0 D2: append-only, never renumber)

# CT_Stylesheet child sequence
CT_STYLESHEET_ORDER = [
    "numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs",
    "cellStyles", "dxfs", "tableStyles", "colors", "extLst",
]
CT_STYLESHEET_INDEX = {tag: i for i, tag in enumerate(CT_STYLESHEET_ORDER)}


def _append_with_count(data, node, blob, added_count):
    """Append ``blob`` inside a container and bump its count attribute,
    as non-overlapping edits (single combined edit when self-closing)."""
    new_count = None
    if "count" in node.attrs:
        try:
            new_count = str(int(node.attrs["count"]) + added_count)
        except ValueError:
            new_count = None      # unparseable producer count: leave it
    if node.self_closing:
        head = data[node.start:node.end]
        assert head.endswith(b"/>")
        head = head[:-2] + b">"
        if new_count is not None:
            head = re.sub(
                br'\scount\s*=\s*("[^"]*"|\'[^\']*\')',
                b' count="%s"' % new_count.encode("ascii"), head, 1)
        return [(node.start, node.end,
                 head + blob + b"</" + node.name + b">")]
    edits = [(node.content_end, node.content_end, blob)]
    if new_count is not None:
        edits.append(_patch_attr(data, node, "count", new_count))
    return edits


def plan_styles_xml(wb, led, original, translator):
    """New styles.xml bytes appending the styles created since arming, or
    None when nothing was added. Never rewrites existing entries.

    Fonts/fills/borders/dxfs come from the MODEL tails (their numbering is
    file-stable at load: seeded in file order, never renumbered). Cell xfs
    and custom number formats come from the :class:`StyleTranslator`, which
    owns the model-to-file numbering translation (PR-0 D2)."""
    (n_fonts, n_fills, n_borders, _n_align, _n_prot,
     _n_numfmts) = led._style_lengths
    additions = {}

    new_fonts = list(wb._fonts)[n_fonts:]
    if new_fonts:
        additions["fonts"] = [tostring(f.to_tree()) for f in new_fonts]
    new_fills = list(wb._fills)[n_fills:]
    if new_fills:
        additions["fills"] = [tostring(f.to_tree()) for f in new_fills]
    new_borders = list(wb._borders)[n_borders:]
    if new_borders:
        additions["borders"] = [tostring(b.to_tree()) for b in new_borders]
    new_dxfs = list(wb._differential_styles.styles)[led.dxfs_len:]
    if new_dxfs:
        additions["dxfs"] = [tostring(d.to_tree(tagname="dxf"))
                             for d in new_dxfs]

    # order matters: xf rendering allocates any missing file numFmt ids
    new_xfs = translator.render_new_xfs()
    if new_xfs:
        additions["cellXfs"] = new_xfs
    new_fmts = translator.render_new_numfmts()
    if new_fmts:
        additions["numFmts"] = new_fmts

    if len(wb._named_styles) > led.named_styles_len:
        raise UnsupportedStructureError(
            "new named styles cannot be written under preserve mode in v0 "
            "(cellStyles/cellStyleXfs coordination); apply direct styles to "
            "cells instead. Nothing was written.")

    if not additions:
        return None

    root = scan_small(original, "styleSheet", max_depth=2)
    by_local = {}
    for child in root.children:
        by_local.setdefault(child.local(), []).append(child)

    edits = []
    for tag, payloads in additions.items():
        blob = b"".join(payloads)
        nodes = by_local.get(tag, [])
        if nodes:
            node = nodes[0]
            edits.extend(_append_with_count(original, node, blob,
                                            len(payloads)))
        else:
            container = (b"<%s count=\"%d\">" % (tag.encode("ascii"),
                                                 len(payloads))
                         + blob + b"</%s>" % tag.encode("ascii"))
            order = CT_STYLESHEET_INDEX[tag]
            for child in root.children:
                idx = CT_STYLESHEET_INDEX.get(child.local(),
                                              len(CT_STYLESHEET_INDEX))
                if idx > order:
                    edits.append((child.start, child.start, container))
                    break
            else:
                edits.append((root.content_end, root.content_end, container))

    return apply_edits(original, edits)
