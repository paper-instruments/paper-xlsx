# paper-xlsx: the worksheet byte scanner

"""Namespace-tracking streaming scanner over ORIGINAL worksheet XML bytes.

Produces the byte spans the splice writer needs — top-level region elements,
rows, cells — plus the shared-formula/array/metadata inventory that gates
edits. Everything it does not understand is left as bytes for the
splice to copy verbatim.

Guard set: DOCTYPE refused;
non-UTF-8 refused; the target grammar is matched only via the EXACT parent
chain worksheet→sheetData→row→c (ancestor containment admits legal decoys in
cell-level extLst and mc:AlternateContent — measured silent wrong-edits);
prefixed or non-main default namespaces refuse (the unguarded failure mode
is silent value deletion accepted by both loaders); r-less rows/cells
refuse. Every refusal happens before any output is written.
"""

import re

from openpyxl.errors import UnsupportedStructureError
from openpyxl.xml.constants import MAX_COLUMN, MAX_ROW, SHEET_MAIN_NS
from openpyxl.utils.cell import range_boundaries

_WS = b" \t\r\n"
_NAME_END = b" \t\r\n/>"

_ATTR_RE = re.compile(
    br'([^\s=/>]+)\s*=\s*("([^"]*)"|\'([^\']*)\')')

_ENCODING_RE = re.compile(br'encoding\s*=\s*["\']([^"\']+)["\']')


class ScanRefusal(UnsupportedStructureError):
    """The original sheet XML uses a construction the splice cannot edit
    safely; the save refuses before writing anything."""


class CellSpan:
    __slots__ = ("row", "column", "start", "end", "_attr_blob", "_attrs",
                 "shared_si", "shared_ref", "array_ref", "has_extlst",
                 "has_formula", "has_unowned_children")

    def __init__(self, row, column, start):
        self.row = row
        self.column = column
        self.start = start
        self.end = None
        self._attr_blob = b""
        self._attrs = None
        self.shared_si = None
        self.shared_ref = None
        self.array_ref = None
        self.has_extlst = False
        self.has_formula = False
        self.has_unowned_children = False

    @property
    def attrs(self):
        # decoded on demand: only the cells an edit actually touches pay
        # for attribute parsing (the splice reads a handful of dirty cells
        # out of the whole sheet)
        if self._attrs is None:
            attrs = {}
            for m in _ATTR_RE.finditer(self._attr_blob):
                key = m.group(1)
                if key == b"xmlns" or key.startswith(b"xmlns:"):
                    continue
                value = m.group(3) if m.group(3) is not None else m.group(4)
                attrs[key.decode("latin-1")] = value.decode("utf-8")
            self._attrs = attrs
        return self._attrs


class RowSpan:
    __slots__ = ("index", "start", "end", "content_start", "content_end",
                 "self_closing", "attrs", "cells")

    def __init__(self, index, start):
        self.index = index
        self.start = start
        self.end = None
        self.content_start = None    # just after the start tag's '>'
        self.content_end = None      # start of '</row>'
        self.self_closing = False
        self.attrs = {}
        self.cells = {}              # column -> CellSpan


class RegionSpan:
    __slots__ = ("tag", "start", "end", "raw")

    def __init__(self, tag, start):
        self.tag = tag
        self.start = start
        self.end = None
        self.raw = None


class SheetScan:
    """Result of scanning one worksheet part."""

    def __init__(self, data):
        self.data = data
        self.rows = {}               # row index -> RowSpan (document order kept separately)
        self.row_order = []          # row indices in document order
        self.regions = {}            # local tag -> [RegionSpan, ...] (top level)
        self.region_order = []       # (local tag, RegionSpan) in document order
        self.sheetdata = None        # RegionSpan for sheetData
        self.sheetdata_content = None  # (content_start, content_end) or None
        self.shared_groups = {}      # si -> ref string (from host cells)
        self.shared_members = {}     # si -> set[(row, col)] seen carrying it
        self.array_refs = []         # ref strings of t="array" formulas
        self.array_bounds = []       # (min_row, min_col, max_row, max_col)
        self.formula_names = {}      # (row, col) -> raw main-ns child names
        self.cache_names = {}        # (row, col) -> raw main-ns child names
        self.rows_monotonic = True
        self.root_end_offset = None  # offset of '</worksheet>'


def _range_bounds(ref):
    min_col, min_row, max_col, max_row = range_boundaries(ref)
    return (
        1 if min_row is None else min_row,
        1 if min_col is None else min_col,
        MAX_ROW if max_row is None else max_row,
        MAX_COLUMN if max_col is None else max_col,
    )


def _decode_name(raw, default_ns, prefixes, what, offset):
    if b":" in raw:
        prefix, local = raw.split(b":", 1)
        ns = prefixes.get(prefix)
        if ns is None:
            raise ScanRefusal(
                "cannot splice: {0} uses undeclared prefix {1!r} at byte "
                "{2}".format(what, prefix.decode("latin-1"), offset))
        return ns, local
    return default_ns, raw


def scan_sheet(data):
    """Scan one worksheet part's bytes into a :class:`SheetScan`.

    Raises :class:`ScanRefusal` for constructions the splice must not touch.
    """
    scan = SheetScan(data)
    main = SHEET_MAIN_NS.encode("ascii")

    pos = 0
    if data[:3] == b"\xef\xbb\xbf":
        pos = 3

    # stack entries: [clark_ns, local, default_ns, prefixes, start_offset]
    stack = []
    current_row = None
    current_cell = None

    n = len(data)
    while pos < n:
        lt = data.find(b"<", pos)
        if lt == -1:
            break
        nxt = data[lt + 1]

        # fast paths for the sheetData hot loop. Guard-equivalent by
        # construction: anything but the three plain shapes — '<c' with an
        # r attribute and no namespace declarations, text-only '<v>',
        # an exact '</c>' closing the open cell — falls through to the
        # generic machinery below, refusals and all.
        if current_cell is not None:
            if nxt == 0x76 and lt + 2 < n and data[lt + 2] == 0x3E:  # <v>
                close = data.find(b"<", lt + 3)
                if close != -1 and data.startswith(b"</v>", close):
                    # text-only content: a raw '<' cannot appear in
                    # well-formed character data, so this IS the end tag;
                    # CDATA/comments/nested markup miss the check and take
                    # the generic path
                    pos = close + 4
                    continue
            elif nxt == 0x2F and data.startswith(b"</c>", lt) \
                    and stack[-1][5] is current_cell:
                stack.pop()
                current_cell.end = lt + 4
                current_cell = None
                pos = lt + 4
                continue
        elif (nxt == 0x63 and current_row is not None  # '<c'
                and lt + 2 < n and data[lt + 2] in b" \t\r\n/>"
                and stack[-1][5] is current_row
                and stack[-1][2] == main):
            gt = _find_tag_end(data, lt)
            self_closing = data[gt - 1] == 0x2F
            attr_blob = data[lt + 2: gt - 1 if self_closing else gt]
            if b"xmlns" not in attr_blob:
                # r extraction MUST tokenize the whole blob (quoted values
                # consumed in order, last r wins) — a bare ' r="..."'
                # pattern search would match decoys inside OTHER
                # attributes' values, on files openpyxl loads fine
                rb = None
                for m in _ATTR_RE.finditer(attr_blob):
                    if m.group(1) == b"r":
                        rb = m.group(3) if m.group(3) is not None \
                            else m.group(4)
                if rb is None:
                    raise ScanRefusal(
                        "cannot splice: a <c> element carries no r "
                        "attribute (implicit cell numbering); editing such "
                        "sheets is not supported in v0")
                coord = rb.decode("ascii")
                col = 0
                i = 0
                n_rb = len(rb)
                while i < n_rb:
                    b0 = rb[i]
                    if 65 <= b0 <= 90:
                        col = col * 26 + (b0 - 64)
                    elif 97 <= b0 <= 122:
                        col = col * 26 + (b0 - 96)
                    else:
                        break
                    i += 1
                if col == 0:
                    raise ScanRefusal(
                        "cannot splice: malformed cell reference "
                        "{0!r}".format(coord))
                while i < n_rb and (rb[i] == 36 or 65 <= rb[i] <= 90
                                    or 97 <= rb[i] <= 122):
                    i += 1
                digits = coord[i:]
                if digits and int(digits) != current_row.index:
                    raise ScanRefusal(
                        "cannot splice: cell {0!r} sits inside row {1} "
                        "(its own reference disagrees with its parent "
                        "row)".format(coord, current_row.index))
                cell = CellSpan(current_row.index, col, lt)
                cell._attr_blob = attr_blob
                if self_closing:
                    cell.end = gt + 1
                else:
                    current_cell = cell
                    stack.append([main, b"c", stack[-1][2], stack[-1][3],
                                  lt, cell])
                current_row.cells[col] = cell
                pos = gt + 1
                continue

        if nxt == 0x3F:  # '?': processing instruction
            end = data.find(b"?>", lt)
            if end == -1:
                raise ScanRefusal("cannot splice: unterminated processing "
                                  "instruction at byte {0}".format(lt))
            if current_cell is not None:
                current_cell.has_unowned_children = True
            if data[lt:lt + 5] == b"<?xml":
                m = _ENCODING_RE.search(data[lt:end])
                if m and m.group(1).lower().replace(b"_", b"-") not in (
                        b"utf-8", b"utf8"):
                    raise ScanRefusal(
                        "cannot splice: declared encoding {0!r} is not "
                        "UTF-8".format(m.group(1).decode("latin-1")))
            pos = end + 2
            continue
        if nxt == 0x21:  # '!': comment / CDATA / DOCTYPE
            if data.startswith(b"<!--", lt):
                end = data.find(b"-->", lt)
                if end == -1:
                    raise ScanRefusal("cannot splice: unterminated comment")
                if current_cell is not None:
                    current_cell.has_unowned_children = True
                pos = end + 3
                continue
            if data.startswith(b"<![CDATA[", lt):
                end = data.find(b"]]>", lt)
                if end == -1:
                    raise ScanRefusal("cannot splice: unterminated CDATA")
                pos = end + 3
                continue
            if data.startswith(b"<!DOCTYPE", lt):
                raise ScanRefusal(
                    "cannot splice: the sheet XML carries a DOCTYPE "
                    "declaration")
            raise ScanRefusal(
                "cannot splice: unrecognized markup at byte {0}".format(lt))
        if nxt == 0x2F:  # '/': end tag
            gt = data.find(b">", lt)
            if gt == -1:
                raise ScanRefusal("cannot splice: unterminated end tag")
            if not stack:
                raise ScanRefusal("cannot splice: unbalanced end tag at "
                                  "byte {0}".format(lt))
            entry = stack.pop()
            _close_element(scan, entry, lt, gt + 1, len(stack))
            if current_cell is not None and entry[5] is current_cell:
                current_cell.end = gt + 1
                current_cell = None
            elif current_row is not None and entry[5] is current_row:
                current_row.end = gt + 1
                current_row.content_end = lt
                current_row = None
            pos = gt + 1
            continue

        # start tag
        gt = _find_tag_end(data, lt)
        self_closing = data[gt - 1:gt] == b"/"
        tag_end = gt + 1
        head = data[lt + 1: gt - 1 if self_closing else gt]
        name_end = _scan_name_end(head)
        raw_name = head[:name_end]
        attr_blob = head[name_end:]

        parent_default = stack[-1][2] if stack else None
        parent_prefixes = stack[-1][3] if stack else {}
        default_ns = parent_default
        prefixes = parent_prefixes

        # fast path: elements the splice never inspects (cell values,
        # inline strings, extension payloads...) skip attribute parsing
        # entirely — unless they declare namespaces, which must be tracked
        depth_now = len(stack)
        needs_attrs = (
            depth_now <= 1
            or (depth_now == 2 and raw_name == b"row")
            or (depth_now == 3 and raw_name == b"c")
            or depth_now == 4
            or b"xmlns" in attr_blob
        )
        if not needs_attrs:
            if not self_closing:
                stack.append([None, raw_name, default_ns, prefixes, lt, None])
            pos = tag_end
            continue

        attrs = {}
        ns_declared = False
        for m in _ATTR_RE.finditer(attr_blob):
            key = m.group(1)
            value = m.group(3) if m.group(3) is not None else m.group(4)
            if key == b"xmlns":
                if not ns_declared:
                    prefixes = dict(parent_prefixes)
                    ns_declared = True
                default_ns = value
            elif key.startswith(b"xmlns:"):
                if not ns_declared:
                    prefixes = dict(parent_prefixes)
                    ns_declared = True
                prefixes[key[6:]] = value
            else:
                attrs[key] = value

        ns, local = _decode_name(raw_name, default_ns, prefixes,
                                 "element", lt)
        depth = len(stack)

        if depth == 0:
            # root guards: must be an unprefixed main-namespace
            # worksheet, else fragments we emit would land in no namespace —
            # the measured silent-value-deletion failure mode
            if local != b"worksheet" or ns != main:
                raise ScanRefusal(
                    "cannot splice: root element is not a spreadsheetml "
                    "worksheet (found {0!r} in namespace {1!r})".format(
                        raw_name.decode("latin-1"),
                        (ns or b"").decode("latin-1")))
            if b":" in raw_name or default_ns != main:
                raise ScanRefusal(
                    "cannot splice: the worksheet uses a prefixed or "
                    "non-default main namespace; editing it byte-wise would "
                    "silently detach edited cells from the schema. Reopen "
                    "without preserve=True to rewrite the sheet lossily.")

        obj = None
        if depth == 1 and ns == main:
            span = RegionSpan(local.decode("ascii"), lt)
            if self_closing:
                # never reaches _close_element (no stack entry): the span
                # must close here or region edits splice with end=None —
                # the whole-document-duplication corruption
                span.end = tag_end
            scan.regions.setdefault(span.tag, []).append(span)
            scan.region_order.append((span.tag, span))
            if span.tag == "sheetData":
                scan.sheetdata = span
                if not self_closing:
                    scan.sheetdata_content = [tag_end, None]
            obj = span
        elif depth == 2 and ns == main and local == b"row" \
                and stack[-1][1] == b"sheetData" and stack[-1][0] == main:
            r_attr = attrs.get(b"r")
            if r_attr is None:
                raise ScanRefusal(
                    "cannot splice: a <row> element carries no r attribute "
                    "(implicit row numbering); editing such sheets is not "
                    "supported in v0")
            index = int(r_attr)
            if scan.row_order and index <= scan.row_order[-1]:
                scan.rows_monotonic = False
            row = RowSpan(index, lt)
            row.attrs = {k.decode("latin-1"): v.decode("utf-8")
                         for k, v in attrs.items()}
            if self_closing:
                row.end = tag_end
                row.self_closing = True
            else:
                row.content_start = tag_end
                current_row = row
            scan.rows[index] = row
            scan.row_order.append(index)
            obj = row
        elif depth == 3 and ns == main and local == b"c" \
                and isinstance(current_row, RowSpan) \
                and stack[-1][5] is current_row:
            # EXACT parent chain worksheet→sheetData→row→c: decoys inside
            # extLst / AlternateContent never reach depth 3 with a row parent
            r_attr = attrs.get(b"r")
            if r_attr is None:
                raise ScanRefusal(
                    "cannot splice: a <c> element carries no r attribute "
                    "(implicit cell numbering); editing such sheets is not "
                    "supported in v0")
            coord = r_attr.decode("ascii")
            col = _column_index(coord)
            digits = coord.lstrip(
                "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz$")
            if digits and int(digits) != current_row.index:
                # openpyxl's reader places the cell by ITS r while the splice
                # keys spans by the parent row: a mismatch (off-spec but
                # loadable) would make an edit insert a duplicate reference
                raise ScanRefusal(
                    "cannot splice: cell {0!r} sits inside row {1} (its own "
                    "reference disagrees with its parent row)".format(
                        coord, current_row.index))
            cell = CellSpan(current_row.index, col, lt)
            cell._attrs = {k.decode("latin-1"): v.decode("utf-8")
                           for k, v in attrs.items()}
            if self_closing:
                cell.end = tag_end
            current_row.cells[col] = cell
            if not self_closing:
                current_cell = cell
            obj = cell
        elif current_cell is not None and depth == 4:
            if ns != main or local not in (b"f", b"v", b"is", b"extLst"):
                current_cell.has_unowned_children = True
            elif local == b"f":
                current_cell.has_formula = True
                if raw_name != b"f":
                    coordinate = (current_cell.row, current_cell.column)
                    names = scan.formula_names.setdefault(coordinate, ())
                    if raw_name not in names:
                        scan.formula_names[coordinate] = names + (raw_name,)
                t = attrs.get(b"t")
                si = attrs.get(b"si")
                ref = attrs.get(b"ref")
                if t == b"shared" and si is not None:
                    si = si.decode("ascii")
                    current_cell.shared_si = si
                    scan.shared_members.setdefault(si, set()).add(
                        (current_cell.row, current_cell.column))
                    if ref is not None:
                        current_cell.shared_ref = ref.decode("ascii")
                        scan.shared_groups[si] = ref.decode("ascii")
                elif t == b"array" and ref is not None:
                    ref_text = ref.decode("ascii")
                    current_cell.array_ref = ref_text
                    scan.array_refs.append(ref_text)
                    scan.array_bounds.append(_range_bounds(ref_text))
            elif local == b"v":
                if raw_name != b"v":
                    coordinate = (current_cell.row, current_cell.column)
                    names = scan.cache_names.setdefault(coordinate, ())
                    if raw_name not in names:
                        scan.cache_names[coordinate] = names + (raw_name,)
            elif local == b"extLst":
                current_cell.has_extlst = True

        if self_closing:
            pos = tag_end
            continue
        stack.append([ns, local, default_ns, prefixes, lt, obj])
        pos = tag_end

    if stack:
        raise ScanRefusal("cannot splice: document ended with unclosed "
                          "elements")
    if scan.sheetdata is None:
        raise ScanRefusal("cannot splice: the worksheet has no sheetData "
                          "element")
    return scan


def _close_element(scan, entry, lt, end, depth_after):
    obj = entry[5]
    if isinstance(obj, RegionSpan):
        obj.end = end
        if obj.tag == "sheetData" and scan.sheetdata_content is not None:
            scan.sheetdata_content[1] = lt
    if depth_after == 0:
        scan.root_end_offset = lt


def _find_tag_end(data, lt):
    """Offset of the '>' closing a start tag, honouring quoted attributes."""
    gt = data.find(b">", lt)
    if gt == -1:
        raise ScanRefusal("cannot splice: unterminated start tag at byte "
                          "{0}".format(lt))
    seg = data[lt:gt]
    # fast path: balanced quotes before '>' mean it is a real tag end
    if seg.count(b'"') % 2 == 0 and seg.count(b"'") % 2 == 0:
        return gt
    # slow path: a '>' sits inside a quoted attribute value
    pos = lt + 1
    n = len(data)
    quote = None
    while pos < n:
        ch = data[pos:pos + 1]
        if quote is not None:
            if ch == quote:
                quote = None
        elif ch in (b'"', b"'"):
            quote = ch
        elif ch == b">":
            return pos
        pos += 1
    raise ScanRefusal("cannot splice: unterminated start tag at byte "
                      "{0}".format(lt))


def _scan_name_end(head):
    for i, byte in enumerate(head):
        if byte in _NAME_END:
            return i
    return len(head)


def _column_index(coord):
    """Column index from an A1 coordinate ('BC12' -> 55)."""
    col = 0
    for ch in coord:
        if ch.isalpha():
            col = col * 26 + (ord(ch.upper()) - 64)
        else:
            break
    if col == 0:
        raise ScanRefusal(
            "cannot splice: malformed cell reference {0!r}".format(coord))
    return col
