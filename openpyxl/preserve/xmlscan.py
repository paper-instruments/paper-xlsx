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


class _SheetScanner:
    """Stateful tokenizer whose methods each own one XML grammar phase."""

    def __init__(self, data):
        self.data = data
        self.scan = SheetScan(data)
        self.main = SHEET_MAIN_NS.encode("ascii")
        self.pos = 3 if data[:3] == b"\xef\xbb\xbf" else 0
        # Entries are [namespace, local name, default namespace, prefixes,
        # start offset, modeled span object].
        self.stack = []
        self.current_row = None
        self.current_cell = None

    def run(self):
        # Keep the worksheet cell loop on locals. On CPython 3.9 and 3.10,
        # routing common cell shapes through bound methods adds material
        # overhead on large sheets. Uncommon forms still delegate below.
        data = self.data
        scan = self.scan
        main = self.main
        stack = self.stack
        pos = self.pos
        current_row = self.current_row
        current_cell = self.current_cell
        n = len(data)
        while pos < n:
            lt = data.find(b"<", pos)
            if lt == -1:
                break
            nxt = data[lt + 1]
            if current_cell is not None:
                if nxt == 0x76 and lt + 2 < n \
                        and data[lt + 2] == 0x3E:
                    close = data.find(b"<", lt + 3)
                    if close != -1 and data.startswith(b"</v>", close):
                        pos = close + 4
                        continue
                elif nxt == 0x66 and lt + 2 < n \
                        and data[lt + 2] == 0x3E:
                    close = data.find(b"<", lt + 3)
                    if close != -1 and data.startswith(b"</f>", close):
                        current_cell.has_formula = True
                        pos = close + 4
                        continue
                elif nxt == 0x69 and data.startswith(b"<is><t>", lt):
                    close = data.find(b"<", lt + 7)
                    if close != -1 and data.startswith(b"</t></is>", close):
                        pos = close + 9
                        continue
                elif nxt == 0x2F and data.startswith(b"</c>", lt) \
                        and stack[-1][5] is current_cell:
                    stack.pop()
                    current_cell.end = lt + 4
                    current_cell = None
                    pos = lt + 4
                    continue
            elif current_row is not None and nxt == 0x2F \
                    and data.startswith(b"</row>", lt) \
                    and stack[-1][5] is current_row:
                stack.pop()
                current_row.end = lt + 6
                current_row.content_end = lt
                current_row = None
                pos = lt + 6
                continue
            elif current_row is None and nxt == 0x72 \
                    and data.startswith(b"<row", lt) \
                    and lt + 4 < n and data[lt + 4] in b" \t\r\n/>" \
                    and stack \
                    and stack[-1][5] is scan.sheetdata \
                    and stack[-1][2] == main:
                gt = _find_tag_end(data, lt)
                self_closing = data[gt - 1] == 0x2F
                attr_blob = data[
                    lt + 4:gt - 1 if self_closing else gt]
                if b"xmlns" not in attr_blob:
                    attrs = {}
                    for match in _ATTR_RE.finditer(attr_blob):
                        value = match.group(3) \
                            if match.group(3) is not None \
                            else match.group(4)
                        attrs[match.group(1)] = value
                    reference = attrs.get(b"r")
                    if reference is None:
                        raise ScanRefusal(
                            "cannot splice: a <row> element carries no r "
                            "attribute (implicit row numbering); editing "
                            "such sheets is not supported in v0")
                    row_index = int(reference)
                    if scan.row_order \
                            and row_index <= scan.row_order[-1]:
                        scan.rows_monotonic = False
                    row = RowSpan(row_index, lt)
                    row.attrs = {
                        key.decode("latin-1"): value.decode("utf-8")
                        for key, value in attrs.items()}
                    if self_closing:
                        row.end = gt + 1
                        row.self_closing = True
                    else:
                        row.content_start = gt + 1
                        current_row = row
                        stack.append([
                            main, b"row", stack[-1][2], stack[-1][3], lt,
                            row])
                    scan.rows[row_index] = row
                    scan.row_order.append(row_index)
                    pos = gt + 1
                    continue
            elif (nxt == 0x63 and current_row is not None
                    and lt + 2 < n and data[lt + 2] in b" \t\r\n/>"
                    and stack[-1][5] is current_row
                    and stack[-1][2] == main):
                gt = _find_tag_end(data, lt)
                self_closing = data[gt - 1] == 0x2F
                attr_blob = data[
                    lt + 2:gt - 1 if self_closing else gt]
                if b"xmlns" not in attr_blob:
                    reference = None
                    for match in _ATTR_RE.finditer(attr_blob):
                        if match.group(1) == b"r":
                            reference = match.group(3) \
                                if match.group(3) is not None \
                                else match.group(4)
                    if reference is None:
                        raise ScanRefusal(
                            "cannot splice: a <c> element carries no r "
                            "attribute (implicit cell numbering); editing "
                            "such sheets is not supported in v0")
                    coordinate = reference.decode("ascii")
                    column = 0
                    index = 0
                    while index < len(reference):
                        value = reference[index]
                        if 65 <= value <= 90:
                            column = column * 26 + value - 64
                        elif 97 <= value <= 122:
                            column = column * 26 + value - 96
                        else:
                            break
                        index += 1
                    if column == 0:
                        raise ScanRefusal(
                            "cannot splice: malformed cell reference "
                            "{0!r}".format(coordinate))
                    while index < len(reference) and (
                            reference[index] == 36
                            or 65 <= reference[index] <= 90
                            or 97 <= reference[index] <= 122):
                        index += 1
                    row_digits = coordinate[index:]
                    if row_digits and int(row_digits) != current_row.index:
                        raise ScanRefusal(
                            "cannot splice: cell {0!r} sits inside row {1} "
                            "(its own reference disagrees with its parent "
                            "row)".format(coordinate, current_row.index))
                    cell = CellSpan(current_row.index, column, lt)
                    cell._attr_blob = attr_blob
                    if self_closing:
                        cell.end = gt + 1
                    else:
                        current_cell = cell
                        stack.append([
                            main, b"c", stack[-1][2], stack[-1][3], lt,
                            cell])
                    current_row.cells[column] = cell
                    pos = gt + 1
                    continue

            self.pos = pos
            self.current_row = current_row
            self.current_cell = current_cell
            if nxt == 0x3F:
                self._scan_processing_instruction(lt)
            elif nxt == 0x21:
                self._scan_declaration(lt)
            elif nxt == 0x2F:
                self._scan_end_tag(lt)
            else:
                self._scan_start_tag(lt)
            pos = self.pos
            current_row = self.current_row
            current_cell = self.current_cell

        self.pos = pos
        self.current_row = current_row
        self.current_cell = current_cell
        if stack:
            raise ScanRefusal("cannot splice: document ended with unclosed "
                              "elements")
        if scan.sheetdata is None:
            raise ScanRefusal(
                "cannot splice: the worksheet has no sheetData element")
        return scan

    def _scan_processing_instruction(self, lt):
        end = self.data.find(b"?>", lt)
        if end == -1:
            raise ScanRefusal(
                "cannot splice: unterminated processing instruction at "
                "byte {0}".format(lt))
        if self.current_cell is not None:
            self.current_cell.has_unowned_children = True
        if self.data[lt:lt + 5] == b"<?xml":
            match = _ENCODING_RE.search(self.data[lt:end])
            if match and match.group(1).lower().replace(b"_", b"-") not in (
                    b"utf-8", b"utf8"):
                raise ScanRefusal(
                    "cannot splice: declared encoding {0!r} is not "
                    "UTF-8".format(match.group(1).decode("latin-1")))
        self.pos = end + 2

    def _scan_declaration(self, lt):
        if self.data.startswith(b"<!--", lt):
            end = self.data.find(b"-->", lt)
            if end == -1:
                raise ScanRefusal("cannot splice: unterminated comment")
            if self.current_cell is not None:
                self.current_cell.has_unowned_children = True
            self.pos = end + 3
            return
        if self.data.startswith(b"<![CDATA[", lt):
            end = self.data.find(b"]]>", lt)
            if end == -1:
                raise ScanRefusal("cannot splice: unterminated CDATA")
            self.pos = end + 3
            return
        if self.data.startswith(b"<!DOCTYPE", lt):
            raise ScanRefusal(
                "cannot splice: the sheet XML carries a DOCTYPE declaration")
        raise ScanRefusal(
            "cannot splice: unrecognized markup at byte {0}".format(lt))

    def _scan_end_tag(self, lt):
        gt = self.data.find(b">", lt)
        if gt == -1:
            raise ScanRefusal("cannot splice: unterminated end tag")
        if not self.stack:
            raise ScanRefusal(
                "cannot splice: unbalanced end tag at byte {0}".format(lt))
        entry = self.stack.pop()
        _close_element(self.scan, entry, lt, gt + 1, len(self.stack))
        if self.current_cell is not None and entry[5] is self.current_cell:
            self.current_cell.end = gt + 1
            self.current_cell = None
        elif self.current_row is not None and entry[5] is self.current_row:
            self.current_row.end = gt + 1
            self.current_row.content_end = lt
            self.current_row = None
        self.pos = gt + 1

    def _scan_start_tag(self, lt):
        parsed = self._parse_start_tag(lt)
        if parsed is None:
            return
        (tag_end, self_closing, raw_name, attrs, namespace, local,
         default_ns, prefixes) = parsed
        depth = len(self.stack)
        self._validate_root(
            depth, raw_name, namespace, local, default_ns)
        obj = self._open_modeled_element(
            depth, raw_name, attrs, namespace, local, lt, tag_end,
            self_closing)
        if not self_closing:
            self.stack.append([
                namespace, local, default_ns, prefixes, lt, obj])
        self.pos = tag_end

    def _parse_start_tag(self, lt):
        gt = _find_tag_end(self.data, lt)
        self_closing = self.data[gt - 1:gt] == b"/"
        tag_end = gt + 1
        head = self.data[lt + 1:gt - 1 if self_closing else gt]
        name_end = _scan_name_end(head)
        raw_name = head[:name_end]
        attr_blob = head[name_end:]
        parent_default = self.stack[-1][2] if self.stack else None
        parent_prefixes = self.stack[-1][3] if self.stack else {}
        depth = len(self.stack)
        needs_attrs = (
            depth <= 1
            or (depth == 2 and raw_name == b"row")
            or (depth == 3 and raw_name == b"c")
            or depth == 4
            or b"xmlns" in attr_blob
        )
        if not needs_attrs:
            if not self_closing:
                self.stack.append([
                    None, raw_name, parent_default, parent_prefixes, lt, None])
            self.pos = tag_end
            return None
        attrs, default_ns, prefixes = _parse_attributes(
            attr_blob, parent_default, parent_prefixes)
        namespace, local = _decode_name(
            raw_name, default_ns, prefixes, "element", lt)
        return (tag_end, self_closing, raw_name, attrs, namespace, local,
                default_ns, prefixes)

    def _validate_root(self, depth, raw_name, namespace, local, default_ns):
        if depth != 0:
            return
        if local != b"worksheet" or namespace != self.main:
            raise ScanRefusal(
                "cannot splice: root element is not a spreadsheetml "
                "worksheet (found {0!r} in namespace {1!r})".format(
                    raw_name.decode("latin-1"),
                    (namespace or b"").decode("latin-1")))
        if b":" in raw_name or default_ns != self.main:
            raise ScanRefusal(
                "cannot splice: the worksheet uses a prefixed or non-default "
                "main namespace; editing it byte-wise would silently detach "
                "edited cells from the schema. Reopen without preserve=True "
                "to rewrite the sheet lossily.")

    def _open_modeled_element(self, depth, raw_name, attrs, namespace, local,
                              lt, tag_end, self_closing):
        if depth == 1 and namespace == self.main:
            return self._open_region(local, lt, tag_end, self_closing)
        if depth == 2 and namespace == self.main and local == b"row" \
                and self.stack[-1][1] == b"sheetData" \
                and self.stack[-1][0] == self.main:
            return self._open_row(attrs, lt, tag_end, self_closing)
        if depth == 3 and namespace == self.main and local == b"c" \
                and isinstance(self.current_row, RowSpan) \
                and self.stack[-1][5] is self.current_row:
            return self._open_cell(attrs, lt, tag_end, self_closing)
        if self.current_cell is not None and depth == 4:
            self._record_cell_child(raw_name, attrs, namespace, local)
        return None

    def _open_region(self, local, lt, tag_end, self_closing):
        span = RegionSpan(local.decode("ascii"), lt)
        if self_closing:
            span.end = tag_end
        self.scan.regions.setdefault(span.tag, []).append(span)
        self.scan.region_order.append((span.tag, span))
        if span.tag == "sheetData":
            self.scan.sheetdata = span
            if not self_closing:
                self.scan.sheetdata_content = [tag_end, None]
        return span

    def _open_row(self, attrs, lt, tag_end, self_closing):
        reference = attrs.get(b"r")
        if reference is None:
            raise ScanRefusal(
                "cannot splice: a <row> element carries no r attribute "
                "(implicit row numbering); editing such sheets is not "
                "supported in v0")
        index = int(reference)
        if self.scan.row_order and index <= self.scan.row_order[-1]:
            self.scan.rows_monotonic = False
        row = RowSpan(index, lt)
        row.attrs = {key.decode("latin-1"): value.decode("utf-8")
                     for key, value in attrs.items()}
        if self_closing:
            row.end = tag_end
            row.self_closing = True
        else:
            row.content_start = tag_end
            self.current_row = row
        self.scan.rows[index] = row
        self.scan.row_order.append(index)
        return row

    def _open_cell(self, attrs, lt, tag_end, self_closing):
        reference = attrs.get(b"r")
        if reference is None:
            raise ScanRefusal(
                "cannot splice: a <c> element carries no r attribute "
                "(implicit cell numbering); editing such sheets is not "
                "supported in v0")
        coordinate = reference.decode("ascii")
        column = _column_index(coordinate)
        digits = coordinate.lstrip(
            "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz$")
        if digits and int(digits) != self.current_row.index:
            raise ScanRefusal(
                "cannot splice: cell {0!r} sits inside row {1} (its own "
                "reference disagrees with its parent row)".format(
                    coordinate, self.current_row.index))
        cell = CellSpan(self.current_row.index, column, lt)
        cell._attrs = {key.decode("latin-1"): value.decode("utf-8")
                       for key, value in attrs.items()}
        if self_closing:
            cell.end = tag_end
        else:
            self.current_cell = cell
        self.current_row.cells[column] = cell
        return cell

    def _record_cell_child(self, raw_name, attrs, namespace, local):
        cell = self.current_cell
        if namespace != self.main or local not in (b"f", b"v", b"is", b"extLst"):
            cell.has_unowned_children = True
            return
        if local == b"f":
            cell.has_formula = True
            if raw_name != b"f":
                coordinate = (cell.row, cell.column)
                names = self.scan.formula_names.setdefault(coordinate, ())
                if raw_name not in names:
                    self.scan.formula_names[coordinate] = names + (raw_name,)
            formula_type = attrs.get(b"t")
            shared_index = attrs.get(b"si")
            reference = attrs.get(b"ref")
            if formula_type == b"shared" and shared_index is not None:
                shared_index = shared_index.decode("ascii")
                cell.shared_si = shared_index
                self.scan.shared_members.setdefault(shared_index, set()).add(
                    (cell.row, cell.column))
                if reference is not None:
                    cell.shared_ref = reference.decode("ascii")
                    self.scan.shared_groups[shared_index] = \
                        reference.decode("ascii")
            elif formula_type == b"array" and reference is not None:
                reference = reference.decode("ascii")
                cell.array_ref = reference
                self.scan.array_refs.append(reference)
                self.scan.array_bounds.append(_range_bounds(reference))
        elif local == b"v" and raw_name != b"v":
            coordinate = (cell.row, cell.column)
            names = self.scan.cache_names.setdefault(coordinate, ())
            if raw_name not in names:
                self.scan.cache_names[coordinate] = names + (raw_name,)
        elif local == b"extLst":
            cell.has_extlst = True


def _parse_attributes(attr_blob, parent_default, parent_prefixes):
    attrs = {}
    default_ns = parent_default
    prefixes = parent_prefixes
    ns_declared = False
    for match in _ATTR_RE.finditer(attr_blob):
        key = match.group(1)
        value = match.group(3) if match.group(3) is not None \
            else match.group(4)
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
    return attrs, default_ns, prefixes


def scan_sheet(data):
    """Scan one worksheet part's bytes into a `SheetScan`.

    Raises `ScanRefusal` for constructions the splice must not touch.
    """
    return _SheetScanner(data).run()


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
