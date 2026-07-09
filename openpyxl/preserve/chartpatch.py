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


def plan_chart_updates(wb, sheet_title, operation, index, amount,
                       overrides=None):
    """Plan every chart/drawing part patch for a shift on ``sheet_title``.

    Returns ({part_name: new_payload}, blockers). Used twice: as a dry run
    when the edit is attempted (blockers refuse before any mutation) and for
    real at save time. ``overrides`` supplies already-planned payloads so a
    part touched by an earlier shift is patched incrementally.
    """
    from .structural import _charts_referencing

    source = getattr(wb, "_paper_source", None)
    overrides = overrides or {}
    plans = {}
    blockers = []
    if not source:
        return plans, blockers
    chart_parts = _charts_referencing(wb, sheet_title)
    with zipfile.ZipFile(io.BytesIO(source)) as z:
        for part in chart_parts:
            payload = overrides.get(part) or z.read(part)
            new_payload, changed, part_blockers = patch_chart(
                payload, sheet_title, operation, index, amount)
            for blocker in part_blockers:
                blockers.append("{0}: {1}".format(part, blocker))
            if changed and not part_blockers:
                plans[part] = new_payload
        drawing_part = _sheet_drawing_part(z, sheet_title)
        if drawing_part is not None:
            payload = overrides.get(drawing_part) or z.read(drawing_part)
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


def patch_chart_renames(payload, mapping):
    """Rewrite chart formula texts through a SIMULTANEOUS title mapping
    (title swaps must never merge reference classes — Batch-3 gate).
    Returns the patched payload or None when nothing matched."""
    from openpyxl.errors import UnsupportedStructureError

    from .rewrite import rename_sheets_in_formula

    hit = False
    edits = []
    for ns, local, _parent, t_start, t_end in _walk_leaf_texts(payload):
        if local != b"f" or ns != CHART_NS:
            continue
        raw = payload[t_start:t_end]
        text = _unescape(raw).decode("utf-8")
        new_text, changed = rename_sheets_in_formula("=" + text, mapping)
        if not changed:
            continue
        hit = True
        edits.append((t_start, t_end,
                      _escape(new_text[1:].encode("utf-8"))))
    if not hit:
        return None
    for marker, label in ((b"c15:", "c15 filtered-series machinery"),
                          (b"AlternateContent", "alternate-content blocks")):
        if marker in payload:
            raise UnsupportedStructureError(
                "renaming sheets: a chart referencing them carries {0}, "
                "whose references the rename patch cannot see. Nothing "
                "was written.".format(label))
    from .crosspart import apply_edits

    return apply_edits(payload, edits)


def patch_chart_rename(payload, old_title, new_title):
    """Rewrite every chart formula text (<c:f>) referencing ``old_title``
    to ``new_title`` (PLAN-v0.1 3.2). Returns the patched payload, or None
    when nothing referenced the old title. Refuses when the chart carries
    machinery whose references the patch cannot see (the same blocker set
    as shift patching)."""
    from openpyxl.errors import UnsupportedStructureError

    from .rewrite import rename_sheet_in_formula

    hit = False
    edits = []
    for ns, local, _parent, t_start, t_end in _walk_leaf_texts(payload):
        if local != b"f" or ns != CHART_NS:
            continue
        raw = payload[t_start:t_end]
        text = _unescape(raw).decode("utf-8")
        new_text, changed = rename_sheet_in_formula(
            "=" + text, old_title, new_title)
        if not changed:
            continue
        hit = True
        edits.append((t_start, t_end,
                      _escape(new_text[1:].encode("utf-8"))))
    if not hit:
        return None
    for marker, label in ((b"c15:", "c15 filtered-series machinery"),
                          (b"AlternateContent", "alternate-content blocks")):
        if marker in payload:
            raise UnsupportedStructureError(
                "renaming sheet {0!r}: a chart referencing it carries "
                "{1}, whose references the rename patch cannot see. "
                "Nothing was written.".format(old_title, label))
    from .crosspart import apply_edits

    return apply_edits(payload, edits)


# ---------------------------------------------------------------------
# Batch 4 (PR-1 §3): per-property chart edits expressed as byte patches

DRAWING_MAIN_NS = b"http://schemas.openxmlformats.org/drawingml/2006/main"


def _text_sequences(data):
    """([(start, end, text)] for <c:f> leaves, same for <a:t> leaves), in
    document order — the two property families chartpatch can express."""
    fs, ts = [], []
    for ns, local, _parent, start, end in _walk_leaf_texts(data):
        if local == b"f" and ns == CHART_NS:
            fs.append((start, end, data[start:end]))
        elif local == b"t" and ns == DRAWING_MAIN_NS:
            ts.append((start, end, data[start:end]))
    return fs, ts


def _neutralized(data, fs, ts):
    """The document with every expressible text span replaced by a
    placeholder — what remains is the INexpressible surface."""
    spans = sorted([(s, e) for s, e, _ in fs] + [(s, e) for s, e, _ in ts],
                   reverse=True)
    for start, end in spans:
        data = data[:start] + b"#" + data[end:]
    return data


def _property_near(neutral_armed, neutral_current):
    """A best-effort name for the first differing property between two
    neutralized renders (for the refusal message)."""
    limit = min(len(neutral_armed), len(neutral_current))
    diff = limit
    for i in range(limit):
        if neutral_armed[i] != neutral_current[i]:
            diff = i
            break
    lt = neutral_current.rfind(b"<", 0, diff + 1)
    if lt == -1:
        return "unknown"
    m = re.match(br"</?([\w:]+)", neutral_current[lt:lt + 64])
    return m.group(1).decode("ascii", "replace") if m else "unknown"


_SHEET_RANGE_RE = re.compile(
    r"^(?:'((?:[^']|'')+)'|([^'!\[\],{}]+))!"
    r"(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$")


def parse_series_range(text):
    """(sheet_title, range_part) for a sheet-qualified single-area range;
    raises ValueError on anything else (external refs, multi-area, array
    literals, unqualified ranges)."""
    m = _SHEET_RANGE_RE.match(text)
    if m is None:
        raise ValueError(
            "series ranges must be single-area, sheet-qualified A1 ranges "
            "like \"'Data'!$B$2:$B$13\" (got {0!r})".format(text))
    sheet = m.group(1).replace("''", "'") if m.group(1) is not None \
        else m.group(2)
    from openpyxl.utils.cell import range_boundaries
    range_boundaries(m.group(3).replace("$", ""))   # bounds check
    return sheet, m.group(3)


def plan_property_edits(wb, ws, key, armed, current, original):
    """A loaded chart's model drifted since arm: express the drift as byte
    patches on the ORIGINAL part bytes, or refuse naming the first
    property chartpatch cannot express (PLAN-v0.1 4.3). Expressible:
    series/axis formula texts (<c:f>) and text runs (<a:t> — titles,
    axis titles). Cached series values are left as-is: Excel re-reads
    series from cells when it renders the chart."""
    from openpyxl.errors import UnsupportedStructureError

    def _refuse(detail):
        raise UnsupportedStructureError(
            "chart {0} on sheet {1!r} was modified, but the edit cannot "
            "be expressed as a byte patch: {2}. Only title/axis text and "
            "series ranges are editable on loaded charts. Nothing was "
            "written.".format(key, ws.title, detail))

    try:
        armed_f, armed_t = _text_sequences(armed)
        current_f, current_t = _text_sequences(current)
        original_f, original_t = _text_sequences(original)
    except ScanRefusal as exc:
        _refuse(str(exc))
    if len(armed_f) != len(current_f):
        _refuse("series or formula references were added or removed "
                "({0} -> {1})".format(len(armed_f), len(current_f)))
    if len(armed_t) != len(current_t):
        _refuse("text runs were added or removed ({0} -> {1}) — adding or "
                "deleting a title is whole-element surgery".format(
                    len(armed_t), len(current_t)))
    neutral_armed = _neutralized(armed, armed_f, armed_t)
    neutral_current = _neutralized(current, current_f, current_t)
    if neutral_armed != neutral_current:
        _refuse("a property near <{0}> changed".format(
            _property_near(neutral_armed, neutral_current)))

    f_changes = {i: (a[2], c[2])
                 for i, (a, c) in enumerate(zip(armed_f, current_f))
                 if a[2] != c[2]}
    t_changes = {i: (a[2], c[2])
                 for i, (a, c) in enumerate(zip(armed_t, current_t))
                 if a[2] != c[2]}
    if not f_changes and not t_changes:
        return original

    sheetnames = {t.casefold() for t in wb.sheetnames}
    for i, (_old, new) in f_changes.items():
        try:
            text = _unescape(new).decode("utf-8")
        except ScanRefusal as exc:
            _refuse(str(exc))
        try:
            sheet, _rng = parse_series_range(text)
        except ValueError as exc:
            _refuse(str(exc))
        if sheet.casefold() not in sheetnames:
            _refuse("the new range {0!r} references sheet {1!r}, which "
                    "does not exist in this workbook".format(text, sheet))
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    for i, (_old, new) in t_changes.items():
        try:
            text = _unescape(new).decode("utf-8")
        except (ScanRefusal, UnicodeDecodeError) as exc:
            _refuse(str(exc))
        if ILLEGAL_CHARACTERS_RE.search(text):
            _refuse("the new text contains characters that cannot be "
                    "written to XML")

    if len(original_f) != len(armed_f):
        _refuse("the original part carries {0} formula references but the "
                "model loaded {1} — the positional mapping is not "
                "trustworthy".format(len(original_f), len(armed_f)))
    if t_changes and len(original_t) != len(armed_t):
        _refuse("the original part carries {0} text runs but the model "
                "loaded {1} — the positional mapping is not "
                "trustworthy".format(len(original_t), len(armed_t)))

    edits = []
    for i, (old, new) in f_changes.items():
        o_start, o_end, o_text = original_f[i]
        if _unescape(o_text) != _unescape(old):
            _refuse("formula reference {0} in the original part ({1!r}) "
                    "does not match the model's arm state ({2!r}) — "
                    "another edit already rewrote it this session; do "
                    "these edits in separate sessions".format(
                        i, o_text, old))
        edits.append((o_start, o_end, _escape(_unescape(new))))
    for i, (old, new) in t_changes.items():
        o_start, o_end, o_text = original_t[i]
        if _unescape(o_text) != _unescape(old):
            _refuse("text run {0} in the original part does not match the "
                    "model's arm state".format(i))
        edits.append((o_start, o_end, _escape(_unescape(new))))

    for start, end, replacement in sorted(edits, reverse=True):
        original = original[:start] + replacement + original[end:]
    return original
