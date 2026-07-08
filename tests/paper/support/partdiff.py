"""Part-level package comparison for the contract harness.

Phase-1 test-local implementation; becomes a thin wrapper over
``openpyxl.package`` once the kernel lands (Phase 2). The byte-identity
invariant is defined on part *payloads*, never whole-archive bytes
(CONVENTIONS §7): zip entry metadata (timestamps, attrs) is out of scope here.
"""
from __future__ import annotations

import io
import zipfile
from xml.etree import ElementTree as ET


def part_payloads(source):
    """Return {part_name: payload_bytes} for an xlsx package.

    ``source`` may be a filesystem path or a bytes object.
    """
    if isinstance(source, (bytes, bytearray)):
        fh = io.BytesIO(source)
    else:
        fh = open(source, "rb")
    try:
        with zipfile.ZipFile(fh) as zf:
            return {info.filename: zf.read(info.filename) for info in zf.infolist()}
    finally:
        fh.close()


class PartsDiff:

    def __init__(self, added, removed, changed, identical):
        self.added = added        # part names only in B
        self.removed = removed    # part names only in A
        self.changed = changed    # payload differs
        self.identical = identical

    @property
    def clean(self):
        return not (self.added or self.removed or self.changed)

    def __repr__(self):
        return (
            "PartsDiff(added={0}, removed={1}, changed={2}, identical={3} parts)".format(
                sorted(self.added), sorted(self.removed), sorted(self.changed),
                len(self.identical),
            )
        )


def diff_parts(a, b):
    """Payload-level diff of two packages (paths or bytes)."""
    pa = part_payloads(a)
    pb = part_payloads(b)
    added = set(pb) - set(pa)
    removed = set(pa) - set(pb)
    common = set(pa) & set(pb)
    changed = {name for name in common if pa[name] != pb[name]}
    return PartsDiff(added, removed, changed, common - changed)


def _significant_text(text):
    """None and whitespace-only collapse to '' for inter-element whitespace;
    anything else is preserved exactly (never normalize cell text content)."""
    if text is None:
        return ""
    if text.strip() == "":
        return ""
    return text


def _xml_nodes_equal(a, b, path, diffs, max_diffs):
    if len(diffs) >= max_diffs:
        return
    if a.tag != b.tag:
        diffs.append("{0}: tag {1!r} != {2!r}".format(path, a.tag, b.tag))
        return
    if dict(a.attrib) != dict(b.attrib):
        diffs.append("{0}: attrib {1!r} != {2!r}".format(path, dict(a.attrib), dict(b.attrib)))
    if _significant_text(a.text) != _significant_text(b.text):
        diffs.append("{0}: text {1!r} != {2!r}".format(path, a.text, b.text))
    if _significant_text(a.tail) != _significant_text(b.tail):
        diffs.append("{0}: tail {1!r} != {2!r}".format(path, a.tail, b.tail))
    a_children = list(a)
    b_children = list(b)
    if len(a_children) != len(b_children):
        diffs.append(
            "{0}: child count {1} != {2} (A: {3}; B: {4})".format(
                path, len(a_children), len(b_children),
                [c.tag for c in a_children], [c.tag for c in b_children],
            )
        )
        return
    for i, (ca, cb) in enumerate(zip(a_children, b_children)):
        _xml_nodes_equal(ca, cb, "{0}/{1}[{2}]".format(path, ca.tag, i), diffs, max_diffs)


def xml_semantic_diff(a_bytes, b_bytes, max_diffs=25):
    """Compare two XML payloads semantically; returns a list of differences.

    Semantics: element tree structure, tags (namespace-aware via Clark
    notation), attributes (order-insensitive), and text content must match.
    Inter-element whitespace is insignificant; any non-whitespace text is
    compared exactly — cell text content is never normalized. Namespace
    *prefixes* are not compared (Clark names make them irrelevant), matching
    the kernel contract of ``xml_equivalent``.
    """
    a = ET.fromstring(a_bytes)
    b = ET.fromstring(b_bytes)
    diffs = []
    _xml_nodes_equal(a, b, "/" + a.tag.split("}")[-1], diffs, max_diffs)
    return diffs


def xml_equivalent(a_bytes, b_bytes):
    return not xml_semantic_diff(a_bytes, b_bytes, max_diffs=1)
