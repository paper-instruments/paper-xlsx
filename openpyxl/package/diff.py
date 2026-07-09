# paper-xlsx: semantic XML comparison and package diff

import hashlib
import io
import zipfile
from xml.etree import ElementTree as ET


def _read_payload_source(source):
    """Accept a filesystem path, bytes, or a binary file-like; return bytes."""
    if isinstance(source, (bytes, bytearray)):
        return bytes(source)
    if hasattr(source, "read"):
        return source.read()
    with open(source, "rb") as f:
        return f.read()


def _looks_like_xml(payload):
    head = payload[:256].lstrip(b"\xef\xbb\xbf \t\r\n")
    return head.startswith(b"<")


def _significant_text(text):
    """Inter-element whitespace is insignificant; any non-whitespace text is
    compared exactly — cell text content is never normalized."""
    if text is None:
        return ""
    if text.strip() == "":
        return ""
    return text


def _walk(a, b, path, diffs, max_diffs):
    if len(diffs) >= max_diffs:
        return
    if a.tag != b.tag:
        diffs.append("{0}: tag {1!r} != {2!r}".format(path, a.tag, b.tag))
        return
    if dict(a.attrib) != dict(b.attrib):
        diffs.append("{0}: attrib {1!r} != {2!r}".format(
            path, dict(a.attrib), dict(b.attrib)))
    if _significant_text(a.text) != _significant_text(b.text):
        diffs.append("{0}: text {1!r} != {2!r}".format(path, a.text, b.text))
    if _significant_text(a.tail) != _significant_text(b.tail):
        diffs.append("{0}: tail {1!r} != {2!r}".format(path, a.tail, b.tail))
    a_children = list(a)
    b_children = list(b)
    if len(a_children) != len(b_children):
        diffs.append("{0}: child count {1} != {2}".format(
            path, len(a_children), len(b_children)))
        return
    for i, (ca, cb) in enumerate(zip(a_children, b_children)):
        _walk(ca, cb, "{0}/{1}[{2}]".format(path, ca.tag.split('}')[-1], i),
              diffs, max_diffs)


def xml_semantic_diff(a, b, max_diffs=25):
    """Semantic differences between two XML payloads (paths/bytes/file-likes).

    Compared: element structure, Clark-qualified tags (namespace *prefixes*
    are insignificant), attributes (order-insensitive), and text content —
    which is never normalized. Returns a list of human-readable differences,
    empty when equivalent.
    """
    ta = ET.fromstring(_read_payload_source(a))
    tb = ET.fromstring(_read_payload_source(b))
    diffs = []
    _walk(ta, tb, "/" + ta.tag.split("}")[-1], diffs, max_diffs)
    return diffs


def xml_equivalent(a, b):
    """True when two XML payloads are semantically equivalent.

    Never normalizes cell text content."""
    return not xml_semantic_diff(a, b, max_diffs=1)


class PartChange:
    """One changed part in a package diff."""

    def __init__(self, part, kind, detail):
        self.part = part
        self.kind = kind          # "xml" or "binary"
        self.detail = detail      # list of strings

    def to_dict(self):
        return {"part": self.part, "kind": self.kind, "detail": list(self.detail)}

    def __repr__(self):
        return "PartChange({0!r}, {1}, {2} details)".format(
            self.part, self.kind, len(self.detail))


class PackageDiff:
    """Part-by-part diff of two OOXML packages.

    XML parts are compared semantically; binary parts by size and SHA-256.
    ``identical`` counts parts whose payloads are byte-identical (a stricter
    condition than semantic equivalence; byte-identical XML parts are never
    parsed at all).
    """

    SCHEMA = "package_diff"
    VERSION = 1

    def __init__(self, added, removed, changed, identical, equivalent):
        self.added = sorted(added)
        self.removed = sorted(removed)
        self.changed = sorted(changed, key=lambda c: c.part)
        self.identical = sorted(identical)
        # XML parts that differ in bytes but are semantically equivalent
        self.equivalent = sorted(equivalent)

    @property
    def clean(self):
        """No parts added, removed, or semantically changed."""
        return not (self.added or self.removed or self.changed)

    def to_dict(self):
        return {
            "schema": self.SCHEMA,
            "version": self.VERSION,
            "added": list(self.added),
            "removed": list(self.removed),
            "changed": [c.to_dict() for c in self.changed],
            "byte_identical": list(self.identical),
            "semantically_equivalent": list(self.equivalent),
        }

    def __repr__(self):
        return ("PackageDiff(added={0}, removed={1}, changed={2}, "
                "byte_identical={3}, equivalent={4})".format(
                    self.added, self.removed,
                    [c.part for c in self.changed],
                    len(self.identical), self.equivalent))


def _payloads(source):
    data = _read_payload_source(source)
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        return {info.filename: zf.read(info.filename) for info in zf.infolist()}


def diff_package(a, b, max_detail=25):
    """Diff two packages (paths, bytes, or binary file-likes) part by part."""
    pa = _payloads(a)
    pb = _payloads(b)
    added = set(pb) - set(pa)
    removed = set(pa) - set(pb)
    identical = []
    equivalent = []
    changed = []
    for name in sorted(set(pa) & set(pb)):
        if pa[name] == pb[name]:
            identical.append(name)
            continue
        if _looks_like_xml(pa[name]) and _looks_like_xml(pb[name]):
            try:
                detail = xml_semantic_diff(pa[name], pb[name], max_diffs=max_detail)
            except ET.ParseError as exc:
                changed.append(PartChange(name, "xml", ["unparseable: {0}".format(exc)]))
                continue
            if detail:
                changed.append(PartChange(name, "xml", detail))
            else:
                equivalent.append(name)
        else:
            detail = [
                "size {0} -> {1}".format(len(pa[name]), len(pb[name])),
                "sha256 {0} -> {1}".format(
                    hashlib.sha256(pa[name]).hexdigest()[:16],
                    hashlib.sha256(pb[name]).hexdigest()[:16]),
            ]
            changed.append(PartChange(name, "binary", detail))
    return PackageDiff(added, removed, changed, identical, equivalent)
