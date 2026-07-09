# paper-xlsx: content-level loss inventory

"""Scan a source archive for content the stock save path cannot preserve.

The inventory must be CONTENT-level, not part-list-level: a stock round trip
of a sparkline-bearing file removes zero parts while gutting the sparklines
. Scans are cheap byte searches over payloads,
never XML parses — the inventory may not crash on files the reader accepts.

Built at load time (the archive is gone by save time on the stock path) and
stashed on the workbook; the stock save path warns from it.
"""

import re

from openpyxl.xml.constants import ARC_APP, EXT_TYPES

# worksheet extLst families we can name (URI -> human name), from upstream
_EXT_URI_RE = re.compile(br'<ext[^>]+uri="(\{[^"]+\})"')

# chart auxiliary parts as real producers name them (colors1.xml, style5.xml)
_CHART_AUX_RE = re.compile(r"(?:colors|style)\d*\.xml$")

# a rich-text run inside sharedStrings: <r> holding an <rPr> (plain <r> with
# no properties flattens losslessly)
_RICH_RUN_RE = re.compile(br"<r\b[^>]*>\s*<rPr")

# drawing anchors that are NOT plain chart/image content: shapes, connectors,
# group shapes, and alternate-content blocks all die in stock's rebuild
_DRAWING_LOSS_MARKERS = (
    (b":sp>", "shape"),
    (b":sp ", "shape"),
    (b"<sp>", "shape"),
    (b":cxnSp", "connector shape"),
    (b":grpSp", "group shape"),
    (b"AlternateContent", "alternate-content block"),
)


class LossInventory:
    """What a stock save would rebuild lossily or drop, as
    ``{"kind", "location", "detail"}`` entries."""

    def __init__(self):
        self.losses = []

    def add(self, kind, location, detail):
        self.losses.append({"kind": kind, "location": location, "detail": detail})

    def __bool__(self):
        return bool(self.losses)

    def __len__(self):
        return len(self.losses)

    def kinds(self):
        return sorted({loss["kind"] for loss in self.losses})

    def render(self):
        lines = [
            "This workbook contains content that will be REBUILT LOSSILY or "
            "DROPPED by this save:"
        ]
        for loss in sorted(self.losses,
                           key=lambda l: (l["kind"], l["location"])):
            lines.append("  - [{0}] {1}: {2}".format(
                loss["kind"], loss["location"], loss["detail"]))
        lines.append(
            "Open the file with load_workbook(..., preserve=True) for a "
            "lossless save."
        )
        return "\n".join(lines)


def _app_properties_are_default(payload):
    """True when app.xml already matches what a stock save regenerates
    (a default ExtendedProperties) — no information is lost by the reset."""
    from openpyxl.package.diff import xml_equivalent
    from openpyxl.packaging.extended import ExtendedProperties
    from openpyxl.xml.functions import tostring

    try:
        default = tostring(ExtendedProperties().to_tree())
        return xml_equivalent(payload, default)
    except Exception:
        return False  # unparseable app.xml: assume it carries information


def _ext_names(payload):
    names = []
    for uri in _EXT_URI_RE.findall(payload):
        try:
            label = EXT_TYPES.get(uri.decode("ascii").upper(), "Unknown")
        except UnicodeDecodeError:
            label = "Unknown"
        names.append(label)
    return names


def scan_archive(archive, valid_files, keep_vba=False, rich_text=False):
    """Build a :class:`LossInventory` from an open source ZipFile.

    ``valid_files`` is the archive namelist (already computed by the reader).
    Read failures on individual entries are recorded loudly, never swallowed
    into silence.
    """
    inv = LossInventory()

    def read(name):
        try:
            return archive.read(name)
        except Exception as exc:  # record loudly; scanning must not crash load
            inv.add("unreadable-part", name,
                    "could not scan: {0}".format(exc))
            return b""

    for name in valid_files:
        if name.endswith("/"):
            continue

        if name == "xl/vbaProject.bin":
            if not keep_vba:
                inv.add("vba", name,
                        "VBA project will be dropped (load with keep_vba=True "
                        "or preserve=True)")
        elif name.startswith("xl/worksheets/") and name.endswith(".xml"):
            payload = read(name)
            if b"<extLst" in payload:
                for label in _ext_names(payload) or ["Unknown"]:
                    inv.add("worksheet-extension", name,
                            "{0} extension will be removed".format(label))
            if not rich_text and b"<is>" in payload \
                    and _RICH_RUN_RE.search(payload):
                inv.add("rich-text", name,
                        "in-cell formatting runs (inline strings) will be "
                        "flattened to plain text (load with rich_text=True "
                        "to model them, or preserve=True to keep them "
                        "verbatim)")
            if b"<protectedRanges" in payload \
                    or b"<protectedRange " in payload:
                # a WORKSHEET-level element (ECMA-376 18.3.1.99), so it is
                # scanned in the sheet payload, not workbook.xml; allow-edit
                # ranges carry password hashes
                inv.add("worksheet-content", name,
                        "protected ranges (allow-edit ranges, incl. their "
                        "password hashes) will be dropped")
        elif name.startswith("xl/drawings/") and name.endswith(".xml"):
            payload = read(name)
            found = sorted({label for marker, label in _DRAWING_LOSS_MARKERS
                            if marker in payload})
            for label in found:
                inv.add("drawing-content", name,
                        "{0} will be dropped when the drawing is rebuilt".format(label))
        elif name.startswith("xl/charts/"):
            if name.endswith(".xml") and "/_rels/" not in name:
                payload = read(name)
                if b"<c:extLst" in payload or b"<extLst" in payload:
                    inv.add("chart-extension", name,
                            "chart-internal extensions will be removed when "
                            "the chart is rebuilt")
            # real producers number these parts (colors1.xml/style1.xml):
            # a bare endswith never matched them
            if _CHART_AUX_RE.search(name):
                inv.add("chart-auxiliary", name,
                        "chart colors/style part will be dropped")
        elif name == ARC_APP:
            payload = read(name)
            if payload and not _app_properties_are_default(payload):
                inv.add("app-properties", name,
                        "extended document properties are reset to defaults on save")
        elif name.startswith("customXml/"):
            inv.add("custom-xml", name, "customXml part will be dropped")
        elif name.startswith("xl/printerSettings/"):
            inv.add("printer-settings", name, "printer settings will be dropped")
        elif name.startswith("xl/threadedComments/"):
            inv.add("threaded-comments", name,
                    "threaded comments part will be dropped")
        elif name == "xl/workbook.xml":
            payload = read(name)
            if b"<fileSharing" in payload:
                inv.add("workbook-content", name,
                        "fileSharing (read-only recommendation/reservation) "
                        "will be dropped")
        elif name == "xl/sharedStrings.xml":
            payload = read(name)
            if not rich_text and _RICH_RUN_RE.search(payload):
                inv.add("rich-text", name,
                        "in-cell formatting runs will be flattened to plain "
                        "text (load with rich_text=True to model them, or "
                        "preserve=True to keep them verbatim)")

    return inv
