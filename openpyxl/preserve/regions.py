# paper-xlsx: worksheet satellite regions (PR-0 D5 Tier 2 / D15)

"""Fully-modeled worksheet satellite elements: faithful serialization,
change detection, and the write-policy gates.

Change detection compares the MODEL's serialization at arm time against the
model's serialization at save time — never model-vs-original-bytes, which
would false-positive on every producer quirk openpyxl normalizes at load
(xr:uid drops, attribute defaults, ...). Self-consistent serialization means
a region is dirty exactly when the USER changed it. When a region is dirty,
the splice replaces the whole element with the model serialization — legal
because these elements are fully modeled (round-trip identity proven in
Phase 0) — subject to the D15 gates.

Serializers mirror openpyxl/worksheet/_writer.py's per-element logic
exactly; a serializer returning ``None`` means "stock would not emit this
element" (absent).
"""

from openpyxl.errors import UnsupportedStructureError
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import tostring


def _nonempty(obj):
    return obj.to_tree() if obj else None


def _sheet_pr(ws):
    return ws.sheet_properties.to_tree()               # _writer.write_properties


def _views(ws):
    return ws.views.to_tree()                          # _writer.write_views


def _sheet_format(ws):
    # mirror _writer.write_format's outline sync, but PURELY: upstream
    # reads DimensionHolder.max_outline, a value to_tree() only refreshes
    # DURING a cols render — order-dependent state that false-dirtied
    # sheetFormatPr on every cols-bearing sheet (Batch-0 gate). Compute
    # the same quantity the way holder.to_tree() does (reindex is the
    # idempotent min/max normalization upstream runs on every render;
    # membership = renders-to-something), without touching max_outline
    # or the model's sheet_format.
    import copy as _copy

    outlines = set()
    for dim in ws.column_dimensions.values():
        dim.reindex()
        if dim.to_tree() is not None:
            outlines.add(dim.outlineLevel)
    fmt = ws.sheet_format
    if outlines:
        fmt = _copy.copy(fmt)
        fmt.outlineLevelCol = max(outlines)
    return fmt.to_tree()


def _cols(ws):
    # a pure READ (ws.column_dimensions['J']) materializes a default entry
    # whose render differs from absence — filter those out or the diff
    # mis-reads a read as a user edit (reads must never dirty)
    holder = ws.column_dimensions
    default_attrs = None
    materialized = []
    for key, dim in list(holder.items()):
        attrs = dict(dim)
        if default_attrs is None:
            from openpyxl.worksheet.dimensions import ColumnDimension

            probe = ColumnDimension(ws, index=dim.index)
            probe_attrs = dict(probe)
            default_attrs = {k: v for k, v in probe_attrs.items()
                             if k not in ("min", "max")}
        significant = {k: v for k, v in attrs.items()
                       if k not in ("min", "max")}
        if significant == default_attrs:
            materialized.append(key)
    if not materialized:
        return holder.to_tree()
    import copy as _copy

    filtered = {k: v for k, v in holder.items() if k not in materialized}
    if not filtered:
        return None
    trimmed = type(holder)(worksheet=ws,
                           default_factory=holder.default_factory)
    trimmed.update(filtered)
    return trimmed.to_tree()


def _protection(ws):
    return _nonempty(ws.protection)


def _scenarios(ws):
    return _nonempty(ws.scenarios)


def _auto_filter(ws):
    return _nonempty(ws.auto_filter)


def _merged(ws):
    # mirror _writer.write_merged_cells
    from openpyxl.worksheet.merge import MergeCell, MergeCells

    if not ws.merged_cells:
        return None
    cells = [MergeCell(str(ref)) for ref in ws.merged_cells]
    return MergeCells(mergeCell=cells).to_tree()


def _validations(ws):
    return _nonempty(ws.data_validations)


def _print_options(ws):
    return _nonempty(ws.print_options)


def _margins(ws):
    return _nonempty(ws.page_margins)


def _page_setup(ws):
    return _nonempty(ws.page_setup)


def _header(ws):
    return _nonempty(ws.HeaderFooter)


def _row_breaks(ws):
    return _nonempty(ws.row_breaks)


def _col_breaks(ws):
    return _nonempty(ws.col_breaks)


class Region:

    def __init__(self, tag, serialize):
        self.tag = tag              # local name in the main namespace
        self.serialize = serialize  # ws -> Element or None

    def render(self, ws):
        el = self.serialize(ws)
        if el is None:
            return None
        return tostring(el)


# CT_Worksheet child sequence (ECMA-376 §18.3.1.99); the splice uses it to
# place a region that did not exist in the original document.
CT_WORKSHEET_ORDER = [
    "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols",
    "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges",
    "scenarios", "autoFilter", "sortState", "dataConsolidate",
    "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting",
    "dataValidations", "hyperlinks", "printOptions", "pageMargins",
    "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
    "customProperties", "cellWatches", "ignoredErrors", "smartTags",
    "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects",
    "controls", "webPublishItems", "tableParts", "extLst",
]
CT_ORDER_INDEX = {tag: i for i, tag in enumerate(CT_WORKSHEET_ORDER)}

# regions the splice can rewrite in this build stage (PR-0 D15 Tier 1 set,
# minus the ones needing cross-part coordination which land in 2d)
SPLICEABLE_REGIONS = [
    Region("sheetPr", _sheet_pr),
    Region("sheetViews", _views),
    Region("sheetFormatPr", _sheet_format),
    Region("cols", _cols),
    Region("sheetProtection", _protection),
    Region("scenarios", _scenarios),
    Region("autoFilter", _auto_filter),
    Region("mergeCells", _merged),
    Region("dataValidations", _validations),
    Region("printOptions", _print_options),
    Region("pageMargins", _margins),
    Region("pageSetup", _page_setup),
    Region("headerFooter", _header),
    Region("rowBreaks", _row_breaks),
    Region("colBreaks", _col_breaks),
]
REGION_BY_TAG = {r.tag: r for r in SPLICEABLE_REGIONS}

# regions whose USER change is detected but written only via the saver's
# own planners (the model render is a detection signature, not writable
# bytes): the tableParts element is rebuilt by preserve.tables and rides
# the region splice as crafted bytes (Batch 2)
DETECT_ONLY_REGIONS = []

# regions whose replacement bytes come from the SAVER's own planner (never
# from a model render): the splice accepts them as-is
SAVER_CRAFTED_REGIONS = frozenset(["tableParts"])


def _render_cf(ws):
    """Serialize conditional formatting WITHOUT the stock writer's dxfId
    side effects — detection only."""
    parts = []
    for cf in ws.conditional_formatting:
        parts.append(tostring(cf.to_tree()))
    return tuple(parts)


def render_cf_for_write(ws):
    """Serialize conditional formatting FOR WRITING, mirroring the stock
    writer's dxf handling (worksheet/_writer.py write_formatting): rules
    carrying a dxf get a dxfId allocated in the workbook's differential
    styles; the new dxfs are appended to styles.xml by the styles planner."""
    from openpyxl.styles.differential import DifferentialStyle

    empty = DifferentialStyle()
    wb = ws.parent
    parts = []
    for cf in ws.conditional_formatting:
        for rule in cf.rules:
            if rule.dxf and rule.dxf != empty:
                rule.dxfId = wb._differential_styles.add(rule.dxf)
        parts.append(tostring(cf.to_tree()))
    return b"".join(parts)


def render_hyperlinks_for_write(ws):
    """The hyperlinks element from the cells' link objects (ids must already
    be assigned for external links)."""
    from openpyxl.worksheet.hyperlink import HyperlinkList

    links = [cell._hyperlink for (_r, _c), cell in sorted(ws._cells.items())
             if getattr(cell, "_hyperlink", None) is not None]
    if not links:
        return b""
    return tostring(HyperlinkList(links).to_tree())


def hyperlink_signatures(ws):
    """Per-cell hyperlink signatures, excluding the relationship id (ids for
    new links are allocated at save time and must not affect detection)."""
    sig = {}
    for (row, col), cell in ws._cells.items():
        link = getattr(cell, "_hyperlink", None)
        if link is not None:
            sig[(row, col)] = (link.target, link.location, link.tooltip,
                               link.display)
    return sig


def _render_tables(ws):
    return tuple(sorted(ws.tables.keys()))


def snapshot_regions(ws):
    """Serialize every tracked region of one worksheet (arm time / save
    time; comparing the two detects user changes)."""
    snap = {}
    for region in SPLICEABLE_REGIONS:
        snap[region.tag] = region.render(ws)
    snap["conditionalFormatting"] = _render_cf(ws)
    snap["hyperlinks"] = hyperlink_signatures(ws)
    snap["tableParts"] = _render_tables(ws)
    return snap


def snapshot_row_attrs(ws):
    """Row display attributes (they serialize as attributes of <row>
    elements inside sheetData, not as a separate element)."""
    snap = {}
    for idx, dim in ws.row_dimensions.items():
        attrs = dict(dim)
        if attrs:
            snap[idx] = tuple(sorted(attrs.items()))
    return snap


def diff_regions(ws, armed_snapshot):
    """Return {tag: new_serialization} for regions the user changed.

    Rendered twice, second pass kept: the arm snapshot is the settled
    render (ledger double-render, PLAN-v0.1 0.3), so the comparison must
    be settled-vs-settled or an impure serializer's first-pass output
    false-dirties the region."""
    snapshot_regions(ws)
    current = snapshot_regions(ws)
    changed = {}
    for tag, rendered in current.items():
        if rendered != armed_snapshot.get(tag):
            changed[tag] = rendered
    return changed


def diff_row_attrs(ws, armed_snapshot):
    """Return {row_index: {attr: value}} for changed rows; a row present in
    the arm snapshot but now attribute-free maps to an empty dict."""
    current = snapshot_row_attrs(ws)
    changed = {}
    for idx in set(current) | set(armed_snapshot):
        if current.get(idx) != armed_snapshot.get(idx):
            changed[idx] = dict(current.get(idx, ()))
    return changed
