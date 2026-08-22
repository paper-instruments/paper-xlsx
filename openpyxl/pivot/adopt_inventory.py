# paper-xlsx: closed-world core-schema inventory for foreign pivot adoption

"""Classify every core-schema element and attribute on a selected pivot graph.

Unclassified content and classified properties carrying an unrecognized
nondefault value refuse adoption. Local names, filenames, and familiar URIs
are not identity.
"""

from __future__ import annotations

from openpyxl.pivot.graph import _local, _parse_xml
from openpyxl.pivot.adopt_evidence import excel_equivalence_proved
from openpyxl.pivot.qualify import (
    EXTENSION_ALLOWLIST,
    QualificationReason,
    _extension_payloads_are_benign,
)
from openpyxl.xml.constants import REL_NS, SHEET_MAIN_NS


INVENTORY_VERSION = 1
REPRESENTED = "represented"
SCHEMA_DEFAULT = "schema-default"
IRRELEVANT = "irrelevant"
UNSUPPORTED = "unsupported"

_SUPPORTED_SUBTOTALS = frozenset((
    "sum", "count", "countNums", "average", "min", "max",
))
_ITEM_TYPES = frozenset(("data", "default", "grand", "blank"))
_IGNORED_ATTR_LOCALS = frozenset((
    "Ignorable",
    "uid",
))
_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_REVISION_NS = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"

# Elements that belong to the released classic-pivot model.
_ELEMENTS = {
    "pivotTableDefinition": REPRESENTED,
    "location": REPRESENTED,
    "pivotFields": REPRESENTED,
    "pivotField": REPRESENTED,
    "items": REPRESENTED,
    "item": REPRESENTED,
    "rowFields": REPRESENTED,
    "colFields": REPRESENTED,
    "rowItems": REPRESENTED,
    "colItems": REPRESENTED,
    "pageFields": REPRESENTED,
    "pageField": REPRESENTED,
    "dataFields": REPRESENTED,
    "dataField": REPRESENTED,
    "field": REPRESENTED,
    "i": REPRESENTED,
    "x": REPRESENTED,
    "pivotTableStyleInfo": REPRESENTED,
    "extLst": REPRESENTED,
    "ext": REPRESENTED,
    "pivotCacheDefinition": REPRESENTED,
    "cacheSource": REPRESENTED,
    "worksheetSource": REPRESENTED,
    "cacheFields": REPRESENTED,
    "cacheField": REPRESENTED,
    "sharedItems": REPRESENTED,
    "s": REPRESENTED,
    "n": REPRESENTED,
    "b": REPRESENTED,
    "d": REPRESENTED,
    "m": REPRESENTED,
    "e": REPRESENTED,
    "pivotCacheRecords": REPRESENTED,
    "r": REPRESENTED,
}

_UNSUPPORTED_ELEMENTS = frozenset((
    "formats", "conditionalFormats", "chartFormats", "pivotHierarchies",
    "filters", "pivotFilter", "autoSortScope", "rowHierarchiesUsage",
    "colHierarchiesUsage", "fieldGroup", "discretePr", "rangePr",
    "cacheHierarchies", "kpis", "tupleCache", "calculatedItems",
    "calculatedMembers", "dimensions", "measureGroups", "maps",
    "calculatedItem", "calculatedMember",
))

_DEFAULTS = {
    ("pivotTableDefinition", "asteriskTotals"): "0",
    ("pivotTableDefinition", "chartFormat"): "0",
    ("pivotTableDefinition", "colGrandTotals"): "1",
    ("pivotTableDefinition", "dataOnRows"): "0",
    ("pivotTableDefinition", "compact"): "1",
    ("pivotTableDefinition", "compactData"): "1",
    ("pivotTableDefinition", "gridDropZones"): "0",
    ("pivotTableDefinition", "disableFieldList"): "0",
    ("pivotTableDefinition", "editData"): "0",
    ("pivotTableDefinition", "enableDrill"): "1",
    ("pivotTableDefinition", "enableFieldProperties"): "1",
    ("pivotTableDefinition", "enableWizard"): "1",
    ("pivotTableDefinition", "fieldListSortAscending"): "0",
    ("pivotTableDefinition", "fieldPrintTitles"): "0",
    ("pivotTableDefinition", "immersive"): "1",
    ("pivotTableDefinition", "indent"): "1",
    ("pivotTableDefinition", "itemPrintTitles"): "0",
    ("pivotTableDefinition", "mdxSubqueries"): "0",
    ("pivotTableDefinition", "mergeItem"): "0",
    ("pivotTableDefinition", "outline"): "0",
    ("pivotTableDefinition", "outlineData"): "0",
    ("pivotTableDefinition", "pageOverThenDown"): "0",
    ("pivotTableDefinition", "pageWrap"): "0",
    ("pivotTableDefinition", "preserveFormatting"): "1",
    ("pivotTableDefinition", "printDrill"): "0",
    ("pivotTableDefinition", "published"): "0",
    ("pivotTableDefinition", "rowGrandTotals"): "1",
    ("pivotTableDefinition", "showCalcMbrs"): "1",
    ("pivotTableDefinition", "showDataDropDown"): "1",
    ("pivotTableDefinition", "showDataTips"): "1",
    ("pivotTableDefinition", "showDrill"): "1",
    ("pivotTableDefinition", "showDropZones"): "1",
    ("pivotTableDefinition", "showEmptyCol"): "0",
    ("pivotTableDefinition", "showEmptyRow"): "0",
    ("pivotTableDefinition", "showError"): "0",
    ("pivotTableDefinition", "showHeaders"): "1",
    ("pivotTableDefinition", "showItems"): "1",
    ("pivotTableDefinition", "showMemberPropertyTips"): "1",
    ("pivotTableDefinition", "showMissing"): "1",
    ("pivotTableDefinition", "showMultipleLabel"): "1",
    ("pivotTableDefinition", "subtotalHiddenItems"): "0",
    ("pivotTableDefinition", "useAutoFormatting"): "0",
    ("pivotTableDefinition", "visualTotals"): "1",
    ("pivotTableDefinition", "applyNumberFormats"): "0",
    ("pivotTableDefinition", "applyBorderFormats"): "0",
    ("pivotTableDefinition", "applyFontFormats"): "0",
    ("pivotTableDefinition", "applyPatternFormats"): "0",
    ("pivotTableDefinition", "applyAlignmentFormats"): "0",
    ("pivotTableDefinition", "applyWidthHeightFormats"): "0",
    ("pivotTableDefinition", "multipleFieldFilters"): "0",
    ("pivotField", "compact"): "1",
    ("pivotField", "outline"): "1",
    ("pivotField", "dragOff"): "1",
    ("pivotField", "dragToCol"): "1",
    ("pivotField", "dragToData"): "1",
    ("pivotField", "dragToPage"): "1",
    ("pivotField", "dragToRow"): "1",
    ("pivotField", "itemPageCount"): "10",
    ("pivotField", "showDropDowns"): "1",
    ("pivotField", "sortType"): "manual",
    ("pivotField", "subtotalTop"): "1",
    ("pivotField", "topAutoShow"): "1",
    ("pivotField", "showAll"): "1",
    ("pivotField", "defaultSubtotal"): "1",
    ("item", "t"): "data",
    ("item", "sd"): "1",
    ("i", "t"): "data",
    ("i", "i"): "0",
    ("i", "r"): "0",
    ("x", "v"): "0",
    ("dataField", "showDataAs"): "normal",
    ("dataField", "subtotal"): "sum",
    ("dataField", "baseField"): "-1",
    ("dataField", "baseItem"): "1048832",
    ("pageField", "item"): "-1",
    ("pivotCacheDefinition", "backgroundQuery"): "0",
    ("pivotCacheDefinition", "enableRefresh"): "1",
    ("pivotCacheDefinition", "saveData"): "1",
    ("cacheField", "numFmtId"): "0",
    ("sharedItems", "count"): "0",
}

_REPRESENTED_ATTRS = {
    ("pivotTableDefinition", "name"): REPRESENTED,
    ("pivotTableDefinition", "cacheId"): REPRESENTED,
    ("pivotTableDefinition", "dataCaption"): REPRESENTED,
    ("pivotTableDefinition", "dataOnRows"): REPRESENTED,
    ("pivotTableDefinition", "compact"): REPRESENTED,
    ("pivotTableDefinition", "outline"): REPRESENTED,
    ("pivotTableDefinition", "compactData"): REPRESENTED,
    ("pivotTableDefinition", "outlineData"): REPRESENTED,
    ("pivotTableDefinition", "gridDropZones"): REPRESENTED,
    ("pivotTableDefinition", "rowGrandTotals"): REPRESENTED,
    ("pivotTableDefinition", "colGrandTotals"): REPRESENTED,
    ("pivotTableDefinition", "tag"): REPRESENTED,
    ("pivotTableDefinition", "createdVersion"): IRRELEVANT,
    ("pivotTableDefinition", "updatedVersion"): IRRELEVANT,
    ("pivotTableDefinition", "minRefreshableVersion"): IRRELEVANT,
    ("location", "ref"): REPRESENTED,
    ("location", "firstHeaderRow"): REPRESENTED,
    ("location", "firstDataRow"): REPRESENTED,
    ("location", "firstDataCol"): REPRESENTED,
    ("location", "rowPageCount"): REPRESENTED,
    ("location", "colPageCount"): REPRESENTED,
    ("pivotFields", "count"): IRRELEVANT,
    ("rowFields", "count"): IRRELEVANT,
    ("colFields", "count"): IRRELEVANT,
    ("rowItems", "count"): IRRELEVANT,
    ("colItems", "count"): IRRELEVANT,
    ("pageFields", "count"): IRRELEVANT,
    ("dataFields", "count"): IRRELEVANT,
    ("items", "count"): IRRELEVANT,
    ("field", "x"): REPRESENTED,
    ("pivotField", "name"): REPRESENTED,
    ("pivotField", "axis"): REPRESENTED,
    ("pivotField", "dataField"): REPRESENTED,
    ("pivotField", "showAll"): REPRESENTED,
    ("pivotField", "defaultSubtotal"): REPRESENTED,
    ("pivotField", "compact"): REPRESENTED,
    ("pivotField", "outline"): REPRESENTED,
    ("pivotField", "subtotalTop"): REPRESENTED,
    ("item", "x"): REPRESENTED,
    ("item", "t"): REPRESENTED,
    ("item", "h"): REPRESENTED,
    ("item", "n"): REPRESENTED,
    ("i", "t"): REPRESENTED,
    ("i", "i"): REPRESENTED,
    ("i", "r"): REPRESENTED,
    ("x", "v"): REPRESENTED,
    ("pageField", "fld"): REPRESENTED,
    ("pageField", "item"): REPRESENTED,
    ("pageField", "hier"): IRRELEVANT,
    ("dataField", "name"): REPRESENTED,
    ("dataField", "fld"): REPRESENTED,
    ("dataField", "subtotal"): REPRESENTED,
    ("dataField", "numFmtId"): REPRESENTED,
    ("dataField", "showDataAs"): REPRESENTED,
    ("dataField", "baseField"): IRRELEVANT,
    ("dataField", "baseItem"): IRRELEVANT,
    ("pivotTableStyleInfo", "name"): REPRESENTED,
    ("pivotTableStyleInfo", "showRowHeaders"): IRRELEVANT,
    ("pivotTableStyleInfo", "showColHeaders"): IRRELEVANT,
    ("pivotTableStyleInfo", "showRowStripes"): IRRELEVANT,
    ("pivotTableStyleInfo", "showColStripes"): IRRELEVANT,
    ("pivotTableStyleInfo", "showLastColumn"): IRRELEVANT,
    ("ext", "uri"): REPRESENTED,
    ("pivotCacheDefinition", "recordCount"): REPRESENTED,
    ("pivotCacheDefinition", "refreshOnLoad"): REPRESENTED,
    ("pivotCacheDefinition", "enableRefresh"): REPRESENTED,
    ("pivotCacheDefinition", "saveData"): REPRESENTED,
    ("pivotCacheDefinition", "backgroundQuery"): IRRELEVANT,
    ("pivotCacheDefinition", "createdVersion"): IRRELEVANT,
    ("pivotCacheDefinition", "refreshedVersion"): IRRELEVANT,
    ("pivotCacheDefinition", "minRefreshableVersion"): IRRELEVANT,
    ("pivotCacheDefinition", "refreshedBy"): IRRELEVANT,
    ("pivotCacheDefinition", "refreshedDate"): IRRELEVANT,
    ("pivotCacheDefinition", "refreshedDateIso"): IRRELEVANT,
    ("cacheSource", "type"): REPRESENTED,
    ("worksheetSource", "ref"): REPRESENTED,
    ("worksheetSource", "sheet"): REPRESENTED,
    ("worksheetSource", "name"): REPRESENTED,
    ("cacheFields", "count"): IRRELEVANT,
    ("cacheField", "name"): REPRESENTED,
    ("cacheField", "numFmtId"): IRRELEVANT,
    ("sharedItems", "count"): IRRELEVANT,
    ("sharedItems", "containsBlank"): REPRESENTED,
    ("sharedItems", "containsDate"): REPRESENTED,
    ("sharedItems", "containsInteger"): REPRESENTED,
    ("sharedItems", "containsMixedTypes"): REPRESENTED,
    ("sharedItems", "containsNonDate"): REPRESENTED,
    ("sharedItems", "containsNumber"): REPRESENTED,
    ("sharedItems", "containsSemiMixedTypes"): REPRESENTED,
    ("sharedItems", "containsString"): REPRESENTED,
    ("sharedItems", "minValue"): REPRESENTED,
    ("sharedItems", "maxValue"): REPRESENTED,
    ("sharedItems", "minDate"): REPRESENTED,
    ("sharedItems", "maxDate"): REPRESENTED,
    ("sharedItems", "longText"): IRRELEVANT,
    ("s", "v"): REPRESENTED,
    ("n", "v"): REPRESENTED,
    ("b", "v"): REPRESENTED,
    ("d", "v"): REPRESENTED,
    ("e", "v"): REPRESENTED,
    ("pivotCacheRecords", "count"): IRRELEVANT,
}

_CUSTOM_SUBTOTAL_ATTRS = frozenset((
    "sumSubtotal", "countASubtotal", "avgSubtotal", "maxSubtotal",
    "minSubtotal", "productSubtotal", "countSubtotal", "stdDevSubtotal",
    "stdDevPSubtotal", "varSubtotal", "varPSubtotal",
))


def classify_selected_graph(workbook, node, cache):
    """Return reasons for unclassified or unsupported selected-graph content."""
    reasons = []
    package = getattr(workbook, "_paper_source", None)
    if package is None:
        return (_reason("foreign-core-semantics-unclassified",
                        part=node.identity.pivot_part),)
    parts = [node.identity.pivot_part]
    if cache is not None and cache.definition_part:
        parts.append(cache.definition_part)
    if cache is not None and cache.records_part:
        parts.append(cache.records_part)
    import io
    import zipfile

    try:
        with zipfile.ZipFile(io.BytesIO(package)) as archive:
            names = set(archive.namelist())
            for part in parts:
                if part not in names:
                    reasons.append(_reason(
                        "foreign-cache-records-unavailable"
                        if part == getattr(cache, "records_part", None)
                        else "invalid-pivot-graph",
                        part=part,
                    ))
                    continue
                root = _parse_xml(archive.read(part))
                if root is None:
                    reasons.append(_reason(
                        "foreign-core-semantics-unclassified", part=part))
                    continue
                reasons.extend(_classify_tree(root, part))
    except (KeyError, OSError, ValueError, zipfile.BadZipFile):
        reasons.append(_reason(
            "foreign-core-semantics-unclassified",
            part=node.identity.pivot_part,
        ))
    reasons.extend(_classify_extensions(workbook, node, cache))
    return tuple(reasons)


def _classify_tree(root, part):
    reasons = []
    skip = set()
    for element in root.iter():
        if element in skip:
            continue
        local = _local(element.tag)
        if local == "extLst":
            for descendant in element.iter():
                skip.add(descendant)
            continue
        ns = _namespace(element.tag)
        if ns and ns != SHEET_MAIN_NS:
            reasons.append(_reason(
                "foreign-core-semantics-unclassified",
                part=part, element=local, namespace=ns,
            ))
            continue
        if local in _UNSUPPORTED_ELEMENTS:
            reasons.append(_reason(
                "unsupported-pivot-feature",
                part=part, element=local,
            ))
            continue
        if local not in _ELEMENTS:
            reasons.append(_reason(
                "foreign-core-semantics-unclassified",
                part=part, element=local,
            ))
            continue
        reasons.extend(_classify_attributes(element, local, part))
    return reasons


def _classify_attributes(element, local, part):
    reasons = []
    for raw, value in element.attrib.items():
        name = _local(raw)
        ns = _namespace(raw)
        if _is_ignored_attribute(raw, name, ns):
            continue
        if ns and ns not in (SHEET_MAIN_NS, REL_NS, ""):
            reasons.append(_reason(
                "foreign-core-semantics-unclassified",
                part=part, element=local, attribute=name, namespace=ns,
            ))
            continue
        if name in _CUSTOM_SUBTOTAL_ATTRS and value not in ("0", "false"):
            reasons.append(_reason(
                "unsupported-pivot-feature",
                part=part, element=local, attribute=name, value=value,
            ))
            continue
        key = (local, name)
        default = _DEFAULTS.get(key)
        classification = _REPRESENTED_ATTRS.get(key)
        if classification is None and default is not None and value == default:
            continue
        if classification in (REPRESENTED, IRRELEVANT, SCHEMA_DEFAULT):
            extra = _attribute_value_reason(local, name, value, part)
            if extra is not None:
                reasons.append(extra)
            continue
        if default is not None and value != default:
            reasons.append(_reason(
                "foreign-core-semantics-unclassified",
                part=part, element=local, attribute=name, value=value,
            ))
            continue
        reasons.append(_reason(
            "foreign-core-semantics-unclassified",
            part=part, element=local, attribute=name,
        ))
    return reasons


def _attribute_value_reason(local, name, value, part):
    if local == "dataField" and name == "showDataAs" and value != "normal":
        return _reason(
            "unsupported-pivot-feature",
            part=part, element=local, attribute=name, value=value,
        )
    if local == "dataField" and name == "subtotal" \
            and value not in _SUPPORTED_SUBTOTALS:
        return _reason(
            "unknown-aggregate",
            part=part, element=local, attribute=name, value=value,
        )
    if local in ("item", "i") and name == "t" and value not in _ITEM_TYPES:
        return _reason(
            "unsupported-pivot-feature",
            part=part, element=local, attribute=name, value=value,
        )
    if local == "cacheSource" and name == "type" and value != "worksheet":
        return _reason(
            "unsupported-source",
            part=part, element=local, attribute=name, value=value,
        )
    return None


def _classify_extensions(workbook, node, cache):
    reasons = []
    fingerprints = tuple(node.extension_fingerprints)
    if cache is not None:
        fingerprints = fingerprints + tuple(cache.extension_fingerprints)
    if not fingerprints:
        return reasons
    known = tuple(
        item for item in fingerprints
        if item.uri in EXTENSION_ALLOWLIST
    )
    unknown = tuple(
        item for item in fingerprints
        if item.uri not in EXTENSION_ALLOWLIST
    )
    if unknown:
        reasons.append(_reason(
            "unsupported-extension",
            part=node.identity.pivot_part,
            uri=unknown[0].uri,
        ))
        return reasons
    if known and not _extension_payloads_are_benign(workbook, node, cache):
        reasons.append(_reason(
            "unsupported-extension",
            part=node.identity.pivot_part,
            uri=known[0].uri,
            detail="nonempty-or-unrecognized-payload",
        ))
        return reasons
    if known:
        # Identified empty Excel compatibility payloads remain ineligible
        # until desktop Excel evidence proves they are semantically orthogonal.
        if not excel_equivalence_proved():
            reasons.append(_reason(
                "foreign-extension-unproved",
                part=node.identity.pivot_part,
                uri=known[0].uri,
            ))
    return reasons


def _is_ignored_attribute(raw, name, ns):
    if raw.startswith("xmlns") or name.startswith("xmlns"):
        return True
    if ns in (_MC_NS, _REVISION_NS):
        return True
    if name in _IGNORED_ATTR_LOCALS:
        return True
    if ns == REL_NS and name == "id":
        return True
    return False


def _namespace(tag):
    if isinstance(tag, str) and tag.startswith("{"):
        return tag[1:].split("}", 1)[0]
    return ""


def _reason(code, **context):
    items = tuple(sorted(
        (key, value) for key, value in context.items() if value is not None))
    return QualificationReason(None, code, items)
