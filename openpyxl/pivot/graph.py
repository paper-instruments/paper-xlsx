# paper-xlsx: relationship-resolved pivot graph

"""Read-only inventory of a workbook's PivotTable package graph.

Identity and ownership come from OPC relationships and workbook registry
entries, never from conventional filenames or module-level id counters.
Foreign pivot parts are scanned, not deserialized through inherited
serializers.
"""

from __future__ import annotations

import hashlib
import io
import zipfile
from dataclasses import dataclass, field
from types import MappingProxyType
from xml.etree.ElementTree import ParseError, fromstring

from openpyxl.errors import (
    RelationshipPolicyError,
    UnsupportedStructureError,
)
from openpyxl.xml.constants import REL_NS, SHEET_MAIN_NS

_PIVOT_TABLE_REL = "/pivotTable"
_PIVOT_CACHE_REL = "/pivotCacheDefinition"
_PIVOT_RECORDS_REL = "/pivotCacheRecords"
_REL_ID_NS = "{%s}id" % REL_NS

_SHARED_ITEM_TAGS = ("m", "n", "b", "e", "s", "d", "x")
_DATA_FIELD_FUNCS = {
    "sum": "sum",
    "count": "count",
    "countNums": "count_numbers",
    "average": "average",
    "min": "min",
    "max": "max",
    "product": "product",
    "stdDev": "std_dev",
    "stdDevP": "std_dev_p",
    "var": "var",
    "varP": "var_p",
}


def _local(name):
    if name.startswith("{"):
        return name.rsplit("}", 1)[-1]
    if ":" in name:
        return name.split(":", 1)[1]
    return name


def _attr(element, *names):
    attrib = element.attrib
    for name in names:
        if name in attrib:
            return attrib[name]
    for raw, value in attrib.items():
        local = _local(raw)
        if local in names:
            return value
    return None


def _rid(element):
    attrib = element.attrib
    if _REL_ID_NS in attrib:
        return attrib[_REL_ID_NS]
    for name, value in attrib.items():
        if name.startswith("{") and _local(name) == "id":
            return value
    return None


def _children(element, tag):
    return [child for child in list(element) if _local(child.tag) == tag]


def _first(element, tag):
    for child in list(element):
        if _local(child.tag) == tag:
            return child
    return None


def _iter_local(element, tag):
    for child in element.iter():
        if _local(child.tag) == tag:
            yield child


def _rels_path(part_name):
    folder, _, base = part_name.rpartition("/")
    return "{0}/_rels/{1}.rels".format(folder, base) if folder \
        else "_rels/{0}.rels".format(base)


def _resolve_target(from_part, target):
    if not target:
        return ""
    if target.startswith("/"):
        return target[1:]
    base = from_part.rpartition("/")[0].split("/") if "/" in from_part else []
    for piece in target.split("/"):
        if piece == "..":
            base = base[:-1]
        elif piece != ".":
            base.append(piece)
    return "/".join(base)


def _owner_of_rels(rels_part):
    if rels_part.startswith("_rels/"):
        base = rels_part[len("_rels/"):]
        return base[:-5] if base.endswith(".rels") else base
    folder, _, base = rels_part.rpartition("/_rels/")
    name = base[:-5] if base.endswith(".rels") else base
    return "{0}/{1}".format(folder, name) if folder else name


def _parse_xml(payload):
    try:
        return fromstring(payload)
    except (ParseError, ValueError, TypeError):
        return None


def _sha256(payload):
    return hashlib.sha256(payload).hexdigest()


def _reason(code, **context):
    items = tuple(sorted(
        (key, value) for key, value in context.items() if value is not None))
    return GraphReason(code, items)


@dataclass(frozen=True)
class GraphReason:
    code: str
    context: tuple = ()

    def to_dict(self):
        return {
            "code": self.code,
            "context": {key: value for key, value in self.context},
        }


@dataclass(frozen=True)
class IncomingRelationship:
    owner_part: str
    relationship_id: str
    rel_type: str
    target: str


@dataclass(frozen=True)
class SourceDescriptor:
    kind: str
    cache_source_type: str | None = None
    sheet: str | None = None
    ref: str | None = None
    name: str | None = None

    def to_dict(self):
        payload = {"kind": self.kind}
        if self.cache_source_type:
            payload["type"] = self.cache_source_type
        if self.sheet is not None:
            payload["sheet"] = self.sheet
        if self.ref is not None:
            payload["ref"] = self.ref
        if self.name is not None:
            payload["name"] = self.name
        return payload


@dataclass(frozen=True)
class ExtensionFingerprint:
    uri: str | None
    namespaces: tuple = ()

    def to_dict(self):
        return {"uri": self.uri, "namespaces": list(self.namespaces)}


@dataclass(frozen=True)
class PivotIdentity:
    worksheet_part: str
    pivot_part: str
    relationship_id: str
    name: str


@dataclass(frozen=True)
class PivotCacheNode:
    cache_id: str | None
    definition_part: str | None
    records_part: str | None
    source_descriptor: SourceDescriptor | None
    referenced_by: tuple
    records_relationship_id: str | None = None
    field_names: tuple = ()
    shared_item_kinds: tuple = ()
    declared_record_count: int | None = None
    actual_record_count: int | None = None
    has_grouping: bool = False
    has_calculated: bool = False
    extension_fingerprints: tuple = ()
    payload_sha256: str | None = None
    reasons: tuple = ()
    valid: bool = True

    def to_dict(self):
        return {
            "cache_id": self.cache_id,
            "definition_part": self.definition_part,
            "records_part": self.records_part,
            "source": None if self.source_descriptor is None
            else self.source_descriptor.to_dict(),
            "referenced_by": [
                "{0}!{1}".format(sheet, name)
                if sheet else name
                for sheet, name in self.referenced_by
            ],
            "field_names": list(self.field_names),
            "shared_item_kinds": [list(kinds) for kinds in self.shared_item_kinds],
            "declared_record_count": self.declared_record_count,
            "actual_record_count": self.actual_record_count,
            "has_grouping": self.has_grouping,
            "has_calculated": self.has_calculated,
            "extension_fingerprints": [
                item.to_dict() for item in self.extension_fingerprints
            ],
            "valid": self.valid,
            "reasons": [item.to_dict() for item in self.reasons],
        }


@dataclass(frozen=True)
class PivotNode:
    identity: PivotIdentity
    sheet_title: str
    cache_id: str | None
    cache_definition_part: str | None
    cache_records_part: str | None
    output_range: str | None
    source_descriptor: SourceDescriptor | None
    cache_relationship_id: str | None = None
    extension_fingerprints: tuple = ()
    tag: str | None = None
    created_version: str | None = None
    updated_version: str | None = None
    min_refreshable_version: str | None = None
    field_count: int | None = None
    row_fields: tuple = ()
    column_fields: tuple = ()
    page_fields: tuple = ()
    data_fields: tuple = ()
    payload_sha256: str | None = None
    parse_error: str | None = None
    reasons: tuple = ()
    valid: bool = True

    def to_dict(self):
        return {
            "name": self.identity.name,
            "sheet": self.sheet_title,
            "worksheet_part": self.identity.worksheet_part,
            "pivot_part": self.identity.pivot_part,
            "relationship_id": self.identity.relationship_id,
            "cache_id": self.cache_id,
            "cache_definition_part": self.cache_definition_part,
            "cache_records_part": self.cache_records_part,
            "output_range": self.output_range,
            "source": None if self.source_descriptor is None
            else self.source_descriptor.to_dict(),
            "tag": self.tag,
            "created_version": self.created_version,
            "updated_version": self.updated_version,
            "min_refreshable_version": self.min_refreshable_version,
            "field_count": self.field_count,
            "row_fields": list(self.row_fields),
            "column_fields": list(self.column_fields),
            "page_fields": list(self.page_fields),
            "data_fields": [dict(item) for item in self.data_fields],
            "extension_fingerprints": [
                item.to_dict() for item in self.extension_fingerprints
            ],
            "valid": self.valid,
            "reasons": [item.to_dict() for item in self.reasons],
        }


@dataclass(frozen=True)
class PivotGraph:
    workbook_part: str
    caches_by_id: object
    caches_by_part: object
    pivots_by_identity: object
    incoming_relationships: object
    registered_cache_parts: tuple
    reasons: tuple = ()
    _pivots: tuple = field(default=(), repr=False)
    _caches: tuple = field(default=(), repr=False)

    @property
    def pivots(self):
        return self._pivots

    @property
    def caches(self):
        return self._caches

    def raise_for_refresh_index(self):
        """Preserve the refresh helper's historical hard-fail on bad pivot XML."""
        for node in self._pivots:
            if node.parse_error == "malformed":
                raise UnsupportedStructureError(
                    "pivot table definition XML is malformed",
                    kind="unsupported-pivot-table")
            if node.parse_error == "unexpected-root":
                raise UnsupportedStructureError(
                    "pivot table definition has an unexpected root element",
                    kind="unsupported-pivot-table")

    def qualified_name_map(self, current_by_original=None):
        """Name -> ((current-title, cache-part), ...) for refresh targeting.

        Matches the historical index: only named pivots whose cacheId resolves
        through the workbook registry to an existing cache part.
        """
        titles = current_by_original or {}
        qualified = {}
        registered = {
            cache_id: node.definition_part
            for cache_id, node in self.caches_by_id.items()
            if node.definition_part
        }
        for node in self._pivots:
            if node.parse_error:
                continue
            cache_part = registered.get(node.cache_id)
            if not node.identity.name or not cache_part:
                continue
            title = titles.get(node.sheet_title, node.sheet_title)
            qualified.setdefault(node.identity.name, []).append(
                (title, cache_part))
        return qualified

    def to_dict(self):
        return {
            "workbook_part": self.workbook_part,
            "caches": [node.to_dict() for node in self._caches],
            "pivots": [node.to_dict() for node in self._pivots],
            "reasons": [item.to_dict() for item in self.reasons],
        }


def load_pivot_graph(source, workbook=None):
    """Parse a relationship-resolved pivot graph from package bytes."""
    if hasattr(source, "read"):
        payload = source.read()
        if hasattr(source, "seek"):
            source.seek(0)
    else:
        payload = source
    with zipfile.ZipFile(io.BytesIO(payload)) as zin:
        return _load_from_zip(zin, workbook)


def load_workbook_pivot_graph(wb):
    """Parse the retained preserve-mode package of ``wb``."""
    source = getattr(wb, "_paper_source", None)
    if source is None:
        raise RelationshipPolicyError(
            "pivot graph inspection requires a preserve-mode workbook that "
            "retained its source package. Nothing was changed.",
            kind="invalid-pivot-graph")
    return load_pivot_graph(source, workbook=wb)


def _load_from_zip(zin, workbook):
    from openpyxl.preserve.saver import _package_info
    from openpyxl.preserve.xmlscan import ScanRefusal

    names = set(zin.namelist())
    try:
        workbook_part, sheets = _package_info(zin)
    except Exception as exc:
        raise RelationshipPolicyError(
            "cannot resolve the workbook part graph for pivot inventory: "
            "{0}. Nothing was changed.".format(exc),
            kind="invalid-pivot-graph") from exc

    incoming, rels_by_owner = _scan_relationships(zin, names)
    graph_reasons = []
    graph_reasons.extend(_duplicate_incoming_reasons(incoming))

    try:
        registered, registry_reasons = _workbook_cache_registry(
            zin, names, workbook_part, rels_by_owner)
    except ScanRefusal as exc:
        raise RelationshipPolicyError(
            "workbook.xml is not a usable pivot registry: {0}. Nothing was "
            "changed.".format(exc),
            kind="invalid-pivot-graph") from exc
    graph_reasons.extend(registry_reasons)

    cache_parts = set()
    for cache_id, part, _rid, _ok in registered:
        if part and part in names:
            cache_parts.add(part)
    for owner_part, rels in rels_by_owner.items():
        for rel in rels:
            if rel.rel_type.endswith(_PIVOT_CACHE_REL) and rel.target in names:
                cache_parts.add(rel.target)

    named_kinds = _named_source_kinds(zin, names, workbook_part, rels_by_owner)
    cache_nodes = []
    for part in sorted(cache_parts):
        cache_ids = [
            cache_id for cache_id, cache_part, _rid, _ok in registered
            if cache_part == part
        ]
        cache_nodes.append(_parse_cache(
            zin, names, part, cache_ids, rels_by_owner, incoming, workbook,
            named_kinds))

    pivots = []
    for sheet_title, sheet_part in sheets.items():
        sheet_rels = [
            rel for rel in rels_by_owner.get(sheet_part, ())
            if rel.rel_type.endswith(_PIVOT_TABLE_REL)
        ]
        for rel in sheet_rels:
            pivots.append(_parse_pivot(
                zin, names, sheet_title, sheet_part, rel, cache_nodes,
                incoming, rels_by_owner, workbook))

    caches_by_id, caches_by_part, _id_reasons = _index_caches(cache_nodes)
    pivots = tuple(_rebind_pivot_caches(node, caches_by_id) for node in pivots)
    cache_nodes = _attach_cache_references(cache_nodes, pivots)
    caches_by_id, caches_by_part, id_reasons = _index_caches(cache_nodes)
    graph_reasons.extend(id_reasons)
    identities = {}
    for node in pivots:
        identities[node.identity] = node

    registered_parts = tuple(sorted({
        part for _cache_id, part, _rid, ok in registered if ok and part
    }))
    return PivotGraph(
        workbook_part=workbook_part,
        caches_by_id=MappingProxyType(caches_by_id),
        caches_by_part=MappingProxyType(caches_by_part),
        pivots_by_identity=MappingProxyType(identities),
        incoming_relationships=MappingProxyType({
            part: tuple(items) for part, items in incoming.items()
        }),
        registered_cache_parts=registered_parts,
        reasons=tuple(graph_reasons),
        _pivots=tuple(pivots),
        _caches=tuple(cache_nodes),
    )


def _scan_relationships(zin, names):
    incoming = {}
    by_owner = {}
    for rels_part in sorted(name for name in names if name.endswith(".rels")):
        owner = _owner_of_rels(rels_part)
        root = _parse_xml(zin.read(rels_part))
        if root is None or _local(root.tag) != "Relationships":
            continue
        rels = []
        for child in list(root):
            if _local(child.tag) != "Relationship":
                continue
            if child.attrib.get("TargetMode") == "External":
                continue
            rel_id = child.attrib.get("Id", "")
            rel_type = child.attrib.get("Type", "")
            target = _resolve_target(owner, child.attrib.get("Target", ""))
            rel = IncomingRelationship(owner, rel_id, rel_type, target)
            rels.append(rel)
            incoming.setdefault(target, []).append(rel)
        by_owner[owner] = tuple(rels)
    return incoming, by_owner


def _resolve_internal_relationship(
        root, owner_part, rels_by_owner, expected_type, link,
        *, allow_implicit=False, required=True):
    """Resolve one part-level relationship without guessing past ``r:id``."""
    rels = rels_by_owner.get(owner_part, ())
    rid = _rid(root)
    if rid is not None:
        if not rid:
            return None, (_reason(
                "missing-relationship-id", part=owner_part, link=link),)
        matches = [rel for rel in rels if rel.relationship_id == rid]
        if not matches:
            return None, (_reason(
                "missing-internal-relationship", part=owner_part, rid=rid,
                link=link),)
        if len(matches) != 1:
            return None, (_reason(
                "duplicate-relationship-id", part=owner_part, rid=rid,
                link=link),)
        rel = matches[0]
        if not rel.rel_type.endswith(expected_type):
            return None, (_reason(
                "relationship-type-mismatch", part=owner_part, rid=rid,
                link=link, rel_type=rel.rel_type),)
        return rel, ()

    typed = [rel for rel in rels if rel.rel_type.endswith(expected_type)]
    if allow_implicit and len(typed) == 1:
        rel = typed[0]
        if not rel.relationship_id:
            return None, (_reason(
                "missing-relationship-id", part=owner_part, link=link),)
        matches = [
            item for item in rels
            if item.relationship_id == rel.relationship_id
        ]
        if len(matches) != 1:
            return None, (_reason(
                "duplicate-relationship-id", part=owner_part,
                rid=rel.relationship_id, link=link),)
        return rel, ()
    if allow_implicit and len(typed) > 1:
        relationship_ids = [rel.relationship_id for rel in typed]
        if len(set(relationship_ids)) != len(relationship_ids):
            return None, (_reason(
                "duplicate-relationship-id", part=owner_part, link=link),)
        return None, (_reason(
            "ambiguous-internal-relationship", part=owner_part, link=link),)
    if allow_implicit and required:
        if rels:
            return None, (_reason(
                "relationship-type-mismatch", part=owner_part, link=link),)
        return None, (_reason(
            "missing-internal-relationship", part=owner_part, link=link),)
    if required:
        return None, (_reason(
            "missing-relationship-id", part=owner_part, link=link),)
    return None, ()


def _duplicate_incoming_reasons(incoming):
    reasons = []
    for part, rels in sorted(incoming.items()):
        owners = {}
        for rel in rels:
            if rel.rel_type.endswith((_PIVOT_TABLE_REL, _PIVOT_CACHE_REL,
                                      _PIVOT_RECORDS_REL)):
                owners.setdefault(rel.rel_type, []).append(rel)
        for rel_type, typed in owners.items():
            owners_of_type = {rel.owner_part for rel in typed}
            if rel_type.endswith(_PIVOT_CACHE_REL):
                continue
            pivot_owned_twice = (
                rel_type.endswith(_PIVOT_TABLE_REL)
                and len(owners_of_type) > 1)
            records_owned_twice = (
                rel_type.endswith(_PIVOT_RECORDS_REL) and len(typed) > 1)
            if pivot_owned_twice or records_owned_twice:
                reasons.append(_reason(
                    "duplicate-incoming",
                    part=part,
                    rel_type=rel_type,
                    owners=",".join(sorted(owners_of_type)),
                ))
    return reasons


def _workbook_cache_registry(zin, names, workbook_part, rels_by_owner):
    from openpyxl.preserve import crosspart

    root = crosspart.scan_small(zin.read(workbook_part), "workbook",
                                max_depth=3)
    cache_targets = {
        rel.relationship_id: rel.target
        for rel in rels_by_owner.get(workbook_part, ())
        if rel.rel_type.endswith(_PIVOT_CACHE_REL)
    }
    registered = []
    reasons = []
    seen_ids = {}
    for group in root.children:
        if group.local() != "pivotCaches":
            continue
        for child in group.children:
            if child.local() != "pivotCache":
                continue
            cache_id = child.attrs.get("cacheId")
            rid = child.attrs.get("id") or child.attrs.get("r:id")
            target = cache_targets.get(rid)
            exists = bool(target) and target in names
            if cache_id is None:
                reasons.append(_reason(
                    "missing-cache-id", workbook_part=workbook_part, rid=rid))
            elif cache_id in seen_ids:
                reasons.append(_reason(
                    "duplicate-cache-id", cache_id=cache_id,
                    parts=",".join(sorted({seen_ids[cache_id], target or ""}))))
            else:
                seen_ids[cache_id] = target or ""
            if rid and target is None:
                reasons.append(_reason(
                    "dangling-workbook-cache", cache_id=cache_id, rid=rid))
            elif target and not exists:
                reasons.append(_reason(
                    "missing-part", cache_id=cache_id, part=target,
                    link="workbook-to-cache"))
            registered.append((cache_id, target, rid, exists))
    return registered, reasons


def _parse_cache(zin, names, part, cache_ids, rels_by_owner, incoming,
                 workbook, named_kinds=None):
    payload = zin.read(part)
    reasons = []
    root = _parse_xml(payload)
    if root is None:
        return PivotCacheNode(
            cache_id=cache_ids[0] if len(cache_ids) == 1 else None,
            definition_part=part,
            records_part=None,
            source_descriptor=None,
            referenced_by=(),
            payload_sha256=_sha256(payload),
            reasons=(_reason("malformed-xml", part=part),),
            valid=False,
        )
    if _local(root.tag) != "pivotCacheDefinition":
        return PivotCacheNode(
            cache_id=cache_ids[0] if len(cache_ids) == 1 else None,
            definition_part=part,
            records_part=None,
            source_descriptor=None,
            referenced_by=(),
            payload_sha256=_sha256(payload),
            reasons=(_reason("unexpected-root", part=part, tag=root.tag),),
            valid=False,
        )

    declared = _attr(root, "recordCount")
    try:
        declared_count = int(declared) if declared is not None else None
    except (TypeError, ValueError):
        declared_count = None
        reasons.append(_reason("invalid-record-count", part=part, value=declared))

    has_records_link = _rid(root) is not None or any(
        rel.rel_type.endswith(_PIVOT_RECORDS_REL)
        for rel in rels_by_owner.get(part, ()))
    records_rel, relationship_reasons = _resolve_internal_relationship(
        root,
        part,
        rels_by_owner,
        _PIVOT_RECORDS_REL,
        "cache-to-records",
        required=has_records_link or bool(declared_count),
    )
    reasons.extend(relationship_reasons)
    records_part = records_rel.target if records_rel is not None else None
    if records_part and records_part not in names:
        reasons.append(_reason(
            "dangling-cache-records", part=part, records_part=records_part))
        records_actual = None
    elif records_part:
        records_actual = _count_records(zin.read(records_part))
        if records_actual is None:
            reasons.append(_reason("malformed-xml", part=records_part))
    else:
        records_actual = None

    incoming_records = [
        rel for rel in incoming.get(records_part or "", ())
        if rel.rel_type.endswith(_PIVOT_RECORDS_REL)
    ]
    if records_part and len(incoming_records) > 1:
        reasons.append(_reason(
            "duplicate-incoming", part=records_part, owner=part))

    if (declared_count is not None and records_actual is not None
            and declared_count != records_actual):
        reasons.append(_reason(
            "record-count-mismatch", part=part,
            declared=str(declared_count), actual=str(records_actual)))

    source = _source_descriptor(root, workbook, named_kinds)
    fields, kinds, grouping, calculated, field_reasons = _cache_fields(root)
    reasons.extend(field_reasons)
    fingerprints = _extension_fingerprints(root)
    cache_id = cache_ids[0] if len(cache_ids) == 1 else None
    if len(cache_ids) > 1:
        reasons.append(_reason(
            "duplicate-cache-id", part=part,
            cache_ids=",".join(item or "" for item in cache_ids)))
    valid = not reasons
    return PivotCacheNode(
        cache_id=cache_id,
        definition_part=part,
        records_part=records_part if records_part in names else None,
        source_descriptor=source,
        referenced_by=(),
        records_relationship_id=(
            records_rel.relationship_id if records_rel is not None else None),
        field_names=tuple(fields),
        shared_item_kinds=tuple(kinds),
        declared_record_count=declared_count,
        actual_record_count=records_actual,
        has_grouping=grouping,
        has_calculated=calculated,
        extension_fingerprints=tuple(fingerprints),
        payload_sha256=_sha256(payload),
        reasons=tuple(reasons),
        valid=valid,
    )


def _count_records(payload):
    root = _parse_xml(payload)
    if root is None or _local(root.tag) != "pivotCacheRecords":
        return None
    return sum(1 for child in list(root) if _local(child.tag) == "r")


def _source_descriptor(cache_root, workbook=None, named_kinds=None):
    source = _first(cache_root, "cacheSource")
    if source is None:
        return SourceDescriptor(kind="unknown")
    source_type = _attr(source, "type")
    if source_type == "external":
        return SourceDescriptor(kind="external", cache_source_type=source_type)
    if source_type == "consolidation":
        return SourceDescriptor(
            kind="consolidation", cache_source_type=source_type)
    if source_type == "scenario":
        return SourceDescriptor(kind="scenario", cache_source_type=source_type)
    worksheet = _first(source, "worksheetSource")
    if worksheet is None:
        return SourceDescriptor(
            kind="unknown", cache_source_type=source_type or "unknown")
    sheet = _attr(worksheet, "sheet")
    ref = _attr(worksheet, "ref")
    name = _attr(worksheet, "name")
    if ref and sheet and not name:
        return SourceDescriptor(
            kind="range", cache_source_type=source_type or "worksheet",
            sheet=sheet, ref=ref)
    if name and not ref:
        kind = (named_kinds or {}).get(name.casefold())
        if kind is None:
            kind = _classify_named_source(workbook, name)
        return SourceDescriptor(
            kind=kind, cache_source_type=source_type or "worksheet",
            sheet=sheet, name=name)
    if name and ref:
        return SourceDescriptor(
            kind="unknown", cache_source_type=source_type or "worksheet",
            sheet=sheet, ref=ref, name=name)
    return SourceDescriptor(
        kind="unknown", cache_source_type=source_type or "worksheet",
        sheet=sheet, ref=ref, name=name)


def _named_source_kinds(zin, names, workbook_part, rels_by_owner):
    kinds = {}
    root = _parse_xml(zin.read(workbook_part))
    if root is not None:
        for defined in _iter_local(root, "definedName"):
            declared = defined.attrib.get("name")
            if declared:
                kinds.setdefault(declared.casefold(), set()).add("defined-name")
    for owner, rels in rels_by_owner.items():
        for rel in rels:
            if not rel.rel_type.endswith("/table") or rel.target not in names:
                continue
            table = _parse_xml(zin.read(rel.target))
            if table is None:
                continue
            for attr in ("name", "displayName"):
                value = _attr(table, attr)
                if value:
                    kinds.setdefault(value.casefold(), set()).add("table")
    resolved = {}
    for folded, found in kinds.items():
        if found == {"table"}:
            resolved[folded] = "table"
        elif found == {"defined-name"}:
            resolved[folded] = "defined-name"
        else:
            resolved[folded] = "named"
    return resolved


def _classify_named_source(workbook, name):
    if workbook is None:
        return "named"
    folded = name.casefold()
    table_hit = False
    defined_hit = False
    for ws in getattr(workbook, "worksheets", ()):
        for key, table in (getattr(ws, "tables", None) or {}).items():
            declared = (key, getattr(table, "name", None),
                        getattr(table, "displayName", None))
            if any(isinstance(item, str) and item.casefold() == folded
                   for item in declared):
                table_hit = True
    holders = [getattr(workbook, "defined_names", {})]
    holders.extend(ws.defined_names for ws in getattr(workbook, "worksheets", ()))
    for holder in holders:
        for key, defn in (holder or {}).items():
            declared = getattr(defn, "name", None)
            if ((isinstance(key, str) and key.casefold() == folded)
                    or (isinstance(declared, str)
                        and declared.casefold() == folded)):
                defined_hit = True
    if table_hit and not defined_hit:
        return "table"
    if defined_hit and not table_hit:
        return "defined-name"
    if table_hit and defined_hit:
        return "named"
    return "named"


def _cache_fields(root):
    container = _first(root, "cacheFields")
    names = []
    kinds = []
    grouping = False
    calculated = False
    reasons = []
    if container is None:
        return names, kinds, grouping, calculated, reasons
    declared = _attr(container, "count")
    fields = _children(container, "cacheField")
    if declared is not None:
        try:
            if int(declared) != len(fields):
                reasons.append(_reason(
                    "field-count-mismatch", declared=declared,
                    actual=str(len(fields))))
        except (TypeError, ValueError):
            reasons.append(_reason("invalid-field-count", value=declared))
    for field in fields:
        names.append(_attr(field, "name") or "")
        shared = _first(field, "sharedItems")
        field_kinds = []
        if shared is not None:
            for child in list(shared):
                tag = _local(child.tag)
                if tag in _SHARED_ITEM_TAGS:
                    field_kinds.append(tag)
        kinds.append(tuple(field_kinds))
        if _first(field, "fieldGroup") is not None:
            grouping = True
        if _attr(field, "formula"):
            calculated = True
    if _first(root, "calculatedItems") is not None \
            or _first(root, "calculatedMembers") is not None:
        calculated = True
    if _first(root, "cacheHierarchies") is not None:
        grouping = True
    return names, kinds, grouping, calculated, reasons


def _extension_fingerprints(root):
    fingerprints = []
    for ext_lst in _iter_local(root, "extLst"):
        for ext in _children(ext_lst, "ext"):
            namespaces = []
            uri = _attr(ext, "uri")
            for key, value in ext.attrib.items():
                if key.startswith("xmlns"):
                    namespaces.append(value)
                elif key.startswith("{http://www.w3.org/2000/xmlns/}"):
                    namespaces.append(value)
            for child in ext.iter():
                if child is ext:
                    continue
                if child.tag.startswith("{"):
                    namespaces.append(child.tag[1:].split("}", 1)[0])
            fingerprints.append(ExtensionFingerprint(
                uri, tuple(sorted(set(namespaces)))))
    return fingerprints


def _parse_pivot(zin, names, sheet_title, sheet_part, rel, cache_nodes,
                 incoming, rels_by_owner, workbook):
    identity = PivotIdentity(
        worksheet_part=sheet_part,
        pivot_part=rel.target,
        relationship_id=rel.relationship_id,
        name="",
    )
    if rel.target not in names:
        return PivotNode(
            identity=identity,
            sheet_title=sheet_title,
            cache_id=None,
            cache_definition_part=None,
            cache_records_part=None,
            output_range=None,
            source_descriptor=None,
            reasons=(_reason(
                "dangling-sheet-pivot", sheet=sheet_title,
                part=rel.target, rid=rel.relationship_id),),
            valid=False,
        )

    payload = zin.read(rel.target)
    root = _parse_xml(payload)
    if root is None:
        return PivotNode(
            identity=identity,
            sheet_title=sheet_title,
            cache_id=None,
            cache_definition_part=None,
            cache_records_part=None,
            output_range=None,
            source_descriptor=None,
            payload_sha256=_sha256(payload),
            parse_error="malformed",
            reasons=(_reason("malformed-xml", part=rel.target),),
            valid=False,
        )
    if _local(root.tag) != "pivotTableDefinition":
        return PivotNode(
            identity=identity,
            sheet_title=sheet_title,
            cache_id=None,
            cache_definition_part=None,
            cache_records_part=None,
            output_range=None,
            source_descriptor=None,
            payload_sha256=_sha256(payload),
            parse_error="unexpected-root",
            reasons=(_reason("unexpected-root", part=rel.target, tag=root.tag),),
            valid=False,
        )

    reasons = []
    name = _attr(root, "name") or ""
    identity = PivotIdentity(
        worksheet_part=sheet_part,
        pivot_part=rel.target,
        relationship_id=rel.relationship_id,
        name=name,
    )
    if not name:
        reasons.append(_reason("missing-name", part=rel.target, sheet=sheet_title))

    cache_id = _attr(root, "cacheId")
    cache_rel, relationship_reasons = _resolve_internal_relationship(
        root,
        rel.target,
        rels_by_owner,
        _PIVOT_CACHE_REL,
        "pivot-to-cache",
        allow_implicit=True,
    )
    reasons.extend(relationship_reasons)
    cache_from_rel = cache_rel.target if cache_rel is not None else None
    if cache_from_rel and cache_from_rel not in names:
        reasons.append(_reason(
            "dangling-pivot-cache", part=rel.target, cache_part=cache_from_rel))

    cache_node = _lookup_cache(cache_nodes, cache_id, cache_from_rel)
    if cache_id is None:
        reasons.append(_reason("missing-cache-id", part=rel.target))
    if cache_node is None and (cache_id or cache_from_rel):
        reasons.append(_reason(
            "dangling-pivot-cache", part=rel.target, cache_id=cache_id,
            cache_part=cache_from_rel))
    if (cache_node is not None and cache_from_rel
            and cache_node.definition_part != cache_from_rel):
        reasons.append(_reason(
            "cache-relationship-mismatch", part=rel.target,
            cache_id=cache_id, rel_target=cache_from_rel,
            registry_part=cache_node.definition_part))
    if (cache_node is not None and cache_id is not None
            and cache_node.cache_id != cache_id):
        reasons.append(_reason(
            "cache-relationship-mismatch", part=rel.target,
            cache_id=cache_id, rel_target=cache_from_rel,
            registry_cache_id=cache_node.cache_id))

    location = _first(root, "location")
    output_range = _attr(location, "ref") if location is not None else None
    field_count, row_fields, column_fields, page_fields, data_fields, \
        field_reasons = _pivot_fields(root)
    reasons.extend(field_reasons)
    owners = [
        item for item in incoming.get(rel.target, ())
        if item.rel_type.endswith(_PIVOT_TABLE_REL)
    ]
    if len(owners) > 1:
        reasons.append(_reason(
            "duplicate-incoming", part=rel.target,
            owners=",".join(sorted(item.owner_part for item in owners))))

    source = cache_node.source_descriptor if cache_node is not None else None
    if source is not None and source.kind == "named" and workbook is not None:
        source = _source_descriptor(
            fromstring(zin.read(cache_node.definition_part)), workbook)

    valid = not reasons
    return PivotNode(
        identity=identity,
        sheet_title=sheet_title,
        cache_id=cache_id,
        cache_definition_part=(
            cache_node.definition_part if cache_node is not None else (
                cache_from_rel if cache_from_rel in names else None)),
        cache_records_part=(
            cache_node.records_part if cache_node is not None else None),
        output_range=output_range,
        source_descriptor=source,
        cache_relationship_id=(
            cache_rel.relationship_id if cache_rel is not None else None),
        extension_fingerprints=tuple(_extension_fingerprints(root)),
        tag=_attr(root, "tag"),
        created_version=_attr(root, "createdVersion"),
        updated_version=_attr(root, "updatedVersion"),
        min_refreshable_version=_attr(root, "minRefreshableVersion"),
        field_count=field_count,
        row_fields=tuple(row_fields),
        column_fields=tuple(column_fields),
        page_fields=tuple(page_fields),
        data_fields=tuple(data_fields),
        payload_sha256=_sha256(payload),
        reasons=tuple(reasons),
        valid=valid,
    )


def _lookup_cache(cache_nodes, cache_id, cache_part):
    if cache_part:
        for node in cache_nodes:
            if node.definition_part == cache_part:
                return node
    if cache_id is not None:
        matches = [node for node in cache_nodes if node.cache_id == cache_id]
        if len(matches) == 1:
            return matches[0]
    return None


def _pivot_fields(root):
    reasons = []
    container = _first(root, "pivotFields")
    fields = _children(container, "pivotField") if container is not None else []
    declared = _attr(container, "count") if container is not None else None
    if declared is not None:
        try:
            if int(declared) != len(fields):
                reasons.append(_reason(
                    "field-count-mismatch", declared=declared,
                    actual=str(len(fields))))
        except (TypeError, ValueError):
            reasons.append(_reason("invalid-field-count", value=declared))

    def indexes(tag, child_tag="field", attr="x"):
        node = _first(root, tag)
        if node is None:
            return []
        values = []
        for child in _children(node, child_tag):
            raw = _attr(child, attr)
            if raw is None:
                continue
            try:
                values.append(int(raw))
            except (TypeError, ValueError):
                reasons.append(_reason(
                    "invalid-field-index", parent=tag, value=raw))
        return values

    row_fields = indexes("rowFields")
    column_fields = indexes("colFields")
    page_fields = []
    page = _first(root, "pageFields")
    if page is not None:
        for child in _children(page, "pageField"):
            raw = _attr(child, "fld")
            try:
                page_fields.append(int(raw))
            except (TypeError, ValueError):
                reasons.append(_reason(
                    "invalid-field-index", parent="pageFields", value=raw))

    data_fields = []
    data = _first(root, "dataFields")
    if data is not None:
        for child in _children(data, "dataField"):
            raw = _attr(child, "fld")
            func = _attr(child, "subtotal") or "sum"
            try:
                field_index = int(raw)
            except (TypeError, ValueError):
                reasons.append(_reason(
                    "invalid-field-index", parent="dataFields", value=raw))
                continue
            data_fields.append(MappingProxyType({
                "field": field_index,
                "aggregate": _DATA_FIELD_FUNCS.get(func, func),
                "name": _attr(child, "name"),
                "number_format_id": _attr(child, "numFmtId"),
                "show_data_as": _attr(child, "showDataAs") or "normal",
                "base_field": _attr(child, "baseField"),
                "base_item": _attr(child, "baseItem"),
            }))
    return len(fields), row_fields, column_fields, page_fields, data_fields, reasons


def _attach_cache_references(cache_nodes, pivots):
    refs = {node.definition_part: [] for node in cache_nodes if node.definition_part}
    for pivot in pivots:
        part = pivot.cache_definition_part
        if part in refs and pivot.identity.name:
            refs[part].append((pivot.sheet_title, pivot.identity.name))
    attached = []
    for node in cache_nodes:
        referenced = tuple(sorted(refs.get(node.definition_part, ())))
        attached.append(PivotCacheNode(
            cache_id=node.cache_id,
            definition_part=node.definition_part,
            records_part=node.records_part,
            source_descriptor=node.source_descriptor,
            referenced_by=referenced,
            records_relationship_id=node.records_relationship_id,
            field_names=node.field_names,
            shared_item_kinds=node.shared_item_kinds,
            declared_record_count=node.declared_record_count,
            actual_record_count=node.actual_record_count,
            has_grouping=node.has_grouping,
            has_calculated=node.has_calculated,
            extension_fingerprints=node.extension_fingerprints,
            payload_sha256=node.payload_sha256,
            reasons=node.reasons,
            valid=node.valid,
        ))
    return attached


def _index_caches(cache_nodes):
    by_id = {}
    by_part = {}
    reasons = []
    grouped = {}
    for node in cache_nodes:
        if node.definition_part:
            by_part[node.definition_part] = node
        if node.cache_id is None:
            continue
        grouped.setdefault(node.cache_id, []).append(node)
    for cache_id, nodes in grouped.items():
        if len(nodes) == 1:
            by_id[cache_id] = nodes[0]
        else:
            reasons.append(_reason(
                "duplicate-cache-id",
                cache_id=cache_id,
                parts=",".join(
                    sorted(node.definition_part or "" for node in nodes)),
            ))
    return by_id, by_part, reasons


def _rebind_pivot_caches(node, caches_by_id):
    cache = caches_by_id.get(node.cache_id)
    if cache is None:
        return node
    if (node.cache_definition_part
            and cache.definition_part
            and node.cache_definition_part != cache.definition_part):
        return node
    return PivotNode(
        identity=node.identity,
        sheet_title=node.sheet_title,
        cache_id=node.cache_id,
        cache_definition_part=cache.definition_part,
        cache_records_part=cache.records_part,
        output_range=node.output_range,
        source_descriptor=cache.source_descriptor or node.source_descriptor,
        cache_relationship_id=node.cache_relationship_id,
        extension_fingerprints=node.extension_fingerprints,
        tag=node.tag,
        created_version=node.created_version,
        updated_version=node.updated_version,
        min_refreshable_version=node.min_refreshable_version,
        field_count=node.field_count,
        row_fields=node.row_fields,
        column_fields=node.column_fields,
        page_fields=node.page_fields,
        data_fields=node.data_fields,
        payload_sha256=node.payload_sha256,
        parse_error=node.parse_error,
        reasons=node.reasons,
        valid=node.valid,
    )
