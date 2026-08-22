# paper-xlsx: read-only foreign PivotTable adoption qualification

"""Answer whether one loaded foreign pivot can become Paper-managed.

This module does not add ``adopt()`` and does not mutate the workbook.
Ordinary inspection remains metadata-bounded; cache reconstruction and
dependency analysis run only here.
"""

from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass

from openpyxl.pivot.adopt_depend import (
    scan_selected_dependencies,
    source_requires_calculation,
)
from openpyxl.pivot.adopt_evidence import excel_equivalence_proved
from openpyxl.pivot.adopt_inventory import classify_selected_graph
from openpyxl.pivot.adopt_ownership import prove_foreign_output
from openpyxl.pivot.api_types import PivotAdoptionQualification
from openpyxl.pivot.qualify import (
    PAPER_TAG,
    QualificationReason,
    _graph_reasons,
    _unambiguous_cache,
)


@dataclass(frozen=True)
class AdoptionAnalysis:
    """Internal adoption-qualification result for tests and later PRs."""

    public: PivotAdoptionQualification
    pivot_part: str | None
    cache_part: str | None
    records_part: str | None
    payload_hashes: tuple
    cache_dependents: tuple
    ownership: object
    origin: str


def qualify_adoption(handle):
    """Return the public read-only adoption qualification for ``handle``."""
    return analyze_adoption(handle).public


def analyze_adoption(handle):
    """Qualify one selected pivot without mutating the workbook."""
    state = handle._state()
    worksheet = handle._worksheet
    workbook = worksheet.parent
    session = getattr(workbook, "_paper_pivot_session", None)
    graph = None if session is None else session.graph
    node = None
    if graph is not None:
        node = graph.pivots_by_identity.get(handle._identity)
    origin = state.qualification.origin
    if origin == "paper" or (node is not None and node.tag == PAPER_TAG):
        return _already_managed(handle, node)
    reasons = []
    if node is None or graph is None:
        reasons.append(_reason("invalid-pivot-graph",
                               part=handle._identity.pivot_part))
        return _result(
            handle, reasons, (), None, False, None, None, origin=origin)

    cache = None
    if node.cache_definition_part:
        cache = graph.caches_by_part.get(node.cache_definition_part)
    if cache is None and node.cache_id:
        cache = graph.caches_by_id.get(node.cache_id)

    reasons.extend(_graph_reasons(node, cache, graph))
    valid = node.valid and (cache is None or cache.valid)
    if not valid or not _unambiguous_cache(node, cache, graph):
        reasons.append(_reason(
            "invalid-pivot-graph",
            part=node.identity.pivot_part,
            cache_id=node.cache_id,
        ))

    reasons.extend(_package_format_reasons(workbook))

    projection = state.projection
    descriptor = node.source_descriptor
    if descriptor is None and cache is not None:
        descriptor = cache.source_descriptor
    source_kind = None if projection.source is None else projection.source.kind
    if source_kind == "defined-name" or (
            descriptor is not None
            and descriptor.kind in ("defined-name", "named")):
        reasons.append(_reason(
            "unsupported-pivot-source",
            kind="defined-name",
            part=None if cache is None else cache.definition_part,
        ))
    if not projection.complete or projection.spec is None:
        reasons.append(_reason(
            "foreign-semantic-incomplete",
            part=node.identity.pivot_part,
        ))
        for item in projection.reasons:
            reasons.append(QualificationReason(None, item.code, item.context))

    if cache is not None and cache.has_grouping:
        reasons.append(_reason(
            "unsupported-grouping",
            part=cache.definition_part,
        ))
    if cache is not None and cache.has_calculated:
        reasons.append(_reason(
            "unsupported-calculated",
            part=cache.definition_part,
        ))

    if cache is None or not cache.records_part \
            or cache.actual_record_count in (None, 0):
        reasons.append(_reason(
            "foreign-cache-records-unavailable",
            part=node.identity.pivot_part,
        ))

    reasons.extend(classify_selected_graph(workbook, node, cache))

    ownership, ownership_reasons = prove_foreign_output(
        workbook, node, projection)
    reasons.extend(ownership_reasons)

    strategy = _strategy(cache)
    if strategy == "shared-isolation" and cache is not None \
            and not cache.referenced_by:
        reasons.append(_reason(
            "foreign-cache-isolation-unproved",
            part=cache.definition_part,
        ))

    dep_blockers, constraints, dependents = scan_selected_dependencies(
        workbook, node, cache, graph, projection, ownership)
    reasons.extend(dep_blockers)

    requires_calc, engine, excluded = source_requires_calculation(
        workbook, projection)
    reasons.extend(excluded)

    if not excel_equivalence_proved():
        reasons.append(_reason(
            "foreign-managed-equivalence-unproved",
            part=node.identity.pivot_part,
        ))

    hashes = []
    if node.payload_sha256:
        hashes.append(("pivot", node.payload_sha256))
    if cache is not None and cache.payload_sha256:
        hashes.append(("cache", cache.payload_sha256))
    return _result(
        handle, reasons, constraints, strategy, requires_calc, engine,
        AdoptionAnalysis(
            public=None,
            pivot_part=node.identity.pivot_part,
            cache_part=None if cache is None else cache.definition_part,
            records_part=None if cache is None else cache.records_part,
            payload_hashes=tuple(hashes),
            cache_dependents=tuple(
                "%s!%s" % item for item in (
                    () if cache is None else cache.referenced_by)
            ),
            ownership=ownership,
            origin=origin,
        ),
        origin=origin,
        dependents=dependents,
    )


def _already_managed(handle, node):
    public = PivotAdoptionQualification(
        eligible=False,
        strategy=None,
        requires_calculation=False,
        calculation_engine=None,
        operation_constraints=(),
        reasons=(_reason(
            "already-managed",
            part=handle._identity.pivot_part,
        ),),
    )
    return AdoptionAnalysis(
        public=public,
        pivot_part=handle._identity.pivot_part,
        cache_part=None if node is None else node.cache_definition_part,
        records_part=None if node is None else node.cache_records_part,
        payload_hashes=(),
        cache_dependents=(),
        ownership=None,
        origin="paper",
    )


def _strategy(cache):
    if cache is None or not cache.definition_part:
        return None
    if len(cache.referenced_by) > 1:
        return "shared-isolation"
    return "dedicated-replacement"


def _package_format_reasons(workbook):
    reasons = []
    if getattr(workbook, "template", False) \
            or getattr(workbook, "is_template", False):
        reasons.append(_reason("foreign-format-unqualified", format="template"))
    source = getattr(workbook, "_paper_source", None)
    if not source:
        return reasons
    try:
        with zipfile.ZipFile(io.BytesIO(source)) as archive:
            names = set(archive.namelist())
            content = archive.read("[Content_Types].xml")
    except (KeyError, OSError, ValueError, zipfile.BadZipFile):
        return reasons
    if b"macroEnabled" in content or "xl/vbaProject.bin" in names:
        code = "foreign-vba-dependency-unproved" \
            if "xl/vbaProject.bin" in names else "foreign-format-unqualified"
        reasons.append(_reason(code, format="xlsm"))
    if b"strict" in content.lower() or b"purl.oclc.org/ooxml" in content:
        reasons.append(_reason("foreign-format-unqualified", format="strict"))
    return reasons


def _dedupe(reasons):
    seen = set()
    unique = []
    for item in reasons:
        key = (item.capability, item.code, item.context)
        if key in seen:
            continue
        seen.add(key)
        unique.append(item)
    unique.sort(key=lambda item: (
        item.capability or "",
        item.code,
        item.context,
    ))
    return unique


def _reason(code, **context):
    items = tuple(sorted(
        (key, value) for key, value in context.items() if value is not None))
    return QualificationReason(None, code, items)


def _result(handle, reasons, constraints, strategy, requires_calc, engine,
            analysis, origin="foreign", dependents=()):
    reasons = _dedupe(reasons)
    constraints = _dedupe(list(constraints))
    public = PivotAdoptionQualification(
        eligible=not reasons,
        strategy=strategy,
        requires_calculation=bool(requires_calc),
        calculation_engine=engine if requires_calc else None,
        operation_constraints=tuple(constraints),
        reasons=tuple(reasons),
    )
    if analysis is None:
        return AdoptionAnalysis(
            public=public,
            pivot_part=handle._identity.pivot_part,
            cache_part=None,
            records_part=None,
            payload_hashes=(),
            cache_dependents=(),
            ownership=None,
            origin=origin,
        )
    return AdoptionAnalysis(
        public=public,
        pivot_part=analysis.pivot_part,
        cache_part=analysis.cache_part,
        records_part=analysis.records_part,
        payload_hashes=analysis.payload_hashes,
        cache_dependents=analysis.cache_dependents,
        ownership=analysis.ownership,
        origin=origin,
    )
