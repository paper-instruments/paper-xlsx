# paper-xlsx: formula-backed pivot source calculation

"""Harvest current formula results for a pivot source without publishing
LibreOffice's rewritten workbook.

Literal sources never invoke the oracle. Formula-backed sources render a
disposable preserve-mode candidate of ordinary workbook edits, then reuse
``openpyxl.oracle`` at its existing seam.
"""

from __future__ import annotations

import hashlib
import io
from dataclasses import dataclass

from openpyxl.errors import (
    OracleTimeoutError,
    OracleUnavailableError,
    UnsupportedStructureError,
)
from openpyxl.pivot.source import snapshot_from_workbook, typed_value


@dataclass(frozen=True)
class PivotCalculationArtifact:
    candidate_sha256: str
    engine: str
    engine_version: str | None
    source: object
    source_identity: str
    values_by_coordinate: object
    excluded_coordinates: object
    errors: tuple

    def to_dict(self):
        return {
            "candidate_sha256": self.candidate_sha256,
            "engine": self.engine,
            "engine_version": self.engine_version,
            "source_identity": self.source_identity,
            "value_count": len(self.values_by_coordinate),
            "excluded": list(self.excluded_coordinates),
            "errors": list(self.errors),
        }


def snapshot_for_pivot(workbook, source):
    """Snapshot a source, calculating formula cells when required."""
    snapshot = snapshot_from_workbook(workbook, source)
    if not snapshot.formula_coordinates:
        return snapshot
    artifact = calculate_pivot_source(workbook, snapshot)
    return apply_calculated_values(snapshot, artifact)


def calculate_pivot_source(workbook, snapshot):
    """Return calculated source values bound to one candidate identity."""
    from openpyxl.oracle import _recalculate_bytes, _recalc_scan, find_soffice

    if find_soffice() is None:
        raise UnsupportedStructureError(
            "formula-backed pivot sources require stock LibreOffice; "
            "soffice was not found. Nothing was changed.",
            kind="unsupported-pivot-source",
            options=list(snapshot.formula_coordinates),
        )
    candidate = render_ordinary_candidate(workbook)
    candidate_sha = hashlib.sha256(candidate).hexdigest()
    try:
        recalculated = _recalculate_bytes(candidate, timeout=120.0)
    except OracleUnavailableError as exc:
        raise UnsupportedStructureError(
            "formula-backed pivot sources require stock LibreOffice. "
            "Nothing was changed.",
            kind="unsupported-pivot-source",
            options=list(snapshot.formula_coordinates),
        ) from exc
    except OracleTimeoutError as exc:
        raise UnsupportedStructureError(
            "LibreOffice timed out while calculating the pivot source. "
            "Nothing was changed.",
            kind="unsupported-pivot-source",
            options=list(snapshot.formula_coordinates),
        ) from exc
    _wb_formulas, wb_values, _scanned, _formulas, errors = _recalc_scan(
        recalculated)
    if errors:
        raise UnsupportedStructureError(
            "pivot source formulas evaluated to Excel errors. "
            "Nothing was changed.",
            kind="unsupported-pivot-source",
            options=list(errors)[:8],
        )
    values = {}
    excluded = {}
    for address in snapshot.formula_coordinates:
        sheet, _sep, coord = address.partition("!")
        if sheet not in wb_values:
            excluded[address] = "missing-sheet"
            continue
        cell = wb_values[sheet][coord]
        if cell.value is None:
            excluded[address] = "no-computed-value"
            continue
        if getattr(cell, "data_type", None) == "e":
            excluded[address] = "formula-error"
            continue
        values[address] = cell.value
    if excluded:
        raise UnsupportedStructureError(
            "pivot source formulas could not be calculated for every "
            "required cell. Nothing was changed.",
            kind="unsupported-pivot-source",
            options=sorted(excluded),
        )
    return PivotCalculationArtifact(
        candidate_sha256=candidate_sha,
        engine="libreoffice",
        engine_version=None,
        source=snapshot.source,
        source_identity=snapshot.identity,
        values_by_coordinate=values,
        excluded_coordinates=excluded,
        errors=tuple(errors),
    )


def apply_calculated_values(snapshot, artifact):
    if artifact.source_identity != snapshot.identity:
        raise UnsupportedStructureError(
            "pivot calculation artifact does not match the current source.",
            kind="unsupported-pivot-source",
        )
    records = []
    for record in snapshot.records:
        values = []
        for index, value in enumerate(record.values):
            address = _record_address(snapshot, record.source_row, index)
            if address in artifact.values_by_coordinate:
                values.append(typed_value(artifact.values_by_coordinate[address]))
            else:
                values.append(value)
        from openpyxl.pivot.source import TypedRecord
        records.append(TypedRecord(tuple(values), record.source_row))
    from openpyxl.pivot.source import SourceSnapshot, _source_identity
    catalogs = [[] for _ in snapshot.fields]
    seen = [set() for _ in snapshot.fields]
    for record in records:
        for index, item in enumerate(record.values):
            if item not in seen[index]:
                seen[index].add(item)
                catalogs[index].append(item)
    shared = tuple(tuple(items) for items in catalogs)
    identity = _source_identity(
        snapshot.source, snapshot.fields, records,
        snapshot.formula_coordinates, generation=None)
    return SourceSnapshot(
        source=snapshot.source,
        fields=snapshot.fields,
        records=tuple(records),
        shared_items=shared,
        formula_coordinates=snapshot.formula_coordinates,
        identity=identity,
        bounds=snapshot.bounds,
        warnings=snapshot.warnings,
    )


def render_ordinary_candidate(workbook):
    """Preserve-save ordinary edits to bytes without staged pivot graph ops."""
    from openpyxl.preserve.saver import save_preserved

    ledger = workbook._paper_ledger
    ops = dict(getattr(ledger, "pivot_operations", {}) or {})
    ledger.pivot_operations = {}
    buf = io.BytesIO()
    try:
        save_preserved(workbook, buf)
        return buf.getvalue()
    finally:
        ledger.pivot_operations = ops


def assert_candidate_unused_as_publication(candidate, published):
    if candidate == published:
        raise UnsupportedStructureError(
            "LibreOffice-produced workbook bytes cannot be the published "
            "artifact.",
            kind="unsupported-pivot-operation",
        )


def _record_address(snapshot, source_row, field_index):
    sheet = snapshot.bounds[0]
    min_col = snapshot.bounds[1]
    min_row = snapshot.bounds[2]
    if min_col is None:
        min_col = 1
        min_row = 1
    from openpyxl.utils import get_column_letter
    row = min_row + source_row
    column = min_col + field_index
    return "%s!%s%s" % (sheet, get_column_letter(column), row)
