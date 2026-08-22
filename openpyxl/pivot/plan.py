# paper-xlsx: immutable pivot build plan

"""Compose a qualified PivotSpec and source snapshot into one build plan.

The planner is a pure function. It does not touch a workbook, ledger, ZIP,
filesystem, or LibreOffice process. Layout coordinates and blank-item captions
remain provisional until Excel transcripts approve them.
"""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass

from openpyxl.errors import BoundaryViolationError
from openpyxl.pivot.aggregate import aggregate_snapshot
from openpyxl.pivot.layout import layout_result
from openpyxl.pivot.source import DEFAULT_LIMITS, SourceSnapshot


@dataclass(frozen=True)
class PivotBuildPlan:
    spec: object
    source_identity: str
    fields: tuple
    shared_items: tuple
    records: tuple
    field_indexes: object
    aggregate: object
    output: object
    cache_estimate_bytes: int
    output_cell_count: int
    warnings: tuple
    provenance: object
    digest: str

    def to_dict(self):
        return {
            "source_identity": self.source_identity,
            "fields": list(self.fields),
            "output_range": self.output.ref,
            "destination": self.output.destination,
            "first_header_row": self.output.first_header_row,
            "first_data_row": self.output.first_data_row,
            "first_data_col": self.output.first_data_col,
            "row_count": self.output.row_count,
            "column_count": self.output.column_count,
            "cache_estimate_bytes": self.cache_estimate_bytes,
            "output_cell_count": self.output_cell_count,
            "included_row_count": self.aggregate.included_row_count,
            "digest": self.digest,
            "warnings": list(self.warnings),
        }


def plan_pivot(spec, snapshot, limits=DEFAULT_LIMITS):
    """Return an immutable plan for ``spec`` over ``snapshot``."""
    if not isinstance(snapshot, SourceSnapshot):
        raise TypeError("plan_pivot requires a SourceSnapshot")
    _estimate_or_refuse(spec, snapshot, limits)
    result = aggregate_snapshot(snapshot, spec, limits=limits)
    if result.included_row_count == 0:
        raise BoundaryViolationError(
            "filters excluded every source row",
            kind="invalid-pivot-source",
        )
    output = layout_result(
        spec, result, spec.destination, limits=limits)
    field_indexes = {
        name: index for index, name in enumerate(snapshot.fields)
    }
    cache_estimate = _estimate_cache_bytes(snapshot)
    if cache_estimate > limits.cache_xml_bytes:
        raise BoundaryViolationError(
            "projected cache XML is %s bytes; the limit is %s"
            % (cache_estimate, limits.cache_xml_bytes),
            kind="pivot-cache-too-large",
            options=[str(cache_estimate), str(limits.cache_xml_bytes)],
        )
    warnings = tuple(result.warnings) + (
        "layout-coordinates-provisional",
        "blank-caption-provisional",
    )
    provenance = {
        "source_row_count": result.source_row_count,
        "included_row_count": result.included_row_count,
        "measure_input_counts": list(result.measure_input_counts),
        "formula_coordinates": list(snapshot.formula_coordinates),
        "calculation": "literal-source",
    }
    digest = _plan_digest(spec, snapshot, output)
    return PivotBuildPlan(
        spec=spec,
        source_identity=snapshot.identity,
        fields=snapshot.fields,
        shared_items=snapshot.shared_items,
        records=snapshot.records,
        field_indexes=field_indexes,
        aggregate=result,
        output=output,
        cache_estimate_bytes=cache_estimate,
        output_cell_count=output.row_count * output.column_count,
        warnings=warnings,
        provenance=provenance,
        digest=digest,
    )


def _estimate_or_refuse(spec, snapshot, limits):
    from openpyxl.pivot.aggregate import _validate_spec_fields

    _validate_spec_fields(spec, snapshot.field_index)
    row_card = 1
    for item in spec.rows:
        row_card *= max(1, len(snapshot.shared_items[
            snapshot.field_index[item.field]]))
    col_card = 1
    for item in spec.columns:
        col_card *= max(1, len(snapshot.shared_items[
            snapshot.field_index[item.field]]))
    states = row_card * max(1, col_card) * len(spec.values)
    if states > limits.aggregate_states:
        raise BoundaryViolationError(
            "projected %s aggregate states exceed the limit of %s"
            % (states, limits.aggregate_states),
            kind="pivot-cardinality-too-large",
            options=[str(states), str(limits.aggregate_states)],
        )
    header_rows = 2
    label_cols = 1 if spec.layout == "compact" or not spec.rows \
        else max(1, len(spec.rows))
    projected_cells = (header_rows + row_card + 1) * (
        label_cols + max(1, col_card) * len(spec.values))
    if projected_cells > limits.output_cells:
        raise BoundaryViolationError(
            "projected %s output cells exceed the limit of %s"
            % (projected_cells, limits.output_cells),
            kind="pivot-output-too-large",
            options=[str(projected_cells), str(limits.output_cells)],
        )


def _estimate_cache_bytes(snapshot):
    # Conservative uncompressed estimate: tags + typed values, not a serialize.
    return 80 * max(1, len(snapshot.records)) * max(1, len(snapshot.fields))


def _plan_digest(spec, snapshot, output):
    payload = {
        "spec": spec.to_dict(),
        "source": snapshot.identity,
        "cells": [
            {
                "r": cell.row,
                "c": cell.column,
                "v": _json_value(cell.value),
                "role": cell.role,
            }
            for cell in output.cells
        ],
    }
    encoded = json.dumps(payload, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def _json_value(value):
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return value
