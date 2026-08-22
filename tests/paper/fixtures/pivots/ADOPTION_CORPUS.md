# Foreign adoption fixture corpus

This directory holds hash-pinned Excel-authored workbooks used by foreign
adoption qualification. Synthetic ZIP surgery must never be labeled
Excel-authored.

## Honesty rule

A fixture is Excel-authored only when desktop Excel wrote the bytes. Until
those binaries exist, PR 9 exercises synthetic packages from
`tests/paper/test_pivot_graph.py` and Paper-created-then-stripped packages.
Those are not release evidence.

`eligible=True` is forbidden for every foreign fixture until
`RELEASE_MATRIX.json` records pinned Excel builds and marker survival. The
qualifier implements that gate as `foreign-managed-equivalence-unproved`.

## Required Excel-authored adoption files

| Fixture | Purpose | Status |
|---|---|---|
| `excel_adopt_dedicated_table.xlsx` | Basic dedicated-cache table source | pending human/Excel authoring |
| `excel_adopt_table_totals_on.xlsx` | Table with `totalsRowShown=true` | pending |
| `excel_adopt_table_totals_off.xlsx` | Table with `totalsRowShown=false` | pending |
| `excel_adopt_range_nested.xlsx` | Range source, nested rows, filters, two measures | pending |
| `excel_adopt_stale_source.xlsx` | Visible output matches persisted cache after source edit | pending |
| `excel_adopt_shared_cache.xlsx` | Two ordinary pivots sharing one cache | pending |
| `excel_adopt_shared_unadoptable.xlsx` | Three siblings; only one semantically adoptable | pending |
| `excel_adopt_formula_source.xlsx` | Formula-backed source with current Excel caches | pending |
| `excel_adopt_custom_ids.xlsx` | Custom part names and relationship IDs | pending |
| `excel_adopt_missing_records.xlsx` | Pivot without saved cache records | pending |
| `excel_adopt_style_only_cells.xlsx` | Blank and style-only cells in the footprint | pending |
| `excel_adopt_manual_format.xlsx` | Manual/custom output formatting | pending |
| `excel_adopt_nondefault_core.xlsx` | Nondefault core-schema properties | pending |
| `excel_adopt_consumers.xlsx` | GETPIVOTDATA, formula, name, and chart consumers | pending |
| `excel_adopt_grouping.xlsx` | Grouped fields | pending |
| `excel_adopt_calculated.xlsx` | Calculated fields/items | pending |
| `excel_adopt_pivotchart.xlsx` | PivotChart dependency | pending |
| `excel_adopt_slicer.xlsx` | Slicer/timeline dependency | pending |
| `excel_adopt_data_model.xlsx` | Data Model or OLAP pivot | pending |
| `excel_adopt_extension.xlsx` | Known Excel compatibility extension payload | pending |
| `excel_macro_pivot.xlsm` | VBA-bearing `.xlsm` refusal | pending |

Sidecars must record expected graph, projection, old output role map,
`qualify_adoption()` result, cache strategy, reasons, and dependency
identities. Schema name is `pivot_adoption_qualification` version 1.
