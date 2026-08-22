# Pivot fixture corpus

This directory is the Paper PivotTable evidence corpus. Binary files are
hash-pinned in `tests/paper/fixtures/MANIFEST.sha256` and each binary has a
JSON sidecar.

## Honesty rule

This machine can run desktop Excel, but a fixture is labeled Excel-authored
only when desktop Excel actually wrote the bytes. Synthetic ZIP surgery and
LibreOffice exports are labeled as such. Nothing in this directory may be
relabeled as Excel-authored.

## Required Excel-authored files

These names are the PR 1 acceptance corpus. A missing binary is an explicit
external-corpus gap, not a license to fabricate Excel provenance.

| Fixture | Purpose | Status |
|---|---|---|
| `excel_basic_table.xlsx` | One table source, one row field, one sum measure, dedicated cache | pending human/Excel authoring |
| `excel_cross_tab.xlsx` | Two row fields, one column field, two measures, totals | pending human/Excel authoring |
| `excel_filtered.xlsx` | Report filter with multiple selected items | pending human/Excel authoring |
| `excel_shared_cache.xlsx` | Two pivots sharing one cache | pending human/Excel authoring |
| `excel_formula_source.xlsx` | Source rows containing formulas with current cached results | pending human/Excel authoring |
| `excel_semantic_edges.xlsx` | Blank rows, formula empty strings, typed item identity, counts, ordering, caption conflicts, and floating-point cases | pending human/Excel authoring |
| `excel_extension_pivot.xlsx` | Ordinary Excel extension payloads on a non-Data-Model pivot | pending human/Excel authoring |
| `excel_unsupported_grouping.xlsx` | Native grouped field whose mutation capabilities must be disabled | pending human/Excel authoring |
| `excel_data_model.xlsx` | Data Model or OLAP-backed pivot | metadata-only until a licensed Excel Data Model workbook is supplied |
| `excel_macro_pivot.xlsm` | Transitional macro-enabled compatibility and unrelated VBA preservation | pending human/Excel authoring |
| `excel_template_pivot.xltx` | Template targeted-pivot API refusal | pending human/Excel authoring |
| `excel_strict_basic.xlsx` | OOXML Strict targeted-pivot API refusal | pending human/Excel authoring |

When a binary is added, commit its sidecar and `MANIFEST.sha256` in the same
change. Sidecars must include producer, exact version, authoring steps,
`expected_pivot_graph`, `expected_pivot_to_dict`,
`expected_pivot_qualification`, expected visible values, and independent
verification status. Qualification sidecars must record origin, validity, and
every operation capability. Foreign Excel/LibreOffice pivots may have at most
`can_refresh_on_open=true` in v1.

## Required LibreOffice-authored file

| Fixture | Purpose | Status |
|---|---|---|
| `libreoffice_basic_pivot.xlsx` | Producer-variance coverage for a simple DataPilot exported as XLSX | authored in this PR when soffice is available |

## Pending Excel semantic transcripts

PR 1 cannot freeze these rules until an Excel-authored sidecar records them:

- fully blank source rows versus formula cells returning `""`
- typed item identity for numeric `1`, text `"1"`, Boolean `True`, dates, case variants, and blank
- `count` and `count_numbers` coercion
- default versus manual item ordering
- floating-point aggregation versus Excel
- duplicate and conflicting captions
- display captions and locale for blank, Values, subtotals, and grand totals
- empty filtered groups

## Graph-invalid cases

Dangling relationships, duplicate cache IDs, custom part names, and record
count mismatches are constructed in `tests/paper/test_pivot_graph.py`. Those
packages are test-local and are not labeled Excel-authored.
