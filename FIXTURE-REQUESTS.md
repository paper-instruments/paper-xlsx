# Fixture Requests

The frozen corpus still needs packages authored by applications that the test
environment cannot reproduce. Do not synthesize or relabel these files.

## Requested Packages

- Excel-authored workbook with styled charts and chart caches
- Excel-authored pivot table with pivot cache and source data
- Excel-authored macro workbook containing a real VBA project
- Excel-authored workbook using the 1904 date system
- Excel-authored workbook with external links
- Real binary `.xlsb` workbook for refusal coverage
- Google Sheets export containing formulas, formatting, charts, and defined names

## Acceptance

Each contribution must include a JSON sidecar matching the corpus convention:

- exact authoring application and version
- features intentionally present
- known cached values or other ground truth
- human verifier and verification date
- any sanitization performed before contribution

After review, update `tests/paper/fixtures/MANIFEST.sha256` in the same commit.
Fixture bytes must never be generated or rewritten by code under test.
