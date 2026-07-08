# FIXTURE-REQUESTS — real-Excel fixtures a human must produce

**Why this file exists.** This environment cannot run desktop Excel, and Phase-0 evidence shows
several load-bearing claims are only testable against genuinely Excel-authored files (stock
openpyxl 3.1.5 round-trips its OWN charts/images byte-identically, so synthetic fixtures
understate the carnage — see `agent_docs/OPEN-QUESTIONS.md` Q11). Every fixture below should be
authored in real desktop Excel (note the exact version), saved normally, and delivered with a
filled sidecar per CONVENTIONS §4 (`verified_by` = the human who authored it). Never label a
fixture with provenance it lacks.

Priority order: items 1–4 block battery baselines and PR-0 evidence; the rest gate specific
tests scheduled for Phases 2–6.

1. **Real client-model-shaped workbook (the gauntlet bucket).** A realistic financial model:
   multiple sheets, cross-sheet formulas, defined names, charts, at least one pivot table with
   cache, conditional formatting (including a color scale or data bar → produces x14 twins),
   data validation dropdowns, merged headers, freeze panes, hidden rows/sheets, years of style
   accretion. This is the load-bearing bucket for the whole fork. Sanitized/synthetic data is
   fine; the STRUCTURE must be Excel-authored.
2. **Feature-isolated: chart + shapes.** One sheet with an Excel-authored chart AND a textbox or
   shape on the same drawing. Purpose: demonstrate stock chart/drawing loss (battery jobs 1–2
   baseline — Excel charts carry styling/extLst content openpyxl's model drops; shapes kill the
   whole drawing on the TypeError path, `reader/drawings.py:30-34`).
3. **Feature-isolated: shared formulas.** A column of ≥20 filled formulas (Excel writes
   `<f t="shared" si ref>` host+followers) plus one array formula (`Ctrl+Shift+Enter` or a
   dynamic-array spill like `=SORT(...)` producing `cm`/`vm` metadata). Purpose: gaps item G1 —
   the splice must detect and handle/refuse shared-group intersections; we need the real byte
   shapes. Please do NOT edit the file with anything but Excel after authoring.
4. **Feature-isolated: real .xlsm with macros.** A workbook with a trivial VBA module (e.g. one
   MsgBox sub) and, ideally, a form control button wired to it. Purpose: battery job 4 with real
   vbaProject.bin + ctrlProps/activeX satellites.
5. **Feature-isolated: pivot table + cache** (separate from item 1's). Purpose: pivot-preservation
   and staleness-confession tests (Phase 4).
6. **Feature-isolated: sparklines + x14 conditional formatting** authored by Excel (our injected
   synthetic versions load, but Excel's real emission includes xr:uid attrs, dyDescent, etc.).
   Purpose: splice-completeness trap test with authentic producer bytes.
7. **Excel-read-back checks (a human opens files we generate and reports):**
   - the four Q7 string variants (`scratch/results/q7_strings/variant_*.xlsx`): do they open
     without the repair dialog? Do values read correctly? What does Excel's next manual save do
     to `t="inlineStr"` cells and to stale sst `count`/`uniqueCount` attributes?
   - a spliced-output file from the Phase-2 prototype (when it exists): repair-dialog check.
   - a cell with leading/trailing spaces in `<t>` WITHOUT `xml:space="preserve"`: does Excel
     strip the padding? (Q4 verifier follow-up.)
8. **Producer-variance probes:** does current Excel ever emit r-less `<row>`/`<c>` elements
   (ECMA-376 allows them; openpyxl/LO never emit them)? Does Excel tolerate/repair stale
   `<dimension>` and row `spans` after external edits? One small Excel-authored file plus one
   observation each is enough.
9. **Feature-isolated: 1904 date system** (`Workbook Options → Use 1900/1904 date system` on
   Mac Excel, or via workbookPr). Purpose: gaps item G11 — splice epoch inheritance + certify()
   comparison. Include a few date cells with known values in the sidecar.
10. **External links + legacy formats:** a workbook referencing another workbook (external link
    chain), one `.xlsb`, and one legacy `.xls` (the latter two for typed-refusal tests only).
11. **Non-canonical package paths (if encountered in the wild):** any real file from a
    third-party producer whose parts don't live at `xl/worksheets/sheetN.xml`-style canonical
    paths, or that uses absolute `/xl/...` rel targets. Purpose: gaps item G10 (rels-driven part
    resolution). Opportunistic — grab one if a real integration surfaces it.

**Sidecar template** (CONVENTIONS §4, copy per fixture):

```json
{
  "fixture": "<name>.xlsx",
  "provenance": {"app": "Excel <version>", "version": "<build>", "notes": "<how authored>"},
  "features": ["..."],
  "ground_truth": {
    "cached_values": {"Sheet!Cell": 0.0},
    "chart_count": 0, "pivot_count": 0, "vba_present": false,
    "formula_count": 0
  },
  "verified_by": "<human name>", "date": "YYYY-MM-DD"
}
```
