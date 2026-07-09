# FIXTURE-REQUESTS — real-Excel open-check queue

Outputs that MUST be opened in desktop Excel (the oracle can only prove
LibreOffice acceptance) before the preserve-by-default release gate flips
for public/pandas surfaces. Each entry names the producing test — run it,
take the artifact from tmp, open in Excel, confirm no repair dialog and
the described content renders.

Queued by Batch 4 (PLAN-v0.1: "every 4.x output joins the real-Excel
open-check queue — this is the producer-sensitive surface"):

1. **Chart + image on an ADDED sheet** —
   `test_drawings.py::TestAddedSheetDrawings::test_chart_and_image_on_added_sheet`.
   Expect: bar chart and 1px image render on sheet "Report"; no repair.
2. **Fresh drawing spliced into a LOADED sheet** —
   `test_drawings.py::TestLoadedSheetFreshDrawing::test_chart_on_machinery_free_sheet`.
   Expect: chart renders on Sheet1 alongside untouched original content.
3. **Anchor append into an EXCEL-AUTHORED drawing** —
   `test_drawings.py::TestExistingDrawingAppend::test_append_chart_preserves_existing_anchors`.
   NOTE: chart_image.xlsx is openpyxl-authored (default-ns drawing). A
   REAL Excel-authored fixture (xdr:-prefixed drawing) is wanted here —
   the append path injects xmlns on the appended anchors for that case
   and only a real-Excel open proves it.
4. **Repointed series** —
   `test_drawings.py::TestChartPropertyEdits::test_repoint_patches_series_bytes`.
   Expect: chart follows the new range; stale cached values refresh on
   open.
5. **Retitled chart** —
   `test_drawings.py::TestChartPropertyEdits::test_title_edit_lands_and_reloads`.
   Expect: title shows "New & Improved <Title>" with literal angle
   brackets.

Standing wants (from earlier batches, unblocked whenever a Windows/Mac
Excel session is available): x14 twin-sync outputs (gauntlet CF edits),
table append_row with totals row, comment creation on machinery-free
sheets.
