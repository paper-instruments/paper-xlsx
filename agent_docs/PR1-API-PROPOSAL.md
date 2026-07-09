# PR-1 — API Proposal for Batches 2–7 (PLAN-v0.1)

**Status:** the design contract for the remainder of the v0.1 wave, grounded in the
live code as of Batch 1 (commit 69f7b6ffe). CONVENTIONS remains the law; PR-0 remains
the contract for the v0 surface and carries the running amendment register. Everything
here follows the house pattern proven in v0: additive APIs, typed refusals with
what/why/what-instead, snapshot-vs-snapshot detection, splice-or-refuse (never a second
whole-file writer), pinned JSON schemas, and the three-legal-outcomes doctrine.
Deviations discovered during implementation are amended HERE in the same commit,
never silently. Every new exception/result/return named below joins the
pinned-surface CI check the day it lands.

## 1. Batch 2 — the part-lifecycle engine and its unlocks

### 1.1 The engine (internal primitive, not public API)

`openpyxl/preserve/lifecycle.py` — one primitive both directions:

```python
def add_part(build, name, payload, content_type=None, default_ext=None,
             relate_from=None, rel_type=None, rel_id=None) -> str
    # writes the part; appends the CT override (or Default by extension);
    # appends the relationship on relate_from's rels part (append-only ids,
    # max+1 — the crosspart discipline); returns the allocated rId.
def remove_part(build, name, referencing_rels=()) -> None
    # drops the part from the copy loop, removes its CT override, removes
    # exactly the named rels; refuses (internal invariant) if any OTHER
    # rel still targets the part.
```

Everything in Batches 2–7 that creates or deletes parts routes through these two —
no bespoke cascades. The calcChain drop and added-sheet creation migrate onto the
engine as its first regression-guarded consumers.

### 1.2 Tables

Model mutations of loaded `Table` objects become SUPPORTED at save (lifting the
Batch-1 refusal): the table part re-renders from the fully-modeled `Table` via the
engine, with guards:

- `tbl.ref` resize must keep the header row fixed and cover >= 1 data row; the
  totals row, if present, stays last — else typed refusal naming the constraint.
- column count changes must match `tableColumns` (add/remove columns explicitly).
- `ws.add_table(Table(...))` on loaded and added sheets: part created via the
  engine; `tableParts` spliced (offsets exist via CT_ORDER_INDEX). Battery 18.
- `del ws.tables[name]`: part removed via the engine; `tablePart` element excised.
- autoFilter interplay: a table's own autoFilter lives inside the table part —
  sheet-level autoFilter untouched.

**Row discipline** (battery 10) — one public verb, module
`openpyxl.preserve.tables`:

```python
def append_row(ws, table_name, values: dict|list) -> AddressRemap|None
    # appends below the last data row: shifts the totals row down when present
    # (via the Phase-6b machinery when rows below exist — same guards), writes
    # the values, extends tbl.ref, re-derives calculated-column formulas from
    # the column's shared pattern (Translator), leaves non-calculated cells
    # empty unless given. Refuses: values for calculated columns that disagree
    # with the pattern; heterogeneous/broken calculated columns (pattern not
    # inferable); tables ending at the sheet limit (BoundaryViolationError).
```

`insert_row(ws, table_name, index, values)` and `delete_row(ws, table_name, index)`
ride the same machinery. All three return the shift's `AddressRemap` when rows
actually moved, else `None`.

### 1.3 Comments

Add/edit/remove on sheets whose original has NO comment parts (the 80% case):
comments part + legacy VML part created via the engine; `<legacyDrawing r:id>`
spliced into the sheet. `cell.comment = Comment(...)` just works; battery 19.
Sheets that ALREADY carry comment parts keep refusing in this wave (editing
preserved VML is Batch-4-class work that earns its own probe) — refusal message
updated to say exactly that. Threaded comments: preserved verbatim, authoring
stays deferred (Appendix A footnote).

### 1.4 Engine dividends

- `docProps/custom.xml` creation (first custom prop on a file without the part)
  and deletion (removing the last prop) — both directions via the engine.
- `xl/styles.xml` creation for styled writes into styles-less packages.
- `wb.replace_part(name: str, payload: bytes) -> None` — the raw escape hatch
  `mark_dirty` pointed at but refused: byte-for-byte part replacement for
  media swaps. Guards: the part must exist (else `TargetNotFoundError`); parts
  the model actively manages (sheets, workbook.xml, styles.xml, sharedStrings)
  refuse (`RelationshipPolicyError` — replacing them would desync the model);
  the receipt/confession records every replaced part.

## 2. Batch 3 — region and structural completion

### 2.1 x14 twin-sync (no new public API — refusals lift)

`openpyxl/preserve/x14.py`: scan the sheet `extLst` for `x14:conditionalFormattings`
/ `x14:dataValidations`; parse ONLY sqref/pivot attributes (never the payloads);
on a classic CF/DV edit, rewrite the twin's sqref set in lockstep inside the
extLst bytes (targeted byte patch, chartpatch-style). Rules whose classic side is
deleted delete the twin; twins with no classic counterpart (x14-only DV) keep
refusing with the precise reason. Battery 20.

### 2.2 Sheet lifecycle (upstream signatures, preserve semantics)

- **Rename** (`ws.title = ...`): cascade rewrite — model formulas + defined names
  (Excel-quoting aware, casefold matching), byte rewrite of `<c:f>` texts in
  chart parts and of pivot/table source references (the chartpatch text-walker
  already locates them), workbook.xml sheet entry. Refuses when the old name is
  textually irresolvable (INDIRECT strings referencing the sheet — dependency
  sketch names the cells). Battery 8.
- **Delete** (`wb.remove(ws)`): reference audit first — formulas/names/charts on
  OTHER sheets referencing the victim → typed refusal enumerating them;
  clean removals proceed via the engine (part + rels + CT + `localSheetId`
  remap + calcChain-style cascade for exclusively-referenced children:
  drawings, comments, tables). Returns a `RemovalReport` (pinned shape:
  `{"removed_parts": [...], "remapped_names": int}`).
- **Reorder** (`wb.move_sheet`): `<sheet>` entry splice + `localSheetId` remap +
  `activeTab` fixup.
- **Copy** (`wb.copy_worksheet`): the copy registers as an ADDED sheet (generated
  whole at save — the ledger already models added sheets); charts/images on the
  source do NOT copy in this wave (upstream's copier skips them anyway); battery 11.

### 2.3 Structural widening

- Multiple shifts per session: ledger snapshots/spans rebase after each shift
  (retiring one-shift-per-session); the byte renumber composes transforms in
  recorded order. Each shift still returns its own `AddressRemap`; remaps
  compose via `remap_b.map(remap_a.map(addr))`.
- Spanning merges: insert/delete inside a merge EXPANDS/SHRINKS it (Excel
  semantics) instead of refusing.
- `move_range`: tracked delete+insert with the 6b rewrite machinery + blocker
  analysis; refusals keep the victim list.
- Per-blocker unlocks: rowBreaks/colBreaks shift with the rows; sheets whose
  extLst carries ONLY known-shiftable families (sparkline sqrefs, x14 CF/DV
  sqrefs via 2.1's parser) get their sqrefs rewritten instead of refusing.

### 2.4 Dynamic arrays

cm/vm metadata bookkeeping (`xl/metadata.xml`): a plain value overwrite of a
rich-value/spill cell drops the cell's `cm`/`vm` attributes AND the
corresponding metadata entries (semantically correct: the cell no longer spills)
— lifting the Batch-0 refusal for the common case (battery 21). Writing INTO a
spill range (non-anchor member) keeps refusing, now with `in_spill` context
naming the anchor (battery 7). Structured references (`Table1[@Col]`) in every
shifting path: rewritten via table-extent awareness or refused by name — never
mis-shifted.

## 3. Batch 4 — charts and images under preserve

- **Phase one:** `ws.add_chart`/`ws.add_image` on ADDED sheets — stock drawing
  writer output routed through the engine (parts + rels + CT). Zero splice risk.
- **Phase two:** on LOADED sheets — new chart/drawing parts via the engine, one
  `<drawing r:id>` element spliced into the original sheet bytes (CT-ordered
  insertion; sheets that already have a drawing get the chart appended to the
  EXISTING drawing part only if that drawing is anchor-only, else refusal).
  Battery 22 complete.
- **Chart editing:** the Batch-1 mutation refusal lifts per-property via
  chartpatch: title text, series `<c:f>` ranges. Everything chartpatch cannot
  express keeps refusing with the property named. Convenience verb:
  `chart.repoint(series_index, new_range)` — validates the range, patches bytes.
- Every 4.x output joins the FIXTURE-REQUESTS real-Excel open-check queue (1.8).

## 4. Batch 5 — computation layer

### 4.1 Scenario runner

```python
Workbook.evaluate(set: dict[str, Any], read: list[str], *,
                  timeout=120.0) -> Evaluation
# temp copy -> spine-applied inputs -> oracle recalc -> harvested outputs.
# The ORIGINAL file and the live workbook are untouched (asserted in tests).
# Evaluation (pinned): .inputs, .outputs {addr: value}, .status
# ("ok"|"errors"), .error_cells, .certification (the CertificationResult of
# the recalced copy), .to_dict() schema "evaluation" v1.
```

Addresses are sheet-qualified A1 or defined names (resolved via the model map
when Batch 6 lands; until then names + A1). Batch mode:

```python
oracle.evaluate_many(source, cases: list[dict], read: list[str], *,
                     pool_size=2, timeout=120.0) -> list[Evaluation]
```

**Pool lifecycle (delegated decision, resolved):** the pool is an
implementation detail of `evaluate_many`, not a public object — `pool_size`
warm profiles created lazily, each per-profile-isolated (the v0 oracle
discipline), crash-replaced once, destroyed before return; never reused across
calls. A public persistent pool earns its API only with a measured latency
case (Appendix-A-5 spirit).

### 4.2 Formula pre-flight linter

```python
openpyxl.formula.lint.lint_formula(text, *, workbook=None, sheet=None)
    -> list[LintFinding]   # {"code", "message", "anchor"}
```

Checks (all tokenizer-based, no evaluation): unknown function names against the
pinned catalog (Excel function list + `_xlfn` handling); unbalanced parens;
unresolvable sheet/range/name references (when workbook given); structured-ref
validity against table columns; `;` used as an argument separator (the
locale-canonical trap — storage is ALWAYS comma-canonical). Wired into the
value-bind chokepoint under preserve via `wb.formula_lint = "off"|"warn"|"refuse"`
(default "warn"; refusal raises before the bind, like strict protection).
`LintWarning` joins the pinned surface.

### 4.3 Oracle v2 — certification-gated write-back

```python
oracle.write_back(source, *, timeout=120.0, allow_uncertified=False)
    -> WriteBackResult
```

Recalcs a temp copy, then splices the computed `<v>` values into the ORIGINAL
package via the spine (values-only cell edits through the standard splice; the
ledger records them; macro-safe by construction — LO bytes never enter the
output). **Hard conditions (per the plan):** only values whose cells the
certification pass verified (CERTIFIED) or that were previously cache-less are
written; on DIVERGED or BASELINE_UNVERIFIABLE the call refuses unless
`allow_uncertified=True`, and then the result and receipt carry a loud
`uncertified=True` stamp. `fullCalcOnLoad` is CLEARED only when write-back
covered every formula cell (excluded classes leave it set). Every write-back
emits a package-diff confession in the result. Battery 24.

### 4.4 Fresh-generation recalc flag

`Workbook.save` on NEVER-loaded workbooks containing formulas sets
`calcPr fullCalcOnLoad` (the model object, stock writer path — no splice
involved) unless the caller set `wb.calculation` explicitly. Stock-visible
change, recorded in PAPER.md; it removes the one stale-cache gap outside
preserve custody.

## 5. Batch 6 — perception and the agent experience

```python
Worksheet.locate(label: str, *, prefer="right") -> Cell
# exact-then-normalized text match over the sheet; the value cell is the
# nearest non-label neighbour (prefer= "right" or "below"). ZERO matches ->
# TargetNotFoundError; >1 candidate labels or ambiguous value cell ->
# AmbiguousTargetError listing every candidate address (the pinned class
# finally earns its keep). Battery 23.
Workbook.search(text_or_regex, *, regex=False, values=True, formulas=True)
    -> list[dict]      # {"address", "match", "kind"}
openpyxl.preserve.scan_errors(wb) -> list[dict]
# LibreOffice-free: cached error tokens (both load views) + broken refs
# (#REF! in formulas) — the cheap first-pass check for blind environments.
Worksheet.allowed_values(cell) -> list|None   # DV vocabulary (list-type DVs)
Workbook.validate() -> None                   # the saver's validation pass
# without writing: every refusal a save WOULD raise, raised now.
```

- **Model map (6.2):** `wb.model_map()` — classification per formula-bearing
  sheet: inputs (no formula, referenced), calculations, outputs (formula,
  unreferenced), constants; corroborated by fill-color convention when present.
  Pinned schema "model_map" v1. `set_input()` (Batch 7) consumes it.
- **Manifest enrichment (6.5):** formula addresses per sheet, inputs/outputs
  counts from the sketch, `certifiable` flag (cached values present),
  per-sheet part names, the 1.6 protection summary.
- **Edit receipt (6.6):**
  `openpyxl.preserve.receipt(before, after, *, recalc=None) -> EditReceipt` —
  composes cells-diff + package-diff + confession + recalc/certification
  status into one pinned artifact (schema "edit_receipt" v1); `wb.save(path,
  receipt=True)` returns one for the save it just performed.
- **Structured refusals (6.7):** `PaperRefusal` gains `.kind` (stable string),
  `.anchor` (sheet-qualified address or part name), `.options` (list of
  suggested remedies) — populated progressively; message text unchanged.
- **Hygiene findings (6.8):** `openpyxl.preserve.findings(wb) -> list[Finding]`
  (pinned taxonomy, resolved here): `hardcode-in-formula`,
  `inconsistent-row-formula`, `error-cell`, `orphaned-name`, `external-link`,
  `hidden-sheet`, `hidden-rows`, `merged-hazard`, `volatile`,
  `magnitude-outlier` (ADVISORY lint only — flags, never decides; the fences
  stand). Measurements, never judgments: every finding carries evidence
  addresses.
- **Workbook diff report (6.9):** `diff_workbooks(a, b, remaps=()) -> Report`
  classifying content-changed vs shifted-by-structural-edit (consumes
  `AddressRemap` chains).

## 6. Batch 7 — delivery, hardening, adoption

- **Style verbs:** `copy_format(ws, src_cell, dst_range)` (style-array reuse via
  the D2 translator); `apply_profile(ws, profile)` with profiles as DATA
  (`{"inputs": {...}, "formulas": {...}, "links": {...}}`, number-format library
  included, per-customer overrideable); both preserve-safe (styles append-only).
- **Ergonomics:** `wb.set_input(name_or_label, value)` (defined names → model
  map → locate); `wb.protect_for_delivery(password=None)` (lock all but
  classified inputs; protection REPORTED in the receipt); `wb.scrub(remove=
  ("comments", "metadata", "personal", "hidden-sheets"))` returning a scrub
  report (hidden sheets: report-or-remove, never silent); pivot
  `wb.set_pivot_refresh_on_load()` — byte patch of `refreshOnLoad` on every
  pivotCacheDefinition via the engine.
- **Hardening:** fsync before `os.replace` + directory fsync (durability);
  spool-to-disk archive builds for path targets (~1× file-size peak memory);
  decompression caps (refuse parts inflating past a pinned ratio+size);
  zip-confusion checks (central-vs-local header agreement) in the raw-copy
  path; `mark_dirty` range clamps to the populated extent.
- **Adoption:** README "the paper API in 90 seconds" + `doc/paper.rst`
  documenting preserve, manifest, refusal taxonomy, diffs, oracle, receipts,
  and the five pinned JSON schemas; the public-default release gate documented
  with its Appendix-A conditions and the internal flip noted.

## 7. Refusal-condition summary (new surface)

Every verb above refuses (typed, atomic, what/why/what-instead) rather than
guessing: table verbs on patterns they cannot infer; rename on textually
irresolvable references; delete on live references; drawing append into
non-anchor-only drawings; write-back on failed certification; locate on
ambiguity; scrub never removes silently. No new verb ever adds a fourth
outcome.

## 8. Delegated decisions — resolved in this document

1. Un-share default: settled by the Batch-0 probe (PR-0 amendment 6).
2. `evaluate()` interface + pool lifecycle: §4.1 (pool is internal, bounded,
   per-call).
3. Findings taxonomy: §5 (6.8) — ten kinds, pinned.
4. Receipt schema: "edit_receipt" v1 = `{schema, version, cells_diff,
   package_diff, confession, recalc, certification, replaced_parts,
   uncertified}` — golden-tested like every other schema.
