# paper-xlsx Fork Ledger

Based on upstream tag `3.1.5`, forked 2026-07-07, marker tag `paper-base`.

Upstream source is the official Mercurial repository at
`https://foss.heptapod.net/openpyxl/openpyxl`. The GitHub repository for this
fork was bootstrapped by cloning that Mercurial repository and converting it to
Git with `hg-fast-export`, then checking out release tag `3.1.5` as `main`.
The upstream tag check showed newer branch commits after `3.1.5`, but no newer
stable release tag.

## Baseline Test Results

- Python 3.9.6 with upstream `requirements.txt` (`lxml==5.0.1`):
  `2592 passed, 6 skipped, 7 xfailed in 17.18s`.
- Python 3.13.3 with CI dependency constraint `lxml<6` (resolved to
  `lxml==5.4.0`): `2592 passed, 6 skipped, 7 xfailed in 3.94s`.
- Environment note: Python 3.13.3 with latest unconstrained `lxml==6.1.1`
  produced four pre-existing upstream failures in
  `openpyxl/xml/tests/test_functions.py::test_iterparse`; lxml now raises
  `TypeError` for the `BytesIO` input where the test expects `ValueError`.
  CI intentionally uses `lxml<6` until upstream handles that dependency change.

## Packaging Smoke Results

- Built with `python -m build`: `paper_xlsx-0.1.0.tar.gz` and
  `paper_xlsx-0.1.0-py2.py3-none-any.whl`.
- Wheel listing starts with `openpyxl/__init__.py` and
  `openpyxl/_constants.py`, confirming the import package was not renamed.
- Wheel smoke and sdist smoke both printed `0.1.0` from
  `openpyxl.__paper_version__`.

## Sanctioned Deviations From Upstream Behavior

1. **Lossy-save warning (CONVENTIONS §1.1; since Phase 2a).** The stock save path
   emits `openpyxl.errors.LossySaveWarning` (a `UserWarning`) when the workbook was
   loaded from a file containing content the regenerating save cannot preserve
   (worksheet extensions, shapes, VBA without `keep_vba`, chart extensions/aux parts,
   non-default app.xml, customXml, printer settings). Warning only — never an
   exception — and silent for files with nothing to lose.
2. **Preserve mode (CONVENTIONS §1.1; since Phase 2a).** `load_workbook(path,
   preserve=True)` opts into the spine: source bytes retained, lossless splice save,
   typed refusals. Pure opt-in; upstream code never enters it. Under preserve, the
   stock load-time "extension is not supported and will be removed" warning is
   suppressed (it would be false — the splice preserves extensions).
3. **No `properties.modified` auto-stamp under preserve (PR-0 D3).** Preserve-mode
   saves raw-copy `docProps/core.xml` unless the user explicitly changed
   `wb.properties`; the stock path keeps stock stamping. Required by the pinned
   no-op payload-identity invariant.

## Future Breaking-Change Candidates

1. Flip `preserve=True` to the default — only after the fixture corpus (including
   the real-Excel bucket) proves the spine.
2. Make the `data_only`+save refusal apply on the stock path by default.

## Phase 0 — Orientation (2026-07-07)

- Baseline re-verified on the development machine (`.venv`, Python 3.13.3, lxml 5.4.0,
  pandas 3.0.3): `2592 passed, 6 skipped, 7 xfailed in 2.77s` — matches the fork-point baseline
  above. Raw log: `scratch/results/baseline_pytest.txt` (gitignored spike area).
- Provenance re-verified: full converted history (9,142 commits, 123 tags);
  `paper-base` == `3.1.5` == `c4986390b`; PyPI's latest openpyxl is still 3.1.5 as of
  2026-07-07, so the fork base is current upstream stable.
- Deliverables: `agent_docs/ARCHITECTURE-NOTES.md` (source tour),
  `agent_docs/OPEN-QUESTIONS.md` (ten open questions answered with evidence, cross-cutting
  gaps, and flags against pinned shapes for human decision), `FIXTURE-REQUESTS.md`
  (real-Excel fixtures a human must author).
- Performance seeds for the Phase-2 guardrail (large synthetic fixture, 3.39 MB / 600k cells):
  stock load 2.505 s, stock save 2.174 s, LibreOffice warm convert 2.09 s.
- Hygiene note: a leftover `soxhub` git remote points at `/tmp/soxhub-openpyxl` (the
  hg-conversion staging clone); candidate for removal, left untouched pending owner decision.

## Phase 1 — Test infrastructure (2026-07-08)

- Fixture corpus frozen: 18 fixtures under `tests/paper/fixtures/` with pinned-schema
  sidecars and `MANIFEST.sha256` (enforced by `tests/paper/test_manifest.py`). All
  provenance is openpyxl-authored / zip surgery / LibreOffice conversion — honestly
  labeled; the real-Excel bucket is requested in `FIXTURE-REQUESTS.md`.
- Five-job battery in `tests/paper/test_battery.py`: `TestStockCarnageBaseline`
  (passes today; regression-guards the damage model with the Phase-0-corrected claims)
  and `TestBatterySafety` (the forever criterion as strict xfails, each naming the
  phase that must flip it).
- Contract-harness helpers in `tests/paper/support/` (part-payload diff, semantic XML
  diff that never normalizes cell text, refusal-atomicity assertion, LibreOffice test
  driver with per-invocation profile isolation and temp-copy discipline).
- `pytest.ini`: registered the `lo_smoke` marker. CI: added a `test-libreoffice` job
  (ubuntu, Python 3.13, `libreoffice-calc`, `PAPER_REQUIRE_LO=1` promotes skips to
  failures).
- `setup.py`: `find_packages` exclude extended with `"tests", "tests.*"` so the new
  top-level test package cannot ship in the wheel.
- Full suite after Phase 1: 2617 passed, 6 skipped, 12 xfailed (2592 upstream tests
  unchanged and green).


## Phase 1.5 — PR-0 API proposal (2026-07-08)

- `agent_docs/PR0-API-PROPOSAL.md`: the v0 design contract. Freezes the delegated
  decisions (inline strings everywhere; per-operation-class collateral sets; no
  core.xml auto-stamp under preserve; performance budget 1.5x stock save, evidence
  0.16x composed prototype; frozen three-tier chokepoint inventory; splice guard
  set; shared-formula dissolve-on-touch; sheet delete/rename/reorder refuse in v0;
  mixed-chart semantics; rels append-only policy; calcChain deletion cascade).
- Sanctioned deviations register grows by one (recorded in PR-0 §10 and below):
  preserve-mode save does not auto-stamp `properties.modified` (stock path
  unchanged) — required by the pinned no-op payload-identity invariant.
- G6/G9 evidence spikes: `scratch/probes/pr0_composed_save.py` (preserve save
  prototype 0.381s vs stock 2.325s on 600k cells; untouched payloads verified
  byte-identical), `scratch/probes/pr0_g9_chokepoints.py` (chart mutation -> chart
  part; ws._rels.append discarded by stock save; code_name -> workbook.xml;
  template -> [Content_Types]).

## Phase 2a — Retention, kernel, lossy-save warning (2026-07-08)

- `openpyxl/errors.py`: the pinned refusal taxonomy (`PaperRefusal` + seven
  subclasses) and `LossySaveWarning` with structured `.losses`.
- `openpyxl/package/`: the kernel — `xml_equivalent` (semantic, never normalizes
  cell text), `diff_package` → `PackageDiff.to_dict()` (schema `package_diff` v1).
- `openpyxl/preserve/`: `zipio` (deterministic entries; raw compressed-stream copy
  with D10 guards + recompression fallback; atomic path delivery via temp +
  `os.replace` with failure-injection test; file-like seek/write/truncate delivery),
  `inventory` (content-level loss scan built at load), `saver` (Phase-2a stub: a
  typed atomic refusal until the splice lands in 2c).
- `load_workbook(..., preserve=True)`: eager byte retention (file-likes rewound —
  the pandas handle case), `wb.preserve`, `ValueError` on `preserve+read_only`
  raised before any handle opens, `.xls`/`.xlsb` extension refusals kept.
- Battery: jobs 3 and 5 flip to green-by-refusal (the blanket preserve-save refusal
  satisfies the criterion; Phases 6a/3 narrow it and must keep them green); jobs 1,
  2, 4 remain strict xfails for 2c/2d.
- Full suite: 2661 passed, 6 skipped, 10 xfailed (upstream 2592 green).

## Phase 2b — The dirty ledger (2026-07-08)

- `openpyxl/preserve/ledger.py`: per-cell dirt keyed by worksheet, formula-change
  flag (drives the calcChain cascade and recalc-on-load), added-sheet tracking,
  `mark_dirty` target parsing, and a style-registry fingerprint that converts the
  StyleProxy nested-mutation leak (upstream silent fan-out corruption) into a
  typed save-time refusal.
- Chokepoints instrumented (Tier 1): `Cell._bind_value` (marks only after
  validation succeeds), the four style descriptors + number-format + named-style
  descriptors, hyperlink/comment setters, `cell.data_type` (converted to a
  property — direct assignment silently demotes formulas and is now a chokepoint;
  the reader writes the backing slot directly to keep the load hot path fast),
  `Worksheet.__delitem__`.
- Refusals installed at the chokepoints PR-0 D7/D8 pin: `insert_rows/insert_cols/
  delete_rows/delete_cols/move_range` on loaded sheets (refined by Phase 6),
  sheet remove/rename/reorder/copy for loaded sheets. All raised before any
  mutation; in-session sheets are exempt (they are generated whole at save).
- Tier-2 satellite regions need no load-time snapshots: the retained blob IS the
  snapshot; the splice save compares faithful re-serializations of fully-modeled
  elements against the original bytes (2c/2d).
- The ledger arms only after load completes; reads/materialization never dirty
  (tested). Perf: large-fixture load 2.522s vs 2.505s baseline (+0.7%, noise).
- Considered and deferred: type-check overrides for `DefinedNameDict.update/|=/
  setdefault` (would be a hard behavior change; upstream-bug note instead).
- Full suite: 2695 passed, 6 skipped, 10 xfailed.

## Phase 2c — The splice writer (2026-07-08)

- `openpyxl/preserve/xmlscan.py`: namespace-tracking byte scanner over original
  sheet XML — spans for regions/rows/cells, shared-formula/array/cm-vm inventory,
  and the full D6 guard set (DOCTYPE, non-UTF-8, prefixed/non-main namespaces,
  r-less rows/cells, exact-parent-chain matching so extLst/AlternateContent decoys
  can never be edited). Optimized hot loop (byte dispatch, balanced-quote fast
  tag-end, selective attribute parsing).
- `openpyxl/preserve/regions.py`: faithful per-region serializers mirroring the
  stock writer; arm-vs-save model-serialization diffing (zero producer-quirk false
  positives); pinned CT_Worksheet order for inserting regions that did not exist.
- `openpyxl/preserve/emit.py`: cell emission through upstream's write_cell with
  the two side effects owned (hyperlink append guarded; style interning wanted);
  D6 attribute-carry rule (ph and foreign attrs survive, cm/vm refused upstream).
- `openpyxl/preserve/splice.py`: the byte-range splice — replace/insert/delete
  cells at scanned spans, dissolve shared-formula groups on touch (D7), region
  replacement with extLst and x14-DV gates, new-row insertion at sorted positions.
- `openpyxl/preserve/crosscheck.py`: ledger cross-check (PAPER_LEDGER_CROSSCHECK=1,
  on for the whole paper test suite): a splice-changed cell the ledger never
  recorded raises hard — corruption inside the safety tooling.
- `saver.py` rewritten: full validation before the first output byte (data_only,
  style-registry guard, workbook-level/custom-props/chartsheet/comment change
  detection), rels-driven sheet-part resolution, core.xml raw-copied unless
  wb.properties changed (D3), theme sync, raw copy for everything untouched.
  Cross-part operations (added sheets, new styles, CF, hyperlinks, tables,
  calcChain cascade, workbook.xml, mark_dirty parts) refuse loudly pending 2d.
- **PR-0 D4 amended with evidence** (never silently): performance budget is now
  2x stock save — the production scanner measures 1.82x (600k cells) / 1.87x
  (150k); the 1.5x pin was seeded by a spike-grade scanner. expat/lxml span
  acceleration recorded as the non-semantic contingency.
- No-op round trips are byte-identical on every fixture class incl. LO-authored
  and .xlsm; the splice-completeness trap (sparklines + x14 twins + drawing ref
  survive a one-cell edit with exactly one part changed) is green; battery jobs
  1 and 4 flip to green (2, awaiting sheet-add support, stays xfail for 2d).
- Full suite: 2731 passed, 6 skipped, 8 xfailed.

## Phase 2d — Cross-part handling (2026-07-08)

- `openpyxl/preserve/crosspart.py`: targeted byte edits against original
  payloads — small-part span scanner; [Content_Types] append/remove; rels
  append (max-numeric+1 rIds, PR-0 D11) and removal; workbook.xml per-element
  splice ({definedNames, calcPr, bookViews} + sheets-element state patches and
  appended entries); styles.xml append-only planner (fonts/fills/borders/
  numFmts/cellXfs/dxfs with count bumps; indices computed from original
  counts, never IndexedList.index — the measured wrong-index bug).
- Added sheets (the pandas case): part generated by the stock writer, part
  name allocated max+1 from the retained namelist, sheetId/rId appended, CT
  override appended; own rels part generated for hyperlink-bearing sheets.
  Battery job 2 goes green — ALL FIVE battery jobs now pass.
- calcChain cascade (D13): part + content-type override + workbook rel all
  removed on formula change; kept byte-identical otherwise.
- Conditional formatting lifted from detect-only: stock-writer-mirrored dxf
  allocation, x14 twin gates (refusal when a rule carries the twin pointer or
  the sheet extLst holds x14 CF — measured against the gauntlet).
- Hyperlink ADDITION on loaded sheets (element render + appended rels;
  removal/modification refuses — it would rewrite preserved relationships).
- v0 refusals kept typed and atomic: comment changes, table lifecycle,
  charts/images/comments/tables on added sheets (PR-0 D9 partially deferred:
  fresh charts on preserved workbooks refuse in v0 — recorded amendment),
  new named styles, workbook.xml elements outside the spliceable set,
  custom-props part creation, non-worksheet mark_dirty parts.
- Full suite: 2754 passed, 6 skipped, 7 xfailed (only upstream xfails remain).

## Phase 3 — Honesty organs (2026-07-08)

- data_only trap (PLAN Phase 3): preserve-mode save refuses with the typed
  error naming `wb.save(path, allow_formula_loss=True)`; with the override,
  ONLY edited cells lose formulas (untouched cells keep them in the preserved
  bytes — the trap is defused, not just fenced). Stock path warns loudly
  (`LossySaveWarning`, "PERMANENTLY replaces every formula"); the override
  silences it. `Workbook.save`/`save_workbook` gain the keyword-only flag.
- Recalc-on-load: any formula-affecting edit under preserve forces the calcPr
  splice with `fullCalcOnLoad="1"` (the model defaults the flag, so it is
  forced into the plan rather than snapshot-diffed) — a human opener's Excel
  always computes fresh numbers. Value-only edits leave workbook.xml
  byte-identical.
- `.xls`/`.xlsb` under preserve raise `UnsupportedStructureError` naming the
  format and the LibreOffice conversion command; the stock path keeps
  upstream's `InvalidFileException` unchanged.
- LossySaveWarning enumeration is now deterministically ordered.
- **Correctness fix uncovered by the honesty tests** (and the reason
  PR-0 D2 pinned explicit style indices): model style numbering drifts from
  the file's on non-openpyxl producers (`_normalise_numbers` rewrites
  numFmtIds in place; the Normal-style bootstrap appends arrays), so emitting
  `cell.style_id` corrupts s indices on LO-authored files (measured
  IndexError on reload). New `openpyxl/preserve/styletrans.py`: parses the
  ORIGINAL styles.xml through upstream's own Stylesheet machinery and
  translates every emitted style to FILE xf numbering, allocating appended
  xfs/numFmts in file numbering; added-sheet parts get their s attributes
  rewritten through the same table; `emit.py` is now the PR-0-D2 thin variant
  of the upstream cell writer with the style index explicit. Regression
  tests on the LO-authored fixtures.
- Full suite: 2769 passed, 6 skipped, 7 xfailed.

## Phase 4 — Perception (2026-07-08)

- `Workbook.manifest()` -> `WorkbookManifest.to_dict()` (schema
  `workbook_manifest` v1): per-sheet dimensions/formula counts/tables/merges/
  CF/DV/freeze/local defined names; workbook defined names; volatile-function
  detection per the pinned §3.7 table (nondeterministic vs deterministic
  reported separately); a confession block enumerated from the PACKAGE (the
  retained bytes under preserve, the loss inventory otherwise — never the
  model, which under-reports exactly the at-risk content); and a preservation
  block stating the active mode's guarantee. Golden-tested against the
  gauntlet (goldens update only via explicit command).
- `openpyxl.package.diff_cells(a, b)` -> `CellsDiff` (schema `cells_diff` v1):
  address + old/new value + old/new formula, deterministic order, using the
  formula view and the cached-value view like the ecosystem does.
- `openpyxl.preserve.perception.dependency_sketch(wb)`: tokenizer-based
  which-cells-feed-which map with cross-sheet resolution, defined-name
  expansion via destinations, and conservative handling of structured/table
  references (unresolved => intersects everything) — the Phase 6a guard input.
- Full suite: 2782 passed, 6 skipped, 7 xfailed.

## Phase 5 — The LibreOffice oracle (2026-07-08)

- `openpyxl/oracle.py`: driver per the measured Q10 rules — temp copies only
  (tested invariant: the caller's path never appears in the soffice argv),
  per-invocation `-env:UserInstallation` profiles, success = rc 0 AND output
  exists, stderr never parsed, process-group kill on timeout
  (`OracleTimeoutError`), typed absence (`OracleUnavailableError`, detection
  soffice -> libreoffice -> macOS app bundle).
- `recalc(source, *, output_path/in_place, timeout)` -> `RecalcResult`
  (schema `oracle_recalc` v1, skill-compatible shape: status/cells_scanned/
  formula_cells/error_cells/errors with sheet+cell+token).
- `certify(source)` -> CERTIFIED / DIVERGED (addresses + both values) /
  BASELINE_UNVERIFIABLE (empty `<v></v>` counts as absent), tolerance
  rel 1e-9 / abs 1e-11, text and errors exact; cells downstream of
  NOW/TODAY/RAND/RANDBETWEEN excluded via a taint fixpoint over the
  dependency sketch; INDIRECT/OFFSET stay in (pinned §3.7).
- **Operational discovery beyond Phase 0:** LibreOffice's headless converter
  does NOT honor calcPr fullCalcOnLoad/forceFullCalc for cells that already
  carry cached values — a tampered-cache fixture "certified" against its own
  tamper. The driver now (a) byte-patches calcPr on the temp copy and
  (b) pre-seeds each fresh profile with OOXMLRecalcMode=0 ("always
  recalculate on load") via registrymodifications.xcu — measured to recompute
  over existing caches. Without (b) the oracle premise silently fails.
- Custody never depends on this module (absence only affects oracle APIs).
- Full suite: 2796 passed, 6 skipped, 7 xfailed.

## Phase 6a — The structural-edit guard (2026-07-08)

- `openpyxl/preserve/structural.py`: `analyze_shift` enumerates what a
  row/column shift would strand — formulas via the dependency sketch
  (cross-sheet included), defined names via destinations, merged ranges,
  CF/DV sqrefs, table extents, and series ranges inside PRESERVED chart
  bytes (byte-scan of retained chart parts for the sheet name — raw-copied
  charts cannot be rewritten, so refusal is the only honest v0 answer).
- The preserve-mode refusal (in place since 2b) is now informative: it names
  every victim by address (e.g. 'Schedule'!B12, 'Summary'!B1, defined name
  Growth — exactly Q11's measured 1100/6399/5400-vs-7499/6500 corruption
  set) and the options. Battery job 3 stays green-by-refusal.
- Stock path: loaded workbooks now get `StructuralShiftWarning` on
  insert/delete rows/cols and move_range ("moves cells but updates
  NOTHING that points at them"); fresh Workbook() construction stays
  silent; stock behavior itself is unchanged.
- Full suite: 2803 passed, 6 skipped, 7 xfailed.

## Phase 6b — The reference rewriter (2026-07-08)

- `openpyxl/preserve/rewrite.py`: Excel INSERT/DELETE semantics — endpoint-wise
  shifting (absolutes move too; spanning ranges expand; deleted references
  become `#REF!` exactly as Excel writes), over Tokenizer operands with sheet
  prefixes and quoting handled; `Translator` untouched (fill semantics,
  load-bearing for shared-formula expansion).
- A shift on a fully-modeled sheet now PROCEEDS under preserve: model-side
  fixups (formulas workbook-wide, defined names incl. sheet-scoped and print
  settings, merges/CF/DV/autoFilter ranges, row display attributes, hyperlink
  anchors) + positional arm-snapshot rebasing, then at save a byte-level
  renumber pre-transform (deleted rows cut, shifted r attributes rewritten,
  every other byte verbatim) that becomes the standard splice's baseline.
  Shared-formula groups on shifted sheets dissolve-on-touch.
- Support matrix is honest: sheets carrying extLst (sparklines/x14), array
  formulas, comments, legacy drawings, tables, manual page breaks, or
  referenced by preserved charts/pivots still REFUSE, with blockers and
  victims named; one structural edit per sheet per session (save between).
- Battery job 3 flips from green-by-refusal to green-by-rewrite: the oracle
  computes 7499 / 7873.95 — the correct values where stock silently produced
  1100/6399/5400. Insert-then-delete round-trips clean (PLAN property test).
- Bug found by these tests and fixed: the first hyperlink added to a loaded
  sheet with no rels part left a dangling r:id (the new rels part existed
  only in the plan); regression-tested.
- Full suite: 2834 passed, 6 skipped, 7 xfailed.

## Phase 6c — Chart-range rewriting (2026-07-08)

- `openpyxl/preserve/chartpatch.py`: namespace-aware leaf-text walker over
  preserved chart/drawing bytes; rewrites `c:f` series reference texts (both
  the openpyxl default-namespace form and prefixed producer forms) and
  `xdr:from`/`xdr:to` anchor markers, entity-safe, everything else verbatim.
- Wired as a dry-run in the shift blockers (charts referencing the sheet no
  longer refuse when patchable) and as real part plans at save time. Charts
  carrying c15 filtered-series machinery, extension lists or
  AlternateContent still refuse; a shift that would delete charted data
  refuses rather than write #REF! into a chart — the refusal stands where
  the patch cannot be honest (PLAN's 6c scope rule).
- Verified: value series move with their data while header refs stay; chart
  and drawing bytes are untouched when the shift is entirely below the
  charted region; patched output loads in LibreOffice.
- Full suite: 2839 passed, 6 skipped, 7 xfailed.

## Final adversarial review (2026-07-08)

Before the implementation PR, a 40-agent adversarial review (6 module
reviewers, every finding independently verified with a runnable repro; 0
refuted) confirmed 34 defects. ALL are fixed and regression-tested
(`tests/paper/test_review_regressions.py`); the review artifacts are in
`scratch/results/final_review/confirmed.json`. Highlights, worst first:

- Row/column style indices reached spliced bytes untranslated (dangling xf,
  IndexError on reload — the PR-0 D2 rule applied to cells but not
  dimensions). Row s= and cols style= now route through the StyleTranslator.
- CF + DV inserted together landed in schema-invalid order (same-offset
  insertions now tie-break by the CT_Worksheet sequence).
- A cell whose r disagrees with its parent row (off-spec but loadable)
  spliced a silent duplicate reference; now a typed refusal.
- Pre-shift cell edits were double-remapped and silently lost (ledger dirt
  now rebases BEFORE the shift fixups); split shared-formula groups could
  re-derive wrong formulas (all groups on shifted sheets dissolve); a
  hyperlink on a deleted row re-attached to the row that shifted up (the
  element re-renders whenever the original had hyperlinks); chart patches
  for multi-sheet shifts overwrote each other (incremental overrides).
- Sheet names needing XML escapes ('P&L') were invisible to the sheets-state
  patch and the chart/pivot reference scans (entity-aware compares; sheet
  matching is case-insensitive like Excel).
- Silent-drop refusal gaps closed: add_chart/add_image and create_chartsheet
  under preserve, wb.template toggles, deleting the last custom property
  (was a crash), the ExcelWriter bypass of the preserve dispatch, the
  append(Cell) write-only-compat path (now ledgered), full-column
  mark_dirty ranges (was a TypeError).
- Reads-never-dirty restored for column dimensions (a pure read materialized
  a visible <cols> entry into the output).
- Oracle: volatile taint-seeding is tokenizer-precise (string literals no
  longer shrink the divergence check); booleans never compare equal to
  numbers; recalc output for .xlsm refuses (LibreOffice conversion would
  strip VBA); timeouts kill the whole process group as documented.
- Contract clarifications: diff_cells scope (values/formulas only) stated;
  manifest confession carries a "source" field and says loudly when
  package-level counts are unavailable on stock loads.
- Stock hot path: the ledger hook cost on fresh-generation writes was cut to
  noise with an inline bail (200k-cell build 0.197s, pre-fork parity).

Post-review suite: 2857 passed, 7 skipped, 7 xfailed.

## Scanner fast paths (2026-07-08)

CI measured the splice save at 2.002-2.005x the stock save on GitHub's
shared runners — over D4's 2x budget by the width of a hair (locally it
measured 1.87x). Rather than amend D4 a second time, the worksheet byte
scanner gained guard-equivalent fast paths for the three hot shapes
(`<c ...>` under a main-namespace row, text-only `<v>`, `</c>`); cell
attributes now decode lazily (only cells an edit actually touches pay).
Anything unusual — namespace declarations, CDATA, nested markup, decoy
end tags — falls through to the untouched generic machinery. Measured:
1.19x locally (was 1.79x), scan cost cut ~4x.

Because this touches the most safety-critical hot path, the change was
differentially verified against the previous scanner: 1,352 fuzz cases
(fixture corpus, 82 crafted hazards, 1,241 seeded mutations) plus a
three-lens adversarial review. The verification caught one real bug in
the first draft — a quote-blind r-attribute regex that an ` r="B9"`
lookalike inside another attribute's quoted value could hijack (openpyxl
loads such files; the splice would have keyed the cell at the wrong
column). Fixed by tokenizing the attribute blob with the same regex the
generic path uses; regression test in test_splice.py
(`test_attr_value_r_decoy_scans_true_column`). Sole remaining
divergence, accepted and documented: invalid UTF-8 in a cell attribute
value (unloadable by any XML parser, so unreachable post-load) now
raises its UnicodeDecodeError at first attrs access during edit
planning instead of at scan time — still strictly before any output is
written.

## Batch 0 — Restore the invariant (2026-07-08, PLAN-v0.1)

Corruption fixes outrank planning; this batch merged before PR-1.

- **Item zero (0.1):** the shared-formula probe came back CLEAN — the splice
  is `si=`-group aware (dissolve-on-touch, `splice.resolve_dirty_cells`).
  Probed adversarially: master/follower/literal/delete edits, two-group
  isolation (untouched group byte-verbatim), gap cells in stale refs,
  orphan + ref-less-host refusals, LO loop closure. Battery job 6: CORRECT.
  Un-share default + the enumerated cache-drop side effect recorded as
  PR-0 amendment 6.
- **0.2:** self-closing region corruption FIXED — the scanner never set
  `RegionSpan.end` for self-closing top-level elements; editing a
  self-closing `autoFilter`/`pageMargins`/`pageSetup`/`sheetFormatPr`
  emitted silently malformed XML (document duplicated after the edit).
  Region x self-closing matrix added (6 regions x edit/no-op arms).
  Battery job 15: CORRECT.
- **0.3:** no-op false-dirty FIXED AS A PATTERN — upstream
  `DimensionHolder.to_tree()` mutates `max_outline` at render time,
  perturbing the next `sheetFormatPr` render; a ZERO-EDIT save corrupted
  cols-bearing sheets (our own hidden.xlsx fixture). The ledger now
  double-renders every region at arm: self-disagreeing regions are PINNED
  (settled second render becomes the snapshot; no-op keeps original bytes;
  USER edits to a pinned region refuse). Any impure upstream serializer —
  present or future — lands in "pinned", never in "false dirty". The
  DimensionHolder INSTANCE was then fixed at the root after the batch gate
  measured the pin's collateral (see the gate entry below): the
  sheetFormatPr render now computes the outline sync purely, so NOTHING
  pins on the shipped corpus (asserted corpus-wide), sheetFormatPr edits
  and column grouping work, and the pattern guard stays armed for the next
  impure serializer (proven by a synthetic one in tests). Battery job 14:
  CORRECT.
- **0.4:** permanent property infrastructure — no-op byte-identity across
  EVERY loadable corpus fixture (glob-enumerated; the 0.3 bug survived v0
  because the hand-listed no-op test skipped hidden.xlsx), and the ledger
  cross-check extended to REGION claims (an unclaimed region may never
  differ; wired through save_preserved, active suite-wide).
- **0.5:** battery grown to the 24-job table — jobs 7-13 and 16-24 land as
  today-state tests (each names the batch that flips it); jobs 14/15/6
  flipped to their required states by this batch. `PAPER_PRESERVE_DEFAULT=1`
  env switch shipped (a default, not a mandate: read_only loads fall back
  to stock); paper-internal harness images set it at Batch-0 exit — the
  PUBLIC default stays False behind the Appendix-A release gate (region
  matrix green + battery green + real-file soak).
- **Process amendments:** the pinned-surface CI check
  (tests/paper/test_pinned_surface.py) mechanizes "pinned means produced,
  tested, or ledgered" — the AddressRemap breach class cannot recur
  silently; the adversarial review is a standing per-batch gate (this
  batch's report is in the PR).

## Batch 1 — Honesty completion (2026-07-08, PLAN-v0.1)

The boundary class, closed: **"the dangerous boundary isn't unmodeled
content — it's modeled objects backed by preserved bytes."**

- **1.1 Preserved-object guards:** loaded tables, charts, images, and
  pivots get settled-serialization snapshots at arm (the 0.3 double-render
  discipline); any in-session mutation refuses atomically at save, naming
  the object and its unlock batch (tables: Batch 2; charts/images: Batch
  4). External-link objects snapshot at workbook level. Battery jobs 10
  and 16 flipped from silent-staleness to REFUSE.
- **1.2 fullCalcOnLoad widened:** a VALUE edit intersecting the dependency
  sketch of any formula forces the recalc flag (case-insensitive sheet
  match; structured refs count conservatively). calcChain is untouched
  (still valid for value-only edits). Battery job 17 flipped.
- **1.3 Pinned-surface debts paid:** AddressRemap implemented
  (insert/delete rows/cols return it under preserve; .map() handles
  cells/ranges/sheet-qualified/absolute forms; deleted addresses map to
  None); BoundaryViolationError raised when a shift would push occupied
  cells past row 1048576 / column XFD; RelationshipPolicyError retyped
  onto the hyperlink modify/remove refusal (its pinned domain). All three
  debt entries removed.
- **1.4 Loss-inventory completeness:** chart colors/style parts matched by
  numbered-name regex (the v0 endswith was dead code — verified against
  lo_authored.xlsx which loses both parts on stock save); rich-text run
  flattening detected in sharedStrings AND inline strings; workbook.xml
  fileSharing and protectedRanges; threaded-comment parts.
- **1.5 Input honesty:** encrypted/CFB files get a typed refusal naming
  the condition and the decrypt route on BOTH load arms (battery job 13
  flipped); duplicate zip entry names refuse under preserve (load-vs-copy
  parser differential); diff_cells no longer sprays stock loss warnings
  during read-only diagnostics (and pins preserve=False so the env
  default cannot double-retain).
- **1.6 Protection awareness:** value writes to locked cells on protected
  sheets warn once per sheet (ProtectedWriteWarning, new pinned surface)
  or refuse atomically under wb.strict_protection — checked BEFORE the
  value binds. Manifest gains per-sheet protection + workbook_protection.
  Protection is reported, never enforced or bypassed. Battery job 9
  flipped. Scope note: style/comment writes to locked cells are not
  protection-checked in v0.1.
- **1.7 Certification noise classes:** external-workbook references and a
  pinned ORACLE_UNSUPPORTED_FUNCS catalog (LAMBDA/LET family, RTD,
  STOCKHISTORY, CUBE*, WEBSERVICE, IMAGE, PY) are excluded-with-reason
  with downstream taint inheritance; CertificationResult gains
  external_excluded/unsupported_excluded.
- **1.8 Producer fingerprint pinned:** fresh app.xml bytes are a pinned
  test constant (changing the producer string is a reviewed decision);
  edited preserve saves keep original app.xml byte-identical (explicit
  test beyond the no-op property); real-Excel open checks queued in
  FIXTURE-REQUESTS.md (LibreOffice smoke is provably blind to this class).

Suite: 2946 passed; env-flip arm green.

## Batch 1 — adversarial gate report (2026-07-08)

Four lenses, one critical + eleven majors confirmed with live repros —
all fixed and fixtured (tests/paper/test_gate_regressions.py):

- **CRITICAL — 3-D span refs invisible:** =SUM(Sheet1:Sheet3!A1) was
  recorded under the phantom sheet key 'Sheet1:Sheet3', so the 1.2
  recalc guard, certification taint, and shift victim analysis all
  silently missed 3-D formulas. Fixed: a ':' in the sheet component
  classifies as unresolved (always-intersecting, conservative).
- **Object-guard evasions:** chart MOVES (anchor lives outside
  chart._write()) — anchors now fingerprinted; image DATA swaps with
  identical anchor+path — backing bytes now digested (non-destructively:
  image._data() CLOSES the ref stream and must never be used for
  snapshots — the first digest attempt corrupted no-op saves that way);
  chartsheet-anchored charts were entirely outside the boundary — now
  snapshotted; the dead 'unstable' bookkeeping now feeds diff_objects as
  the oscillating-serializer fallback (skip-compare, never false-refuse).
- **Recalc-guard gaps:** computed-string INDIRECT/OFFSET targets left no
  sketch footprint — such formulas now always count as unresolved;
  case-sensitive sheet compares in the taint walk and
  cells_referencing — casefolded.
- **Protection evasions:** del ws['A1'] evaded what ws['A1']=None
  refused — __delitem__ now protection-checked pre-deletion; structural
  shifts on protected sheets (Excel blocks them) — warn/strict-refuse
  via the same 1.6 discipline; strict_protection documented as
  preserve-only (inert on stock loads).
- **Perf regression:** the pre-bind hook doubled the per-write lookup
  chain (+13% on the fresh-gen hot path) — hoisted to a single resolved
  bail reused by both hooks.
- **Input honesty:** CFB sniff anchored at absolute offset 0 (a valid
  xlsx handed over at an embedded-CFB offset false-refused; a genuine
  CFB via mid-position handle evaded the typed refusal);
  protectedRanges is a WORKSHEET element — the workbook.xml check was
  dead code (the Batch-1 test itself had planted it in the wrong part);
  rich-text loss warnings suppressed under rich_text=True (the stock
  save PRESERVES runs there — the warning was loud-but-wrong).
- **Certify evasions:** Excel writes _xlfn.LET( — prefix now normalized
  before catalog match; RANDARRAY added to VOLATILE_NONDETERMINISTIC
  (CONVENTIONS 3.7 amendment: nondeterministic dynamic-array RNG);
  external refs hiding behind defined names now seed external-link (both
  in certify and the sketch).
- **Boundary guard:** occupancy now includes dimension-only rows and
  merged-range anchors; inserts beyond all content no longer
  false-refuse (and the message no longer overclaims).
- **Accepted with rationale:** style-only edits count as dirty for the
  recalc flag (value/style indistinguishable in led.cells — conservative
  direction, idempotent recalc); the certify LO test's CERTIFIED arm is
  partially tautological (LO-vs-LO) — its exclusion-membership
  assertions carry the signal, now including the _xlfn form; sketch
  rebuild per save is uncached (0.5s on 50k formulas — revisit only with
  a measured case, Appendix-A-5 spirit).

## Batch 2 — The part-lifecycle engine and its unlocks (2026-07-08, PLAN-v0.1)

- **The engine (PR-1 1.1):** openpyxl/preserve/lifecycle.py —
  PartPlan.add_part/remove_part: part + content-type (Override AND
  Default) + relationship planned in lockstep, applied by the build loop
  after ALL registrations (an ordering bug caught mid-batch: plans
  composed before the custom-props registration silently dropped its CT
  entry). reserve_rid gives every planner touching one rels part a shared
  sequential allocator (two independent next_rid computations collide).
  First consumers: the calcChain cascade (migrated), custom-props
  creation AND deletion (both were v0 refusals), styles.xml creation for
  styled writes into styles-less packages (cells write model indices —
  a fresh part shares the model numbering), and wb.replace_part(name,
  bytes) — the raw media-swap escape hatch with call-time guards.
- **Tables (PR-1 1.2):** loaded-table mutation re-renders the part from
  the fully-modeled Table, located by displayName through the ORIGINAL
  sheet rels; geometry guards (anchor fixed, data region non-empty,
  column count == tableColumns, autoFilter inside ref). Add/remove via
  the engine with the sheet tableParts element rebuilt as saver-crafted
  bytes riding the region splice (per-element xmlns:r; surviving
  originals keep their rIds verbatim). preserve.tables.append_row: totals
  row stays last (cells move down), calculated columns re-derive
  (explicit formula or the column pattern via Translator), autoFilter
  synced, ref extended; refuses content below the table. Battery jobs 10
  and 18: CORRECT. Mid-batch catch: a table-mutation-only session
  initially skipped the saver's work gate — the silent fourth outcome
  resurrected for exactly one edit shape; fixed and battery-tested.
- **Comments (PR-1 1.3):** creation on comment-free sheets — comments
  part + legacy-VML part generated whole (CommentSheet/from_cell, the
  stock writer's own machinery), one <legacyDrawing r:id> spliced;
  vml Default appended to [Content_Types] when absent. Sheets already
  carrying comment machinery keep refusing (editing preserved VML is
  Batch-4-class work). Battery job 19: CORRECT.
- Scope refusals kept honest: comments/tables + hyperlink changes on the
  same sheet in one save refuse (the hyperlink planner allocates rIds
  outside the engine); comments+tables coexist via reserve_rid (tested).

Suite: 2975 both arms.

## Batch 2 — adversarial gate report (2026-07-08)

Four lenses, five criticals + six majors confirmed with live repros —
all fixed and fixtured (tests/paper/test_gate_regressions.py Batch-2
sections):

- **CRITICAL, duplicate workbook rIds:** added sheets allocated rIds
  outside the engine while styles.xml creation allocated inside it —
  one save produced two rId4 Relationships (OPC unique-Id violation).
  Added sheets now reserve through the engine's shared allocator.
- **CRITICAL, shadowed hyperlink rels:** table REMOVAL + hyperlink add
  on one sheet — the engine's rels payload was written with `continue`
  before the hyperlink planner's, leaving a dangling r:id (reload
  KeyError, URL lost). Engine rels now compose ON TOP of the planner
  payload.
- **CRITICAL, table extLst dropped:** to_tree() omits Table.extLst, so
  mutating any table silently stripped alt-text/x14 extensions.
  Mutation now refuses when the original part carries extLst or xr
  revision ids (the splice's own region discipline, applied).
- **CRITICAL, basename rel removal:** removing table1.xml also cut a
  sibling mytable1.xml's relationship (suffix matching). Removal now
  resolves each relationship target against the rels owner and removes
  on exact equality.
- **CRITICAL, comment control chars:** comment text/author had no
  illegal-character guard (cells do): under the stdlib serializer the
  save wrote an unparseable part. Typed refusal added.
- **Majors:** single-quoted ref attributes silently disabled the anchor
  guard (both-quote regex + refusal when ref is unlocatable);
  append_row validated calc columns AFTER moving the totals row (now
  validate-then-mutate, atomic, model untouched on refusal — and the
  freed totals slot is restyled as a data row); table @id now
  workbook-unique (scanned package-wide); multi-sheet table/comment
  creation collided on part numbers (allocators consult engine-added
  names); ct Default handling is semantic (either quote style,
  case-insensitive extension; different ContentType refuses);
  replace_part guards widened (rels + table parts managed) and
  contradictory combos (swap + drop/re-render of one part) refuse at
  save; comment height/width joined the snapshot (resizes on
  machinery sheets refused, not dropped); displayName uniqueness
  enforced against defined names and other tables (casefold).
- **Accepted/noted:** replace_part + calcChain-drop combo refuses (the
  cascade wins by refusal, never silently); removed parts may leave
  grandchildren (queryTables) as OPC-legal orphans — dead weight noted
  for the Batch-3 lifecycle audit; crosscheck still verifies worksheet
  parts only (rels/table/comment parts are outside it) — noted as the
  standing tooling gap, revisit with Batch 3's crosscheck extension;
  corpus lacks styles-less/custom-props frozen fixtures (synthesized
  inline in tests; queued for the next corpus regeneration).

## Batch 3 — Region and structural completion (2026-07-08, PLAN-v0.1)

- **3.1 x14 twin-sync** (the highest-traffic refusal, lifted): twin-
  bearing CF composes from ORIGINAL bytes (the model drops <x14:id>
  pointers on re-render, measured) — survivors verbatim, deletions
  remove the GUID-matched twin entry, sqref-only changes patch classic
  attribute AND twin xm:sqref in lockstep, new rules append as model
  renders. DV: the blanket D15 refusal narrows to classic/x14 sqref
  OVERLAP. Battery 20: CORRECT.
- **3.2 sheet lifecycle:** rename cascades (model formulas + defined
  names rewritten at set time, tokenizer-based with 3-D span endpoints;
  chart <c:f> byte-patched; workbook.xml name-patched on ORIGINAL
  bytes; ledger re-keyed so a renamed sheet never masquerades as added;
  INDIRECT-textual and pivot references refuse) — battery 8 CORRECT.
  copy_worksheet registers as an ADDED sheet; comments on added sheets
  generate via the stock writer's anysvml legacyDrawing + the engine —
  battery 11 CORRECT. Delete runs the reference audit (refuse with
  enumeration) then cascades at save (part + rels + exclusive closure
  with reference counting; RemovalReport pinned). Reorder rebuilds the
  sheets element from ORIGINAL entry bytes; definedNames/bookViews
  force-re-render (position-derived).
- **3.3 structural widening:** multiple shifts per session compose (the
  one-shift refusal was pure conservatism — fixups/rebases run at edit
  time in order, the byte renumber replays in order); spanning merges
  already followed Excel semantics (expand/shrink/move — now pinned by
  tests); move_range lands as TRACKED CELL EDITS (no byte renumber) with
  coherence guards (merges/CF/DV/tables intersecting either rectangle,
  and outside formulas referencing the moved block, refuse with
  victims).
- **3.4 dynamic arrays:** ordinary value writes on cm/vm cells are
  CORRECT — the overwrite ends the cell's rich-value/spill role, so the
  attributes drop (never carry; unreferenced metadata records are legal
  dead weight). Writing INTO a spill/array range keeps refusing, now
  with the in_spill context naming the anchor. Battery 21 CORRECT;
  battery 7 refuses with context.
- **3.5 structured references:** never mis-shifted, by construction —
  tables on the shifted sheet block the shift, bracketed operands are
  never rewritten (pinned by tests).

## Batch 4 — Charts and images under preserve (2026-07-08, PLAN-v0.1)

- **4.1 added sheets:** ws.add_chart/add_image on ADDED sheets — the
  stock writer's own drawing serialization (SpreadsheetDrawing._write)
  routed through the lifecycle engine (preserve/drawings.py): chart
  parts, media parts (semantic ct Defaults per image format), drawing +
  drawing-rels parts, and the sheet's drawing rel Target filled into the
  generated payload. Zero splice risk. Charts are single-use across
  sheets (refusal, mirroring stock's InvalidFileException).
- **4.2 loaded sheets:** machinery-free sheets get a FRESH drawing part
  via the engine plus exactly one `<drawing r:id>` element spliced at
  its CT-schema position (inline xmlns:r — the tableParts lesson);
  sheets with an EXISTING drawing get new anchors appended INTO the
  original bytes — only when that drawing is anchor-only (top-level
  children all anchors, no comments/CDATA/PI), with rIds reserved on the
  original drawing rels, cNvPr shape ids bumped past the existing
  maximum, and a default-xmlns declaration injected on appended anchors
  when the host document is prefix-namespaced. Non-anchor-only drawings
  refuse AT ADD TIME (atomic — the object never joins the model).
  Battery 22: CORRECT.
- **4.3 chart editing:** the Batch-1 blanket mutation refusal lifts
  per-property. At save, the armed settled render and the current
  settled render are compared with every expressible text span
  neutralized (<c:f> formula texts + <a:t> text runs, namespace-aware
  via chartpatch's leaf walker); if anything else differs, refuse naming
  the property (e.g. "a property near <style> changed"). Expressible
  drift patches the ORIGINAL part bytes positionally, each patch
  verified against the arm state verbatim (mismatch = another rewrite
  already touched it → refuse, advise separate sessions). New ranges
  validate as sheet-qualified single-area A1 ranges on existing sheets;
  new text validates against ILLEGAL_CHARACTERS_RE. Convenience verb
  `chart.repoint(series_index, new_range)` (works in stock mode too);
  title assignment expresses the same way. Cached series values stay —
  Excel re-reads series from cells at render. Chart property edits +
  shifts in ONE session refuse (composing the two rewrites would
  double-shift the new range). The chart part name is stamped at read
  (`chart._paper_part`) by reader/drawings.py. Battery 16: title edit
  CORRECT, inexpressible mutations refuse (contract: "refuse or
  correct").
- Every 4.x output queued in agent_docs/FIXTURE-REQUESTS.md for
  real-Excel open checks (the producer-sensitive surface).

## Batch 5 — Computation layer (2026-07-08, PLAN-v0.1)

- **5.1 scenario runner:** `wb.evaluate(set={...}, read=[...])` (preserve
  workbooks; runs against the AS-LOADED source bytes) and
  `oracle.evaluate(source, ...)`: inputs applied to a temp copy through
  the SPINE (every preserve guard applies), LibreOffice recalculates,
  outputs harvested; the original file and live workbook untouched
  (asserted). ONE LibreOffice run serves both the outputs and the
  certification: original caches vs computed, with input-downstream
  cells excluded as the new `input_excluded` class (CertificationResult
  gained the field, additive). Addresses: sheet-qualified single-cell A1
  or defined names; everything else refuses typed (TargetNotFoundError).
  `oracle.evaluate_many(source, cases, read, pool_size=2)`: warm
  per-thread LibreOffice profiles, created lazily, crash-replaced once,
  destroyed before return (PR-1 delegated decision). `Evaluation` pinned
  (schema "evaluation" v1). Battery 12: CORRECT.
- **5.2 pre-flight linter:** `openpyxl.formula.lint.lint_formula` —
  tokenizer-based, never evaluates. Codes: parse-error,
  unbalanced-parens, semicolon-separator (the locale-canonical trap; ';'
  outside array constants only), unknown-function (pinned catalog in
  formula/catalog.py; `_xlfn.` stripped; a warning, never a gate — UDFs
  are legal), unknown-sheet/-name/-table/-column (with workbook;
  structured refs against real table columns; 3-D span endpoints).
  LET/LAMBDA formulas skip name checks (locals are invisible without
  evaluation — documented). Wired into the value-bind chokepoint under
  preserve via `wb.formula_lint` = off|warn|refuse (default warn);
  refusal restores the pre-bind type, dirties nothing. `LintWarning`
  pinned.
- **5.3 certification-gated write-back:** `oracle.write_back(path)`
  recalcs a temp copy, then splices computed values into the ORIGINAL
  package — a NEW splice channel (`cache_writes`): the `<f>` bytes stay
  verbatim, only `<v>`/`t` change, cells claimed to the crosscheck like
  dirty cells; LO bytes never enter the output (macro-safe, vbaProject
  byte-identical, tested). Gated: DIVERGED/BASELINE_UNVERIFIABLE refuse
  unless allow_uncertified=True (loud `uncertified` stamp). Only
  verified or previously cache-less cells are written;
  volatile/external/unsupported classes never. fullCalcOnLoad clears
  only on full coverage. Package-diff confession in the result.
  `WriteBackResult` pinned (schema "oracle_write_back" v1). Battery 24:
  CORRECT, certification-gated.
- **5.4 fresh-generation recalc flag:** satisfied by UPSTREAM behavior —
  stock CalcProperties defaults `fullCalcOnLoad=True`, so every
  fresh-generated workbook already saves with the flag (verified;
  no change needed, recorded here per the PR-1 stock-visible note).

## Batch 6 — Perception and the agent experience (2026-07-08, PLAN-v0.1)

- **6.1 verbs:** `ws.locate(label, prefer="right"|"below")` —
  exact-then-normalized label match, value = nearest non-label
  neighbour; zero matches -> TargetNotFoundError, multiple labels or no
  value neighbour -> AmbiguousTargetError listing every candidate (THE
  DEBT IS PAID — the pinned class is produced and tested; ledger entry
  removed). Battery 23: CORRECT-or-Ambiguous. `wb.search(text_or_regex,
  regex=, values=, formulas=)` -> [{"address","match","kind"}].
  `openpyxl.preserve.scan_errors(wb)` — LibreOffice-free: live error
  values, cached t="e" tokens from the preserved bytes (both load views
  from one workbook), and #REF! markers in formulas.
  `ws.allowed_values(cell)` — list-DV vocabulary (literal or
  range-sourced; unreadable sources -> None). `wb.validate()` — replays
  the full preserve save machinery into a discarded buffer: every
  refusal a save WOULD raise, raised now, nothing written.
- **6.2 model map:** `wb.model_map()` (pinned schema "model_map" v1) —
  inputs/calculations/outputs/constants per formula-bearing sheet via
  the dependency sketch (whole-column refs clamped to populated
  extents); fill-color corroboration recorded as a convention when >=80%
  of inputs share a non-default fill; unresolved references flagged in
  conventions.
- **6.5 manifest enrichment:** per-sheet formula_addresses + part_name;
  workbook-level computation counts + certifiable flag (cached values
  present in the preserved bytes) + protection_summary. Manifest golden
  regenerated (diff purely additive).
- **6.6 edit receipt:** `openpyxl.preserve.receipt(before, after,
  recalc=None)` -> EditReceipt (pinned schema "edit_receipt" v1):
  cells-diff per worksheet part (added/removed/changed refs),
  package-diff (parts changed/added/removed), the loss-inventory
  confession of the AFTER package, optional oracle result.
  `wb.save(path, receipt=True)` returns one for that save
  (preserve-only).
- **6.7 structured refusals:** PaperRefusal gains `.kind` / `.anchor` /
  `.options` (kw-only, default empty; message text unchanged, populated
  progressively — locate's refusals carry them fully).
- **6.8 findings:** `openpyxl.preserve.findings(wb)` — the pinned
  ADVISORY taxonomy (hardcode-in-formula, inconsistent-row-formula,
  error-cell, orphaned-name, external-link, hidden-sheet, hidden-rows,
  merged-hazard, volatile, magnitude-outlier); measurements with
  evidence addresses, never judgments — the fences stand.
- **6.9 diff report:** `openpyxl.preserve.diff_workbooks(a, b,
  remaps=())` -> DiffReport (schema "workbook_diff" v1): cell diffs
  classified content-changed vs shifted-by-structural-edit by pushing
  A-side addresses through the AddressRemap chain; sheet
  membership changes listed.

## Batch 7 — Delivery, hardening, adoption (2026-07-08, PLAN-v0.1)

- **Style verbs (preserve.styleverbs):** `copy_format(ws, src, range)`
  (whole-format copy via style-array reuse, ledger-marked; the D2
  translator resolves at save) and `apply_profile(ws, profile)` —
  profiles are DATA ({role: {number_format/fill/bold/italic/font_color/
  locked}}) applied by MODEL-MAP role, with a named number-format
  library. Both ride the splice; styles.xml stays append-only.
- **Ergonomics:** `wb.set_input(name_or_label, value)` — defined names
  first, then locate() across sheets (multi-sheet hits are ambiguous,
  typed); NEVER overwrites a formula cell (kind=input-is-calculation).
  `wb.protect_for_delivery(password=None)` — locks everything except
  classified inputs, enables sheet protection, returns the report
  (protection is advisory and REPORTED, never presented as security).
  `wb.scrub(remove=...)` — comments (in-session; PRESERVED machinery is
  reported, not silently stripped), metadata, personal, hidden-sheets
  (through the audited removal: a refusal lands in "skipped" — never
  silent). `wb.set_pivot_refresh_on_load()` — refreshOnLoad byte-patched
  onto every pivotCacheDefinition via replaced_parts.
- **Hardening:** deliver() fsyncs the temp file BEFORE os.replace and
  the directory after (durability of the rename); path-target saves
  spool the archive DIRECTLY into the delivery temp file (~1x file-size
  peak memory; the crosscheck env keeps the in-memory build);
  decompression caps at load (2 GiB/part absolute; >500x ratio above
  64 MiB refuses — pinned numbers, documented in reader/excel.py);
  the raw-copy path verifies central-vs-local zip header agreement
  (name, method, CRC, sizes) and falls back to recompression on ANY
  disagreement — a zip-confusion payload is normalized to the central
  directory's view, the view zipfile and Excel read; mark_dirty clamps
  bounded ranges to the populated extent (an oversized range would mark
  millions of phantom DELETIONS).
- **Adoption:** README "The paper API in 90 seconds"; doc/paper.rst
  (the contract, perception, editing, oracle, delivery, refusal
  taxonomy, the release gate). The public preserve-by-default flip
  stays release-gated per PLAN-v0.1 (mechanism shipped, internal images
  flip via PAPER_PRESERVE_DEFAULT, public default awaits the
  FIXTURE-REQUESTS real-Excel queue).

## Batch 6 — adversarial gate report (2026-07-08)

Four lenses, 24 findings confirmed with live repros — deduplicated to
fourteen defects, all fixed and fixtured
(tests/paper/test_gate_regressions.py Batch-6 classes):

- **locate() rework (three criticals + two majors, one root cause):**
  every silent-guess branch of the value walk was a lying instrument —
  the merged interior of the label's OWN merge came back as "the value"
  (unwritable MergedCell, next write crashed raw); a cached formula
  string under data_only, a number-stored-as-text, or any adjacent text
  value was skipped as "another label" and an unrelated farther cell
  returned silently. The walk now REFUSES instead of guessing: merged
  interiors are covered cells (skipped), error-typed cells are values,
  and an adjacent string with ANY populated competitor raises
  AmbiguousTargetError naming both candidates; a lone adjacent string
  is the value. prefer= validates before matching.
- **diff_workbooks (critical):** new content written at coordinates a
  shift VACATED was invisible (the B-side pass skipped every key
  present in A); under remaps the skip now keys on consumed images
  only. Bool-aware comparison (1 -> TRUE reported as unchanged).
- **search (critical):** ArrayFormula cells were searched by Python
  repr — fabricated matches (memory addresses!) and missed real
  formula text; .text is searched now (scan_errors too). Invalid regex
  raises ValueError with the pattern.
- **openpyxl.preserve.receipt (major):** the same-named submodule
  import shadowed the lazily-exported FUNCTION — the second access
  returned the module ("'module' object is not callable"). The module
  is now receipts.py; the attribute can only resolve to the function.
- **read_only/write_only (majors):** search/model_map/scan_errors/
  findings crashed raw (or silently returned []) — typed ValueError
  naming the materialized-cells requirement.
- **model_map (major):** cross-sheet inputs on formula-free sheets
  were invisible (the sheet was skipped entirely); referenced-only
  sheets now classify their referenced cells as inputs. Manifest
  computation counts follow (golden regenerated, computation only).
- **manifest part_name (major):** stale None after an in-session
  rename; the rename-aware part mapping is now shared
  (hygiene.current_titles_by_part) by scan_errors and the manifest.
- **allowed_values (major+minor):** whole-column sources crashed raw
  (None bounds now clamp to the populated extent); reversed sources
  returned [] posing as "empty vocabulary" (bounds normalize).
- **merged-hazard (minor):** unfireable from the model (shadowed
  interior values are discarded at load) — a byte-level scan of the
  preserved package now provides the evidence.
- **Found and fixed mid-gate:** scan_errors' cached-error regexes were
  double-quote-only (the FIFTH both-quote lesson); the rename/shift
  model cascades tripped the formula linter (refuse mode would have
  refused a legitimate rename) — machinery-internal rebinds now
  suppress lint.
- **Rejected with rationale:** locate refusing across an unmaterialized
  gap (documented walk rule; typed refusal); receipt/BadZipFile on
  non-package input (upstream convention at the package-open boundary);
  manifest ~10x slower on huge sheets via model_map expansion
  (absolute cost sub-second at 150k cells; revisit only with a measured
  hot path); search's raw PatternError verdict was split — the typed
  ValueError wrap shipped anyway.

Suite: 3101 upstream+paper; 509 paper both env arms.

## Batch 5 — adversarial gate report (2026-07-08)

Four lenses, 20 findings confirmed with live repros plus one found
pre-gate — deduplicated to eleven defects, all fixed and fixtured
(tests/paper/test_gate_regressions.py Batch-5 classes):

- **Lint false positives (majors):** the QUOTED storage form of
  external-workbook references ('[Budget.xlsx]Sheet One'!A1) was judged
  against local sheet names — refuse mode blocked legitimate binds, and
  path-form externals reported TWO phantom sheets via the 3-D split
  (external detection now covers the quoted branch); in-session tables
  have no tableColumns until save, so every structured ref against a
  just-added table was refused (columns unknowable → never unknown);
  Excel's ' escape in column specs ([Col'[1']]) mis-split as nested
  parts (escaped specs are unknowable → never unknown); the catalog
  lacked 2023-25 functions (GROUPBY/PIVOTBY/PERCENTOF/TRIMRANGE/REGEX*/
  TRANSLATE/DETECTLANGUAGE/PY) AND the storage-canonical dynamic-array
  operators _xlfn.ANCHORARRAY/_xlfn.SINGLE (the stored forms of A1# and
  @) — all added; eta-reduced function references (REDUCE(0,A,SUM))
  flagged unknown-name (bare names matching the catalog now pass).
- **Certification taint escapes (majors):** with scenario inputs in
  play, cells fed ONLY through INDIRECT/OFFSET (sketch.unresolved)
  escaped the input taint — the evaluation certification falsely
  DIVERGED on healthy workbooks; unresolved formulas now inherit the
  input taint (always-intersecting). A defined name shaped like column
  letters ("IN", "TAX") was parsed as a whole-COLUMN reference by the
  dependency sketch (range_boundaries permissiveness), so its readers
  escaped every taint walk — pure-alphabetic no-colon tokens now resolve
  as names or land in unresolved, never as phantom columns.
- **Write-back gaps:** BASELINE_UNVERIFIABLE early-returned WITHOUT the
  exclusion classes, so allow_uncertified wrote volatile cells and
  their downstream (NOW() caches written as truth) — the early-return
  result now carries volatile/external/unsupported lists;
  write_back's own date serials were judged DIVERGED by its own
  certification (datetime-vs-serial pairs now compare numerically via
  the workbook epoch in _values_match); string caches with significant
  whitespace gain xml:space="preserve". Found pre-gate: an UNCERTIFIED
  write no longer clears fullCalcOnLoad (Excel must never be told to
  trust caches nobody verified — stricter than the PR-1 wording, which
  pinned coverage as the necessary condition).
- **Chokepoint bypass (minor):** ArrayFormula objects carried their
  text past the lint chokepoint (garbage accepted under refuse and
  saved); the dt=='f' branch now lints .text. Merged-cell interior
  inputs crashed evaluate() with a raw AttributeError → typed
  TargetNotFoundError naming the anchor remedy.
- **Rejected with rationale:** lint at 'warn' costs ~25-35µs/bind
  (~8-11x relative on a microbenchmark, invisible in real sessions;
  set formula_lint='off' for bulk binds — noted here); pre-existing
  #DIV/0! cells make Evaluation.status 'errors' (honest attribution,
  the error IS in the recalced copy); evaluate_many pool hygiene
  (leaked soffice) did not reproduce.

Suite: 3070 upstream+paper; 478 paper both env arms.

## Batch 4 — adversarial gate report (2026-07-08)

Four lenses (Workflow orchestration), 21 findings confirmed with live
repros plus two found pre-gate — deduplicated to twelve defects, all
fixed and fixtured (tests/paper/test_gate_regressions.py Batch-4
classes):

- **CRITICAL, rId remap cross-wire:** appending chart+image into an
  existing drawing remapped local rIds by sequential in-place replace;
  when a reserved id equaled a still-unreplaced local id the anchors
  cross-wired (chart frame → PNG, blip → chart XML; output unreadable).
  Two-pass placeholder remap now.
- **CRITICAL, duplicate rIds (drawing + hyperlink):** the hyperlink
  planner allocated next_rid independently of the engine's allocator —
  one save produced two rId1 relationships (OPC violation). Hyperlinks
  now allocate through part_plan.reserve_rid.
- **CRITICAL, file-object image corruption:** _image_payload read
  seekable streams from their CURRENT position (PIL parks it mid-file):
  garbage media bytes saved silently. Reads from offset 0 now (position
  restored) + image-signature validation.
- **CRITICAL, double-unescape:** chartpatch._unescape chained
  str.replace, so a title containing literal '&lt;' text was decoded
  twice and silently rewritten. Single-pass regex (fixes the shared
  helper used by rename/shift chart patching too).
- **CRITICAL, flat positional a:t mapping:** an original that serializes
  valAx before catAx (schema-legal) got the WRONG axis title patched.
  Text/formula leaves now map within ancestor-path groups; structural
  disagreement refuses.
- **CRITICAL, rename skipped in-session charts:** add chart → rename
  sheet → save left the new chart part referencing the dead title. The
  rename cascade now rewrites non-armed model charts' refs (mirrors the
  shift fix).
- **Pre-gate criticals:** the chart single-use seen-set lived on the
  workbook (second save false-refused; now per-save PartPlan); added
  charts' ranges silently ignored shifts (apply_model_shift now rewrites
  them; stranding deletes block pre-move).
- **Majors:** the drawing tag tokenizer stopped at '>' inside quoted
  attribute values (false refusals; quote-aware _find_tag_end walker
  now) and the cNvPr id scan missed single-quoted ids (duplicate shape
  ids; both-quote scan — the FOURTH both-quote-styles lesson); an
  orphan drawing rel (rel+part present, sheet element absent) swallowed
  added charts invisibly (the element is spliced back, referencing the
  existing rel); the shift+chart-edit refusal keyed on ANY shift in the
  session (false refusal for unrelated sheets; now scoped to chart parts
  a shift actually patches); empty self-closing <wsDr/> false-refused
  (expanded before append).
- **Minors:** numeric charrefs in ORIGINAL chart text leaked a bare
  ScanRefusal (typed refusal now); _OBJECT_UNLOCKS still promised
  "editing lands with Batch 4" after it shipped (messages state the
  actual surface).
- **Rejected with rationale:** same Image object added twice writes two
  media parts (stock parity, wasteful not wrong); multi-line title edits
  refuse (run-count change = whole-element surgery — honest, documented
  here); one verifier initially rejected the rename-skips-added-charts
  finding as stock parity — overruled by the record_rename coherence
  contract.

Suite: 3037 upstream+paper; 445 paper both env arms.

## Batch 3 — adversarial gate report (2026-07-08)

Four lenses, seven criticals + a major set confirmed with live repros —
all fixed and fixtured (tests/paper/test_gate_regressions.py Batch-3
classes):

- **CRITICAL, twin modification = silent twin deletion:** editing a
  twin-bearing CF rule reclassified it as delete+new, dropping the x14
  twin (the exact battery-job-1 carnage the module exists to prevent).
  An unconsumed twin block whose sqref a NEW block claims now refuses
  as MODIFIED, naming the range.
- **CRITICAL, byte-identical cross-steal:** two identical classic
  blocks with different twins could consume each other's originals.
  Matching is now order-preserving (same position first, then earliest
  unconsumed).
- **CRITICAL, single-quoted twin ids invisible:** `id='{...}'` twin
  entries escaped GUID extraction — the third both-quote-styles lesson;
  fallback added. Partial multi-GUID twin entries (rules for both
  deleted and surviving blocks) refuse rather than half-delete.
- **CRITICAL, title swap merged reference classes:** sequential
  pairwise chart patching turned an A<->B swap into all-refs-A. Renames
  now patch through ONE simultaneous mapping (charts, formulas, names).
- **CRITICAL, shift+rename corrupted charts:** the shift's chart lookup
  used the CURRENT title against ORIGINAL bytes; rename-then-shift also
  false-refused. Original titles (led.renames) now thread through every
  byte-level lookup, including the removal audit's.
- **CRITICAL, cm/vm stripped on re-emission:** style-only edits and
  dissolution re-emits dropped rich-value/spill metadata; only VALUE
  changes may drop it. `value_overwrites` (bind chokepoint, deletions,
  move src/dst) now discriminates at carry_attributes.
- **CRITICAL, DV coexistence over-matched:** the xm:sqref scan read
  sparkline/other ext ranges as validations (false refusals, and real
  overlaps missable); the scan is now scoped to the dataValidations
  ext block alone.
- **Majors:** freed-title reuse after removal duplicated sheets entries
  (rebuild membership is now by OBJECT, not title); rename + hide in
  one session tripped the internal overlap guard (both compose into one
  whole-entry edit); chartsheet + sheet-scoped names under forced
  definedNames re-render skewed localSheetId (typed refusal); dataTable
  formulas joined the shift blockers (t="dataTable" is positional,
  unrewritable); move_range gained the guard set (bounds incl. Excel
  maxima, sheet protection, charts referencing either rectangle,
  array/dataTable values inside, defined names into src/dst, outside
  formulas referencing src OR dst); quoted 3-D spans ('A':'B'!) were
  invisible to the rename tokenizer (quoted spans now split
  unconditionally); the removal audit missed sheet-SCOPED defined
  names, CF rule formulas and DV formulas on surviving sheets (all
  probe via the rename machinery now) and false-refused when the only
  "reference" was a chart dying WITH the sheet (exclusive-closure
  exemption); the in_spill refusal advised editing the anchor —
  impossible under the same guard — and now states the honest option
  (reopen without preserve).
- **Accepted/noted:** crosscheck still verifies worksheet parts only —
  the standing tooling gap, carried to Batch 4; twin-bearing CF rule
  REORDERING (same blocks, new order) survives as verbatim originals in
  original order, a cosmetic divergence Excel ignores.

Suite: 3008 upstream+paper; 416 paper both env arms.

## Pinned-surface debt ledger

Debts are pinned surface not yet produced-and-tested; each names its
owning batch. Paying a debt REQUIRES removing its entry (the CI check
enforces both directions).

(no open debts — AmbiguousTargetError was paid by Batch 6's locate())

## Batch 0 — adversarial gate report (2026-07-08)

Four lenses over the batch diff (process amendment 2), all findings
fixed and fixtured before merge:

- **Pinning collateral (major):** the interim pin refused legitimate
  column grouping/ungrouping on every cols-bearing workbook, and a
  width-only edit false-dirtied sheetFormatPr AFTER arm (adding
  outlineLevelCol="0", dropping unmodeled attributes like
  x14ac:dyDescent). Root fix: regions._sheet_format computes the outline
  sync purely (mirrors holder.to_tree membership without reading its
  render-time side effect). Corpus-wide zero-pins asserted; grouping,
  width, and defaultRowHeight edits all splice; explicit outlineLevelCol
  edits follow stock semantics (normalized on cols-bearing sheets,
  verbatim otherwise — it is writer-derived metadata). The pattern guard
  itself is retained and proven with a synthetic impure serializer.
- **Cross-check row blindness (major):** verify_splice was blind to row
  attribute drift/duplication/deletion inside sheetData while the saver
  rewrites row attrs unclaimed. Extended with per-row signatures + row
  claims; unit-tested in both directions; active suite-wide.
- **Env-flip blast radius (major):** the suite could not run under
  PAPER_PRESERVE_DEFAULT=1 (17 stock-arm failures). Root conftest now
  normalizes the env (the suite asserts both arms explicitly; battery
  job 2 covers the flipped default via monkeypatch) — the suite is
  env-invariant, verified both ways. Legacy .xls/.xlsb loads under the
  env default now fall back to stock exceptions ("a default, never a
  mandate").
- **Battery evasions (major):** jobs 10/16 could be satisfied by a
  warn-then-drop implementation; both now arm simplefilter("error")
  like job 9.
- **Pinned-surface checker (minor):** comment mentions could satisfy
  both arms and force debt-entry deletion. Produced-arm now strips
  comments; tested-arm requires pytest.raises/warns or except-clauses.
- **Minors accepted with rationale:** the extLst region guard is a
  substring scan (over-refuses two exotic shapes; pre-existing v0 code,
  NARROWED by the end fix; fails safe); cross-check does not compare
  inter-tag order/root attrs (unreachable from a span-bounded splice;
  documented in the docstring); chart/rels/workbook byte-patch paths sit
  outside verify_splice (safety-tooling coverage gap, noted for the
  Batch-3/4 crosscheck extension); deleting ALL column dimensions leaves
  a stale outlineLevelCol (stock-divergent but Excel-tolerated; owned by
  the region attr-carry work); data_only+env-default flips warn-then-
  destroy into typed refusal (the intended direction); pandas
  if_sheet_exists='replace' under the env refuses loudly where stock
  destroyed silently (uncovered blast radius, noted here); a refused
  pandas append still lets pandas' close() rewrite zip local-header
  timestamps via a no-op save (part payloads byte-identical — the
  per-part invariant holds; raw whole-file bytes were never pinned).
- **Gate side-discovery, fixed in 0.2's addendum:** the self-closing
  <sheetData/> expansion path was a fourth instance of the end=None
  corruption class, repaired by the same fix and pinned with a test.

## Release Safety

The repository is private. The release workflow targets the `pypi` environment
and the publish step is additionally guarded by `vars.PUBLISH_ENABLED == 'true'`.
Create required reviewers on the `pypi` environment in GitHub before any
release. Publishing is intentionally disabled by default.

Do not push upstream release tags to origin. Only the `paper-base` marker tag is
pushed during bootstrap. Future `v*` release tags are pushed deliberately only
when publishing is intended.

## Upstream Merge Policy

Quarterly, clone or pull the official Mercurial upstream, convert the updated
history to Git in a fresh staging repository, identify the newest release tag,
merge that release into this repository, and run the full baseline suite.
Resolve conflicts using this ledger as the map. Merge, never rebase, after the
fork is published.
