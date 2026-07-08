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
