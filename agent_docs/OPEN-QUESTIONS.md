# OPEN-QUESTIONS — paper-xlsx Phase 0 findings

**Status:** Phase 0 deliverable, 2026-07-07. The ten open questions from the Phase-0 charter,
each answered with an evidence-backed finding, a recommendation (phrased as the PR-0 decision it
feeds), and the spike that produced it. Q2 and Q4 were additionally put through adversarial
verification (both verdicts: **amended** — corrections are folded in below and called out).
All spikes live in `scratch/probes/` (gitignored, never merge); verbatim outputs in
`scratch/results/`. A cross-cutting gaps section (from an independent completeness critic) and a
register of flags against PINNED shapes close the document.

Baselines used throughout: `large.xlsx` = 3.39 MB / 600k cells / stock load 2.505 s / stock save
2.174 s; upstream suite 2592 passed, 6 skipped, 7 xfailed.

---

## Q1 — Retention hook: where does load discard the archive, and what does retention cost?

**Finding.** The archive opens at `openpyxl/reader/excel.py:95` (via `:123`) and its bytes die
at `excel.py:306-307` (`archive.close()` for non-read-only loads); the hook point is
`read_workbook()` (`excel.py:150-170`). Three retention mechanisms measured on `large.xlsx`
(fresh subprocess per variant, medians of 3):

| Variant | load cost | save-time access | steady memory |
|---|---|---|---|
| dict of decompressed payloads | 0.016 s | 0 | 23.8 MB (7.0× file size) |
| **whole-file bytes + ZipFile(BytesIO)** | **0.0007 s** | 0.016 s | **3.4 MB (= file size)** |
| keep ZipFile open | 0.0001 s | 0.016 s | ~0 (fd) |
| temp-file copy | 0.0013 s | 0.016 s | 0 RAM / 3.4 MB disk |

Keeping the handle or the path is disqualified by reproduced hazards: stock save truncates the
target in place (`writer/excel.py:291`, same-inode verified), so a retained open handle reads
`BadZipFile` on large files and **silently stale bytes** on small ones (8 KB buffer cache); a
retained path silently reads a different file after the user overwrites the source. `keep_vba`
already retains the entire archive in memory (`excel.py:162-165`) — the in-tree existence proof
that bytes-retention is safe, including save-over-same-path. Raw compressed-stream copy of all
parts: 0.0015 s (235× faster than recompression at 0.353 s), achievable in ~25 lines over
CPython-private `ZipFile` internals (payload identity verified); stdlib has no public API.

**Recommendation (PR-0).** Retain the original archive as one immutable `bytes` blob on a NEW
workbook attribute (never `wb.vba_archive` — flips mime_type, `workbook/workbook.py:360-370`;
never `_archive` — owned by read_only), wrapped in `ZipFile(BytesIO(blob))` for part access.
Hook in `read_workbook()` under `preserve=True`. Untouched-part copy = raw compressed-stream
copy with guard conditions (see gaps G8) and recompression as documented fallback. Preserve-mode
save goes temp-file + `os.replace` for path targets (in-place truncation is the hazard; replace
is handle-safe, measured). Amend §3.2 wording: "retained bytes ≈ file size" is true for
compressed retention; decompressed retention measures 7× on XML-heavy files.

**Spike:** `scratch/probes/q1-retention_{measure,extra,hazards,hazards2}.py`;
`scratch/results/q1-retention*.{md,json,txt}`.

---

## Q2 — Chokepoint completeness (adversarially verified: **amended**)

**Finding.** The public mutation surface decomposes into (a) ~35 clean funnels that method/setter
instrumentation covers (cell value/formula via `Cell._bind_value`; `__setitem__/__delitem__`;
`append`; merge/unmerge; the four style descriptors; NamedStyle registry ops; CF/DV adds;
defined-name dict ops; table add; sheet lifecycle incl. `title`/`active`/`move_sheet`/
`copy_worksheet`; freeze panes; print areas/titles; chart/image/pivot adds; comments;
hyperlinks; protection password setters; autofilter methods; breaks; custom doc properties), and
(b) **~14 by-reference bypass families that no method hook can see**: mutable satellites handed
out by property (`row_dimensions[n].height=`, `sheet_format`, `sheet_view(s)`, `page_setup`/
`page_margins`/`print_options`, `HeaderFooter`, `protection` fields, `wb.calculation.*`,
`wb.properties.*`, post-hoc mutation of held `CellRange`/`Rule`/`DataValidation`/`DefinedName`/
`Table` objects, `CellRichText` in-place edits, nested style-proxy leaks, bare attributes like
`ws.sheet_state`). Verification added five missed paths, all measured: **`ws.sheet_properties`
(sheetPr) family; `cell.data_type` direct set (silently demotes a formula to literal text —
cell-granular, invisible to element snapshots); `wb.loaded_theme`; chartsheet satellite set;
post-hoc image/chart `anchor` mutation → drawing part.** Also measured: pure READS materialize
state that changes stock output (`row_dimensions[5]`, `ws['Z99']`, `iter_rows`); sheet part
paths renumber positionally on delete/reorder; `DefinedNameDict.update()/|=/setdefault` bypass
its own type check; loading itself fires chokepoints (`create_sheet`, `_clean_merge_range`
style writes) and bypasses the value setter (`_reader.py:371-372`).

**Recommendation (PR-0).** Freeze a three-tier ledger architecture: **Tier 1** — instrument the
~35 funnels (plus `cell.data_type` converted to a property). **Tier 2** — for fully-modeled
by-reference satellites, snapshot each satellite XML element at preserve-load and re-serialize +
compare at save (legal because these elements ARE fully modeled — this is not the forbidden
lossy compare-save; catches every bare-attribute path for free); the snapshot list must include
`sheetPr` for both worksheet and chartsheet parts. **Tier 3** — named must-fix bypasses:
styles.xml semantic re-diff at save (StyleProxy leak), `CellRichText` always-dirty,
per-sheet hyperlink-set hash, `wb.loaded_theme` bytes-compare, and `mark_dirty()` as the
documented escape hatch for `wb.vba_archive`/`wb.rels`/`ws._rels`/`wb.shared_strings`.
Ledger keys on semantic mutation, never materialization; arms only after load completes.
`DefinedNameDict` gains the missing mutating-dunder overrides. Sheet part names are pinned under
preserve (no positional renumbering). Four families remain code-cited-only and must be either
measured pre-freeze or marked "unverified" in PR-0: post-hoc in-session chart mutation,
`add_pivot` part mapping, `RelationshipList.append` reachability, `wb.code_name`/`is_template`.

**Spike:** `scratch/probes/q2-chokepoints_*.py`, `scratch/probes/v1_*.py`;
`scratch/results/q2-chokepoints.md`, `scratch/results/v1-chokepoint-critic.md`.

---

## Q3 — The pandas path

**Finding.** `pd.ExcelWriter(path, engine="openpyxl", mode="a")` is chokepoint-pure: pandas
opens the file itself and hands openpyxl an **open `r+b` BufferedRandom, never a path**
(`pandas/io/excel/_base.py:1259-1283`), loads via `load_workbook(handle, **engine_kwargs)`
(`_openpyxl.py:76`), writes through `ws.cell()` + the public `Cell.value` setter + style/number-
format descriptors, uses `create_sheet`/`del book[name]`/`title=`/`merge_cells`/`freeze_panes=`/
`auto_filter.ref=`, and saves ONCE back into the same handle then truncates
(`_openpyxl.py:107-114`). A mechanical `._` grep confirms zero touches of openpyxl-private
state. `engine_kwargs` forwards verbatim, so **`preserve=True` reaches `load_workbook` the day
the kwarg exists** (verified: today it raises the expected TypeError from the forwarding line;
`data_only` control reached it). `pd.read_excel` defaults `read_only=True, data_only=True,
keep_links=False` and never writes (0 save calls; file bytes identical). Battery job 2 vs stock
on the synthetic trap fixture: no parts deleted; the injected sparkline extLst died, drawing XML
was regenerated with a spurious `<a:prstDash>`, and existing workbook rels were renumbered —
but the openpyxl-authored chart survived (see the stale-premise register). Bonus: `mode="a"` on
`.xlsm` **silently deletes `vbaProject.bin` by default** (`KEEP_VBA=False`).

**Recommendation (PR-0).** The pandas integration story is exactly one kwarg:
`engine_kwargs={"preserve": True}`. Consequences to pin: preserve must accept file-like objects
end-to-end (eager byte retention at load; handle-target save via full in-memory build then one
seek(0)/write/truncate — see gaps G4); the ledger instruments `AutoFilter.ref` and arms
post-load; `if_sheet_exists="replace"` maps to first-class sheet-remove/add ledger entries with
no rId renumbering; the Phase-2a lossy warning must fire on handle saves (pandas always uses
one). Battery job 2's stock baseline asserts sparkline death + rels renumbering + drawing
mutation on the synthetic fixture, and reserves "charts gone" for real-Excel fixtures.

**Spike:** `scratch/probes/q3-pandas_*.py`; `scratch/results/q3-pandas.md`.

---

## Q4 — Splice mechanics (adversarially verified: **amended**)

**Finding.** Byte-range splice beats parse-and-echo, decisively. Echo attempts, measured:
stdlib ElementTree **cannot re-serialize the default-namespace worksheet form at all**
(`ValueError` on unprefixed attributes with `default_namespace`); blind ET echo rewrites 237/237
lines and destroys x14/xm prefixes; et_xmlfile-as-echo renames prefixes and bloats declarations;
a hand-rolled sax echo got closest (13/237 lines) but **silently mutated cell text** (`&#13;` →
raw CR → LF on reparse) — a measured instance of the silent-wrongness class inside a would-be
safety mechanism; and in no-lxml installs openpyxl's own `register_namespace('s', ...)` poisons
stdlib output into `<s:worksheet>`. lxml echo is good (14/237 lines) but lxml is optional. The
byte splice — a namespace-tracking scanner over the original bytes computing element spans, raw
copy outside spans, replacement elements serialized via the existing cell-writer machinery —
produced a **1-region, +1-byte diff**, loads correctly in both openpyxl and LibreOffice, handles
t-type changes both ways, and survived attr-order/self-closing/quoting/prefixed-decoy drills.
Verification then **broke the original guard wording** with two legal, loadable decoys: a
`<c>` inside another cell's cell-level `extLst`, and one inside `mc:AlternateContent` before
sheetData — both spliced silently wrong under ancestor-containment matching. It also proved the
prefixed-root failure is **silent value deletion** (both loaders open the file; B8 reads None —
the reader parses row children tag-blind, `worksheet/_reader.py:303`), and that a replacement
built from `r/s/t` alone drops legal extra cell attributes (`ph`, `cm`, `vm`). ECMA-376 makes
row/cell `@r` optional; openpyxl and LO both implement implicit counters and never emit r-less.
Performance: pure-Python scan+splice 1.56 s on the 23.7 MB sheet (vs 2.17 s whole stock save).

**Recommendation (PR-0, incorporating the verifier's corrections verbatim).** v0 mechanism =
byte-range splice; echo is rejected for passthrough and **must never be a silent fallback**.
Mandatory guards, all detectable mid-scan before any write (refusal atomicity free):
(1) target `<c>` matched only via the EXACT parent chain root→sheetData→row→c at depth 3 —
never ancestor containment; (2) refuse when the root/main namespace is prefixed or the in-scope
default ns ≠ spreadsheetml at the target (unguarded failure mode is silent value deletion);
(3) the replacement cell carries over verbatim every target attribute not intentionally
rewritten (`cm`/`vm`/`ph`/foreign-ns) and refuses on unexpected target-cell children (cell-level
extLst); (4) r-less rows/cells → refuse all operations in v0 (implicit counters are a
demonstrated-feasible later upgrade); (5) DOCTYPE / non-UTF-8 → refuse; (6) `TargetNotFoundError`
when the scan finds no target. Inserts/deletes ride the same scanner (sorted-merge insertion
offsets); `<dimension>`/`spans` staleness is tolerated by all testable readers — real-Excel
fixture requested. PLAN §C's "echo events verbatim" phrasing should be restated as span-splice
(the only literally-verbatim reading; no semantic change to the §3.4 invariant).

**Spike:** `scratch/probes/q4-splice_*.py`, `scratch/probes/v2_*.py`;
`scratch/results/q4-splice.md`, `scratch/results/v2-splice-skeptic.md`, `scratch/results/q4/`,
`scratch/results/v2/`.

---

## Q5 — Sheet-internal non-cell regions: the v0 line

**Finding.** Worksheet child order is procedural (no `__elements__`; see ARCHITECTURE-NOTES §3).
Round-tripping the trap fixture is canonically identical on every modeled element except extLst
(warned + dropped). A quirk fixture with Excel-producer patterns loses, additionally and
silently: `xr:uid` attributes everywhere, row `spans`/`x14ac:dyDescent`, `pageSetup r:id` (+ the
printerSettings part and rel), the cfRule→x14 twin pointer (`<extLst><ext><x14:id>`), and
protectedRanges/phoneticPr/cellWatches/ignoredErrors/AlternateContent/picture wholesale. Two
reader bugs found: a colBreaks living only in `mc:Fallback` is **resurrected as a real
top-level element** (depth-blind dispatch, `_reader.py:156-169`), and a schema-legal
`<col phonetic="1"/>` **crashes load** (TypeError — `ColumnDimension.__init__` takes no `**kw`).
A 16-edit blast-radius probe maps each mutation API to exactly one element — except hyperlink
add, which renumbered the sheet's drawing rel (rels rebuild, `writer/excel.py:202`): fatal under
splice unless sheet rels become append-only. x14 twins (CF and DV) are GUID/sqref-linked to
classic elements; editing or deleting the classic half alone orphans or double-applies the twin.

**Recommendation (PR-0).** Pin the CT_Worksheet child-sequence constant for element placement.
**Tier 1 (splice the whole element from the model):** mergeCells, dataValidations (classic,
gated), hyperlinks (with append-only rels), autoFilter (refuse if it carries extLst),
sheetProtection, printOptions, pageMargins, rowBreaks/colBreaks, sheetViews (refuse on
sheetViews-level extLst), cols, sheetFormatPr (with unknown-attr carry-over), sheetPr,
tableParts. **Tier 2 (re-serialize with loud warning):** classic conditionalFormatting only
when the touched block has no cfRule extLst AND intersects no x14 twin sqref; pageSetup with
r:id carried over; headerFooter with empty-children emission suppressed. **Tier 3 (refuse,
`UnsupportedStructureError`):** any CF/DV edit intersecting an x14 twin; all edits to unmodeled
regions (protectedRanges, phoneticPr, customSheetViews, cellWatches, ignoredErrors, smartTags,
picture, oleObjects, controls, AlternateContent, anything in extLst); drawing/legacyDrawing are
passthrough-only. Build a read-only extLst perception pass (URI inventory + xm:sqref/x14:id
extraction) to power the gates and the manifest confession block; suppress/rephrase the stock
"will be removed" warning under preserve (it becomes false). The ledger cross-check differ must
be depth-aware (the mc:Fallback bug class) and must whitelist load-time merge materialization.
Schedule two fixture-backed reader patches: `<col phonetic="1"/>` crash; cfRule TypeError
discard. (Reconciliation of this tiering with Q2's — one matrix, not two taxonomies — is gaps
G5, a PR-0 obligation.)

**Spike:** `scratch/probes/q5regions_*.py`; `scratch/results/q5-regions.md`, `q5_quirk.xlsx`,
blast/diff transcripts under `scratch/results/`.

---

## Q6 — Mixed fresh-on-preserved (new charts/images on loaded rich files)

**Finding.** All drawing/chart part names and rIds are minted at save from per-save session
counters that restart at 1 (`writer/excel.py:132-139`) — guaranteed collisions with preserved
parts; rId allocation `rId{len+1}` collides with non-contiguous preserved ids (measured).
`[Content_Types].xml` and all rels are regenerated wholesale. One drawing per sheet is a hard
invariant (writer emits at most one `<drawing>`; **LibreOffice silently drops a second**,
measured). Anchor-merge INTO an existing preserved drawing is empirically tractable — a 4-edit
splice (new chart part; anchor appended before `</wsDr>`; drawing-rels entry; content-type
override) round-tripped through openpyxl AND LibreOffice, including against an `xdr:`-prefixed
LO-authored drawing — but requires drawing-internal rId/cNvPr bookkeeping. Stock machinery
already handles mixed loaded+fresh sheets correctly. Upstream bug: second save of an
image-bearing workbook crashes (`Image._data()` closes its BytesIO).

**Recommendation (PR-0).** (1) Fresh chart/image onto a preserved sheet WITHOUT an existing
drawing: **support in v0** via a pinned merged-edit set — new parts named
`1 + max existing number per family` (never session counters), one appended sheet-rels
relationship with `rId = max-numeric + 1`, spliced `<drawing r:id>` element at its CT_Worksheet
slot, targeted content-type appends. (2) Onto a sheet WITH a preserved drawing: **v0 refuses**
(`UnsupportedStructureError`, message naming the drawing part and the options); anchor-merge is
the documented v0.5 lift. (3) Charts on new in-session sheets under preserve: support (verified
working). (4) Preserve-load still populates `ws._charts`/`ws._images` for perception but never
re-serializes preserved drawings and never routes preserved images through `Image._data()`.
(5) Phase-2a warning wording: plain charts/images are "rebuilt lossily", not "deleted".

**Spike:** `scratch/probes/q6-mixed_*.py`; `scratch/results/q6-mixed.md` + surgery outputs.

---

## Q7 — Inline strings vs sharedStrings append

**Finding.** Moot in the expected direction: **stock 3.1.5 never writes sharedStrings.xml** —
every string is `t="inlineStr"` (`cell/_writer.py:21-22`); there is no sst writer in the tree;
resaving an sst-bearing file strips part + content-type override + workbook rel. Measurements on
an 11k-string-edit workload (10k new + 1k changed): inline vs shared-append file sizes are a
wash (+2.73% vs +2.72% on the sst-less base; inline WINS +2.14% vs +2.74% on an LO-normalized
`t="s"` base); even a pathological all-repeats bulk load saves only 24 KB with sharing. Both
variants read back correctly in openpyxl, pandas, and LibreOffice, including mixed
`t="s"`+`inlineStr` sheets. Stale sst count/uniqueCount attributes are ignored by all three
readers (openpyxl never reads them). The coupling risk is asymmetric and measured: creating an
sst part on an sst-less file requires BOTH registration edits — omit the content-type override
and **openpyxl fails hard** (IndexError); omit the workbook rel and **LibreOffice silently
drops every `t="s"` cell**. Excel itself is untestable here (see FIXTURE-REQUESTS.md); expected
behavior (expectation, not evidence): inlineStr is schema-legal, normalized to sst on next
manual save.

**Recommendation (PR-0).** **Inline strings for ALL operation classes in v0** — single-cell,
scattered, bulk, and the pandas new-sheet case. `xl/sharedStrings.xml`, where present, is never
modified: raw-copied byte-identical, indices never renumbered; untouched `t="s"` cells keep
resolving. Collateral-set amendment: sharedStrings drops OUT of the sanctioned single-cell set.
Append-only sst remains the sanctioned post-v0 fallback, gated on a real-Excel fixture
demonstrating need; if ever built, part creation must add both registration edits. Kernel note:
`xml_equivalent`/`diff_package` must compare cell values, not string-storage encoding, across
Excel round-trips.

**Spike:** `scratch/probes/q7-strings_0{1..4}*.py`; `scratch/results/q7-strings.md`,
`scratch/results/q7_strings/` (all variant packages).

---

## Q8 — read_only / write_only vs preserve

**Finding.** read_only parks the open ZipFile as `wb._archive` (`excel.py:167-168`), lazy for
cells only (sst and stylesheet parse eagerly), save already refuses (`TypeError("Workbook is
read-only")`, `workbook.py:382-383`). write_only is constructor-only (`Workbook(write_only=True)`);
`load_workbook` has no such kwarg, so preserve+write_only is **unreachable by construction**
(probe-verified both directions). Mode-exclusive modules (`worksheet/_read_only.py`,
`worksheet/_write_only.py`, `cell/read_only.py`) need zero edits; the only shared seams are
`reader/excel.py` (kwarg + guard + retention hook, inert for other modes), one dispatch guard at
the top of `save_workbook` (`writer/excel.py:279`), and call-only reuse of `write_cell`. Flag
plumbing for preserve copies the existing pattern exactly (table in
`scratch/results/q8-modes.md`). preserve+data_only stays loadable (retained bytes keep all
formulas; the Phase-3 refusal is a save-time two-liner reading `wb.data_only`). Cross-charter
flag: `Workbook.mime_type` keys off `vba_archive` truthiness — the preserve save must take
content types from the retained `[Content_Types].xml`, never `wb.mime_type`.

**Recommendation (PR-0).** Signature: keyword-only `preserve=False` on `load_workbook`
(`reader/excel.py:316`) and `ExcelReader.__init__` (`:121`), after a bare `*`; `openpyxl.open`
is covered for free (`openpyxl/__init__.py:8`). preserve+read_only raises **`ValueError`**
(programmer error per CONVENTIONS §2 — correctly-typed flags in an invalid combination; in-repo
precedent `copy_worksheet`'s ValueError; do NOT extend the `save` TypeError wart), raised at the
top of `ExcelReader.__init__` BEFORE `_validate_archive` so no handle is opened on refusal.
No guard needed for write_only; no load-time refusal for preserve+data_only/keep_vba/keep_links/
rich_text. Workbook gains `_preserve = False` class attr + read-only `preserve` property beside
`_read_only`. Retained bytes live under a new attribute, never `_archive`.

**Spike:** `scratch/probes/q8-modes_{readonly,writeonly}.py`; `scratch/results/q8-modes.md`.

---

## Q9 — Repo provenance

**Finding (verified inline, not delegated).** History is complete: 9,142 commits back to the
2010-04-09 upstream initial import, 123 tags (1.0 → 3.1.5), full hg-fast-export conversion of
the Heptapod Mercurial repository with all upstream dev branches present. `paper-base` ==
upstream tag `3.1.5` == `c4986390b`; HEAD is exactly one commit ahead (`d2b5d62e9` "Bootstrap
paper-xlsx fork identity": `__paper_version__ = "0.1.0"`, packaging rename to `paper-xlsx`,
CI + guarded release workflows, PAPER.md, README). PAPER.md's recorded baseline was
**re-verified on this machine today**: 2592 passed, 6 skipped, 7 xfailed (2.77 s, Python 3.13.3,
lxml 5.4.0). PyPI's latest openpyxl is still **3.1.5 (2024-06-28)** as of 2026-07-07 — the fork
base is the current upstream stable. The built wheel ships `openpyxl/` as the import package
(name rule holds in packaging). The Mercurial-sync policy is already recorded in PAPER.md
(quarterly staging-conversion + merge of the newest release tag; never automated history
merges). Hygiene note: a leftover `soxhub` git remote points at `/tmp/soxhub-openpyxl` (the
conversion staging clone) — candidate for removal, left untouched pending owner decision.

**Recommendation.** No action beyond the PAPER.md addendum recording today's re-verification.
Do not delete the `soxhub` remote without the owner's say-so.

**Evidence:** `scratch/results/baseline_pytest.txt`; git/tag/PyPI checks in the session log.

---

## Q10 — Oracle viability HERE

**Finding.** Viable and fast. `/opt/homebrew/bin/soffice` → LibreOffice 26.2.4.2 (Homebrew cask
wrapper → `/Applications/LibreOffice.app`). The vendored recalc script's temp-copy discipline
holds (audited line-by-line; the input path is written only by the final `cp` after success
checks). Live run on `schedule.xlsx`: empty `<v></v>` cached values → `6500`/`6825`/cross-sheet
`6500`; cold 2.80 s (profile creation ~1.6 s), warm 1.19 s; `large.xlsx` (600k cells) warm
2.09 s. Failure modes, all measured: **shared-profile concurrency is nondeterministic** (one
trial: hard `DeploymentException` abort; other trials: silent IPC delegation with exit 0 and
empty stdout) — distinct `-env:UserInstallation` profiles run cleanly in parallel; **soffice
exits 0 on corrupt/unloadable input** (only the output-exists check catches it); the reference
script collapses its intended timeout exit 124 into exit 1 (bash wrapper bug); successful runs
emit stderr noise (`Task policy set failed`); `=1/0` converts successfully with the error cached
as `<c t="e"><v>#DIV/0!</v></c>`; the profile `.lock` file persists after clean exit. CI
(ubuntu-latest) needs `apt-get install --no-install-recommends libreoffice-calc` (pulls core +
common, which ship both `/usr/bin/soffice` and `/usr/bin/libreoffice`); no Java/X needed.
`pytest.ini` has `--strict-markers`, so `lo_smoke` must be registered before first use. macOS
DMG-only installs put nothing on PATH — detection needs a darwin fallback.

**Recommendation (Phase-5 driver spec, PR-0 records it).** Per-invocation unique
`-env:UserInstallation=file://<fresh tempdir>` is mandatory. Success predicate = returncode 0
AND output file exists; never parse stderr. `OracleTimeoutError` from the subprocess timeout
directly with `start_new_session=True` + process-group kill. Detection: `which("soffice")` →
`which("libreoffice")` → `/Applications/LibreOffice.app/Contents/MacOS/soffice` → else
`OracleUnavailableError`. Error scan keys on `t="e"` cells. Treat "cached value absent" as
including EMPTY `<v></v>` (stock openpyxl's shape). CI: LibreOffice on ONE matrix leg with
`PAPER_REQUIRE_LO=1` turning skips into failures there; register `lo_smoke` in pytest.ini;
`soffice --version` print step. LO output is an answer key only — never bytes to splice.

**Spike:** `scratch/probes/q10-oracle_{a_recalc_timing,b_concurrency,b_failures}.py`;
`scratch/results/q10-oracle.md`.

---

## Q11 — Damage-model reproduction (the carnage baseline preview)

| PLAN damage row | Reproduced? | Loss at | Loud/silent |
|---|---|---|---|
| Charts/images/drawings deleted | **Partially — coverage-gated.** Parseable chart+image survived; shapes, drawing `mc:AlternateContent`, chart extLst died 100% | LOAD (lossy projection) + SAVE (regeneration) | silent |
| VBA stripped | Yes — part, CT override, rel all gone without `keep_vba`; byte-identical with it | SAVE | silent |
| data_only trap | Yes, verbatim — `<f>` count 3→0; only literals remain | LOAD discards, SAVE writes | silent |
| Stale cached values | Yes — openpyxl writes `<f>` + **empty `<v></v>`**; pandas reads NaN; recalced twin reads 6500 | SAVE (never calculates) | silent to pipelines |
| insert_rows corruption | Yes, quantified — SUM range unchanged, defined name & cross-sheet ref point at moved cells; **LibreOffice computes 1100 / 6399 / 5400 where correct answers are 7499 / 6500** | mutation (`_move_cells`) | **silent; plausible numbers** |
| Extensions dropped | Yes — sparklines 2→0 | LOAD (parse-to-warn) | **warned** at load |

Fairness inventory (all SURVIVE stock round-trip on the synthetic fixture): merges, classic CF,
DV, table, comments, hyperlinks, freeze panes, hidden row/sheet, named styles, defined names,
and openpyxl-authored charts/images. Stock is genuinely good at everything it fully models; the
carnage is exactly the unmodeled/half-modeled set — which is what real-Excel files are full of.
Also measured: a data_only round-trip renumbered style indices with zero style edits;
`keep_vba` preserves ctrlProps/activeX/customUI/VML/.emf beyond VBA itself.

**Recommendation.** Battery baselines assert what actually dies per fixture class; "charts gone"
claims move to real-Excel/shape-bearing fixtures (FIXTURE-REQUESTS.md is load-bearing). The
Phase-2a lossy-save warning must be **content-level** (part-list comparison under-reports: zero
parts removed while sparklines were gutted) and must account for the load-time-loss timing.
The insert_rows numbers (1100/6399/5400 vs 7499/6500) are the Phase-6a refusal justification
artifact. `wb.calculation.fullCalcOnLoad = True` empirically makes LibreOffice recompute on
load — Phase 3's flag mechanism is validated.

**Spike:** `scratch/probes/q11-damage_p{1,1b,2,3,4,5}*.py`; `scratch/results/q11-damage*.txt`,
`scratch/results/q11-damage.md`.

---

## Cross-cutting gaps (independent completeness critic; severity as assigned)

- **G1 (blocking-for-PR-0) — Shared/array/dataTable formula groups under splice.** Real-Excel
  files write filled ranges as `<f t="shared" si ref>` (host + followers); openpyxl expands them
  at load via Translator (`worksheet/_reader.py:249-272`) and cells retain NO group membership.
  Under preserve, the original bytes keep the group while the model forgets it: splicing a
  shared-formula HOST orphans every follower — silent corruption in exactly the class the fork
  exists to kill. Same for `<f t="array" ref>` spill ranges and dynamic-array `cm`/`vm`
  metadata (which the Q4 carry-verbatim rule would carry STALE onto a changed value).
  **Resolution:** preserve-load records per-sheet si-group and array-ref maps; splice refuses
  (v0) or multi-cell-splices the whole group when a dirty cell intersects one; drop/refuse
  `cm`/`vm` on value-changed cells; real-Excel shared-formula fixture requested.
- **G2 (blocking) — Collateral-set amendment must be reconciled in one place:** calcChain
  deletion cascades to its `[Content_Types].xml` override AND workbook rel (both proven
  load-bearing registration edits); styles.xml joins for style-bearing edits; sharedStrings
  drops out under inline-only. PR-0 publishes per-operation-class collateral sets.
- **G3 (blocking) — `properties.modified` auto-stamp** (`writer/excel.py:292`) would fail the
  pinned no-op payload-identity invariant on every preserve save. PR-0 decides: raw-copy
  core.xml unless `wb.properties` is dirty (sanctioned deviation logged), or stamp via the
  injectable clock and add core.xml to the collateral set.
- **G4 (blocking) — Save-target contract:** CONVENTIONS §7 pins temp-file+atomic-rename;
  pandas always saves into an open r+b handle where rename is impossible. PR-0 pins dual-mode
  semantics: path → temp+rename; handle → build fully in memory, then one seek(0)/write/truncate;
  §7 wording amended; pandas handle-dance fixture test added.
- **G5 (note) — Merge Q2's chokepoint tiers and Q5's region tiers into ONE per-region matrix**
  (dirt-detection mechanism × write mechanism × refusal gate), folding in the Q2-verifier
  additions. E.g. conditional formatting is Q2-Tier-2 snapshot but Q5-Tier-3 refusal-gated —
  both evidence-backed, currently unreconciled.
- **G6 (note) — Performance budget not yet composed:** scan+splice alone is 1.56 s on the 23.7 MB
  sheet vs a 2.17 s stock-save denominator; the end-to-end preserve save was never measured as
  one number. Run one composed prototype before pinning the multiple; otherwise pin 2× with
  expat-byte-offset/lxml acceleration as contingency.
- **G7 (note) — `definedName localSheetId` is a positional index** into the sheets list: sheet
  delete/reorder under preserve (incl. pandas `if_sheet_exists="replace"`) silently invalidates
  sheet-scoped names in preserved workbook.xml bytes; the dependent-part cascade for sheet
  delete is likewise undefined. v0 floor: refuse delete/reorder on workbooks with sheet-scoped
  defined names, or spec the remap + cascade.
- **G8 (note) — Raw compressed-stream copy guards:** the recipe touches CPython-private ZipFile
  internals, verified on 3.13 only, and does not handle data-descriptor entries (GP flag bit 3)
  or Zip64 local headers. Pin guard conditions (flag bit 3, zip64, non-DEFLATE, private-attr
  probe → fall back to recompression) + a fixture for each.
- **G9 (note) — Four chokepoint families are code-cited but unmeasured** (post-hoc chart object
  mutation, `add_pivot` mapping, `RelationshipList.append` reachability,
  `wb.code_name`/`is_template`): measure before the freeze or mark honestly in PR-0.
- **G10 (phase-later) — Part-name resolution must be rels-driven**, never
  pattern-matched (`xl/worksheets/sheetN.xml` assumptions break on non-canonical producers).
- **G11 (phase-later) — 1904 date system** has no fixture and no finding; add a date1904
  workbook to the taxonomy plus one splice + certify test.

---

## Flags against PINNED shapes (human decision required; per the tier rule, not locally patched)

1. **CONVENTIONS §3.5 sanctioned collateral set** — three measured amendments needed (G2):
   styles.xml for style-bearing edits; calcChain deletion cascade (part + CT override + rel);
   sharedStrings removed under the inline-only decision. Also G3's core.xml decision.
2. **CONVENTIONS §3.2 wording** — "retained bytes ≈ file size" holds for compressed retention
   only (decompressed measures 7×); pin the representation (whole-file bytes blob).
3. **CONVENTIONS §7 / PLAN atomic-rename** — impossible for handle targets (pandas); dual-mode
   save semantics needed (G4).
4. **PLAN §C "echo events verbatim"** — literally unachievable (no parser reports empty-element
   form); restate the mechanism as byte-span splice. The §3.4 invariant itself is unchanged.
5. **PLAN damage-model table row 1 and battery jobs 1–2 wording** — "charts gone" is
   coverage-gated in 3.1.5; claims move to real-Excel/shape-bearing fixtures; the 2a warning
   must be content-level and load-time-aware.
6. **CONVENTIONS §4 fixture-corpus expectation vs reality** — the five-job battery's stock
   carnage baseline partially depends on fixtures this environment cannot author (real Excel);
   FIXTURE-REQUESTS.md created as the load-bearing companion.
