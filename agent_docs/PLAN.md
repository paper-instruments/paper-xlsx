# paper-xlsx — v0 Implementation Plan

## What this repository is and why you're here

This is Paper Instruments's hard fork of **openpyxl**, the standard Python library for Excel
files — roughly three hundred million downloads a month, and the engine pandas uses under
`read_excel`/`ExcelWriter`. The import name `openpyxl` is **frozen forever** (pip name:
`paper-xlsx`); this fork must remain a drop-in replacement for every script, pipeline, and model
prior that says `import openpyxl` — including pandas itself, which will route through our save
path without knowing anything changed.

Why fork: openpyxl's problem is the inverse of most libraries'. Its API surface is rich — cells,
formulas, styles, tables, named ranges, conditional formatting, chart creation — but its
**persistence core is destructive**. On save it regenerates the entire file from its in-memory
model, so everything it does not model is deleted: charts, images, drawings, VBA behavior
without a flag, pivot machinery in part, sparklines, modern extensions. Its own documentation
admits it ("shapes will be lost from existing files if they are opened and saved"). Three more
behavioral traps compound it: it **never calculates** (written formulas carry no cached values,
so downstream readers see stale or empty results); `data_only=True` + save **permanently
replaces every formula with its last cached value**; and `insert_rows`/`delete_rows` move cells
while updating **nothing** — not formulas, not defined names, not chart ranges — so one inserted
row silently corrupts every SUM below it. All of this is silent. The ecosystem's coping
mechanisms prove the pain is real: the production LLM xlsx skill in our environment makes a
LibreOffice recalculation script **mandatory** after any formula write, carries a verbatim prose
warning about the `data_only` self-destruct, demands "EXACTLY match existing format" (which
stock openpyxl cannot honor on rich files) — and demonstrates `insert_rows` with no warning at
all, because the scariest landmine isn't even fenced.

The mission: keep everything that is excellent — the reader, the object model, the formula
tokenizer, fifteen years of absorbed producer quirks — and **transplant the spine**: under
preserve mode, the original package becomes the source of truth and the object model becomes a
source of edits to it. Untouched content survives byte-identical *by construction*, not by
coverage. On top of the spine: honesty guards, perception (manifest + semantic diff), the
LibreOffice oracle (recalc + error scan + certification), and reference-aware structural edits.
The playbook follows a house pattern proven on prior Office-library forks; the difference here
is that **no battle-tested reference helper folder exists** — this plan carries the algorithms
itself. Where prior plans said "mine the reference material," this one specifies mechanisms
inline. Read them as binding design intent; ground exact signatures in PR-0.

**Read first, in order:** `CONVENTIONS.md` (governing; it wins over anything here) → this plan →
the environment's xlsx skill (SKILL.md plus its `scripts/recalc.py` and `scripts/office/`
LibreOffice driver — read as user stories AND as mining material for Phase 5: it already solves
headless-LibreOffice setup, sandboxed-socket workarounds, timeout handling, and error-scan JSON)
→ an upstream source tour (Phase 0 tells you where to look) → upstream's documented warnings
(tutorial "Warning" blocks; the `data_only` note).

**End state.** This package replaces stock openpyxl in our agent
environments — and, because the import name is frozen, under pandas —
with zero code changes anywhere. Everything that worked before still
works. What changes is custody: an agent, or any program, can be handed
a real human-authored workbook (charts, pivots, macros, years of
formatting), make a surgical change, and hand back the same file with
only that change in it — provably, via the package diff — or receive a
typed refusal explaining why not. Every operation has exactly three
legal outcomes: done correctly, refused loudly, or done with a warning
enumerating what could not be preserved. "Silently wrong," stock
openpyxl's signature failure, ceases to exist as an outcome. The
customer for this guarantee is agent-driven knowledge work on
client-supplied financial models, where today's choice is "only touch
spreadsheets we generated ourselves" or "risk returning a damaged file."

---

## The damage model (why each organ exists)

| Failure | Trigger | Loud or silent | Organ that kills it |
|---|---|---|---|
| Charts/images/drawings deleted | load+save any rich file | silent | Spine (Phase 2) |
| VBA stripped | .xlsm load+save without flag | silent | Spine raw-copies the part (Phase 2) |
| All formulas → dead values | `data_only=True` then save | silent | Guard (Phase 3) |
| Stale cached values | any formula write | silent to pipelines | Recalc-on-load flag (Phase 3) + Oracle (Phase 5) |
| References corrupted | insert/delete rows/cols | **silent; numbers look plausible** | Refuse (6a), rewrite (6b) |
| Modern extensions dropped | load+save | sometimes a warning | Spine passthrough (Phase 2) |
| Wrong numbers trusted in-pipeline | any recalc consumer | silent | Certification (Phase 5) |

Two orientation facts: the damage is **deterministic and content-gated**, not flaky — a
chart-bearing file loses its charts 100% of the time under stock save; and the loud failure
(charts) is the least dangerous because someone notices, while the silent ones (wrong sums,
stale values) are the ones that reach a board deck.

---

## Architecture: the spine transplant (binding design)

Read CONVENTIONS §3 first; this section adds mechanism detail.

**A. Byte retention (preserve mode).** `load_workbook(path, preserve=True)` retains every
archive part's payload bytes in memory (≈ file size; cheap). Parts the reader never parses —
drawings, chart XML, chart embedded workbooks, `vbaProject.bin`, pivot caches, media, custom
XML — are *never parsed here either*; they exist only as retained bytes and are raw-copied at
save. This single mechanism kills the headline losses, with one catch that motivates mechanism C.

**B. The dirty ledger.** Instrument every public mutation chokepoint to record
(part, address/region): the cell value and formula setters, style assignment, number formats,
row/column dimension objects, merged-cell operations, sheet add/remove/rename/reorder, defined
names, tables, conditional formatting, data validation, workbook properties. Phase 0's job is to
produce the *exhaustive inventory* of these chokepoints from the real code; PR-0 freezes it.
Provide `mark_dirty(...)` as the documented escape hatch. Debug mode cross-checks the ledger
against a region-level semantic diff on every harness run — a missed chokepoint is corruption
inside the safety tooling and blocks release. The ledger is load-bearing: a compare-based
patch-save is impossible here because openpyxl cannot serialize a faithful candidate to compare
— serialization *is* the lossy act.

**C. The splice writer.** The catch in A: worksheet parts are always parsed, and stock
re-serialization of a sheet drops what the model didn't load — including the `<drawing r:id>`
element that attaches the chart, and sparkline/x14 `extLst` blocks. So touched sheets are not
re-serialized; they are **spliced**: stream-parse the *original* sheet part, echo events
verbatim, and merge in the dirty set by coordinate order — replace `<c>` elements whose address
is dirty, insert new `<c>`/`<row>` elements at their sorted position, delete where the ledger
says deleted. Two sorted streams, O(sheet), O(1) memory. Unmodeled sheet content passes through
untouched because it is never interpreted. For *untouched* sheets, skip even this: raw copy.

**D. Cross-part strategies (pinned in CONVENTIONS §3.5; summary).** New/changed string cells:
inline strings or append-only sharedStrings — PR-0 picks per operation class; never renumber
existing indices. New styles: append-only xf. `calcChain.xml`: delete on any formula change;
Excel rebuilds it. Sanctioned collateral set for a single-cell edit: that sheet part,
sharedStrings (append, if used), calcChain (deleted), workbook calcPr. The changed-part-budget
test enforces this list literally.

**E. What the spine does NOT change.** Fresh-generation (`Workbook()` → save) never enters
preserve mode and is untouched. `read_only`/`write_only` streaming modes are orthogonal and
untouched. Stock load+save keeps stock behavior plus the new loud warning (CONVENTIONS §1.1).

---

## Phase 0 — Orientation (no code)

Produce `ARCHITECTURE-NOTES.md` answering, with file/class references from the actual tree
(locate by grepping class names, never by remembered paths):

1. Where the archive is opened at load, and where its bytes are discarded — the retention hook
   point.
2. The full path of a cell write: which setters exist on `Cell`, where styles attach, where
   row/column dimensions live — the ledger chokepoint inventory, exhaustively.
3. How the worksheet writer emits sheet XML today (the module that streams `<row>`/`<c>`), and
   which event-writing machinery it uses — the splice writer will reuse it.
4. How sharedStrings is represented in memory and written; how style indices are assigned.
5. Where `calcPr` lives and how workbook.xml is written; where calcChain is handled.
6. How charts/images added *in-session* are tracked and written — the spine must not break
   fresh-chart creation, and mixed cases (loaded rich file + newly added chart) need a defined
   answer in PR-0.
7. The `Serialisable` descriptor framework: read `openpyxl/descriptors/`, summarize the pattern
   in ten lines — it is the house style for any new XML vocabulary.
8. The formula tooling: the tokenizer and `Translator` — Phase 6's building blocks.
9. Run the upstream pytest suite; record baseline results in `PAPER.md` (rule one of forking:
   never operate on a patient whose baseline vitals you didn't record).

## Phase 1 — Test infrastructure (first-class; nothing merges before it)

Implement CONVENTIONS §4: fixture corpus + sidecars + `MANIFEST.sha256`, the five-assertion
contract harness, `lo_smoke` marker, injectable clock, `FIXTURE-REQUESTS.md` for real-Excel
fixtures (real client-model-shaped workbooks are the load-bearing bucket — request them
explicitly and precisely).

**Freeze the five brownfield jobs as the standing acceptance battery**, and run them against
STOCK openpyxl first, recording the carnage as the baseline that both proves the thesis and
regression-guards the fix:
1. **Assumption flip:** chart-bearing model, change one input cell, save. (Stock: charts gone,
   values stale.)
2. **pandas append:** `pd.ExcelWriter(path, engine="openpyxl", mode="a")` a new sheet onto a
   charted report. (Stock: charts gone — and note this exercises our save path *through
   pandas*, the superset payoff test.)
3. **Schedule row:** insert a row above a SUM-bearing schedule. (Stock: silent reference
   corruption.)
4. **xlsm round-trip:** load+save a macro workbook without flags. (Stock: macros gone.)
5. **data_only trap:** load `data_only=True`, read, save a copy. (Stock: formulas destroyed.)

Pass criterion forever after: each job ends **correct or loudly refused — never silently
wrong**. Job 3 is expected to pass via refusal after Phase 6a and via correctness after 6b.

## Phase 1.5 — PR-0: API Proposal (CONVENTIONS §8)

Signatures, refusal conditions, and examples for every organ below, grounded in Phase 0's
findings; plus the delegated decisions: chokepoint inventory, inline-string vs. sharedStrings
policy, mixed fresh-chart-on-rich-file semantics, the performance budget number. Human approval
gates implementation.

## Phase 2 — The spine (the long pole; land in sub-stages, each shippable)

- **2a — Retention + raw copy + honesty warning.** `preserve=True` retains payloads; save
  raw-copies every never-parsed part; and *both* save paths gain the lossy-save warning that
  enumerates unpreserved content by inspecting the archive's part list against the modeled set
  ("3 charts, 1 pivot cache, VBA present"). 2a alone upgrades the honesty story shippably.
- **2b — The dirty ledger** per the chokepoint inventory, with the debug-mode cross-check wired
  into the harness from day one.
- **2c — The splice writer** for touched sheets, including the splice-completeness trap test
  (one-cell edit on a sheet carrying sparklines, x14 CF, and a drawing reference — everything
  survives).
- **2d — Cross-part handling** per the pinned strategies; the sanctioned-collateral budget test
  goes green here.
- **Acceptance:** battery jobs 1, 2, and 4 pass; no-op round trip part-payload-identical;
  performance guardrail met; upstream suite green throughout.

## Phase 3 — Honesty organs

`data_only`+save: refuses under preserve mode (typed `UnsupportedStructureError` naming the
trap and the `allow_formula_loss=True` override), warns loudly on the stock path. Any
formula-affecting edit auto-sets the workbook's full-recalc-on-load flag so a human opener's
Excel always computes fresh numbers — stale cached values can never masquerade as current to a
person. Typed load-time refusals for `.xlsb` and legacy `.xls` naming the format and suggesting
LibreOffice conversion. The lossy-save warning from 2a graduates to a structured, testable
message.

## Phase 4 — Perception

`workbook.manifest()`: sheets with dimensions, tables, defined names, formula counts,
volatile-function detection, external links, and a **confession block** — charts present, pivots
present, VBA present, extensions present — with a `preservation` field stating what survives
under the active mode. Golden-tested JSON per CONVENTIONS §2. `diff_cells(a, b)`: cell-level
semantic diff (address, old/new value, old/new formula) built on the same machinery the ledger
cross-check uses. A dependency sketch built on the tokenizer (which cells feed which) — coarse
is fine; it feeds Phase 6's guards and the agent's planning.

## Phase 5 — The oracle (LibreOffice driver + recalc + certification)

Mine the environment skill's recalc script and its LibreOffice bootstrap for the hard-won
operational knowledge (profile setup, sandbox socket restrictions, first-run behavior), then
productize:

- **Driver:** locate `soffice`; headless invocation against a **temp copy only** (the original
  path is never handed to LibreOffice — tested invariant); timeout with kill-and-retry-once;
  typed failures `OracleUnavailableError`/`OracleTimeoutError`; no parallel instances per
  profile.
- **`recalc()`:** recompute all sheets, scan every cell for Excel error tokens (#REF!, #DIV/0!,
  #VALUE!, #NAME?, #N/A), return structured results (status, totals, error locations) — mirror
  the skill's JSON shape so existing agent muscle memory transfers.
- **`certify()`:** the divergence check. Pre-flight on an *untouched* copy: LibreOffice recalc,
  then compare computed values against the file's own cached values — Excel's answer key for
  its current inputs — cell by cell, under the pinned tolerance and volatile-exclusion table
  (CONVENTIONS §3.7). Returns `CERTIFIED` / `DIVERGED` (with addresses and both values) /
  `BASELINE_UNVERIFIABLE` (cached values absent or stale — e.g., the file was last written by
  stock openpyxl). No judgment anywhere: measurements and typed states; the caller decides.
- Custody never depends on this phase: everything in Phases 2–4 works with no LibreOffice
  installed.

## Phase 6 — Reference-aware structural edits

- **6a — The guard (ship first).** Under preserve mode, `insert_rows`/`delete_rows`/
  `insert_cols`/`delete_cols` refuse (`UnsupportedStructureError`, precise message) when the
  shift would strand: formulas referencing shifted ranges (tokenizer-detected, cross-sheet
  included), defined names, conditional-formatting or data-validation ranges, merged ranges,
  tables, or **chart series ranges** — remembering that chart XML is raw-preserved bytes, so a
  row shift makes preserved charts point at wrong rows: refusal is the only honest v0 answer on
  chart-referenced sheets. Stock path keeps stock behavior + a loud warning. 6a converts the
  scariest silent corruption into a typed refusal — battery job 3 goes green-by-refusal here.
- **6b — The rewriter.** Excel-semantics shifting built on the existing `Translator`/tokenizer:
  relative and absolute references, cross-sheet references, defined names, CF/DV sqrefs, merged
  ranges, table extents, hyperlink ranges — implemented as ledger entries the splice writer
  applies, never as whole-sheet re-serialization. Property-style tests: insert-then-delete
  round-trips to the original; sums recomputed by the oracle match pre-edit values.
- **6c — Chart-range rewriting** inside preserved chart parts (targeted XML patch of series
  references), lifting the 6a refusal for charts. Scope 6c honestly; if it slips, the refusal
  stands — never the silent third option.

---

## Order, dependencies, and the gate

0 → 1 → 1.5 → 2a → 2b → 2c → 2d → 3 → 4 → 5 → 6a → 6b → 6c. Phase 3's guards depend only on
2a; Phase 5 is independent of 2b–2d and may interleave; 6a needs Phase 4's dependency sketch.
The standing gate between every phase: the five-job battery, upstream suite green, ledger
cross-check clean.

## Prohibitions (repo-specific, beyond CONVENTIONS)

- **No formula calculation engine, ever** — not "a few common functions," not "just SUM." A
  partial engine returns confident wrong numbers on the models that matter most; we route to
  real implementations (the oracle) and report divergence, we never approximate Excel's math.
- No rendering; no bundling or pip-depending on LibreOffice; oracle receives temp copies only.
- No new whole-workbook re-serialization code paths; touched sheets go through the splice, or
  the operation refuses.
- No behavior changes to `read_only`/`write_only` modes or the fresh-generation path.
- No string-formatted XML anywhere; no renumbering of existing sharedStrings/style indices.
- Never weaken a refusal to make a test pass; never label a fixture with provenance it lacks;
  never claim a test ran without running it.

## Ask-for-help triggers

Real-Excel fixture needs (via `FIXTURE-REQUESTS.md`); any upstream test that goes red; the
mixed fresh-chart-on-preserved-file semantics if PR-0's answer proves wrong in practice; any
place the sanctioned collateral set seems to need widening; performance guardrail misses; any
temptation to add "just a small" calculation capability — that one is a hard stop and a
conversation.
