# paper-xlsx — v0.1: The Completion Wave (final, supersedes all prior drafts)

## What this is

v0 proved the thesis: the spine (retention → ledger → splice → cross-part discipline), the
oracle, structural rewriting with Excel reference semantics, and the three-legal-outcomes
contract all exist and work — ~5,650 lines of library code in `openpyxl/preserve/`,
`openpyxl/package/`, and `openpyxl/oracle.py`, with upstream's suite untouched and green. The
post-implementation adversarial review then did its job: it confirmed the architecture survived
contact with reality, and it found that the remaining third of the mission lives at the
**boundaries** — satellite sheet regions, modeled objects backed by preserved bytes, and the
perception layer agents actually navigate by. It also found two silent-corruption bugs and one
honesty-gap family **on main right now**.

v0.1 is one wave that finishes the mission. Per explicit decision: **nothing is milestone-split
into a later version unless it independently earned deferral** — the five deferred items and
their triggers are in Appendix A; everything else is in scope here, organized as batches for
review cadence, not calendar. Estimated volume: roughly 2–3× v0's library code plus tests. The
process that carried v0 carries this: one batch = one commit series = one review gate.

`CONVENTIONS.md` remains the law. Two process amendments ship alongside Batch 0 (see Process
Amendments below). The named threat class this wave adds to the vocabulary, verbatim from the
review: **"the dangerous boundary isn't unmodeled content — it's modeled objects backed by
preserved bytes."** Design every Batch 1–4 decision against it.

---

## The battery — expected-state transitions (the wave's acceptance spine)

The standing battery grows to 24 jobs. Each row is implemented as a test at the state marked
"today" (three of them assert *corruption* — those tests exist to be flipped by Batch 0). The
pass criterion for every job, forever: **correct, or loudly refused — never silently wrong.**
Expected states transition only when the implementing batch merges; weakening a state to make a
batch "pass" is a firing offense for the change, per the standing prohibition.

| # | Job | Today (post-v0, per review) | Required after v0.1 | Batch |
|---|---|---|---|---|
| 1 | Assumption flip on charted model | correct | correct | — |
| 2 | pandas `mode="a"` onto charted report | correct with `engine_kwargs` | correct; internal default flip | 0-exit |
| 3 | Insert row above SUM schedule | correct | correct | — |
| 4 | .xlsm round-trip | correct | correct | — |
| 5 | data_only trap | refuse | refuse | — |
| 6 | Edit inside shared-formula group | **UNKNOWN — probe item zero** | correct or refuse | 0 |
| 7 | Write into dynamic-array spill | refuse | refuse w/ `in_spill` context | 1 |
| 8 | Rename sheet referenced by formulas/names | refuse | correct (cascade rewrite) | 3 |
| 9 | Write to locked cell, protected sheet | **silently complies** | warn/refuse per mode | 1 |
| 10 | Table append (calc columns + totals row) | refuse / silent non-extension | correct | 2 |
| 11 | Copy sheet within charted workbook | refuse | correct | 3 |
| 12 | Scenario: set 3 inputs, read 5 outputs | hand-rolled loop | one `evaluate()` call, certified | 5 |
| 13 | Load encrypted (CFB) workbook | cryptic zip error | typed refusal + decrypt route | 1 |
| 14 | **No-op save, `cols`/hidden fixture** | **CORRUPTS (confirmed)** | byte-identical | 0 |
| 15 | **Edit self-closing region (autoFilter/pageSetup)** | **CORRUPTS (confirmed)** | correct | 0 |
| 16 | Mutate loaded `table.ref` / `chart.title` | **silent staleness (confirmed)** | refuse or correct | 1 |
| 17 | Value edit feeding formulas → human opens file | stale masquerade, no flag | `fullCalcOnLoad` set | 1 |
| 18 | "Make this range a table" | refuse | correct | 2 |
| 19 | Add comment to comment-free sheet | refuse | correct | 2 |
| 20 | Edit x14 CF on professionally formatted book | refuse (highest-traffic refusal) | correct (twin-sync) | 3 |
| 21 | Ordinary value write on Excel-365 file | refuse (cm/vm) | correct | 3 |
| 22 | Add chart to loaded workbook | refuse | correct | 4 |
| 23 | Locate cell by header/label ("WACC input") | no API | localized or `AmbiguousTargetError` | 6 |
| 24 | Oracle write-back of recalculated values | n/a | correct, **certification-gated** | 5 |

---

## Batch 0 — Restore the invariant (merges before everything; starts immediately)

Corruption on main outranks all planning. Fixtures-first for every defect, per the standing
rule.

- **0.1 Item zero: the shared-formula probe.** Determine whether the splice is `si=`-group
  aware. Excel stores one master formula plus pointer followers; a group-blind splice editing a
  master silently rewrites every follower, and editing a follower desyncs the group. If v0
  handles it: flip battery job 6 to "correct," document where, done. If not: it is corruption
  in the same class as 0.2/0.3 and gets fixed in this batch — detect membership at load; on any
  edit touching a group, un-share correctly (materialize followers' translated formulas via the
  Translator) or refuse with the group extent named. Record the un-share-vs-refuse default
  decision in the amendment register.
- **0.2 Self-closing region corruption.** The scanner never sets `RegionSpan.end` for
  self-closing top-level elements; editing `<autoFilter …/>`, `pageSetup`, `sheetPr`,
  `sheetFormatPr` under preserve emits malformed XML. Fix the scanner; then add the **region ×
  self-closing matrix** as a property test across every satellite region and both element
  forms.
- **0.3 No-op false-dirty — fix the pattern, not the instance.** Upstream's
  `DimensionHolder.to_tree()` mutates state as a side effect, so arm-time and save-time renders
  differ, false-dirtying `<cols>` and corrupting via 0.2 on a **zero-edit save**. Patch the
  instance, then harden the class: double-render every region at arm time; on self-mismatch,
  pin the region to original bytes and mark it refuse-on-edit. Any impure upstream serializer —
  present or future — must land in "pinned," never in "false dirty."
- **0.4 Property tests as permanent infrastructure:** no-op byte-identity across **all**
  fixtures × **all** regions; ledger cross-check extended to region claims (a region the ledger
  didn't claim may never differ).
- **0.5 Batch-0 exit actions:** flip `preserve=True` as the default **in our own harness
  images** (we control that blast radius; the battery covers it); record in PAPER.md. The
  *public* default remains a release gate — Appendix A, item 1.

## Process Amendments (commit alongside Batch 0)

1. **Pinned-surface CI check.** Every exception class, result state, and return type pinned in
   CONVENTIONS or an approved API proposal must be raised/produced by at least one test, or
   carry an explicit ledger entry. (The AddressRemap-returns-None and never-raised-refusals
   findings, mechanized away.)
2. **Adversarial review is a standing gate.** It has now out-caught the 2,800-test suite three
   times; it graduates from event to infrastructure. Every batch below ends with an adversarial
   pass over its diff + the new surface; findings become fixtures before the batch merges.

## PR-1 — API Proposal gate (between Batch 1 and Batch 2)

Signatures, refusal conditions, and examples for all new surface in Batches 2–7 (lifecycle
engine, table verbs, `evaluate()`, write-back, findings, receipts, validate, style verbs),
grounded in the live code; carries the delegated decisions (un-share default if still open;
batch-`evaluate()` interface and pool lifecycle; findings taxonomy; receipt schema). Human
approval gates mass implementation. Amendments via the register, never silently.

## Batch 1 — Honesty completion (the boundary class, closed)

- **1.1 Preserved-part-backed live objects** — the review's central discovery. Loaded tables,
  charts, images, pivots, external links are mutable with no ledger hooks; edits vanish
  silently on save (the forbidden fourth outcome). Close it: arm read-only guards or ledger
  hooks on every such object class. **Refusal is fully acceptable; silence is not.** Where
  Batches 2–4 later add real support (tables, charts), they lift these refusals; Batch 1's job
  is only to make the boundary honest *now*.
- **1.2 `fullCalcOnLoad` widened** to value edits whose address hits the dependency sketch —
  the single most common agent edit must never save stale caches unflagged (battery 17).
- **1.3 Pinned-surface debts paid:** implement `AddressRemap` (shifts return it, period); raise
  or formally ledger the three never-raised refusal classes.
- **1.4 Loss-inventory completeness:** chart `colors1.xml`/`style1.xml` (fix the dead
  `endswith`), rich-text run stripping, `fileSharing`, `protectedRanges` — all enumerated in
  the confession or preserved.
- **1.5 Input honesty:** duplicate zip entry names → typed refusal (load-vs-copy parser
  differential is an attack surface); encrypted CFB signature → typed refusal naming the
  condition and the decrypt route; `diff_cells` read-path warning spray silenced.
- **1.6 Protection awareness:** consult sheet protection and cell `locked`; preserve-mode
  writes to locked cells warn by default, refuse under strict flag; manifest confession gains a
  protection summary. We report protection; we never bypass it.
- **1.7 Certification noise classes:** external-workbook links and oracle-unsupported functions
  classified excluded-with-reason (like volatiles), so `DIVERGED` keeps meaning "genuine
  disagreement."
- **1.8 Producer-fingerprint pinning:** pin what we write into app.xml for fresh content; add
  the fixture asserting preserved files keep original app.xml byte-identical; add fresh-chart
  outputs to `FIXTURE-REQUESTS.md` for a human real-Excel open check (the field incident on
  record: Excel rendered charts differently on the producer string alone; LibreOffice smoke is
  blind to this class).

## Batch 2 — The part-lifecycle engine and its unlocks

- **2.1 The engine:** generalize v0's added-sheet create/delete cascade (part + content-type +
  rels + confession) into one primitive. Every subsequent "add a part" feature in this plan
  routes through it; no bespoke cascades.
- **2.2 Tables:** create ("make this range a table" — battery 18), remove, and the full row
  discipline: append/insert/delete that extends the range, respects calculated columns, and
  keeps the totals row last (battery 10). Guards refuse structures we can't keep coherent.
- **2.3 Comments:** add/edit on comment-free sheets (the 80% case); legacy-vs-threaded handling
  per the confession; threaded authoring stays triggered (Appendix A footnote).
- **2.4 Engine dividends:** docProps/styles.xml creation, custom-props deletion,
  `wb.replace_part()` for media swaps.

## Batch 3 — Region and structural completion

- **3.1 x14 twin-sync** for conditional formatting and data validation — the highest-traffic
  refusal in professionally formatted workbooks (battery 20). Legacy and x14 twins edited in
  lockstep or the operation refuses.
- **3.2 Sheet lifecycle cascade:** rename with formula/defined-name/chart-text rewrite (the
  chartpatch text-walker already locates the sites — battery 8); delete with the
  calcChain-drop cascade as template plus a reference audit (enumerate what breaks;
  refuse or proceed-with-report); reorder with `localSheetId` remap; copy-as-added-sheet
  preserving charts (battery 11).
- **3.3 Structural widening:** multiple shifts per session via ledger rebase (retire
  one-shift-per-session); spanning-merge expansion; `move_range` as tracked delete+insert;
  per-blocker unlocks replacing coarse refusals.
- **3.4 Dynamic arrays:** cm/vm metadata bookkeeping so ordinary value writes on Excel-365
  files stop refusing (battery 21); spill-range writes keep refusing, now with `in_spill`
  context.
- **3.5 Structured references** in every shifting path: extend the Translator or refuse —
  `Table1[@Revenue]` is never mis-shifted.

## Batch 4 — Charts and images under preserve

- **4.1 Phase one:** charts/images on **added** sheets — stock writer territory, zero splice
  risk (lifts part of battery 22).
- **4.2 Phase two:** splice `<drawing r:id>` into loaded sheets — new drawing part via the
  lifecycle engine, one spliced element into original sheet bytes (battery 22 complete).
- **4.3 Chart editing:** chartpatch-based mutation of loaded chart objects (lifting the Batch-1
  refusal), plus `chart.repoint(series → range)` — "the chart now covers Q1–Q4."
- Every 4.x output joins the real-Excel open-check queue (1.8) — this is the producer-sensitive
  surface.

## Batch 5 — Computation layer

- **5.1 The scenario runner.** `wb.evaluate(set={...}, read=[...]) → typed result`: temp copy →
  spine-applied inputs → oracle recalc → harvested outputs → original untouched (asserted).
  Batch mode takes input-vector grids for sensitivity tables on a **warm LibreOffice pool**
  (per-profile isolation, crash-replace, bounded size — the one performance vertical this wave
  funds). Results carry certification state. Battery 12.
- **5.2 Formula pre-flight linter.** Before any formula string reaches a cell: tokenize;
  function catalog; balanced parens; references resolve (sheets/ranges/names); structured-ref
  validity; locale separators (`;`-for-`,` is a silent #NAME? factory). Wired into setters
  under preserve as warn-or-refuse per strictness flag. Milliseconds instead of an oracle
  round-trip.
- **5.3 Oracle v2 — write-back, certification-gated.** Splice recalculated `<v>` values into
  the *original* package (macro-safe in-place recalc; closes the stale-cache story inside
  custody). **Hard design conditions:** only `CERTIFIED` values are ever written back;
  `DIVERGED`/`BASELINE_UNVERIFIABLE` write-backs refuse, or proceed only under an explicit
  flag with the receipt stamped loudly — we do not launder LibreOffice's disagreements into
  Excel's cache slots under our signature. `fullCalcOnLoad` managed coherently (cleared only
  when write-back covers the full dependency set). Package-diff confession on every
  write-back. Battery 24.
- **5.4 Recalc-flag support for fresh-generation files** (the one oracle gap stock-path users
  hit).

## Batch 6 — Perception and the agent experience

- **6.1 Localization:** find-cell-by-header/label (raise the pinned `AmbiguousTargetError` at
  last — battery 23); workbook-wide content search with the family's normalization discipline.
- **6.2 Model map:** input/calculation/output classification (no-formula-and-referenced,
  corroborated by color convention where present); label-relative addressing ("the cell right
  of 'WACC'").
- **6.3 LibreOffice-free static error scan** (cached error tokens, broken-ref detection) so
  blind environments get a cheap first-pass check.
- **6.4 DV vocabulary lookup** ("what values does this cell accept").
- **6.5 Manifest enrichment:** formula addresses, inputs/outputs from the sketch, `certifiable`
  flag, part names, protection summary (1.6).
- **6.6 The edit receipt:** one composed artifact per session — semantic diff + package diff +
  confession + recalc/certification status. This is the custody guarantee made legible; agents
  and humans consume the same object.
- **6.7 `wb.validate()` preflight** and **structured refusal fields** (machine-readable kind,
  anchor, remedy on every `PaperRefusal`).
- **6.8 Hygiene findings — `inspect.findings()`:** hardcoded constants in formulas;
  inconsistent formulas across a projection row; error cells; orphaned names; external-link
  inventory; hidden sheets/rows; merged-cell hazards; volatile census; *advisory* magnitude
  lint (flags a 20-where-0.20-expected pattern, never decides — see fences).
- **6.9 Workbook diff report:** human-readable model_v3→v4 across values/formulas/names/
  structure, classifying *content changed* vs *shifted by structural edit* (rides on
  `AddressRemap`).

## Batch 7 — Delivery, hardening, adoption

- **7.1 Style verbs + convention profiles:** `copy_format(src, dst_range)`; profile application
  (blue inputs / black formulas / green links); the number-format library
  (`$#,##0`, `0.0%`, `0.0x`, parenthesized negatives) as data, per-customer overrideable.
- **7.2 Small ergonomics:** `set_input(name_or_label, value)` via defined names then the model
  map; protect-for-delivery (lock all but classified inputs); pivot **refresh-on-load flag**;
  scrub/sanitize (strip comments, metadata, personal info; hidden-sheet report-or-remove) — the
  compliance verb before external delivery.
- **7.3 Hardening:** fsync durability on atomic rename; spool-to-disk saves (measured peak is
  ~3× file size; take it to ~1×); decompression caps; zip-confusion checks; `mark_dirty` range
  clamps.
- **7.4 Adoption:** README + docs for the paper API — preserve, manifest, refusals, oracle,
  receipts (the review's own highest-leverage-per-line finding: the fork is currently
  invisible); document the public-default release gate (Appendix A item 1) and the internal
  flip already taken at Batch-0 exit.

---

## Order, dependencies, sizing

0 → amendments → 1 → PR-1 → 2 → 3 → 4 → 5 → 6 → 7, with allowed overlaps: 5.2 (linter) any
time after PR-1; 6.x independent of 4; 7.4 docs may trail each batch it documents. Batch 0 has
absolute precedence and should be underway before this document is even approved — corruption
fixes never wait on planning. Standing gate between batches: full battery at expected states,
upstream suite green, ledger + region cross-checks clean, adversarial pass complete with
findings fixtured. Long poles: 0.1 if the probe comes back bad; 3.1 (x14); 4.2 (drawing
splice); 5.3 (write-back). Nothing here is tens-of-thousands of lines; the wave totals roughly
2–3× v0.

## Fences (prohibitions, not deferrals — violating one is a conversation, never a commit)

No calculation engine — not one function; residual agent errors are semantic (sign conventions,
period alignment) which an engine cannot touch and a partial engine worsens. No rendering. No
pivot *creation* (preservation + refresh-on-load covers brownfield; agents build SUMIFS
summaries for greenfield). No semantic input validation — the 20-vs-0.20 hazard is judgment;
6.8's lint may flag, the library never decides. No VBA editing (preserve always, touch never).
No R1C1, no ODS, no legacy .xls beyond refusal, no concurrency/xlsx-as-database, no three-way
merge. No silent third option, anywhere, ever.

## Pitfall register (fixture + test on first contact)

Locale-canonical formula storage (5.2 owns); 1904 date epoch — date writes respect workbook
date mode; sheet-scoped duplicate defined names → refuse naming both candidates; rich-text
cells replaced with plain text lose in-cell runs (legal + warned; inventory per 1.4); impure
upstream serializers (0.3's pattern guard owns); producer-fingerprint sensitivity (1.8 owns);
`append()`-path ledger granularity — region entries, not per-cell (audit under 0.4's
cross-check).

---

## Appendix A — Deferred, with triggers (each earned its deferral; nothing else did)

1. **Public preserve-by-default flip (incl. pandas' default on the published package).** Zero
   code withheld — mechanism ships in this wave; internal images flip at Batch-0 exit. The
   *public* default is a release gate because two silent-corruption bugs were confirmed on main
   the week this plan was written; flipping the default routes every `pd.ExcelWriter(mode="a")`
   on earth through that code. **Gate:** full region matrix green + battery green + soak on
   real customer files with zero custody incidents. Then flip, pandas-append path first.
2. **Cross-workbook sheet import.** The hard 80% is collision reconciliation (style tables,
   duplicate names, numFmt IDs, themes) — plausibly a v0-sized project — and demand is zero:
   no battery job, no customer ask, and the 40-agent frequency census never surfaced it.
   **Trigger:** first assembly/library customer.
3. **.xlsb reading.** A different binary format (BIFF12): thousands of parser lines or a new
   runtime dep, to save one conversion step that already works via the typed-refusal +
   LibreOffice route. **Trigger:** first xlsb-heavy customer; the decision then is
   route-vs-dependency, not build.
4. **Streaming edit mode.** An architecture change for a file that has never appeared; 7.3's
   spool-to-disk removes most of the motivation (~1× file-size memory). **Trigger:** the first
   file that cannot load — whose actual shape then determines the right lazy design.
5. **Incremental (dependency-scoped) certification.** Optimization of an unmeasured cost, with
   a subtle new honesty risk (an under-built graph silently under-certifies). **Trigger:**
   measured recalc latency pain on a real model, profile attached.

*(Footnote: threaded-comment authoring rides trigger 2's class — first review-workflow
customer.)*

## Definition of done for the wave

All 24 battery jobs at their required states; upstream suite green; pinned-surface CI check
green; no-op byte-identity across all fixtures × regions; every adversarial finding fixtured
and fixed or ledgered; README/docs shipped; internal default flipped; public-default gate
documented with its conditions; PAPER.md ledger current. At that point the package meets the
program goal as defined: every operation an agent performs on a real workbook ends correct,
loudly refused, or loudly enumerated — with a receipt — and the residual error budget lives
above the library, in the skill, the evals, and the training data.
