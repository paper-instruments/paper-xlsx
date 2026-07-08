# Fork Engineering Conventions — paper-xlsx

**Status:** v1 — governs all development in this repository. Read completely before writing any
code. This document is deliberately prescriptive: where it pins a name, a schema, or a rule, use
it verbatim — these are decisions, not suggestions. When you hit a situation it doesn't cover and
the resolution would shape public API, stop and escalate to a human rather than invent.

**Context in four sentences.** This repository is a hard fork of **openpyxl**, the standard
Python library for Excel files (~300M downloads/month; the Excel engine under pandas). Unlike
other Office libraries, openpyxl's weakness is not a thin API — its verb surface is rich — it is
the **persistence core**: on save it regenerates the entire file from its in-memory model, so
everything it does not model (charts, drawings, VBA, pivot caches, modern extensions) is
destroyed, silently. The fork exists to transplant that spine: **the original package becomes
the source of truth; the object model becomes a source of edits to it.** Everything else —
the reader, the object model, the formula tooling — is excellent and stays.

---

## 1. Prime directives

**1.1 — Strict superset, with two sanctioned soft deviations.**
Upstream's own pytest suite stays green on every PR — that is the mechanical proof that every
existing caller keeps working. v0 makes **zero hard behavior changes** to existing public API,
with exactly two sanctioned soft deviations, both logged in `PAPER.md`:

- **Lossy-save warnings.** The default (stock-behavior) save path gains loud runtime warnings
  when it is about to destroy content it cannot preserve, enumerating what dies ("this workbook
  contains 3 charts, 1 pivot cache, and a VBA project that will be dropped; open with
  `preserve=True`"). Warnings, never exceptions, on the stock path.
- **Preserve mode.** `load_workbook(path, preserve=True)` opts a workbook into the new spine:
  lossless save, refusals active (e.g., `data_only`+save refuses; unsafe structural edits
  refuse). Pure opt-in; upstream code never enters it. Our agents always use it.

Flipping `preserve=True` to the default, and making the `data_only`-save refusal default, are
the first two entries in PAPER.md's **Future breaking-change candidates** ledger — executed only
after the fixture corpus proves the spine, never silently.

**1.2 — The name rule.**
PyPI distribution name: `paper-xlsx`. Python import name: **`openpyxl`, frozen forever.** Yes,
this breaks cosmetic symmetry with any rename instinct — the rule was never about matching the
pip name; it is about the import line the world already writes. `import openpyxl` appears in
millions of scripts, in pandas itself, and in every model's training prior. Never rename the
package directory; never write `import paper_xlsx` anywhere, including tests and docs. The
sentinel `__paper_version__` on the package root distinguishes this fork from stock and from any
name-reservation stub on PyPI.

**1.3 — Refusal atomicity.**
A refused operation leaves the in-memory model, the dirty ledger, and any file on disk exactly
as they were. Validate-fully-then-mutate. Every documented refusal condition gets a test
asserting (a) the typed refusal is raised and (b) output bytes equal input bytes.

**1.4 — The reopen rule.**
Every test assertion about workbook content goes save → reopen → assert; never assert on the
in-memory object you just mutated. Where LibreOffice is available, high-risk paths additionally
assert the saved file loads in it (independent implementation).

**1.5 — Fail loudly. Silence is the enemy.**
This library's historic failure mode is the file that opens fine and is quietly wrong. Nothing
in this fork may silently drop, silently skip, or silently corrupt. The three legal outcomes of
any operation are: done correctly; refused with a typed error saying what was found and why it
was unsafe; or done with a loud warning enumerating exactly what could not be preserved.

---

## 2. API design (pinned)

- **Additive placement.** New capability = new methods/properties on existing classes, or new
  modules under `openpyxl.` Never repurpose or shadow an existing name.
- **Keyword-only options** after primary positionals; boolean defaults are the safest behavior.
- **Addressing (pinned).** Public APIs and JSON payloads address cells and ranges as
  sheet-qualified A1 (`"Model!B7"`, `"Assumptions!B2:D10"`) — the format's native, human-legible
  anchor. Structural edits (row/col insert/delete) invalidate addresses by nature; any API that
  performs one returns an address-remap object, and post-edit operations must re-resolve rather
  than reuse pre-edit addresses.
- **Typed returns, stable JSON.** Inspection results are small typed objects with `.to_dict()`:
  snake_case keys, deterministic key order, top-level `"schema"` name and integer `"version"`,
  0-based indices only where the concept is inherently index-like, A1 everywhere else.
- **Exceptions (pinned).** `openpyxl.errors` (new module) defines `PaperRefusal(Exception)` as
  the base for all safe refusals, with subclasses: `AmbiguousTargetError`, `TargetNotFoundError`,
  `UnsupportedStructureError`, `BoundaryViolationError`, `RelationshipPolicyError`, plus the
  oracle pair `OracleUnavailableError` and `OracleTimeoutError`. Programmer errors remain
  `TypeError`/`ValueError`. Certification outcomes are NOT exceptions — see next bullet.
- **Certification results (pinned).** The oracle's divergence check returns a typed result, one
  of three states: `CERTIFIED` (LibreOffice reproduces the file's cached values within
  tolerance), `DIVERGED` (with the list of disagreeing addresses and both values), or
  `BASELINE_UNVERIFIABLE` (the file's cached values are absent/stale, so no answer key exists).
  Callers — agents or humans — decide what to do; the library reports measurements, never
  judgments.
- **No new runtime dependencies.** lxml stays optional as upstream has it. LibreOffice is a
  system tool detected at call time — never a pip dependency, never bundled. Its absence raises
  `OracleUnavailableError` from oracle APIs only; everything custody-related works without it.

---

## 3. Implementation doctrine — the spine

**3.1 The two jobs of the object model.** openpyxl's parallel object model holds two jobs today:
(1) an affordable in-memory representation of a grid — rational, keep it forever; a live XML
tree per cell is forbidden at any scale — and (2) the source from which the file is regenerated
at save — this is where losslessness dies, and this job is terminated. Under preserve mode, the
original archive is the source of truth and the model is a source of edits to it.

**3.2 Byte retention.** At load (preserve mode), retain the original archive's part payloads.
Memory budget: retained bytes ≈ the .xlsx file size — noise next to the object model's own
footprint. Retain bytes, not a path, so save-over-same-file composes with atomic
temp-write-then-rename.

**3.3 The dirty ledger — load-bearing, not an optimization.** Every public mutation records
(part, coordinate/region) at instrumented chokepoints. This differs from a compare-based
patch-save **by necessity**: a compare-based design needs the library to serialize a faithful
candidate to compare against the original, and openpyxl's serialization is the lossy act — there
is nothing faithful to compare. Consequences, all pinned: instrument the chokepoints
exhaustively (cell value/formula setters, style assignment, row/column dimensions, sheet
add/remove/rename, defined names, table and CF/DV mutations, properties); provide a documented
`mark_dirty(part_or_range)` escape hatch for anyone reaching below the public API; and in
test/debug mode, cross-check the ledger against a region-level semantic diff of output vs.
original — a mutation the ledger missed is corruption inside the safety tooling and is a
release-blocking bug class.

**3.4 Save = ordered-stream splice.** Untouched parts: raw payload copy — byte-identical by
construction, never even parsed. Touched worksheet parts: stream-parse the **original** sheet
XML, pass events through verbatim, and at each `<c>` whose address is in the dirty set, emit the
replacement cell; new cells splice into position because rows and cells are coordinate-ordered —
a merge of two sorted streams, O(sheet size), O(1) memory. Everything unmodeled inside the sheet
(sparkline extLst, x14 conditional formatting, the `<drawing>` reference that keeps charts
attached, `mc:AlternateContent`) passes through because the model is never asked to understand
it. For touched parts the invariant is semantic correctness plus preservation of unmodeled
content — not byte identity; byte identity is the invariant for untouched parts only.

**3.5 Cross-part discipline (pinned strategies).**
- `sharedStrings.xml`: **append-only** — new strings become new `<si>` entries at the end with
  count attributes bumped; existing indices never renumber. Sanctioned simplification: changed
  or new string cells may be written as **inline strings** (`t="inlineStr"`) in the spliced
  sheet, avoiding sharedStrings coordination entirely; Excel normalizes on next save. Either
  strategy is legal; pick one per operation class in PR-0 and document it.
- `styles.xml`: append-only for new xf records; existing style indices never renumber.
- `calcChain.xml`: **delete the part** whenever formulas change — legal, and Excel rebuilds it.
  Its removal is part of the sanctioned collateral set.
- `[Content_Types].xml`, rels: updated only when parts are added/removed, via targeted edits.
- `workbook.xml` (calcPr, defined names, sheet list): spliced where cheap; where a small,
  fully-modeled element must be re-serialized, the changed-part budget test is the guard.
- **Sanctioned collateral set (pinned):** a single-cell edit may legally change: that sheet's
  part, sharedStrings (append only, if used), calcChain (deleted), and workbook.xml's calcPr
  (recalc-on-load flag). Anything else changing is a test failure.

**3.6 XML machinery.** New XML vocabulary uses upstream's own declarative mapping framework —
the `Serialisable`/descriptor system in `openpyxl/descriptors/` — never hand-assembled element
trees in API code, and **never string-formatted XML anywhere**. The splice writer emits through
the same event-writing machinery upstream uses (et_xmlfile/lxml), with namespace handling
verified against fixtures from multiple producers.

**3.7 Volatile-function table (pinned).** For certification, cells downstream of
nondeterministic functions — `NOW`, `TODAY`, `RAND`, `RANDBETWEEN` — are excluded from
value comparison and reported separately. `INDIRECT`/`OFFSET` are volatile but deterministic
given inputs: they stay **in** the comparison. Numeric tolerance: relative 1e-9 with absolute
floor 1e-11; text and error values compare exactly. Deviations from this table require PR-0
amendment, not local judgment.

---

## 4. Testing contract (first-class phase — built before organs)

**Fixture corpus** — `tests/paper/fixtures/`, frozen via `MANIFEST.sha256`, sidecars per
fixture (schema below), never regenerated by code under test:
- **Provenance buckets:** authored in real Excel (the load-bearing bucket — see
  `FIXTURE-REQUESTS.md`), exported from Google Sheets, exported from LibreOffice, written by
  stock openpyxl (deliberately: the stale-cached-values case), generated by our own code.
- **Taxonomy:** minimal-clean; feature-isolated (one file per feature: chart+embedded elements,
  pivot table+cache, .xlsm with macros, sparklines/x14 conditional formatting, defined names +
  cross-sheet formulas, tables/ListObjects, merged cells, external links, data validation,
  hidden rows/sheets); gauntlet (a real-model-shaped file combining all of it); an .xlsb and a
  legacy .xls (refusal tests only); corrupt-by-construction; large (≥100k cells, perf smoke).
- **Sidecar schema (pinned):**

```json
{
  "fixture": "example.xlsx",
  "provenance": {"app": "Excel 16.x", "version": "…", "notes": "…"},
  "features": ["chart", "pivot", "vba"],
  "ground_truth": {
    "cached_values": {"Model!B7": 0.09, "Model!D42": 1234567.0},
    "chart_count": 3, "pivot_count": 1, "vba_present": true,
    "formula_count": 214
  },
  "verified_by": "human name", "date": "YYYY-MM-DD"
}
```

- **Provenance honesty:** you cannot run desktop Excel. Bootstrap with LibreOffice-authored
  fixtures labeled truthfully; maintain `FIXTURE-REQUESTS.md` for the real-Excel files a human
  must produce (real client-model-shaped workbooks especially). Never label a fixture with
  provenance it lacks.

**Contract harness** — shared conftest in `tests/paper/`; every mutating API passes all five:
1. Save → reopen → assert (never in memory).
2. Intended effect present in the reopened workbook.
3. Changed-part budget: package diff shows exactly the expected parts changed (per the
   sanctioned collateral set) and every other part payload byte-identical.
4. Independent-loader smoke: LibreOffice headless load/convert exit-code check where `soffice`
   exists; marked `lo_smoke`, skippable where not.
5. Refusal atomicity: every documented refusal input raises the typed refusal AND leaves the
   output byte-identical to the input.

**Invariant suites:**
- No-op load(preserve=True)+save: part list identical, every part payload byte-identical.
- Single-cell edit: exactly the sanctioned collateral set differs.
- Splice completeness trap: a fixture whose sheet carries content openpyxl half-understands
  (sparkline extLst, x14 CF, drawing ref) survives a one-cell edit intact — this is the spine's
  signature test.
- Ledger cross-check (debug mode): ledger claims == semantic region diff, on every harness run.
- Certification determinism: same file → identical certification result, run twice.
- Oracle isolation: the original file path is never handed to LibreOffice — temp copies only,
  asserted.
- Performance guardrail: splice save on the large fixture within a pinned multiple of stock
  save time (set the number in PR-0 after baseline measurement); fresh-generation path
  (`Workbook()` → save) unaffected within noise.

**Discipline:** golden files update only via explicit command with human-reviewed diffs; any
date-stamping takes an injectable clock; **no fix without a fixture** — every bug found
downstream becomes a frozen fixture + failing test before its fix merges.

---

## 5. Scope policy

- **In the package:** the spine, honesty guards, manifest/diff perception, the oracle driver
  (recalc + error scan + certification), reference-aware structural edits, typed refusals.
- **Out of the package, permanently:** a formula calculation engine (a partial engine is a
  silent-wrongness machine — we route to implementations of Excel's semantics, never
  approximate them); rendering of any kind; bundling LibreOffice; .xlsb/.xls support beyond a
  typed refusal that names the format and suggests conversion; pivot-cache refresh (preserve +
  confess staleness instead).
- Workflow QA, styling presets, and skill-level guidance live in the consuming harness, not
  here.

---

## 6. Repo hygiene

- One organ per branch per PR; PR descriptions link the `PAPER.md` entry.
- Never reformat upstream files; formatters scoped to new files only.
- `PAPER.md` ledger: baseline test results at fork point, sanctioned soft deviations (§1.1),
  future breaking-change candidates, upstream-sync policy. **Upstream note:** openpyxl lives on
  Heptapod (Mercurial), releasing roughly annually — sync policy is: diff each upstream release
  against `paper-base`, port relevant changes as reviewed patches; never attempt automated
  history merges across the hg/git boundary.
- Sentinel `__paper_version__` maintained; new tests under `tests/paper/`, CI runs upstream's
  suite and ours.

---

## 7. Package kernel (pinned)

A new submodule `openpyxl.package` exposing `xml_equivalent(a, b) -> bool` (semantic XML
comparison that never normalizes cell text content), `diff_package(path_a, path_b) ->
PackageDiff` (part-by-part; XML parts semantic, binary parts size/hash; typed result with
`.to_dict()`), and the diagnostics the spine uses internally. Note the family difference: here
patch-writing is not an opt-in utility — under preserve mode it **is** the save path; the kernel
module exists so tests and agents can verify what save did. Zip writing is deterministic (fixed
entry order and timestamps — decide, implement, test); the byte-identity invariant is defined on
**part payloads**, not whole-archive bytes; all writes go temp-file-then-atomic-rename with a
failure-injection test proving the original survives a mid-write crash.

---

## 8. PR-0 protocol

After the test-infrastructure phase, the next PR is an **API Proposal**: exact signatures,
return types, refusal conditions, and 2–3 usage examples for every planned organ, grounded in
the actual code — plus the decisions this document delegates (inline-string vs. sharedStrings
strategy per operation class; the performance budget number; the exact chokepoint inventory for
the ledger). It confirms the pinned shapes (§2 exceptions/addressing, §3 collateral set and
volatile table, §4 sidecar schema, §7 kernel) fit the real code, flagging mismatches for human
decision. Humans approve PR-0 before mass implementation. Deviations later are amended via PR,
never silently.

---

## 9. Definition of done (per organ)

- Contract-harness assertions pass on the relevant fixtures, including refusal-atomicity cases.
- Upstream pytest suite green.
- New tests under `tests/paper/`; goldens updated deliberately; ledger cross-check clean.
- `PAPER.md` entry written; docstrings on all new public API.
- No diffs outside the organ's scope; no upstream files reformatted.
