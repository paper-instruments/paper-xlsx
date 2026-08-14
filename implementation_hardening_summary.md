# paper-xlsx pre-open-source implementation summary

Package: `paper-xlsx` on the implementation-hardening branch, compared with
`paper-base` `c4986390`.

Stock openpyxl is the create/edit baseline. This fork is for **agent spreadsheet
work on real files**: inspect, edit, preserve unsupported OOXML, verify the
change, and deliver the result without rewriting unrelated package parts.

**Verdict:** the hard fork has a place. This branch applies the required
pre-open-source cut: it keeps the preservation spine, fixes the defects in
existing Paper features, adds the two narrowly scoped missing operations,
removes accidental API and fork surface, and restores upstream behavior when
preservation is disabled.

This is the concise companion to `implementation_hardening_spec.md`. That file
contains the reproductions, edge cases, acceptance criteria, and implementation
detail.

## What paper-xlsx currently adds: does it have a place?

The rows below group low-level modules by the product behavior they provide.
Read **What it does** first; the API column only identifies the corresponding
surface.

| What it does | API or implementation | Use case | Place? | Decision |
| --- | --- | --- | --- | --- |
| Open an editable workbook in preservation mode by default, retain its source package, inventory its content, and arm change tracking. | `load_workbook(..., preserve=...)`, reader, inventory | Edit existing | Yes | Keep. Fix file-like source detection. |
| Observe direct value, type, style, hyperlink, comment, and protection assignments so ordinary openpyxl code participates in preservation. | Cell/style hooks, preservation ledger | Edit existing | Yes | Core reason for the hard fork. Keep, but replace the quadratic rollback snapshots. |
| Track sheet, name, table, chart, region, row, and column changes and either rewrite dependent references or refuse. | Workbook/worksheet hooks, `AddressRemap`, structural planner | Structured edit | Yes | Keep, but close the untracked surfaces and compose non-conflicting plans. |
| Save supported changes by patching original XML and copying untouched ZIP parts rather than regenerating the entire workbook. | Preserve saver, splice/emit modules | Package preservation | Yes | Core reason for the hard fork. Keep and split the planner into reviewable phases. |
| Clear stale formula caches and set recalculation metadata after formula-affecting edits. | Formula-cache scanner and splicer | Calculation correctness | Yes | Keep. This is directly supported by retained-artifact eval evidence. |
| Refuse an unsupported change before delivery and write the destination atomically. | Typed errors, validation, save delivery | Package safety | Yes | Keep. Make rollback and refusal closed contracts for every supported path. |
| Preview or explain the workbook/package changes that a save will make. | `validate()`, save receipts, package diff, cell diff, `scan_errors()` | Verification | Yes | Keep. Share one pure plan and report preservation-derived effects. |
| Inspect validation choices, search workbook content, or copy formatting through deterministic helpers. | `allowed_values()`, `search()`, `copy_format()` | Agent ergonomics | Yes | Keep. Each operation has an explicit target and observable result. |
| Guess a nearby value cell from a label and optionally mutate it. | `locate()`, `set_input()` | Agent ergonomics | No | Remove. The six-cell neighborhood heuristic can select an unrelated cell; explicit coordinates and defined names already cover safe targeting. |
| Append a row to a loaded Excel table while expanding its range, preserving totals/filter state, and deriving declared calculated columns. | Current internal `append_row()` | Structured table edit | Yes | Keep the job. Replace the hidden helper with atomic `Worksheet.append_table_row()`, remove formula guessing, and keep it separate from generic `Worksheet.append()`. |
| Infer workbook roles from formula references. | `model_map()`, `ModelMap` | Inspection | No | Remove. “Referenced literal = input” and “unreferenced formula = output” are naive semantic claims. Blank inputs are omitted, structured references can remain unresolved, and the feature has almost no direct eval use. Keep the separate internal dependency sketch used for preservation safety. |
| Surface likely hardcodes, formula inconsistencies, hidden content, volatility, and magnitude outliers. | `findings()`, `Finding` | Inspection | No | Remove. Its blessed-literal list, three-row majority rule, and 1000x outlier threshold are application policy, not workbook facts. Keep objective `scan_errors()`. |
| Lint assigned formulas for suspicious functions or references. | `formula_lint`, `lint_formula()`, function catalog | Diagnostics | No | Remove. The hand-maintained semantic checker warns on legal UDFs and Paper's own transient rename state, while providing no demonstrated formula-quality benefit. Keep only parsing needed by preservation internals. |
| Infer which cells should be editable, then protect or format the workbook by role. | `protect_for_delivery()`, `apply_profile()` | Deliver | No | Remove. Workbook-role inference is not reliable enough to drive bulk mutation, and no organic eval task required this convenience layer. Explicit openpyxl protection and style APIs remain available. |
| Remove comments, metadata, personal properties, or hidden sheets through one delivery helper. | `scrub()` | Deliver | No | Remove. The only eval demand came from a synthetic task written to exercise this API. The operations already exist explicitly in openpyxl, and a catch-all helper adds policy rather than preservation. |
| Repoint a chart and edit loaded tables, validations, formatting, filters, comments, drawings, and related XML without rebuilding unrelated content. | `Chart.repoint()`, region and part editors | Structured edit | Yes | Keep, but fix stale caches, lexical normalization, and known XML edge cases. |
| Recalculate, certify, or evaluate an explicit source through optional LibreOffice tooling. | Module-level `openpyxl.oracle` APIs | Calculation and QA | Yes, optional | Keep lazy and explicit-source. Remove package-replacement writes and the misleading `Workbook.evaluate()` wrapper. |
| Declare a low-level dirty range or replace an arbitrary package part. | `mark_dirty()`, `replace_part()` | Escape hatch | No | Remove. A dirty declaration cannot validate unsupported mutation, and raw replacement can shadow planners or globally alter shared parts. |
| Request Excel to refresh selected preserved pivot caches when the workbook opens. | `set_pivot_refresh_on_load()` | Pivot correctness | Yes, targeted | Keep the job. Require named pivots or explicit `all=True`, record typed planner requests, and patch only selected cache-definition attributes. |
| Reject malformed or ambiguous ZIP/OPC structures before preservation. | Duplicate/case-colliding names, overlapping entries, header disagreement, truncation, declared-size mismatch, CRC, encryption, and compression-method validation | Package integrity | Yes | Keep. These checks identify packages that cannot be interpreted or preserved faithfully. |
| Reject otherwise valid workbooks when they exceed fixed resource thresholds. | Paper-defined ZIP entry, per-part byte, aggregate byte, compression-ratio, source-size, and stock-path caps | Package eligibility policy | No | Remove completely. These thresholds are not XLSX validity rules and must not affect either preserve mode or the upstream-compatible stock path. |
| Detect mixed installations where two distributions claim the `openpyxl` import namespace. | Fork sentinel and distribution ownership guard | Installation | Yes | Keep. The hard fork intentionally shares openpyxl's import namespace. |

## Missing product capability

This review found one user job that the current public surface cannot perform
safely. The other required work is repair or completion of features Paper has
already shipped.

| What the agent needs to do | Why the package does not cover it | Evidence | Decision |
| --- | --- | --- | --- |
| Replace one existing worksheet image while keeping its size, position, and every other image unchanged. | `replace_part()` performs a global media-part replacement. If several shapes share that part, all of them change. The package has no relationship-aware single-image operation. | All six Harbor `replace_part()` image attempts missed the isolated-image contract. | Add `Worksheet.replace_image(target, replacement, *, name=None)` before open source. It must allocate a fresh media part and retarget exactly one proven relationship. |

Not a missing product job: composing a row/column shift with a chart edit,
rewriting every supported structural reference, and preserving lexical XML.
Paper already claims those behaviors. Their failures belong under defects.

## Defects in what Paper already added

These are not requests for a larger product. They are correctness, performance,
compatibility, or API-boundary problems in the current Paper implementation.

| Existing Paper feature | Defect today | Change before open source |
| --- | --- | --- |
| Atomic cell mutation | Every bind snapshots growing ledgers and registries, producing sharply superlinear write cost. | Replace whole-ledger snapshots with a cell-local undo journal and nested savepoints. Restore exact target-local deltas on failure. |
| Atomic `Worksheet.append()` | Every append copies `self._cells`; preserve mode also snapshots workbook structural state. Paper work leaks into the stock path. | Use an append-local journal. Roll back only append-owned coordinates, supplied-cell bindings, and `_current_row`; refuse unsafe same-sheet re-entrancy. |
| Loaded-object change detection | Unstable serializers can be skipped, and image fingerprint errors can become `None`, allowing successful in-memory edits to disappear at save. | Produce a stable canonical fingerprint or refuse with a typed untrackable-object error. Never silently skip comparison. |
| Sheet rename | Paper rewrites an internal hyperlink during rename, then the save planner refuses that same change. | Make hyperlink locations rename-owned plan entries and compose them with formula, name, chart, and print-region rewrites atomically. |
| `model_map()` inference | It classifies only populated cells, cannot resolve every formula form, and equates “non-formula cell referenced by a formula” with “input.” It also calls an unreferenced formula an “output” and an unreferenced literal a “constant,” none of which can be proved from dependency structure alone. For `SUM(Sales[Amount])`, it reports unresolved references and no inferred inputs; blank intended inputs never appear. | Remove `Workbook.model_map()`, `ModelMap`, and the now-unused model-map builder. Retain the lower-level conservative dependency sketch used internally for cache invalidation, structural safety, and refusal. |
| Inference-backed protection and profiles | `protect_for_delivery()` and `apply_profile()` turn the heuristic model map into bulk workbook mutation. The dedicated delivery benchmark required the feature by construction; the other Harbor tasks did not. | Remove both APIs before the first public release, along with the model-map layer. Use explicit stock protection/style operations when requested. |
| Catch-all delivery scrubbing | `scrub()` bundles comments, metadata, personal properties, and hidden-sheet deletion into Paper policy. Its benchmark was also constructed specifically around the helper, and the observed metadata path could still serialize `creator="openpyxl"` after attempting to clear personal authorship. | Remove `scrub()`. Callers can make each requested change explicitly through workbook properties, comments, and sheet operations, with normal preservation checks. |
| Workbook findings | The API turns subjective thresholds and layout assumptions into first-party findings: blessed literals, a three-row majority formula pattern, a 1000x outlier rule, and blanket hidden/volatile warnings. | Remove `findings()`, `Finding`, and their public exports. Retain objective error scanning. |
| Nearby-cell targeting | `locate()` searches only up to six cells right or below a label; `set_input()` mutates the guessed result. Common layouts can make that target unrelated or ambiguous. | Remove both. Use explicit coordinates or workbook defined names. |
| Table-row append | The current helper is hidden, lacks a complete rollback transaction, imports the formula linter, and copies a formula from the preceding row even when the table column declares no calculated-column formula. Generic `Worksheet.append()` does not replace its table-range, totals-row, filter, and calculated-column behavior. | Publish `Worksheet.append_table_row()`. Plan the entire table mutation, commit through a table-local undo journal, use only declared calculated-column formulas, and remove linter dependencies. |
| `Chart.repoint()` | The series formula changes while `numCache` or `strCache` can retain data from the old range. | Remove the matching cache in the same chart patch and report it as a derived effect. |
| Package-part planning | Two supported operations can be refused merely because they touch the same XML part, even when their exact nodes do not conflict. | Build one composable plan per part with exact-node conflict detection. This fixes structural-shift plus chart-edit failures. |
| Region editors | A narrow validation, formatting, filter, or table edit can reserialize the whole element and normalize omitted defaults, ordering, prefixes, or extensions. | Use token-level patches that preserve unrelated lexical XML. |
| Formula linting | The hand-maintained catalog and partial reference analysis cannot reliably judge UDFs, add-ins, future functions, external names, structured references, or dynamic formulas. It also warns on Paper's own transient rename state. | Remove the entire public linter, mode, bind hook, warning, and catalog. Retain only formula parsing used by preservation. |
| Public API boundary | `openpyxl.preserve.__all__` exports ledger/scanner/saver internals. `Workbook.remove()` returns an inaccurate fork-specific report instead of upstream-compatible `None`. | Add public-surface snapshots; remove internal exports and `RemovalReport`; restore `Workbook.remove()` return behavior. |
| Structural safety | Preflight does not centrally enumerate every coordinate- and formula-bearing surface. Some operations, including `unmerge_cells()`, can bypass the claimed contract. | Add a closed-world supported-surface registry and operation matrix. Every affected surface must rewrite correctly or refuse before mutation. |
| `preserve=False` compatibility | Paper snapshots, guards, warnings, parser policy, and structural behavior can still affect workbooks after preservation is disabled. | Gate every Paper hook on an armed ledger and add differential tests against `paper-base`. |
| XML part editors | Known cases involving rich text, self-closing cells, CDATA/namespaces, x14 twins, drawing rIds, table IDs/formulas, VML, and chart extensions can be missed or normalized. | Fix each known editor hole and add positive and adversarial producer fixtures. Remove broad exception fallbacks and substring matching. |
| Raw mutation escape hatches | `replace_part()` can shadow planner output or globally alter shared parts. `mark_dirty()` cannot prove an unsupported mutation is safe merely by adding ledger coordinates. | Remove both public APIs. Keep part ownership and replacement state internal to dedicated closed-contract planners such as isolated image replacement. |
| Pivot refresh helper | The current method rewrites every pivot cache through generic raw replacement state, even when the caller needs one pivot. | Require explicit pivot names or `all=True`; resolve names to cache parts, record typed requests, patch only root `refreshOnLoad`, disclose shared-cache effects, and compose through package-part ownership. |
| Oracle evaluation and write-back | `recalc(in_place=True)` can broadly replace the package; `Workbook.evaluate()` silently evaluates retained source bytes rather than unsaved live edits. | Make `write_back()` the only cache mutation path, remove in-place recalc, remove the workbook wrapper, and keep explicit-source module evaluation. |
| Save receipts | Receipts do not distinguish requested edits from automatic cache removal, calc-chain removal, recalculation metadata, relationship, and content-type effects. | Add versioned `derived_effects`, produced from the actual plan and delivered archive without a second archive build. |
| `validate()` and the save planner | `validate()` duplicates archive construction, and large saver/scanner/structural functions mix responsibilities. | Share one immutable pure plan between validation and save; split large functions by ownership without changing output. |
| Fork hygiene | Five upstream files differ only in formatting or docstrings, increasing sync and review cost without changing behavior. | Restore those files byte-for-byte to `paper-base`. |
| Preserve-default source classification | Valid OOXML in a file-like object can miss preservation because its filename is absent or misleading. | Sniff seekable file-like content. Explicit `preserve=True` or `False` remains authoritative. |
| Paper resource caps | Paper added fixed entry, per-part byte, aggregate byte, compression-ratio, 512 MiB source, and 2 GiB stock-path caps. They reject otherwise valid workbooks and are not OOXML validity rules. | Remove every cap everywhere, including the stock path. Do not replace them with configurable package-eligibility thresholds. |

## Why these defects made the cut

The Harbor export is useful for API and artifact evidence, not as a clean
package-only score comparison because the Paper treatments also changed the
spreadsheet skill.

| Observed evidence | What it supports |
| --- | --- |
| Runtime-only preservation passed 727/730 checks versus 542/730 for the bare runtime; package integrity was 390/390. | Keep preserve-by-default and the automatic preservation spine. |
| Formula-cache repair recovered 12 binary task passes and lost none; one task moved from 60/149 to 149/149 rubric checks. | Keep automatic stale-cache invalidation. |
| 27 trajectories hit the structural-shift/chart-edit same-part refusal. | Fix planner composition; do not make agents save/reload around it. |
| Six image replacement attempts used `replace_part()` and all missed the isolated-image contract. | Add the safe image operation and remove raw replacement. |
| 22 of 24 Paper outputs on the revenue-row task normalized unrelated validation defaults. | Replace whole-element region serialization with lexical patches. |
| Default lint warned during Paper's own rename flow, with no demonstrated formula-quality gain. | Remove the semantic linter. |
| Protection and scrubbing appeared only in a synthetic task whose prompt explicitly required both operations; no other Harbor task required workbook locking. | Remove the delivery convenience layer rather than treating benchmark-by-construction usage as product demand. |
| `model_map()` appeared 13 times in the full arm and 4 times in the light arm, versus 12 and 3 protection calls respectively. | Almost all observed use was caused by the delivery helper; direct demand was approximately one trajectory per arm. Remove the public semantic classifier. |
| `locate()` appeared once; `set_input()` and the hidden `append_row()` had no calls. `replace_part()` and `mark_dirty()` appeared 7 and 4 times, but the demonstrated raw-part use was incorrect. | Remove guessed targeting and false-safety escape hatches. Do not infer that table expansion is redundant from a hidden API's zero use: code review proves it is a distinct job, so promote and harden it. |
| Paper-full agents used receipts in 162/180 trajectories and `validate()` in 129/180; light-skill agents used them far less with comparable outcomes. | Keep the APIs, but make them cheap and optional rather than mandatory rituals. |

## Exact pre-open-source change cut

### Add

- `Worksheet.replace_image(target, replacement, *, name=None)`.
- `Worksheet.append_table_row(table_name, values)` as the supported public form
  of the existing table-expansion job.
- Cell-local and append-local undo journals with nested savepoints.
- A table-local undo journal covering cells, totals, styles, comments,
  hyperlinks, table/filter references, registries, caches, and ledger deltas.
- Typed pivot-refresh request state and a planner-owned root-attribute patch.
- Stable fingerprints or typed untrackable-object refusals.
- Exact-node package-plan conflict detection and shared part ownership.
- A closed-world structural-surface registry and operation matrix.
- Receipt `derived_effects`.
- Content sniffing for seekable file-like sources.
- Public API snapshots, upstream differential tests, failure injection,
  scaling benchmarks, adversarial XML fixtures, and large-valid-package tests.

### Edit

- Make cell and append rollback proportional to the mutation while retaining
  the change-nothing-on-failure guarantee.
- Compose sheet rename with internal hyperlink rewriting.
- Make `Chart.repoint()` clear the corresponding cached values.
- Compose supported non-conflicting edits that share a package part.
- Make region editors patch only the intended tokens.
- Replace hidden `append_row()` with atomic `Worksheet.append_table_row()`;
  retain totals/filter expansion and declared calculated columns, but remove
  preceding-row formula inference and formula-linter hooks.
- Require `set_pivot_refresh_on_load(pivots=[...])` or explicit `all=True`;
  target resolved cache definitions, preserve all unrelated pivot XML, and
  report shared-cache and receipt effects.
- Make structural operations closed-world: rewrite or refuse before mutation.
- Make every `preserve=False` fork point match `paper-base` behavior.
- Fix the enumerated XML editor edge cases.
- Make `write_back()` the only oracle cache-write path; keep evaluation as an
  explicit-source module operation.
- Generate validation and receipts from the same pure save plan.
- Split the large preservation functions by ownership without changing output.
- Classify file-like sources by their actual container.
- Make `Workbook.remove()` return upstream-compatible `None`.

### Remove

- Whole-ledger, whole-sheet, and whole-workbook mutation snapshots.
- Skip-on-instability object comparison and swallowed fingerprint failures.
- Refuse-after-mutate sheet rename behavior.
- Warn-and-continue inferred delivery mutation.
- `Workbook.protect_for_delivery()` and inference-backed `apply_profile()`.
- `Workbook.scrub()` and its catch-all delivery policy.
- `Workbook.model_map()`, `ModelMap`, and the unused model-map builder. Keep the
  separate internal dependency sketch.
- `findings()`, `Finding`, and their subjective workbook-hygiene policies. Keep
  objective `scan_errors()`.
- `formula_lint`, `lint_formula()`, `lint_on_bind()`, `LintWarning`, linter
  modes/hooks, and the hand-maintained formula-function catalog. Keep internal
  preservation parsing.
- `Worksheet.locate()` and `Workbook.set_input()` guessed-target helpers.
- The hidden module-level `append_row()` entry point after its behavior moves to
  `Worksheet.append_table_row()`; do not remove the table-expansion capability.
- `Workbook.mark_dirty()` and `Workbook.replace_part()` raw mutation escape
  hatches. Keep planner-owned replacement state internal.
- Untargeted default-all pivot refresh and its use of generic
  `replaced_parts`; keep the targeted planner-owned operation.
- `Workbook.evaluate()`. Keep module-level explicit-source oracle evaluation.
- Retained stale chart caches after repointing.
- Blanket same-package-part conflict refusal.
- Whole-element normalization for narrow supported region edits.
- `DirtyLedger`, `save_preserved`, `scan_archive`, and `LossInventory` from the
  public preserve namespace.
- Public/inaccurate `RemovalReport` behavior.
- `oracle.recalc(in_place=True)` and generic LibreOffice package replacement.
- Duplicate archive construction in `validate()` and receipt generation.
- Every Paper-defined ZIP/source eligibility cap, including the stock-path cap.
- Formatting/docstring-only forks in `formula/translate.py`, `styles/fills.py`,
  `worksheet/_write_only.py`, `worksheet/filters.py`, and
  `worksheet/header_footer.py`.

## What stays unchanged

- The hard fork and shared `openpyxl` import namespace.
- Preserve-by-default for editable OOXML.
- Source-package retention, mutation tracking, XML splicing, untouched-part
  copying, typed refusal, and atomic delivery.
- Formula-cache invalidation and recalculation metadata updates.
- Package/cell diff, validation, receipts, and objective error scanning.
- Deterministic helpers `allowed_values()`, `search()`, and `copy_format()`.
- Atomic table-aware row append under `Worksheet.append_table_row()`.
- Explicitly scoped pivot refresh-on-load through the preservation planner.
- Optional explicit-source oracle paths that do not replace the package.
- No new facade, distribution split, source abstraction, or general-purpose XML
  editing API.

## Release verification

- The complete non-LibreOffice suite passes: **3,259 passed, 23 skipped,
  25 deselected, and 7 expected failures**.
- Cell-write scaling is linear over the measured range. Preserve mode took
  0.0158 s, 0.0310 s, and 0.0632 s for 5k, 10k, and 20k writes, versus
  0.0062 s, 0.0123 s, and 0.0253 s on the stock path. The preserve overhead
  stayed near 2.5x instead of rising with workbook size.
- Append scaling is also linear. Preserve mode took 0.0644 s, 0.1266 s, and
  0.2543 s for 5k, 10k, and 20k appends, versus 0.0237 s, 0.0472 s, and
  0.0950 s on the stock path. The overhead stayed near 2.7x.
- Release-cut tests cover target-local rollback (including `BaseException`),
  append re-entrancy, validate-without-archive, file-like sniffing, isolated
  image replacement, receipt effects, chart cache removal, rename/hyperlink
  composition, omitted validation defaults, and extension-bearing tables.
- The local LibreOffice executable aborts even for the untouched smoke fixture.
  Oracle and independent-loader tests must therefore be rerun in the known-good
  LibreOffice CI environment; the local abort is not counted as a package pass
  or failure.

Real-producer fixtures from desktop Excel, LibreOffice, and Google Sheets, plus
fixed-identity evals for image replacement, structural edits,
rename/hyperlinks, x14/shared formulas, macros, and custom XML remain useful
release confidence checks, but they are not substitutes for the automated
correctness suite above.

Package correctness, model behavior, treatment compliance, and evaluator
validity must be reported separately.
