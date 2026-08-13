# paper-xlsx code hardening specification

- Status: Proposed
- Scope: Runtime code in `v0.1.3` compared with `paper-base`
- Documentation: Deliberately deferred
- Target: Correctness and compatibility fixes in the next `0.1.x`; internal
  cleanup before `0.2`

## Decision

Keep the hard fork, but make it smaller in behavior and clearer in ownership.

The fork is justified by the preservation spine:

- retain the source OOXML package;
- record direct object-model mutations;
- splice supported edits into original XML;
- copy untouched package parts without regeneration;
- invalidate stale formula caches;
- refuse unsupported edits before mutation; and
- deliver output atomically.

Those guarantees require hooks in `Cell`, style descriptors, worksheet
structural methods, workbook lifecycle methods, the reader, and the writer. A
wrapper around upstream openpyxl could not reliably observe direct assignments
such as `cell.value = ...`, `cell.number_format = ...`, sheet renames, or row
insertion.

The current implementation is not ready to call complete. Four defects are
confirmed and release-blocking:

1. preserve-mode cell assignment becomes quadratic as dirty cells accumulate;
2. `Worksheet.append()` becomes quadratic even with `preserve=False`;
3. `Workbook.protect_for_delivery()` can lock real inputs when dependency
   analysis is incomplete; and
4. `Chart.repoint()` changes the source range while knowingly retaining stale
   chart caches.

The general-purpose benchmark does not prove that the fork improves agent task
scores. It does prove that automatic preservation works even when agents never
call Paper-specific helpers. The implementation should therefore prioritize
the automatic preservation spine over a wider agent-convenience API.

## Review method and baseline

This review compared:

- `paper-base`: `c4986390b313b3770ec60957057d2cca57a57eb2`;
- `v0.1.3`: `39675b59dcbb7fc2097d0b8d078b93f4feb1613c`; and
- the current documentation stack base: `gavin/docs-readme`.

The production-code delta from `paper-base` to `v0.1.3` is 59 files,
16,493 inserted lines, and 145 deleted lines. The review covered every modified
upstream runtime file and every added runtime module. It also inspected the
commit sequence from the initial bootstrap through cache invalidation,
manifest removal, preserve-by-default, and the `0.1.3` release.

The following checks were used:

- direct diff and commit-history review;
- public class, function, and signature comparison;
- AST comparison after stripping docstrings;
- targeted reproductions for delivery protection and formula linting;
- scaling probes for cell assignment and row append;
- the Paper test suite; and
- retained Train evaluation reports and artifact replays.

The current local test run produced 748 passes, 24 failures, and 6 skips. All
24 failures invoked the locally installed LibreOffice, which aborted with exit
code 134 even for the untouched smoke-test fixture. This is an environment or
LibreOffice failure, not evidence that those 24 package behaviors pass. The
independent-loader and oracle gate is currently unverified and must be rerun in
the known-good LibreOffice CI environment.

## Fork inventory and decision

| Area | What changed from upstream | Decision |
| --- | --- | --- |
| Distribution import | Fork sentinel and same-namespace ownership guard | Keep |
| Reader | Preserve default, source retention, ZIP validation, content inventory, ledger arming | Keep; tighten file-like classification later |
| Cell/style hooks | Track values, types, styles, hyperlinks, comments, and protection writes | Keep; replace whole-ledger rollback snapshots |
| Workbook/worksheet lifecycle | Track sheets, names, regions, tables, charts, and structural edits | Keep only where the affected state is closed-world |
| Saver | Plan changes, splice XML, copy untouched parts, validate, and deliver atomically | Keep; split the 727-line planner without changing behavior |
| Formula caches | Remove stale caches and force recalculation metadata | Keep; this is directly eval-backed |
| Package diff/receipt | Compare package and cell effects | Keep; add a small derived-effects field |
| Formula linter | Lint every preserve-mode formula assignment by default | Keep as an opt-in diagnostic; default it off |
| Model map/findings | Heuristic workbook interpretation | Keep advisory and explicit; do not let uncertainty authorize mutation |
| Delivery helpers | Protection, scrubbing, formatting profiles | Keep explicit, but protection and role-based formatting must fail closed |
| Oracle | Optional LibreOffice recalculation and certification | Keep lazy and optional; do not move distributions now |
| Raw escape hatches | `mark_dirty()` and `replace_part()` | Keep advanced and guarded; no namespace move now |

## Severity summary

### Glaring issues

- Quadratic mutation rollback on normal cell writes.
- Quadratic and upstream-incompatible `Worksheet.append()` behavior.
- Destructive protection based on incomplete inference.
- Stale chart cache retention after range changes.

### Important but bounded issues

- Formula linting is default-on despite legal UDF false positives and no
  preservation requirement.
- The public preserve namespace exposes internal machinery as supported API.
- Structural-rewrite guarantees are broader than the enumerated reference
  surfaces the code can prove.
- Several core functions are too large to review safely.
- Five upstream files differ only in formatting or docstrings.

### Not currently defects

- `validate()` proves that preserve save planning succeeds. Its name is broad,
  but its implementation does what its current contract states.
- `Workbook.evaluate()` intentionally evaluates the as-loaded source. That is
  surprising, but changing it would be a compatibility break without evidence
  that callers need different behavior.
- Retaining source bytes in memory has a documented cap. A different source
  abstraction may help later, but no current measurement justifies redesigning
  it first.
- The shared `openpyxl` namespace has deployment costs, but it is what makes
  existing code run through the preservation spine. A second facade would add
  complexity without solving mixed-distribution imports.

## P0: required correctness and performance fixes

### P0.1 Make cell mutation rollback proportional to the edited cell

#### Confirmed failure

`openpyxl/cell/cell.py::_CellBindTransaction` snapshots every entry in:

- `ledger.cells`;
- `ledger.value_overwrites`;
- `ledger.cache_writes`;
- the full number-format registry; and
- the number-format index.

It does this before every preserve-mode cell assignment. The snapshot grows
with all prior writes, so a sequence of otherwise constant-time assignments is
quadratic.

Measured on a minimal loaded workbook:

| Writes | Preserve mode | Stock path | Slowdown |
| ---: | ---: | ---: | ---: |
| 5,000 | 0.183 s | 0.006 s | 29x |
| 10,000 | 0.918 s | 0.013 s | 72x |
| 20,000 | 3.205 s | 0.025 s | 127x |

The rising per-write cost is the important result; absolute time varies by
machine.

The current broad snapshot is also not fully accurate. `_CellBindTransaction`
retains `cell._style` by reference. Automatic date/time formatting mutates that
shared `StyleArray` in place. A failure injected after the ledger mark produced
this state:

```text
value/type restored: 1, "n"
style before: numFmtId=0
style after:  numFmtId=164
number-format registry restored to empty
cell.number_format: IndexError
```

The registry rollback removed format 164, but the cell still referenced it.
Whole-ledger copying is therefore both slow and insufficient as an atomicity
mechanism.

#### Required change

Replace whole-ledger snapshots with a cell-local undo journal. The journal must
record each mutation as it occurs and execute undo actions in reverse order on
failure. It must support savepoints or equivalent nesting because assigning a
hyperlink to an empty cell invokes the value setter inside the hyperlink
setter.

Journal only state the operation can change:

- the target cell's value and data type;
- the target cell's style **by value**, not by reference;
- hyperlink and comment identity, parent, and reference state, including state
  on caller-supplied objects;
- for each relevant ledger mapping, whether the worksheet key existed, the
  original container object, whether the coordinate existed, and its prior
  value when the mapping holds values rather than membership;
- the prior `formulas_changed` flag;
- the prior protection-warning membership for the target sheet; and
- the exact number-format registry delta when automatic date/time formatting
  can add an entry. Explicit number-format setters need the same local-journal
  rule at their style-descriptor chokepoint; they must not rely on a cell-bind
  transaction that they do not enter.

Rollback must preserve mapping and container identity. If the operation made a
new worksheet key, remove that key. If the key already existed, restore the
original set or dictionary in place rather than replacing or clearing
unrelated entries.

Number-format rollback must restore the registry list, `_dict`, and `clean`
state together and restore the cell's copied `StyleArray` before any caller can
observe it. It must be impossible to leave a cell pointing at a truncated
format index.

Treat `ledger.cache_writes` as a separate correctness decision, not incidental
snapshot state. A successful direct edit of a cell with a staged oracle cache
must either invalidate that coordinate's staged cache or refuse with an
intentional public error; it must not reach the current internal
dirty-cell/cache-write overlap refusal at save. If the edit later rolls back,
restore the prior coordinate-local cache entry exactly.

Rollback must restore those exact deltas. It must not copy unrelated dirty
coordinates, registries, sheets, or cache entries. A local journal is also
safer under unexpected re-entrancy because it cannot revert an unrelated edit
that occurred after the transaction began.

Do not weaken atomic refusal to gain speed.

#### Acceptance criteria

- Existing atomicity tests still pass.
- A refused mutation restores model and ledger state exactly.
- Inject failure after automatic datetime and time number-format registration;
  prove the cell style, registry list, registry index, and `clean` flag match
  their pre-operation values and `cell.number_format` remains readable.
- Inject failure at each mutation boundary for value, direct `data_type`,
  formula lint, protection warnings-as-errors, comment binding, hyperlink
  binding, and hyperlink-to-empty-cell nested value binding.
- Cover both pre-existing and newly created worksheet keys in `cells`,
  `value_overwrites`, and `cache_writes`; prove rollback preserves container
  identity and unrelated coordinates.
- Cover a pre-existing staged oracle cache entry on both successful edit and
  rollback paths.
- Run the atomicity matrix for `Exception` and representative `BaseException`
  interruption paths because the public setter currently promises rollback
  for both.
- 5k, 10k, and 20k value-write timings scale linearly within normal benchmark
  noise.
- A 20k-write preserve run is no more than 10x the `preserve=False` write loop
  before save. This is an initial guardrail, not a permanent performance goal.
- Add a benchmark test that fails on superlinear growth, not only a fixed wall
  clock threshold.

### P0.2 Remove whole-sheet snapshots from `Worksheet.append()`

#### Confirmed failure

`Worksheet.append()` copies `self._cells` before every call, regardless of
whether preserve mode is active. Under preserve mode it also captures the full
workbook structural state. Both snapshots grow with the sheet.

Measured for three-cell rows:

| Appends | Preserve mode | `preserve=False` |
| ---: | ---: | ---: |
| 250 | 0.043 s | 0.001 s |
| 500 | 0.139 s | 0.004 s |
| 1,000 | 0.515 s | 0.011 s |
| 2,000 | 2.316 s | 0.037 s |

This is also an upstream-compatibility regression because the stock path now
pays for Paper's rollback behavior.

#### Required change

- Restore the upstream fast path exactly when no armed Paper ledger exists.
- In preserve mode, validate the iterable and pre-built `Cell` bindings before
  committing the row where possible.
- Track only cells added by the current append, the prior `_current_row`, the
  prior bindings of caller-supplied `Cell` objects, and coordinate-local ledger
  deltas.
- On failure, remove only the cells added by the attempted append and restore
  those local values.
- Do not call `_capture_structural_state()` for a simple row append.

#### Acceptance criteria

- The `preserve=False` append implementation has no Paper snapshot or ledger
  work in its hot path.
- Preserve and stock append loops scale linearly through at least 20k rows.
- Invalid generators and cross-sheet `Cell` objects remain atomic.
- Existing append, ledger, style, and formula-lint tests continue to pass.

### P0.3 Refuse unsafe inferred delivery protection

#### Confirmed failure

`model_map()` records `conventions["unresolved_references"] = True` when it
cannot resolve all formula inputs. `protect_for_delivery()` ignores that flag,
unlocks only inferred inputs, locks every other populated cell, and enables
sheet protection.

For a table `Sales` over `A1:A3` and `=SUM(Sales[Amount])`, the current result
contains no inferred inputs, reports unresolved references, and locks the real
data cells. Blank intended input cells are also left locked because the method
only visits populated cells.

#### Required change

Use an explicit-or-proved contract:

```python
wb.protect_for_delivery(
    password=None,
    *,
    inputs=None,
    on_unresolved="refuse",
)
```

- Accept sheet-qualified single cells, ranges, and defined names in `inputs`.
- Resolve every explicit input before mutating anything.
- Allow explicit blank inputs and unlock them.
- If `inputs` is omitted, refuse when the model map reports any unresolved
  reference.
- Do not add a warn-and-continue mode.
- Build the complete plan before changing protection state.
- Preserve protection flags other than `locked`.

Even a fully resolved formula graph does not prove that every referenced
literal is meant to be user-editable. Keep inferred-input mode for
compatibility in `0.1.x`, but make explicit inputs the preferred long-term
contract.

`apply_profile()` has the same inference boundary at lower severity. It must
refuse role-based mutation when `model_map()` reports unresolved references,
unless the caller passes an explicit role-to-cell mapping.

#### Acceptance criteria

- Reproduce and refuse the structured-reference case.
- Explicit inputs work when formula analysis is incomplete.
- Explicit blank input cells remain unlocked after save and reopen.
- Bad and ambiguous targets leave every style and protection setting intact.
- Referenced constants are not silently unlocked when the caller supplies an
  explicit input list.

### P0.4 Remove or refresh chart caches when repointing

#### Confirmed failure

`Chart.repoint()` changes `numRef.f` and explicitly retains `numCache`. The
comment assumes Excel will reread the source cells, but embedded chart caches
are observable workbook state and other renderers may use them. The package
therefore contains two different series definitions after a successful call.

#### Required change

For the selected series:

1. validate the new range;
2. identify the exact formula node and cache node in the chart part;
3. update the formula and remove the corresponding cache in the same patch;
4. refuse before mutation if the mapping is ambiguous or the chart encoding is
   unsupported; and
5. record both the requested formula change and cache removal in the receipt.

Do not generate a new cache without recalculation evidence.

#### Acceptance criteria

- Saved chart XML contains the new range and no stale cache for that series.
- Title, category, secondary-axis, scatter, chartsheet, and multi-series cases
  patch only the selected series.
- Numeric character references and unsupported extension encodings refuse
  atomically.
- Excel and LibreOffice smoke tests open and render the result in the
  independent-loader CI job.

## P1: reduce default behavior and public surface

### P1.1 Make formula linting opt-in

Formula linting is not part of package preservation. It flags unknown
functions, including legal workbook UDFs and add-in functions. The warning
itself acknowledges that those functions can be valid, yet `formula_lint`
defaults to `"warn"` on every preserve-mode workbook.

Required change:

- change the default to `"off"`;
- keep `"warn"` and `"refuse"` as explicit modes;
- bypass formula parsing entirely in the off path; and
- retain direct `lint_formula()` for callers that want diagnostics.

This makes the fork less intrusive without removing the feature.

### P1.2 Define a public/internal API boundary before adding APIs

No public upstream symbol was removed between `paper-base` and `v0.1.3`.
Signature changes are keyword-only additions:

- `load_workbook(..., *, preserve=None)`;
- `Workbook.save(..., *, allow_formula_loss=False, receipt=False)`; and
- `save_workbook(..., *, allow_formula_loss=False)`.

The major compatibility differences are behavioral:

- editable OOXML loads preserve by default;
- preserve saves can raise typed refusals where upstream would save lossily;
- row/column structural methods can return `AddressRemap`;
- preserve-mode `data_only=True` saves refuse by default;
- source bytes remain retained until the workbook is released; and
- only one installed distribution may own the `openpyxl` namespace.

Keep those intentional differences.

The current `openpyxl.preserve.__all__` also advertises internal machinery:

- `DirtyLedger`;
- `save_preserved`;
- `scan_archive`; and
- `LossInventory`.

These are implementation components, not stable user operations. Mark them
internal now and remove them from `__all__` through a deprecation window. Do
not move modules or rename working APIs in this hardening release.

Keep these supported groups:

- core: `load_workbook(..., preserve=...)`, `Workbook.preserve`, save flags,
  typed errors, and `AddressRemap`;
- verification: package diff, cell diff, receipt, and error scanning;
- explicit helpers: `locate`, `allowed_values`, `search`, `set_input`,
  `copy_format`, and `append_row`;
- advisory helpers: model map, findings, formula lint, and role profiles; and
- optional oracle: recalc, certify, evaluate, batch evaluate, and write-back.

The removal of `Workbook.manifest()` between `0.1.2` and `0.1.3` was
intentional: it caused severe large-workbook latency and duplicated cheaper
targeted APIs. Because the package is pre-1.0, removal was defensible, but it
revealed that no compatibility policy exists. From this release forward:

- no public API removals in a patch release;
- deprecations warn for at least one minor release;
- a removed API must have a stated replacement or an explicit “no
  replacement” rationale; and
- the pinned-surface test must distinguish supported public API from internal
  implementation names.

### P1.3 Make structural rewrite coverage closed-world

The structural implementation is intentionally conservative and already
refuses many difficult constructs: arrays, data tables, dynamic and 3D
references, pivots, VML/comments, certain chart encodings, and unsupported
table interactions.

Its guarantee must be limited to the reference surfaces it enumerates and
tests:

- ordinary cell formulas;
- defined names;
- table ranges and modeled table formulas;
- conditional-formatting formulas;
- data-validation formulas;
- chart series formulas; and
- modeled worksheet regions and relationships.

Required code changes:

- centralize the supported-surface registry used by preflight and rewriting;
- make every supported surface produce an explicit plan entry;
- refuse known formula-bearing relationship or extension families that are
  not in the registry;
- do not imply that arbitrary custom XML strings are rewritten; and
- add tests proving an operation either rewrites every registered affected
  surface or leaves model, ledger, and bytes unchanged.

Do not block a workbook merely because it has unrelated unknown parts. The
refusal should be tied to a known affected relationship or formula-bearing
surface, not a coincidental `A1` string in arbitrary XML.

### P1.4 Report safety-derived save effects

The current receipt is an accurate before/after package diff, but it does not
separate requested edits from safety-derived edits. Formula-cache invalidation
is important enough to expose directly because it changed real evaluator
outcomes.

Add a small, versioned `derived_effects` list. Each entry needs:

- `kind`;
- affected part;
- affected cell or formula when applicable; and
- cause, such as `input_changed`, `formula_changed`, or `chart_repointed`.

Required initial kinds:

- `formula_cache_removed`;
- `chart_cache_removed`;
- `calc_chain_removed`;
- `recalculation_metadata_changed`; and
- relationship or content-type changes created by supported lifecycle edits.

Do not redesign the entire receipt schema in this release.

### P1.5 Make the core code reviewable

The main risk is concentrated in functions too large for reliable review:

- `_save_preserved`: 727 lines, about 174 branch points;
- `scan_sheet`: 350 lines, about 101 branch points;
- `apply_model_shift`: 172 lines, about 43 branch points;
- `shift_blockers`: 142 lines, about 43 branch points;
- `begin_move_range`: 138 lines, about 49 branch points;
- `lint_formula`: 144 lines, about 46 branch points; and
- `plan_workbook_xml`: 143 lines, about 47 branch points.

Refactor in behavior-preserving slices:

1. split save planning by sheet, workbook metadata, relationships, lifecycle,
   and delivery;
2. give each phase an explicit immutable plan result;
3. separate XML tokenization from worksheet semantic validation;
4. separate structural blocker discovery from rewrite-plan construction; and
5. keep one final commit phase after all preflight succeeds.

Do not combine this refactor with new feature support. Compare generated
packages across the existing corpus before and after each slice.

AST comparison found five modified upstream files with no executable change
after docstrings are stripped:

- `openpyxl/formula/translate.py`;
- `openpyxl/styles/fills.py`;
- `openpyxl/worksheet/_write_only.py`;
- `openpyxl/worksheet/filters.py`; and
- `openpyxl/worksheet/header_footer.py`.

Revert those diffs to the exact upstream fork-point content. Formatting-only
changes make future upstream comparison harder and have no package value.

### P1.6 Classify file-like sources by content

Path inputs can use the extension to choose preserve mode. File-like inputs
may have no meaningful name or a misleading suffix. The current default
disables preservation for a file-like object named `.xls` or `.xlsb` even if
its bytes are actually OOXML.

For seekable file-like inputs, sniff the container and OOXML content type while
restoring the caller's position. Use the content result for the default. Keep
an explicit `preserve=True` or `False` authoritative.

This is lower priority than the confirmed P0 defects.

## Changes not proposed now

The earlier draft included broader architecture ideas. This code review does
not support doing them now.

- Do not add a second top-level facade.
- Do not split the oracle into another distribution.
- Do not move advisory or raw APIs to new modules solely for aesthetics.
- Do not rename or remove `validate()` in this hardening release.
- Do not change `Workbook.evaluate()` from source-state to live-state
  semantics without a concrete caller and migration plan.
- Do not build a `PackageSource` abstraction before measuring memory and I/O
  on representative large workbooks.
- Do not add a receipt-v2 redesign beyond the narrow `derived_effects` field.
- Do not add new worksheet or workbook convenience methods before the current
  surface is classified and stabilized.

## Evaluation alignment

The evaluations support the preservation spine, not every added API.

### What the evidence supports

- In the dedicated preservation suite, the bare runtime passed 542/730 checks
  and paper-xlsx runtime-only passed 727/730. The runtime-only package-integrity
  score was 390/390.
- In 50 sampled runtime-only traces, agents did not explicitly call
  validation or receipt APIs. Automatic preserve-by-default still delivered
  the integrity gain. This is strong evidence for default preservation and
  weak evidence for expanding convenience APIs.
- Exact artifact replay after the formula-cache fix recovered 12 binary task
  passes and lost none. One task moved from 60/149 to 149/149 rubric checks.
  Cache invalidation is therefore a proven package mechanism.
- The removed manifest produced an 868 KB trace and more than 150 seconds of
  work on a 2.39-million-cell workbook. Its removal was correct.

### What the evidence does not support

- The broad SpreadsheetBench comparison did not show a package-only score
  win. Historical runs also mixed runtime, skills, agent behavior, and
  evaluator conditions.
- A package-plus-skill run cannot identify the package effect.
- A workbook that saves successfully is not thereby calculation-correct,
  render-correct, or task-correct.
- The production oracle example failed closed when source caches were absent
  and required manual XML repair. That supports bounded oracle claims, not
  automatic repair claims.

### Evaluation work required after P0

Run four explicit treatments on the same frozen task identities:

1. upstream openpyxl runtime, no Paper skill;
2. paper-xlsx runtime, no Paper skill;
3. upstream runtime with the same general spreadsheet skill; and
4. paper-xlsx runtime with that skill.

Report separately:

- task rubric score;
- package integrity;
- formula/cache integrity;
- refusal correctness;
- agent completion and tool errors;
- load, edit, save, and total runtime; and
- treatment compliance.

Add a mutation-heavy slice containing bulk writes, row appends, formulas,
tables, charts, pivots, and large workbooks. The current eval exposed formula
cache value; it was not designed to catch the quadratic mutation paths found
in this review.

## Test and release gates

### Required tests

- Preserve-mode no-op save across the frozen fixture corpus.
- Exact part-level change allowlists for each supported edit.
- Refusal atomicity for model, ledger, and destination bytes.
- Linear-scaling benchmarks for cell writes and appends.
- Upstream behavior tests with `preserve=False`.
- Protection tests with unresolved references and blank explicit inputs.
- Chart repoint tests that inspect formula and cache nodes.
- Public API snapshot and deprecation tests.
- Independent load in a known-good LibreOffice environment.
- Real-producer fixtures from desktop Excel, LibreOffice, and Google Sheets,
  including pivots/caches, external links, macros, rich text, drawings, and
  1904-date workbooks.

### Next `0.1.x` gate

Must complete:

- P0.1 delta-based cell mutation transactions;
- P0.2 append fast-path and local rollback;
- P0.3 fail-closed delivery protection;
- P0.4 chart cache invalidation;
- formula lint default off;
- regression and scaling tests; and
- a green independent-loader CI run.

### `0.2` gate

Must complete:

- supported/internal API classification and deprecation mechanics;
- structural supported-surface registry;
- narrow receipt derived effects;
- behavior-preserving split of the save planner; and
- reversion of formatting-only upstream diffs.

### `1.0` gate

Must complete:

- a published compatibility policy reflected in tests;
- real-producer corpus coverage;
- package-only evaluation with fixed task identities;
- measured performance budgets for load, bulk edit, append, validate, receipt,
  and save; and
- zero unresolved P0 or silent-inconsistency defects.

## Definition of done

This proposal is implemented when:

- every executable difference from upstream has an owner, test, and stated
  compatibility effect;
- normal value writes and appends are linear;
- no heuristic authorizes destructive mutation while reporting uncertainty;
- chart and formula source references cannot disagree with retained caches;
- preserve-off behavior stays close to the upstream fast path;
- internal implementation objects are not accidentally promised as stable
  public API;
- core planners are small enough to review by phase; and
- the release gates report package correctness separately from agent and
  evaluator behavior.
