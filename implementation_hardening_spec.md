# paper-xlsx implementation hardening specification

- Status: Proposed
- Target: P0 fixes in 0.1.x; contract and architecture changes before 1.0
- Scope: Current `main` implementation and the `gavin/docs-readme` documentation branch
- Related: `hardening_proposal.md` on `main`, the historical post-0.1.1 hardening plan that this specification says to preserve

## Executive decision

Keep the hard fork.

The fork is justified by the preservation spine: retained source bytes, the
mutation ledger, XML splicing, guarded structural edits, typed refusals,
formula-cache invalidation, package receipts, and atomic delivery. These
features need hooks inside openpyxl's object model and ownership of its save
path. A wrapper would be more fragile than the fork.

Narrow the product around that spine. Heuristic workbook interpretation,
delivery preparation, raw package escape hatches, and LibreOffice orchestration
must not share the same trust boundary or API prominence as preservation.

Do not publish stronger safety claims or declare 1.0 until the two confirmed
correctness defects in this specification are fixed:

1. `Workbook.protect_for_delivery()` can lock the actual input cells when
   dependency analysis is incomplete.
2. `Chart.repoint()` knowingly retains stale embedded chart caches after it
   changes a series reference.

The package should be described as a source-compatible openpyxl fork with
intentional semantic changes. It is not a strict behavioral superset.

## Goals

- Preserve the strong parts of the current design.
- Remove silent-success paths whose output is internally inconsistent or based
  on unproved inference.
- Separate deterministic preservation behavior from advisory agent behavior.
- Make API names and result types state exactly what they prove.
- Make receipts distinguish requested edits from safety-derived changes.
- Define a closed-world contract for structural reference rewriting.
- Keep the shared `openpyxl` namespace only with explicit operational controls.
- Build release claims from tested producer and feature coverage.

## Non-goals

- Do not implement a partial Excel formula engine.
- Do not claim that worksheet protection is a security boundary.
- Do not add support for legacy `.xls` or `.xlsb` editing.
- Do not make `paper-xlsx` satisfy another distribution's dependency on
  `openpyxl`; Python package metadata cannot express that relationship.
- Do not use a package-plus-skill evaluation as proof of package-only causality.
- Do not expand the public API while the existing contracts are being corrected.

## Review basis

This specification combines a direct repository audit with an independent
Claude Fable review of the verified code excerpts, reproductions, documentation
claims, and test evidence.

The repository audit covered current `main` at `ce57ba77`, the documentation
branch at `f715e2ec`, and the diff from the `paper-base` tag. At the time of the
review:

- the fork added about 31,000 lines across 173 files;
- the Paper-specific suite completed with 772 passed and 6 skipped;
- the full suite completed with 3,353 passed, 23 skipped, and 7 expected
  failures;
- the fixture corpus contained no desktop-Excel-authored workbook and still
  lacked Google Sheets exports, pivots and caches, real VBA, external links,
  real `.xlsb`, and 1904 date-system coverage.

The upstream-compatible test result is evidence of compatibility. It is not a
substitute for preserve-mode or real-producer coverage. Prior task evaluations
are used here only to identify practical mechanisms and failure modes.

## Design rules

The implementation must follow these rules.

1. **Preserve or refuse.** A supported edit must preserve every package detail
   outside its declared ownership. An unproved edit must refuse before mutation.
2. **Inference cannot silently authorize mutation.** Heuristics may advise.
   They may not lock, remove, rewrite, or expose content unless the caller
   confirms the inferred scope or the analysis is complete.
3. **Derived edits are first-class edits.** Cache removal, calculation metadata,
   relationship changes, and other safety work must appear in the receipt.
4. **Validation names must be narrow.** An API that proves saveability must not
   imply business, calculation, rendering, or Excel-compatibility validation.
5. **Current-object methods use current state.** A method on a live `Workbook`
   must not silently run against its original source snapshot.
6. **Coverage claims are closed-world.** The package must enumerate what it can
   rewrite or preserve and refuse when an operation reaches outside that set.
7. **Warnings are not success.** A warning is acceptable only for an explicit
   stock-mode compatibility path. Preserve-mode safety failures must refuse.

## Current implementation assessment

### Keep in the core

- Source-package retention and source identity binding.
- The dirty ledger and mutation transactions.
- Cell and package XML splicing.
- Exact raw copy of untouched ZIP parts.
- Atomic path delivery and concurrent replacement checks.
- Typed `PaperRefusal` errors and refusal atomicity.
- Formula-cache invalidation and recalculation metadata.
- The `data_only=True` formula-loss guard.
- Package and cell diffing.
- Edit receipts.
- ZIP consistency and decompression guards.
- The distribution ownership guard and `paper-xlsx-doctor`.
- Structural editing only where the full affected reference set is proved.

### Keep, but change the contract

- Chart title and range editing.
- Structural reference rewriting.
- `Workbook.validate()`.
- `Workbook.evaluate()`.
- Preserve-mode classification for file-like inputs.
- Receipt schema and semantics.
- Large-source retention.

### Move out of the preservation trust boundary

- `Workbook.model_map()`.
- `openpyxl.preserve.findings()`.
- `Workbook.protect_for_delivery()` inference.
- `Workbook.scrub()` delivery policy.
- Oracle process orchestration.
- `Workbook.mark_dirty()` and `Workbook.replace_part()` raw escape hatches.

Deterministic helpers such as `locate()`, `search()`, `allowed_values()`, and
guarded `set_input()` may remain convenient public APIs. They must still use
honest result types and typed ambiguity failures.

## P0: correctness and public-contract fixes

### P0.1 Refuse unsafe inferred delivery protection

#### Current failure

`openpyxl/preserve/modelmap.py` records
`conventions["unresolved_references"] = True` when dependency analysis cannot
prove which cells are referenced. `Workbook.protect_for_delivery()` ignores the
flag and locks every populated cell not classified as an input.

A workbook containing a table `Sales` over `A1:A3` and the formula
`=SUM(Sales[Amount])` currently produces:

```text
inputs=[]
constants=[A1, A2, A3]
unresolved_references=True
A2_locked=True
A3_locked=True
```

The operation succeeds even though it has made the workbook unusable for its
intended input workflow.

#### Required API

Replace inference-only protection with an explicit or proved contract:

```python
report = wb.protect_for_delivery(
    password=None,
    *,
    inputs=None,              # iterable of sheet-qualified cells or names
    on_unresolved="refuse",   # only supported default
)
```

- If `inputs` is supplied, resolve and validate every target before mutation.
- If `inputs` is omitted, use `model_map()` only when it reports no unresolved
  references.
- If any reference remains unresolved, raise `UnsupportedStructureError` with:
  - `kind="unresolved-input-classification"`;
  - the unresolved formulas or constructs as evidence;
  - remedies to pass explicit inputs or remove the unsupported construct.
- Do not add a `warn-and-continue` mode.
- Return a versioned `ProtectionReport` rather than an unversioned dictionary.
  It must include the input source (`explicit` or `model-map`), locked and
  unlocked cells, unresolved count, sheet-protection settings, and a statement
  that OOXML protection is advisory.
- Permit explicit input ranges to include currently blank cells. Enabling sheet
  protection makes blank cells locked by default, so a blank cell intended for
  future user entry must be unlocked deliberately and reported.

#### Classification follow-up

Add structured-reference resolution to the dependency sketch. Cover table
names, column names, current-row references, header/data/totals qualifiers,
escaped column names, and cross-sheet table ownership. Unknown table syntax
must remain unresolved.

Do not assume every referenced literal is an editable input. `model_map()` may
continue to report that measurement, but delivery protection must require
explicit confirmation before unlocking inferred cells. The long-term default
should therefore become explicit `inputs=`.

#### Atomicity

Build the complete protection plan before changing a cell or enabling sheet
protection. A bad address, unresolved name, protected unsupported object, or
password failure must leave cell styles and protection metadata unchanged.

#### Tests

- Reproduce the `Sales[Amount]` failure.
- Cover all supported structured-reference forms.
- Refuse unknown, external, dynamic, and ambiguous references.
- Prove explicit `inputs=` works when inference is incomplete.
- Prove referenced constants are not unlocked without confirmation.
- Prove explicit blank input cells remain editable after save and reopen.
- Inject a failure at each commit boundary and compare serialized output.
- Save, reopen, and inspect protection settings in both openpyxl and
  LibreOffice.

#### Acceptance criteria

No protection operation may succeed with incomplete input classification.
Every unlocked cell must come from an explicit list or a complete, reported
classification accepted by the caller.

### P0.2 Remove stale chart caches when repointing a series

#### Current failure

`Chart.repoint()` changes `numRef.f` and deliberately retains `numCache`. The
reference and embedded cache can then describe different data. Desktop Excel
often refreshes the cache, but non-recalculating renderers, preview systems, and
package consumers may use the stale values.

This contradicts the package's formula-cache rule: stale cached results are not
valid evidence and must not be shipped as if they were current.

#### Required behavior

The default operation must never retain a cache tied to the old range.

```python
chart.repoint(series_index, new_range, *, cache_policy="drop")
```

Supported policies:

- `drop`: remove the matching `numCache` or `strCache`, patch the reference, and
  mark the workbook for full recalculation and chart refresh.
- `refresh`: future optional behavior that may rebuild the cache only from a
  trusted calculation result. Do not infer formula results from the live
  openpyxl model.
- No `keep` policy.

If the splice writer cannot prove that it can remove the exact cache without
changing sibling chart XML, raise `UnsupportedStructureError`. Do not fall back
to regenerating the whole chart part.

Handle numeric, string, category, x-value, and y-value references explicitly.
If only some references in a series are changed, remove only the caches made
stale by those references.

#### Receipt requirements

The receipt must report:

- the changed chart part;
- the old and new series references;
- each removed cache;
- the recalculation or refresh flag added to the workbook.

#### Tests

- Add a chart fixture with real embedded caches and nontrivial formatting.
- Repoint to shorter, longer, blank-containing, text, and formula-backed ranges.
- Cover line, bar, scatter, and category charts.
- Assert that old cache points and `ptCount` do not survive.
- Assert that sibling series, extensions, images, and drawing anchors remain
  byte-identical.
- Render or recalculate with LibreOffice where supported.
- Add a refusal test for an unsupported cache shape.

#### Acceptance criteria

After a successful repoint, no embedded cache may claim to represent the old
range. If that cannot be guaranteed, the operation refuses atomically.

### P0.3 Make receipts distinguish requested and derived changes

#### Current behavior

Changing an input such as `B1` can also invalidate a dependent formula cache in
`B3` and update `xl/workbook.xml`. The current receipt detects those package
changes, but the README example says only `B1` changed. The example is wrong,
and the receipt does not explain the cause of each observed change clearly.

#### Required schema

Introduce `edit_receipt` version 2 while preserving the existing raw diff:

```json
{
  "schema": "edit_receipt",
  "version": 2,
  "requested_changes": [],
  "derived_changes": [],
  "observed_package_diff": {},
  "recalculation": {
    "required": true,
    "reason": "dependent-cache-invalidated"
  },
  "artifact": {
    "before_sha256": "...",
    "after_sha256": "..."
  }
}
```

- `requested_changes` comes from committed ledger operations.
- `derived_changes` records cache invalidations, calculation-property changes,
  relationship updates, content-type changes, and other safety work.
- `observed_package_diff` remains the final byte-derived truth.
- Every requested or derived change must map to an observed package change.
- An unexplained observed change must make receipt creation refuse.
- Record whether the receipt base is the as-loaded source or a previous
  delivery, plus that base artifact's digest. The current cumulative
  as-loaded-source behavior may remain during 0.x, but it must never be
  implicit after multiple saves.
- Preserve a compatibility view for version 1 consumers during the 0.x series.

#### Tests

- Update the README quick-start example from an actual serialized receipt.
- Cover a value edit with dependent cache removal.
- Cover formula edits, style-only edits, chart-cache removal, structural edits,
  relationship changes, and no-op saves.
- Prove that an injected unexplained part change makes receipt creation refuse.

#### Acceptance criteria

A caller can tell what it asked for, what Paper changed for safety, and what
bytes actually changed. Documentation examples exactly match executable tests.

### P0.4 Correct the README and RST safety contract

Update `README.rst`, `doc/paper.rst`, and the proposed `README.md` on
`gavin/docs-readme` together.

#### Required factual corrections

| Current claim | Replacement |
|---|---|
| "strict-superset" | "source-compatible fork with intentional save and safety semantics" |
| "everything upstream works unchanged" | State the tested compatibility surface and list deliberate differences |
| "only two behavioral deltas" | Remove; document preserve default, refusals, cache invalidation, structural return values, stream constraints, memory behavior, and distribution conflict |
| "will not silently corrupt a real Excel file" | "designed to preserve or refuse within its tested contract" |
| "versioned payloads throughout" | List the APIs that return versioned payloads; do not generalize |
| Receipt shows only `B1` | Show `B1`, dependent `B3` cache invalidation, and workbook calculation metadata |
| `021192cf` is `paper-base` | State that `paper-base` is `c4986390`; describe `021192cf` as the preservation bootstrap commit |
| Every added operation has three safe outcomes | Limit the statement to operations covered by the declared Paper contract |

Also:

- Label the 15-task comparison as an internal package-plus-workflow evaluation.
- Do not imply package-only causality.
- Link a reproducible method or public artifact before quoting evaluation
  conclusions.
- Describe worksheet protection as advisory.
- State that `validate()` proves only saveability until it is replaced.
- State that formula-cache invalidation can make `data_only=True` return `None`
  until a trusted engine recalculates.
- Remove the statement in `modelmap.py` that `set_input()` consumes the model
  map, unless the implementation is deliberately changed to make that true.
- Use PyPI-safe absolute logo URLs or prove relative asset rendering in the
  built long description.
- Preserve `hardening_proposal.md` as design history. Do not delete it as part
  of README cleanup.
- Treat the Python 3.9 floor as a separate support-policy decision. Do not hide
  it inside documentation cleanup.

#### Acceptance criteria

Every public claim maps to a current test, an explicit limitation, or a linked
evaluation artifact. README examples are executed in CI.

## P1: contract and architecture changes

### P1.1 Replace `validate()` with an honest save-validation API

`Workbook.validate()` performs a preserve save into `BytesIO` and returns
`None`. It proves that the current preserve saver can emit a coherent package.
It does not prove formulas, cached values, charts, business logic, rendering,
or Excel compatibility.

Add:

```python
report = wb.validate_save()
```

`SaveValidationReport` version 1 must include:

- whether the current state can be serialized;
- planned and derived changed parts;
- refusal checks performed;
- cache invalidations and recalculation requirement;
- artifact size and digest for the staged candidate;
- checks not performed, including calculation, rendering, and business logic.

The implementation should expose the staged candidate internally so a caller
that immediately saves can commit the exact validated bytes after rechecking
source and destination custody. Do not serialize twice.

Deprecate `validate()` as an alias during 0.x, with a message that directs users
to `validate_save()`. Remove it before 1.0 unless compatibility data justifies
retaining the alias.

### P1.2 Make workbook evaluation use current state

The current `Workbook.evaluate(set=..., read=...)` runs against the preserved
as-loaded bytes and ignores unsaved edits. This is surprising for a method on a
live workbook and can return answers for the wrong model.

Change the default contract:

```python
evaluation = wb.evaluate(assignments={...}, read=[...])
```

- Stage the current workbook through the preserve saver.
- Run the oracle against those exact staged bytes.
- Do not modify the live workbook or original path.
- Include the staged artifact digest in the evaluation result.
- Rename `set` to `assignments`.

Preserve the old behavior only as an explicit method:

```python
evaluation = wb.evaluate_source(assignments={...}, read=[...])
```

Tests must prove that unsaved edits affect `evaluate()` and do not affect
`evaluate_source()`.

### P1.3 Define a closed-world structural rewrite contract

Before any row, column, range, table, or sheet-name mutation, build one bounded
inventory of every reference-bearing construct that the operation can affect.

The capability matrix must explicitly classify:

- worksheet formulas;
- array, shared, spill, and data-table formulas;
- global and sheet-scoped defined names;
- print areas and print titles;
- tables and structured references;
- chart series, categories, x values, y values, and caches;
- data validation formulas;
- conditional-formatting formulas;
- sparklines and x14 extensions;
- pivot definitions and caches;
- external links;
- calculation chains;
- three-dimensional, dynamic, indirect, and opaque references.

Each class must be one of:

- `rewrite`: parsed, transformed, and verified;
- `unaffected`: proven outside the mutation scope;
- `refuse`: present and potentially affected, so the operation cannot proceed.

There is no `ignore` state.

Return the capability decision and affected constructs in `AddressRemap` or a
linked versioned structural-edit report. Re-parse the affected constructs in
the output package and prove that each old in-scope reference was rewritten or
was correctly classified as unaffected.

### P1.4 Classify file-like sources by content

Preserve selection for file-like objects must not depend on the `.name` suffix.

- Save and restore the stream cursor.
- Read a bounded prefix.
- Detect OOXML ZIP structure from content.
- Validate `[Content_Types].xml` and the workbook relationship before enabling
  preserve mode.
- Refuse legacy or unsupported content with a typed error.
- Keep `read_only=True` behavior explicit.

A stream named `model.xls` that contains a valid OOXML package must not silently
fall back to the stock save path.

### P1.5 Replace full in-memory source retention with a source abstraction

Preserve mode currently keeps a complete source archive in memory and applies a
512 MiB source cap. This adds a second large representation next to openpyxl's
object model.

Introduce an immutable `PackageSource` abstraction with:

- artifact digest and size;
- source identity and custody information;
- random access to bounded part data;
- an in-memory backend for small byte inputs;
- a spooled or temporary-file backend for large streams and paths;
- deterministic cleanup;
- no dependence on the original path after load.

Keep decompression-bomb and part-size limits separate from source-storage
policy. Do not expose a broad set of tuning flags. Start with an internal spool
threshold and measure it against representative files before making it public.

### P1.6 Formalize the shared-namespace operating model

The fork and the shared import namespace are separate decisions. Keep the
`openpyxl` import for the current product because transparent compatibility with
existing code and pandas is a core use case. Do not describe the arrangement as
equivalent to a normal distribution/import-name split: both `paper-xlsx` and
`openpyxl` own the same files.

Required controls:

- Document that Paper must run in a controlled virtual environment, container,
  or equivalent isolated runtime.
- Run `paper-xlsx-doctor` in deployment and eval environment setup.
- Add installation-order tests for stock then Paper, Paper then stock,
  uninstalling either distribution, editable installs, and wheel installs.
- Add dependency-resolution tests with pandas and another package that declares
  `Requires-Dist: openpyxl`.
- Fail early when both distribution metadata records are present.
- Publish a short operational decision record describing the tradeoff.

Before 1.0, prototype a separate `paper_xlsx` import facade and measure the
compatibility loss. Do not switch namespaces without that evidence. A separate
namespace is the fallback if shared ownership cannot be made operationally
reliable.

## P2: API boundaries and product shape

### P2.1 Separate core, advisory, delivery, oracle, and raw APIs

Organize public APIs by trust level:

```text
openpyxl.preserve    deterministic preservation and save contracts
openpyxl.agent       advisory inspection and heuristic interpretation
openpyxl.delivery    explicit delivery policy and protection reports
openpyxl.oracle      optional external calculation and certification
openpyxl.advanced    raw package and ledger escape hatches
```

Recommended moves:

| Current API | Target |
|---|---|
| `Workbook.model_map()` | `openpyxl.agent.model_map(wb)` |
| `findings(wb)` | `openpyxl.agent.findings(wb)` |
| `Workbook.protect_for_delivery()` | `openpyxl.delivery.protect(wb, ...)` |
| `Workbook.scrub()` | `openpyxl.delivery.scrub(wb, policy=...)` |
| `Workbook.mark_dirty()` | `openpyxl.advanced.mark_dirty(wb, target)` |
| `Workbook.replace_part()` | `openpyxl.advanced.replace_part(wb, name, payload)` |
| Oracle functions | Keep in `openpyxl.oracle`, make support explicitly optional |

Use thin deprecated method shims during the 0.x migration. The core Workbook
class should stop accumulating policy and heuristic methods.

### P2.2 Make advisory results expose uncertainty

Every advisory result must contain:

- schema and version;
- evidence used;
- unresolved or unsupported constructs;
- confidence or completeness, where meaningful;
- a statement that the result is advisory;
- stable identifiers for cells and package parts.

Replace language such as "measurements, never judgments" where a heuristic does
classify or prioritize content. A classification is a judgment even when it is
evidence-backed.

### P2.3 Make raw escape hatches explicit

Raw part replacement and manual dirty marking are necessary for expert use, but
they can undermine the ledger if used incorrectly.

- Move them to `openpyxl.advanced`.
- Require preserve mode.
- Require exact target declarations.
- Record the caller assertion in the ledger and receipt.
- Validate the replacement part and package relationships before commit.
- Prevent a raw edit from producing a clean receipt with unexplained changes.

### P2.4 Keep the oracle optional and evidence-bounded

LibreOffice orchestration is useful but is not part of source preservation.

- Do not import or probe LibreOffice on normal workbook load or save.
- Define a clear availability check and version report.
- Separate baseline certification from recalculation and cache write-back.
- Never convert `BASELINE_UNVERIFIABLE` into implicit permission to write.
- Include engine version, input and output digests, excluded formulas, and
  certification state in every result.
- Consider a separate distribution only if the module creates dependency or
  release coupling; a separate namespace inside the current distribution is
  sufficient initially.

## Testing and evidence plan

### Test reporting

Report these suites separately:

1. Upstream compatibility tests.
2. Preserve-mode contract tests.
3. Refusal atomicity tests.
4. Producer-fixture round trips.
5. LibreOffice calculation and rendering tests.
6. Packaging and installation tests.
7. Performance and memory tests.

Do not use the full upstream-compatible test count as the headline evidence for
preserve-mode coverage.

### Required producer corpus before 1.0

Add legally shareable, frozen fixtures with provenance for:

- desktop Excel on supported Windows and macOS versions;
- Google Sheets exports;
- LibreOffice;
- pivot tables and caches;
- external links;
- real VBA projects;
- cached charts with complex formatting;
- shared, array, spill, and data-table formulas;
- 1904 date-system workbooks;
- comments, threaded comments, x14 validations, conditional formatting, shapes,
  text boxes, and `AlternateContent`;
- large and sparse workbooks.

Keep unsupported `.xlsb` fixtures for detection and refusal only unless product
scope changes.

### Evaluation design

Run package evaluation as a controlled comparison:

- same tasks, files, model, prompt, tools, timeouts, and grader;
- stock openpyxl plus stock workflow;
- Paper package plus the same stock workflow;
- Paper package plus Paper workflow as a separate treatment.

Measure:

- requested task correctness;
- unaffected-part equivalence;
- formulas and formula-cache state;
- chart references and cache state;
- relationship validity;
- renderability;
- refusal quality;
- completion rate and time;
- package and skill treatment compliance.

Treat prior production examples as mechanism evidence, not package-only causal
proof.

### Performance budgets

Establish budgets for:

- load time;
- save time;
- peak resident memory;
- temporary-disk use;
- receipt generation;
- structural-inventory scans;
- oracle startup and batch reuse.

Record both stock and preserve-mode results on small, large, and sparse files.
Regressions beyond an agreed threshold require an explicit release note.

## Delivery sequence

Implement this specification as focused pull requests.

1. **Protection refusal hotfix.** Guard unresolved classifications, add the
   structured-reference reproduction, and return a versioned report.
2. **Chart-cache correctness.** Drop or refresh stale caches and add cached-chart
   fixtures.
3. **Receipt v2 and executable documentation example.** Separate requested,
   derived, and observed changes.
4. **Documentation contract correction.** Update current RST and Gavin's README;
   keep Python-floor and repository-cleanup changes separate.
5. **Save validation and current-state evaluation.** Add `validate_save()`,
   `evaluate_source()`, and current-state `evaluate()`.
6. **Closed-world structural rewriting.** Publish the capability matrix and add
   refusal gates for uncovered constructs.
7. **Source classification and storage.** Content-sniff streams and introduce
   `PackageSource`.
8. **API boundary migration.** Add advisory, delivery, and advanced namespaces
   with deprecation shims.
9. **Packaging operating model.** Add resolution tests and the namespace
   decision record.
10. **Producer corpus and 1.0 evidence.** Fill the real-producer gaps, publish
    coverage, and run the matched evaluation.

Each pull request must start from a failing reproduction, make the smallest
coherent production change, and prove saved-and-reopened artifact behavior.

## Release gates

### Gate for the next 0.1.x release

- P0.1 protection failure fixed.
- P0.2 stale chart caches fixed or the operation temporarily refuses.
- Quick-start receipt example corrected.
- Absolute safety and strict-superset claims removed.
- Full upstream and Paper suites green.

### Gate for 0.2

- Receipt v2 shipped.
- `validate_save()` and current-state `evaluate()` shipped.
- Structural rewrite capability matrix enforced.
- File-like content classification shipped.
- Public APIs grouped by trust level, with migration warnings.

### Gate for 1.0

- Real Excel and Google producer corpus covers the advertised feature set.
- No unresolved P0 or P1 safety defects.
- Shared-namespace deployment model has repeatable installation evidence.
- Preserve-mode performance and memory budgets are published and met.
- Public safety claims map to named tests and producer fixtures.
- Package-only and package-plus-workflow evaluation results are separated.

## Definition of done

This proposal is complete when:

- no successful Paper operation knowingly leaves contradictory cached and live
  representations;
- heuristic uncertainty cannot silently drive a delivery mutation;
- every preserve-mode mutation has a declared ownership and refusal boundary;
- receipts explain requested, derived, and observed changes;
- validation and evaluation APIs operate on the state their names imply;
- raw, advisory, delivery, and calculation APIs are separated from the core
  preservation contract;
- documentation describes intentional incompatibilities and evidence limits;
- the supported producer and feature matrix is tested with frozen artifacts;
- the fork's maintenance and namespace costs are accepted explicitly rather
  than hidden behind "drop-in" language.
