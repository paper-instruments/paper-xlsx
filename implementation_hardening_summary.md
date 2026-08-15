# paper-xlsx pre-open-source implementation summary

- Status: Implemented on `agent/paper-xlsx-implementation-hardening-spec`
- Baseline: Paper v0.1.3, compared with upstream-derived `paper-base`
  `c4986390`
- Release target: paper-xlsx 0.2.0

This document records the completed release cut. “Added,” “changed,” and
“removed” below are deltas from Paper v0.1.3, not unfinished proposals and not
changes made by upstream openpyxl.

## Result

The hard fork remains justified by one coherent job: edit an existing XLSX
through ordinary openpyxl objects while retaining unsupported OOXML, rewriting
only supported structures, invalidating stale calculation caches, refusing
unsafe edits before delivery, and saving atomically.

The release keeps that preservation spine. It removes Paper APIs that guessed
workbook meaning or exposed unsafe mutation escape hatches, adds two narrow
operations required by real workbook tasks, and repairs correctness,
performance, compatibility, and reviewability defects in the retained code.

## Added in 0.2.0

| Addition | What it provides |
| --- | --- |
| `Worksheet.append_table_row(table_name, values)` | Atomically expands a supported named worksheet table, its filter, totals placement, inherited cell formats, and declared calculated columns. Loaded tables are checked against retained source XML and relationships; connected, extended, merged/spill, sorted, protected, ambiguous, or conflicting states refuse before mutation. A refusal is final and must not be bypassed. |
| `Worksheet.replace_image(target, replacement, *, name=None)` | Replaces one worksheet image by allocating a new media part and retargeting exactly one relationship, without changing other shapes that shared the old media part. |
| Targeted pivot refresh requests | `Workbook.set_pivot_refresh_on_load(pivots=[...])` or explicit `all=True` records planner-owned requests and patches only selected pivot-cache root attributes. |
| Local rollback journals | Cell assignment, generic append, and table-row append now restore mutation-owned state without copying whole ledgers, sheets, or workbooks. Nested savepoints preserve exact failure atomicity. |
| Composable part planning | Supported edits to distinct nodes in the same package part now compose; exact-node conflicts still refuse. |
| Stable object fingerprints | Loaded charts, tables, drawings, validations, and related objects are compared deterministically or produce a typed refusal instead of silently skipping change detection. |
| Closed structural-surface registry | Coordinate-changing operations enumerate supported reference-bearing surfaces and either rewrite each affected surface or refuse before mutation. |
| Receipt-derived effects | Version 2 edit receipts distinguish requested changes from automatic cache, calculation, relationship, and content-type effects. |
| Content-based file-like detection | Seekable streams are classified from their OOXML content when preservation is not explicitly enabled or disabled. |

## Changed from Paper v0.1.3

| Existing Paper behavior | Release behavior |
| --- | --- |
| Cell binding copied growing mutation ledgers and registries. | Cell-local journals make writes linear while restoring exact cell, cache, format-registry, warning, and ledger state on failure. |
| `Worksheet.append()` copied the whole cell map and, in preservation mode, workbook structural state. | An append-local journal records only appended coordinates, supplied-cell bindings, and `_current_row`; unsafe same-sheet re-entrancy refuses before mutation. |
| A serializer or fingerprint failure could make a changed object look unchanged. | Comparison is stable or fails with a typed untrackable-object error. |
| Sheet rename rewrote internal hyperlinks and could then refuse its own change. | Rename-owned formula, name, chart, print-region, and internal-hyperlink edits commit as one plan. |
| `Chart.repoint()` could leave cached values from the old range or mutate the model before discovering an unsupported patch. | Repointing preflights the complete prospective patch, changes the model only after success, removes the matching chart cache, and reports that automatic effect. |
| Edits sharing an XML part conflicted at part granularity. | Planner conflict detection operates on exact owned nodes. |
| Narrow table, validation, filter, and formatting edits could normalize unrelated XML. | Region editors patch the intended lexical tokens and preserve omitted defaults, ordering, prefixes, and unknown children. |
| Structural preflight did not centrally cover every modeled reference surface. | A closed operation matrix rewrites all known affected surfaces or refuses before mutation. |
| Paper hooks could affect `preserve=False`. | The stock path bypasses Paper ledgers, snapshots, guards, warnings, parser policy, and preservation behavior. |
| Several XML editors used broad fallback or substring matching. | Rich text, self-closing cells, CDATA/namespaces, x14 twins, drawing relationships, table IDs/formulas, VML, and chart extensions have explicit handling and adversarial coverage. |
| Oracle recalculation could deliver a full LibreOffice rewrite, while meaningful in-place cache write-back required an uncertified override. | `recalc()` is the single calculation API: without an output it returns evidence, and with a separate output it splices eligible caches into a Paper-preserved candidate while keeping full recalculation enabled. Evaluation remains an explicit-source module operation. |
| `allowed_values()` conflated absent and unsupported validation and could return formula text as an allowed input. | It returns `None` only for absent list validation, supports exact literals and static one-dimensional value ranges, and typed-refuses ambiguous or unsupported sources. |
| `scan_errors()` matched `#REF!` substrings inside formula text. | It uses the shared formula tokenizer and reports only actual error operands plus cached error values. |
| `copy_format()` could leave a partly changed range after failure. | It preflights merges and protection, then applies the complete style through a range-local rollback journal. |
| `validate()` and receipt generation repeated archive work. | Validation and save use the same pure preservation plan, and receipts inspect the delivered archive without rebuilding it. |
| Large planner and scanner functions mixed unrelated responsibilities. | Planning is split into named phases. Scanner semantic phases are separated while its measured hot loop remains intentionally inline. |
| Fixed Paper resource caps rejected otherwise valid packages. | All Paper entry-count, part-size, aggregate-size, compression-ratio, source-size, and stock-path caps are gone. Structural ZIP/OPC integrity validation remains. |
| Five upstream files differed only in formatting or docstrings. | Those files are restored byte-for-byte to `paper-base`. |
| `Workbook.remove()` returned a Paper-specific report. | It again returns upstream-compatible `None`. |

## Removed from Paper v0.1.3

These removals shrink the Paper public surface. They do not remove the
underlying explicit openpyxl operations for cell assignment, styling,
protection, workbook properties, comments, sheets, or defined names.

| Removed Paper API or behavior | Reason |
| --- | --- |
| `Workbook.model_map()` and `ModelMap` | Formula-reference structure cannot prove semantic roles such as input, output, or constant. The conservative internal dependency sketch remains for preservation safety. |
| `findings()` and `Finding` | Blessed literals, three-row majorities, magnitude thresholds, and blanket hidden/volatile warnings were subjective policy. Objective `scan_errors()` remains. |
| `Workbook.protect_for_delivery()` and `apply_profile()` | They converted unreliable role inference into broad protection and formatting mutations. Callers can use explicit stock protection and style APIs. |
| `Workbook.scrub()` | It bundled unrelated metadata, comment, property, and sheet-removal policy. Callers can perform each requested operation explicitly. |
| `Worksheet.locate()` and `Workbook.set_input()` | A six-cell neighborhood heuristic could select and mutate an unrelated cell. Explicit coordinates and defined names remain. |
| `formula_lint`, bind-time linting, modes, warnings, and the hand-maintained function catalog | The checker warned on legal UDFs, add-ins, future functions, and Paper’s own transient rename state without demonstrated formula-quality benefit. Preservation formula parsing remains. |
| `Workbook.mark_dirty()` | Ledger membership cannot prove an unsupported mutation is safe. |
| `Workbook.replace_part()` | Raw replacement could shadow planners or alter every consumer of a shared part. Dedicated relationship-aware operations own package changes instead. |
| Untargeted default-all pivot refresh | Refresh now requires pivot names or explicit `all=True`; the pivot-refresh job remains. |
| `Workbook.evaluate()` | The wrapper evaluated retained source bytes rather than unsaved workbook state. Explicit-source `openpyxl.oracle` evaluation remains. |
| `oracle.recalc(in_place=True)` | Broad package replacement conflicted with preservation custody. |
| `oracle.write_back()`, `WriteBackResult`, and `allow_uncertified` | The useful path overwrote the source using LibreOffice values precisely when cache equivalence was unproven. Separate-output preserved recalculation remains through `oracle.recalc()`. |
| `RemovalReport` | The fork-specific return value broke upstream `Workbook.remove()` compatibility. |
| `DirtyLedger`, `save_preserved`, `scan_archive`, and `LossInventory` public exports | These are implementation components, not supported user operations. The ledger and scanners remain internal where preservation needs them. |
| Hidden module-level `append_row()` | Its valid table-expansion job moved to public, atomic `Worksheet.append_table_row()` rather than being deleted. |
| Whole-ledger, whole-sheet, and whole-workbook mutation snapshots | Local journals provide the same change-nothing-on-failure contract without growing copy costs. |

## Deliberately retained

- The shared `openpyxl` import namespace and `openpyxl.__paper_version__`
  sentinel.
- Preserve-by-default for editable OOXML, with explicit `preserve=False` as the
  upstream-compatible path.
- Source-package retention, mutation tracking, XML splicing, untouched-part
  copying, typed refusal, and atomic delivery.
- Formula-cache invalidation and recalculation metadata updates.
- Package and cell diffing, `validate()`, edit receipts, and objective
  `scan_errors()`.
- Deterministic helpers `allowed_values()`, `search()`, and `copy_format()`.
- Atomic table expansion through `Worksheet.append_table_row()`.
- Explicitly scoped pivot refresh through the preservation planner.
- Optional explicit-source oracle recalculation, evaluation, and certification.
- ZIP/OPC integrity checks for duplicate or case-colliding names, overlapping
  entries, header disagreement, truncation, size or CRC mismatch, encryption,
  and unsupported compression methods.

## Eval evidence used to choose the cut

The Harbor treatments bundled package, skill, and harness changes, so they do
not establish a package-only score gain. They were used to find repeated API
use and artifact failures:

- Runtime-only preservation passed 727/730 package checks versus 542/730 for
  the bare runtime; package integrity was 390/390. This supports retaining the
  automatic preservation spine.
- Formula-cache repair recovered 12 binary task passes and lost none. This
  supports automatic cache invalidation.
- Twenty-seven trajectories hit the structural-shift/chart-edit same-part
  refusal. This supports exact-node plan composition.
- All six raw image-replacement attempts missed the isolated-image contract.
  This supports `replace_image()` and removal of `replace_part()`.
- Twenty-two of 24 outputs on the revenue-row task normalized unrelated data
  validation defaults. This supports lexical region patching.
- Formula lint warned during Paper’s own rename flow without demonstrated
  formula-quality benefit.
- Protection and scrubbing appeared only in a synthetic task written to require
  them. `model_map()` use was overwhelmingly induced by that helper.
- Low use did not determine removals: hidden `append_row()` had no calls, but
  code review established a real table-expansion job, so it became the hardened
  public `append_table_row()` operation.

## Verification

- The non-LibreOffice suite passes with 3,323 tests, 7 skips, 24 deliberate
  deselections, and 7 expected failures. Packaging checks and focused
  release-cut tests
  cover rollback under `Exception` and `BaseException`, append re-entrancy,
  pure validation, file-like sniffing, image replacement, receipt effects,
  chart-cache removal, rename/hyperlink composition, omitted validation
  defaults, extension-bearing tables, and valid packages above the removed
  Paper thresholds.
- Cell-write and append benchmarks remain linear from 5,000 through 20,000
  operations; preservation overhead stays approximately 2.5x and 2.7x rather
  than increasing with workbook size.
- GitHub Actions passes on Windows, Python 3.9--3.13, stdlib XML, default
  dependencies, LibreOffice, documentation, CodeQL, and build/install from the
  source and wheel distributions.
- All 24 local Apple Silicon LibreOffice smoke tests pass, including preserved
  recalculation, certification, array results, chart/table delivery, structural
  edits, and corrupt-input refusal.
- Sphinx builds with warnings treated as errors. The source and wheel
  distributions build successfully and both pass strict Twine validation.

Package correctness, model behavior, treatment compliance, and evaluator
validity remain separate claims.
