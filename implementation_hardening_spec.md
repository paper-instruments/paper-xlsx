# paper-xlsx code hardening specification

- Status: Implemented on `agent/paper-xlsx-implementation-hardening-spec`
- Scope: Runtime code and release-facing documentation, compared with
  `paper-base`
- Documentation: Updated to match the release-cut API
- Target: The implementation that will be published as the first open-source
  release

This is a release cut, not a roadmap. Every numbered change in this document is
required before publication. There are no later-version buckets. Code and APIs
not named in a required change stay as they are.

## Decision

Keep the hard fork, but do not publish the current implementation unchanged.

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

The current implementation has confirmed release blockers in four classes:

1. normal mutations can become quadratic;
2. some successful in-memory edits can be silently omitted, become unsavable,
   or be overwritten by a competing package-part plan;
3. structural and relationship operations either leave stale dependent state
   or refuse common combinations that the public surface claims to support;
   and
4. Paper behavior leaks into `preserve=False` and rejects or changes workbooks
   that upstream openpyxl accepts.

The Harbor corpus also exposed two concrete product gaps: no safe way to replace
one image when a media part is shared, and no way to compose an address shift
with an explicit chart-range edit in one save. Those gaps caused agents to
rewrite raw ZIP/XML or split one logical edit across save/reload sessions.

The evaluations do not prove a package-only score gain because the paper
treatments also changed the spreadsheet skill. They do show which APIs models
actually used, which refusals recurred, and which saved artifacts violated the
task contracts. The implementation should prioritize the automatic
preservation spine and these observed failure boundaries over speculative
convenience APIs.

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

This revision also parsed the complete Harbor export at
`/private/tmp/paper-office-evals-harbor-20260812`:

- 722 complete agent trajectories and 854 recorded results;
- 15 workbook tasks;
- upstream with the Anthropic spreadsheet skill;
- paper-xlsx `0.1.3` with the full Paper skill;
- upstream with no spreadsheet skill; and
- paper-xlsx `0.1.3` with the light Paper skill.

Setup failures were kept separate from agent trajectories. API counts came from
tool-call arguments and code written by the agents, not from API names quoted in
skill text. Grader checks were inspected at the individual-trial level where a
treatment gap appeared.

The implemented release cut passes the complete non-LibreOffice suite: 3,259
tests passed, 23 skipped, 25 LibreOffice-tier tests were deselected, and 7 were
expected failures. The isolated LibreOffice tier still invokes a locally
installed executable that aborts with exit code 134 even for its untouched
smoke fixture. This is an environment or LibreOffice failure, not evidence
that those package behaviors pass. The independent-loader and oracle gate was
rerun in the Linux GitHub Actions environment and passed, along with Windows,
Python 3.9--3.13, stdlib XML, documentation, and distribution-build jobs.

## Fork inventory and decision

| Area | What changed from upstream | Decision |
| --- | --- | --- |
| Distribution import | Fork sentinel and same-namespace ownership guard | Keep |
| Reader | Preserve default, source retention, content inventory, ledger arming | Keep; fix file-like classification |
| ZIP/OPC integrity | Detect duplicate or colliding names, overlap, header disagreement, truncation, size/CRC mismatch, encryption, and unsupported compression | Keep; these checks identify malformed or ambiguous packages |
| Resource eligibility caps | Reject packages by entry count, part size, aggregate size, compression ratio, source size, or stock-path part size | Remove completely; these are Paper policy thresholds, not XLSX validity rules |
| Cell/style hooks | Track values, types, styles, hyperlinks, comments, and protection writes | Keep; replace whole-ledger rollback snapshots |
| Workbook/worksheet lifecycle | Track sheets, names, regions, tables, charts, and structural edits | Keep only where the affected state is closed-world; close known hook gaps |
| Saver | Plan changes, splice XML, copy untouched parts, validate, and deliver atomically | Keep; make plans composable by package part and split the 727-line planner |
| Formula caches | Remove stale caches and force recalculation metadata | Keep; this is directly eval-backed |
| Package diff/receipt | Compare package and cell effects | Keep; add a small derived-effects field |
| Formula linter | Lint formula assignments with a hand-maintained catalog and heuristic reference checks | Remove; it is an incomplete semantic checker, not preservation infrastructure |
| Model map | Semantic roles inferred from formula references | Remove; the categories are not provable and direct use is negligible |
| Findings | Flag likely hardcodes, inconsistencies, volatility, hidden content, and magnitude outliers | Remove; the thresholds and classifications are subjective workbook semantics |
| Delivery helpers | Inferred protection, scrubbing, formatting profiles | Remove; these are policy conveniences, not preservation primitives |
| Targeting helpers | Infer a nearby value cell and mutate it | Remove `locate()` and `set_input()`; explicit coordinates and defined names are safer |
| Table-row append | Expand a loaded table, preserve totals/filter state, and derive declared calculated columns | Keep the capability; replace hidden `append_row()` with atomic `Worksheet.append_table_row()` and remove formula guessing |
| Oracle | Optional LibreOffice recalculation and certification | Keep module-level, explicit-source operations lazy and optional; remove package-replacement writes and `Workbook.evaluate()` |
| Raw escape hatches | `mark_dirty()` and `replace_part()` | Remove; neither can prove that the declared low-level change is safe |
| Pivot refresh | Patch selected pivot-cache definitions to refresh on load | Keep the capability; require explicit targets or `all=True` and route it through a dedicated save planner |

## Severity summary

### Release-blocking issues

- Quadratic mutation rollback on normal cell writes.
- Quadratic and upstream-incompatible `Worksheet.append()` behavior.
- Unstable object fingerprints are skipped, allowing loaded object edits to
  save as no-ops.
- Sheet rename rewrites hyperlink locations and then the save planner refuses
  the same hyperlink change.
- Inference-backed protection and formatting can mutate the wrong cells.
- Stale chart cache retention after range changes.
- A structural shift and an explicit chart edit targeting the same chart part
  cannot be composed.
- There is no relationship-aware isolated image replacement operation.
- Raw part replacement can shadow model-, metadata-, theme-, chart-, comment-,
  drawing-, or relationship-owned output.
- `unmerge_cells()` and several formula/region surfaces bypass the preservation
  ledger or structural preflight.
- Known XML-part editors can silently miss or normalize content they do not own.
- Paper-only guards and resource limits alter upstream behavior even when
  preservation is disabled.

### Required bounded fixes

- Formula linting and workbook findings make unsupported semantic claims that
  are unrelated to package preservation.
- Target-inference and raw mutation helpers expose convenient operations whose
  safety the package cannot prove.
- The public preserve namespace exposes internal machinery as supported API.
- Structural-rewrite guarantees are broader than the enumerated reference
  surfaces the code can prove.
- Several core functions are too large to review safely.
- Five upstream files differ only in formatting or docstrings.

## Required implementation changes

The numbering is sequencing, not a later-release priority scheme. All items are
part of the same pre-publication gate.

### R1 Make cell mutation rollback proportional to the edited cell

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
  protection warnings-as-errors, comment binding, hyperlink
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

### R2 Remove whole-sheet snapshots from `Worksheet.append()`

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

- Use an append-local transaction in both modes. When no armed Paper ledger
  exists, keep an upstream-like hot path that records only target coordinates,
  `_current_row`, and caller-supplied `Cell` bindings; it must perform no
  Paper-ledger or whole-sheet snapshot work. This preserves the fork's current
  atomic append behavior without retaining its quadratic cost.
- In preserve mode, validate the iterable and pre-built `Cell` bindings before
  committing the row where possible.
- Do not materialize an arbitrary generator merely to preflight it. Consume it
  lazily and journal each append-owned mutation before it occurs.
- For each target coordinate, capture its pre-operation entry only the first
  time it is touched. This covers duplicate normalized dictionary keys such as
  `{"A": 1, 1: 2}` without treating an intermediate value as the baseline.
- Snapshot each caller-supplied `Cell` object only once by identity. The same
  object may appear more than once in the iterable; rollback must restore its
  original parent, row, and column rather than an intermediate binding.
- Refuse a caller-supplied `Cell` that is already materialized in a worksheet's
  `_cells` mapping. Moving that object would leave two coordinates pointing to
  one mutable cell whose own coordinate no longer matches one mapping.
- In preserve mode, journal coordinate-local changes to `cells`,
  `value_overwrites`, `cache_writes`, `formulas_changed`, protection-warning
  state, and any style registry extended while binding the row.
- On failure, execute undo actions in reverse order. Restore overwritten
  coordinate entries, remove newly created entries, reset `_current_row`,
  restore supplied-cell bindings, and restore ledger deltas without touching
  unrelated workbook state.
- Detect re-entrant `append()` on the same worksheet. The first implementation
  should refuse the inner append before it mutates anything; supporting it
  later would require nested append savepoints and a defined row-order
  contract. Appends on other worksheets and unrelated edits performed by a
  generator are outside the outer append's rollback scope.
- Do not call `_capture_structural_state()` for a simple row append.

The atomicity boundary covers mutations performed by `append()`. It cannot
undo external side effects performed by caller generator code, and it cannot
un-emit a warning already delivered to Python's warning machinery. It must,
however, restore Paper's warning-membership state when that warning is promoted
to an exception.

#### Acceptance criteria

- The `preserve=False` append implementation has no Paper snapshot or ledger
  work in its hot path and retains append-local atomic rollback.
- Preserve and stock append loops scale linearly through at least 20k rows.
- A generator that raises after several values leaves no append-owned changes.
- Invalid, cross-sheet, already-materialized, and repeated `Cell` objects have
  explicit tested behavior; every refusal is atomic.
- Duplicate dictionary targets retain upstream final-write-wins behavior on
  success and restore the original coordinate on failure.
- Re-entrant same-sheet append refuses before mutation. Generator edits to a
  different sheet survive an outer append failure because they are unrelated
  operations.
- Existing coordinate and ledger membership is restored rather than removed;
  mapping, set, and caller-supplied object identities remain stable.
- Failure injection covers value/type binding, automatic number-format
  registration, protection warnings-as-errors, ledger marking,
  `Exception`, and `BaseException` interruption points.
- Exceeding Excel's row or column bounds after partial generator consumption is
  atomic.
- When the generator performs no deliberate unrelated edit, serializing before
  and after each refused append produces equivalent package state.
- Existing append, ledger, and style tests continue to pass.

### R3 Never skip change detection for unstable loaded objects

#### Confirmed failure

`openpyxl/preserve/ledger.py::diff_objects()` unions the objects whose table,
chart, image, or pivot serializer was unstable at arm time and save time, then
skips comparison for every object in that set. The code calls this a way to
avoid a false refusal. It also means that a real edit to such an object can be
omitted while save copies the original backing part.

The image fingerprint helper compounds the problem by converting any read
failure into `None`. Two failures can therefore compare equal. A successful
in-memory edit must never become an apparently successful no-op save.

#### Required change

- Replace unstable serializer fingerprints with a stable, purpose-built
  canonical fingerprint for every loaded object type the package claims can be
  observed.
- If a stable fingerprint cannot be produced, mark that exact object
  untrackable and refuse its mutation or the save. Do not skip it.
- Preserve object identity and the original source package on refusal.
- Convert image stream/read/fingerprint failures into a typed refusal with the
  worksheet and object index. Do not use `except Exception: return None` as a
  state value.
- Keep no-op saves working: an untrackable but untouched object may be copied
  verbatim only when a separate stable guard proves it was not mutated.

#### Acceptance criteria

- Inject nondeterministic serializers for one table, chart, image, and pivot;
  real edits either land exactly or refuse, never disappear.
- A consumed or unreadable image stream produces a typed refusal before
  destination replacement.
- Unrelated editable objects on the same sheet remain usable.
- A refused save leaves the destination and source identities unchanged.

### R4 Compose sheet rename with internal hyperlink rewrites

#### Confirmed failure

The rename reference pass rewrites internal `Hyperlink.location` values that
mention the renamed sheet. Later, `_plan_hyperlinks()` compares those values
with the arm snapshot and refuses every changed hyperlink because only additions
are supported. The rename succeeds in memory but leaves the session unsavable.

#### Required change

- Treat hyperlink-location rewrites derived from the same sheet rename as part
  of the rename plan.
- Patch internal locations without rewriting external hyperlink relationships
  or unrelated hyperlink attributes.
- Continue to refuse unsupported user-initiated hyperlink removal or target
  replacement, but distinguish it from a planner-derived rename effect.
- Run the hyperlink preflight before committing the in-memory rename, or keep a
  complete rename savepoint that restores the workbook if planning refuses.

#### Acceptance criteria

- Rename a sheet referenced by internal hyperlinks and save in one session.
- Reopen and prove formulas, defined names, chart references, print regions,
  hyperlink locations, and external hyperlink targets are correct.
- An unsupported manual hyperlink edit still refuses atomically.
- A failed rename leaves sheet titles, references, ledger state, source bytes,
  and destination bytes unchanged.

### R5 Remove the heuristic model map and delivery convenience layer

#### Confirmed failure and product decision

`model_map()` records `conventions["unresolved_references"] = True` when it
cannot resolve all formula inputs. `protect_for_delivery()` ignores that flag,
unlocks only inferred inputs, locks every other populated cell, and enables
sheet protection. `apply_profile()` uses the same heuristic role map to drive
bulk style and protection changes.

For a table `Sales` over `A1:A3` and `=SUM(Sales[Amount])`, the current map
contains no inferred inputs and reports unresolved references. Protection then
locks the real data cells. Blank intended input cells are absent because the
map classifies only populated cells. Even a fully resolved formula graph does
not prove that every referenced literal is meant to be user-editable.

The problem is not limited to downstream mutation. The public map itself gives
semantic names to structural observations:

- a populated non-formula cell referenced by a parsed formula becomes an
  `input`;
- a referenced formula becomes a `calculation`;
- an unreferenced formula becomes an `output`; and
- an unreferenced populated literal becomes a `constant`.

Those categories are not provable from formula references. A referenced literal
can be a fixed constant. An unreferenced formula can be an intermediate reached
through an unresolved, external, structured, dynamic, or 3-D reference. A
blank cell can be an intended future input. The public names make the result
look more authoritative than it is.

The Harbor call counts do not establish product demand. Protection and
scrubbing were concentrated in a synthetic task named
`delivery-scrub-and-protection`; its prompt explicitly required both actions.
The other workbook tasks did not require locking. `apply_profile()` had no
comparable task-level requirement. This is benchmark-by-construction use, not
evidence that agents need a role-inferred delivery layer.

The usage counts also do not justify keeping the classifier independently.
`model_map()` appeared in 13 full-skill trajectories and 4 light-skill
trajectories, while protection appeared in 12 and 3 respectively. After
accounting for the delivery helper's internal call, direct use was approximately
one trajectory per arm.

`scrub()` is deterministic rather than inference-backed, but it bundles policy
choices about comments, metadata, personal properties, and hidden sheets into
one catch-all operation. Those edits are already available through explicit
openpyxl workbook, property, comment, and sheet APIs. In one retained delivery
trajectory, clearing personal properties still saved `creator="openpyxl"`, so
the helper did not meet the intuitive promise of removing personal authorship.
It is not part of the preservation spine and does not justify another public
contract.

#### Required change

- Remove `Workbook.protect_for_delivery()`.
- Remove `openpyxl.preserve.apply_profile()` and its public export. Retain
  `copy_format()` and ordinary openpyxl style APIs.
- Remove `Workbook.scrub()`.
- Remove `Workbook.model_map()`, `openpyxl.preserve.modelmap.ModelMap`, and the
  now-unused model-map builder/module.
- Do not add replacement inferred-role, delivery-profile, or catch-all scrub
  APIs before publication.
- Keep the separate internal `dependency_sketch()` used by cache invalidation,
  structural analysis, oracle exclusions, and conservative refusals. It reports
  parsed ranges and unresolved references; it must not expose semantic cell-role
  claims as public truth.
- Retain ordinary openpyxl protection, workbook-property, comment, and sheet
  operations. Preserve mode must track, land, or refuse those explicit edits
  under the same rules as other mutations.
- Remove model-map and delivery-helper documentation, examples, exports, and
  tests that imply these APIs remain supported. Add public-surface tests proving
  they are absent.

There is no deprecation window. These APIs have not been published in an
open-source release, and the purpose of this cut is to avoid promising weak
surface area.

#### Acceptance criteria

- `Workbook` exposes no `model_map()`, `protect_for_delivery()`, or `scrub()`
  method.
- `openpyxl.preserve` does not export `apply_profile()`.
- No public `ModelMap` type or model-map builder/module remains.
- Internal preservation tests continue to exercise `dependency_sketch()` and
  its conservative unresolved-reference behavior.
- `copy_format()` and stock explicit style/protection/property/comment/sheet
  operations remain available.
- A preserve-mode explicit protection or metadata edit either lands with the
  exact requested package delta or refuses before destination delivery.
- Public API snapshots and package documentation contain no removed model-map
  or delivery convenience API.

### R6 Remove or refresh chart caches when repointing

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

### R7 Compose supported edits that target the same package part

#### Eval-confirmed failure

The Harbor corpus contains 27 distinct paper-treatment trajectories where a
row or column shift and a later explicit chart edit targeted the same chart
part:

- 17 `revenue-plan-row-insertion` trajectories; and
- 10 `forecast-quarter-column-insertion` trajectories.

The save refused with: the chart "was edited in the same session as a
row/column shift that patches the same chart part." Agents responded by saving,
reloading, and applying the chart edit in a second session, or by rewriting raw
ZIP/XML. One logical workbook edit should not require either workaround when
both component edits are individually supported.

#### Required change

- Build one plan per package part, not competing whole-part payloads from
  independent planners.
- Apply structural reference remaps and explicit chart formula/title edits to
  the same source token tree in a deterministic order.
- Detect conflicts at the exact XML node or byte span. Refuse only overlapping
  incompatible edits, not every pair of edits that shares a part.
- Use the same composition rule for workbook relationships, sheet
  relationships, content types, comments, tables, and drawing changes.
- Preflight the complete composed plan before mutating the destination.

#### Acceptance criteria

- The revenue-row and forecast-column fixtures complete their structural and
  chart edits in one load/save session.
- The final chart formulas, drawing anchors, defined names, validations,
  conditional formatting, and formulas match the intended post-shift state.
- Two compatible edits to different nodes in one part compose.
- Two incompatible edits to the same node refuse with both owners named and no
  bytes written.

### R8 Add a relationship-aware isolated image replacement operation

#### Eval-confirmed failure

The isolated-image task required replacing the image at `H20` while another
anchor continued to share `xl/media/image1.png`. The correct package delta was:

- preserve `drawing1.xml` and `image1.png` byte-identical;
- add `xl/media/image2.png`; and
- retarget exactly one relationship in `drawing1.xml.rels`.

Agents used `replace_part()` in six paper-full image trajectories. Five scored
`0.625` and one scored `0`: they replaced the shared media globally, edited the
drawing XML unnecessarily, or could not add the relationship-owned part. The
paper-full task mean was `0.519`, compared with `1.000` for upstream/no-skill
and `0.899` for paper-light, whose successful runs used direct ZIP surgery.
This is a missing supported operation, not evidence that raw ZIP surgery is the
desired API.

#### Required change

Add one narrow operation:

```python
ws.replace_image(target, replacement, *, name=None)
```

- `target` accepts an anchor coordinate or a loaded image object. `name` may
  disambiguate multiple images at one anchor.
- `replacement` accepts the same path/file-like/image inputs as the standard
  openpyxl `Image` constructor.
- Resolve the exact drawing anchor and relationship before mutation.
- Allocate a fresh media part, add its content type when required, and retarget
  only the selected image relationship.
- Preserve the drawing XML, anchor geometry, non-visual properties, old media
  part, and all other relationships byte-identical.
- Refuse ambiguous anchors, unsupported drawing encodings, unreadable image
  data, extension/content mismatches, and relationship sharing that cannot be
  proved.
- Do not expose the internal generic `PartPlan.add_part()` as a public escape
  hatch.

#### Acceptance criteria

- The Harbor isolated-image fixture produces exactly one new media part and
  one changed relationship part.
- A control image sharing the old media remains unchanged.
- Anchor position and size, chart parts, worksheet parts, and drawing XML stay
  byte-identical.
- PNG, JPEG, GIF, and BMP replacements validate signatures and content types.
- Ambiguous and unsupported cases refuse atomically.

### R9 Preserve lexical XML state during supported region edits

#### Eval-confirmed failure

On `revenue-plan-row-insertion`, 11 of 12 paper-full outputs and 11 of 12
paper-light outputs failed the data-validation contract. The intended change
was only `sqref="B5:B6"` to `sqref="B5:B7"`; Paper re-emitted omitted false
attributes as explicit `showDropDown="0"`, `showErrorMessage="0"`, and
`showInputMessage="0"`. Eleven of 12 upstream/no-skill outputs preserved the
expected representation.

The explicit false values may be semantically close, but the fork's product
claim is targeted preservation. A range edit must not normalize unrelated
attributes.

#### Required change

- Patch owned attributes and formula text in the original XML token stream for
  loaded data validations, conditional formatting, filters, tables, and other
  supported regions.
- Preserve omitted defaults, attribute order, namespace prefixes, extension
  children, and unrelated lexical content.
- Re-render a whole element only for a supported create/delete operation, or
  when a preflight proves the original element is fully owned.
- If the exact owned field cannot be located uniquely, refuse rather than
  normalizing the region.

#### Acceptance criteria

- Changing only a data-validation `sqref` changes only that attribute's bytes.
- Absent false/default attributes remain absent.
- Equivalent tests cover conditional formatting, auto filters, table ranges,
  and extension-backed validation/formatting twins.
- The revenue-row fixture passes its exact validation contract without raw XML
  repair.

### R10 Remove formula linting

Formula linting is not part of package preservation. It attempts to judge
formula validity with a hand-maintained function catalog and partial reference
analysis. That cannot reliably cover workbook UDFs, add-in functions, future
Excel functions, external names, dynamic arrays, structured references,
`LET`, or `LAMBDA`. A warning mode still trains callers and agents to treat
incomplete diagnostics as meaningful, while refusal can reject legal workbook
content.

The Harbor run produced a concrete self-inflicted false warning during sheet
rename: the rename planner wrote formulas referring to `Operating Model` before
the linter could see that sheet title, so it emitted `unknown-sheet` warnings
for two formulas that were correct in the completed workbook. Paper-full agents
explicitly changed `formula_lint` in 13 trajectories; paper-light agents did
not change it at all. There is no task-score or artifact evidence that the
linter prevented a bad formula repair.

#### Required change

- Remove `Workbook.formula_lint`, `lint_formula()`, `lint_on_bind()`, the
  `"off"`/`"warn"`/`"refuse"` configuration, and their public exports.
- Remove `openpyxl/formula/lint.py` and the hand-maintained
  `openpyxl/formula/catalog.py` when no preservation code imports them.
- Remove `LintWarning` if no remaining supported behavior emits it.
- Remove bind-time lint hooks, linter-specific documentation, and tests that
  assert heuristic diagnostics.
- Keep the tokenizer and conservative dependency/reference parsing that the
  preservation planner actually needs. Removing the linter must not remove or
  weaken formula cache invalidation, structural preflight, or safe rewrites.

#### Acceptance criteria

- No public linter symbol, workbook mode, bind hook, warning type, or function
  catalog remains.
- Assigning any syntactically storable formula has stock-openpyxl behavior
  unless a separate preservation invariant requires a typed refusal.
- Internal formula parsing used by preservation retains its existing positive
  and adversarial tests.
- The sheet-rename fixture emits no transient linter warning because the
  linter no longer exists.

### R11 Finalize the public/internal API boundary now

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

The current `openpyxl.preserve.__all__` advertises internal machinery:

- `DirtyLedger`;
- `save_preserved`;
- `scan_archive`; and
- `LossInventory`.

These are implementation components, not stable user operations. Remove them
from `__all__` before publication. Do not create a deprecation window for an
API the open-source release has not promised.

Keep these supported groups:

- core: `load_workbook(..., preserve=...)`, `Workbook.preserve`, save flags,
  typed errors, and `AddressRemap`;
- verification: package diff, cell diff, receipt, and error scanning;
- deterministic helpers: `allowed_values()`, `search()`, `copy_format()`, and
  `Worksheet.append_table_row()`;
- targeted pivot refresh through
  `Workbook.set_pivot_refresh_on_load(pivots=..., all=...)`; and
- optional module-level oracle operations with an explicit source: recalc,
  certify, evaluate, batch evaluate, and write-back.

Remove these Paper additions before publication:

- `Workbook.model_map()`, `ModelMap`, and the model-map builder;
- `Workbook.protect_for_delivery()`, `openpyxl.preserve.apply_profile()`, and
  `Workbook.scrub()`;
- `findings()`, `Finding`, and the heuristic formula-linting surface in R10;
- `Worksheet.locate()` and `Workbook.set_input()`, which guess a mutation
  target by walking up to six cells right or below a label;
- `Workbook.mark_dirty()` and `Workbook.replace_part()`, whose declarations
  cannot make an otherwise unsupported mutation safe;
- `Workbook.evaluate()`, whose live-workbook receiver is misleading because
  the implementation evaluates retained source bytes; and
- the hidden module-level `append_row()` entry point, after its supported
  behavior moves to `Worksheet.append_table_row()`.

The removals are based on implementation contracts, not only low eval usage:

- `findings()` embeds a hand-picked blessed-literal list, a three-row majority
  rule for formula consistency, a 1000x magnitude-outlier threshold for runs of
  five or more cells, and blanket warnings for hidden or volatile content.
  Those may be useful application-specific policies, but they are not facts the
  workbook package can prove. Keep objective `scan_errors()` checks.
- `locate()` walks only a bounded nearby neighborhood and can select an
  unrelated cell in common label/value layouts. `set_input()` converts that
  guess into a write. Explicit addresses and defined names avoid that semantic
  leap.
- The current `set_pivot_refresh_on_load()` patches every pivot cache through
  raw replacement state. The user job is valid, but the implementation must
  become targeted and planner-owned rather than being removed.
- The current `append_row()` is not an alias for `Worksheet.append()`. It
  expands a table, keeps totals last, updates the filter range, and derives
  calculated columns. That is a real table operation worth keeping, but its
  hidden location, partial atomicity, and fallback formula inference need to be
  corrected before publication.
- The raw escape hatches and workbook oracle wrapper have deeper correctness
  failures specified in R15 and R16.

#### Required table-row API

- Replace the hidden helper with
  `Worksheet.append_table_row(table_name, values)`; do not overload generic
  `Worksheet.append()` with table semantics.
- Preserve table-range and auto-filter expansion, totals-row relocation,
  declared calculated-column formulas, boundary checks, and refusal when
  content below the table would require a structural shift.
- Remove the fallback that copies a formula from the preceding row when the
  table column has no `calculatedColumnFormula`. A missing ordinary-column
  value remains blank unless the caller supplies it.
- Remove every formula-linter dependency from normalization and preflight.
  Normal formula grid/boundary checks and preservation parsing remain.
- Build the complete cell, totals, table, and filter plan before mutation.
  Commit through a table-local undo journal covering affected cells, styles,
  comments, hyperlinks, table/filter references, formula-cache state, style
  registries, and preservation-ledger deltas.
- Reuse the coordinate-local transaction primitives from R1/R2, but keep the
  public table-row contract separate from generic row append.

#### Required pivot-refresh API

- Keep `Workbook.set_pivot_refresh_on_load()`, but require either explicit
  pivot names through `pivots=[...]` or an explicit `all=True`. Supplying
  neither, or both, refuses before state changes.
- Resolve pivot names to cache-definition parts before recording the request.
  Unknown or ambiguous names refuse with typed errors. If named pivots share a
  cache, disclose that the cache-level setting affects every pivot using it.
- Add typed planner state such as `pivot_refresh_requests`; do not write pivot
  bytes through generic `replaced_parts`.
- During save planning, patch only the root `refreshOnLoad` attribute in each
  selected `pivotCacheDefinition`. Preserve every other byte, namespace,
  extension, relationship, and cache-record part.
- Compose the patch through the shared package-part ownership plan. A competing
  unsupported edit to the same cache part refuses before archive assembly.
- Treat an already-enabled cache as an idempotent no-op and report successful
  new patches through R17 `derived_effects`.

Explicit coordinates, defined names, cell assignment, worksheet protection,
styles, properties, comments, and sheet operations remain available through
stock openpyxl APIs. Removing a convenience helper does not remove the
underlying explicit operation. Keep the internal conservative dependency
sketch used for preservation safety; it is not the removed semantic model map.

`Workbook.remove()` must return `None`, matching upstream, instead of returning
a `RemovalReport` whose `removed_parts` is always empty and whose
`remapped_names` counts surviving names rather than remapped names. The save
receipt is the correct place to report package-part removals after planning.
Remove `RemovalReport` from the public surface if no other operation can return
an accurate instance.

The pinned-surface test must distinguish supported public API from internal
implementation names and compare upstream types/signatures under
`preserve=False`.

#### Acceptance criteria

- Star import and generated API snapshots contain no ledger, scanner, planner,
  or loss-inventory implementation types.
- Generated API snapshots prove that every removal listed above is absent from
  its former namespace or class.
- No dead implementation module or internal-only orphan remains solely to
  support a removed public helper.
- `Worksheet.append_table_row()` is public, documented, and has atomicity and
  exact-package-delta tests distinct from generic `Worksheet.append()`.
- Pivot refresh requires explicit scope, performs no raw replacement-ledger
  write, and exposes its cache-level effects in the receipt.
- `Workbook.remove()` matches upstream return behavior.
- Every supported Paper symbol has a direct public-surface test; every internal
  symbol remains importable only by explicit internal module path if runtime
  code needs it.

### R12 Make structural rewrite coverage closed-world

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
- route `merge_cells()` and `unmerge_cells()` through the same ledger,
  protection, `data_only`, savepoint, and preserved-XML hooks;
- enumerate `ArrayFormula` and `DataTableFormula` surfaces on loaded and newly
  added sheets; rewrite a proven case or refuse the structural operation;
- rewrite or refuse affected freeze panes, selections, protected ranges,
  top-level sort state, and other coordinate-bearing worksheet state;
- detect namespaced dynamic functions such as `_xlfn.INDIRECT` wherever the
  unprefixed form is a blocker;
- refuse known formula-bearing relationship or extension families that are
  not in the registry;
- do not imply that arbitrary custom XML strings are rewritten; and
- add tests proving an operation either rewrites every registered affected
  surface or leaves model, ledger, and bytes unchanged.

Do not block a workbook merely because it has unrelated unknown parts. The
refusal should be tied to a known affected relationship or formula-bearing
surface, not a coincidental `A1` string in arbitrary XML.

#### Acceptance criteria

- A matrix covers every registered surface for row insert/delete, column
  insert/delete, move range, sheet rename, sheet add, and sheet remove.
- Each matrix case either rewrites all affected registered surfaces or proves
  byte-identical refusal atomicity.
- Added-sheet arrays/data tables and affected unmodeled regions cannot pass
  preflight silently.
- Unrelated unknown custom parts do not block a supported edit.

### R13 Restore upstream behavior when preservation is disabled

#### Confirmed failure

Paper-specific behavior currently leaks into the stock path:

- `Worksheet.append()` takes Paper snapshots even when no ledger is armed;
- `ws[address] = value` on a `data_only=True` stock workbook refuses while
  `cell.value = value` still writes;
- archive intake and XML policy add Paper refusals to ordinary openpyxl loads;
  and
- some warnings and structural guards key off a loaded workbook rather than an
  armed preserve ledger.

`preserve=False` is the compatibility escape hatch. It must behave like the
fork-point openpyxl except for packaging identity and unavoidable shared bug
fixes.

#### Required change

- Gate every mutation ledger, structural refusal, Paper warning, inventory
  scan, and save-planning hook on an armed preserve ledger.
- Restore upstream `data_only` mutation behavior in stock mode consistently for
  `__setitem__`, `Cell.value`, deletion, merge/unmerge, and structural methods.
- Move Paper ZIP/OPC/XML integrity validation out of shared upstream parser
  functions and into preserve/package-inspection entry points.
- Add a differential suite that runs the same load/edit/save program under
  fork-point openpyxl and paper-xlsx with `preserve=False`.

#### Acceptance criteria

- Stock-mode output, return values, warnings, and exception types match the
  fork point for the covered public API matrix.
- No Paper ledger or whole-state snapshot appears in stock hot paths.
- `preserve=False` accepts every valid fixture accepted by fork-point openpyxl.

### R14 Close the known XML part-editor holes

These are small code paths with silent-loss or false-refusal consequences. All
must be resolved before publication:

- rich-text cell re-emission must preserve unowned children, comments, and
  processing instructions or refuse before emitting;
- worksheet scanning must recognize self-closing cells with whitespace before
  `/>` and must not treat CDATA or namespace variants as ordinary text;
- x14 classic/extension twin matching must compare the same canonical form at
  arm and save, including allocated `dxfId` state;
- drawing relationship-id remapping must replace every exact relationship
  token, not only the first textual `rId1` occurrence;
- table-id detection must match the root `<table>` element, not a
  `<tableColumn>` substring, and table expansion must account for modeled
  calculated/totals formulas;
- comment planning must distinguish comment VML from form controls,
  header/footer VML, and other `vmlDrawing` relationships;
- chart blockers must parse extension nodes and namespaces instead of refusing
  on `extLst` or `c15:` byte substrings; and
- every broad `except Exception` in a fingerprint, scanner, or planner must
  either prove a safe fallback or raise a typed refusal with context.

#### Acceptance criteria

- Add one minimal positive and one adversarial fixture for each bullet.
- Positive cases preserve every unowned byte or semantic node.
- Adversarial cases refuse before model or destination mutation.
- The same canonicalizer is used for baseline and save-time comparisons.

### R15 Remove raw mutation escape hatches

#### Confirmed failure

`replace_part()` can replace core/custom properties, themes, charts, comments,
drawings, media, and other parts that preservation planners or multiple package
relationships may own. The assembly loop gives raw replacements precedence, so
a payload can shadow a valid model edit. Harbor agents treated this generic ZIP
operation as isolated image replacement even though it globally changes every
relationship that shares the media part.

`mark_dirty()` has the inverse problem: it lets a caller mark coordinates after
an unsupported low-level mutation, but that declaration cannot prove the model,
XML, relationships, content types, caches, and preserved source still agree. It
creates a false safety signal without teaching the planner what changed or how
to serialize it.

#### Required change

- Remove public `Workbook.mark_dirty()` and `Workbook.replace_part()`.
- Remove `check_replace_part()` and raw-replacement documentation/tests if no
  supported internal planner still needs them.
- Do not replace them with another generic dirty-range, raw XML, ZIP-part, or
  relationship escape hatch.
- Keep planner-owned internal part-replacement state only where a dedicated
  safe operation has a closed contract. The shared ownership registry required
  by R7, R8, and targeted pivot refresh remains an internal composition
  mechanism, not a public API. Pivot refresh uses its typed request state and
  save-plan patch, not the generic raw replacement map.
- Direct image-replacement callers to the relationship-aware operation in R8.
  Other unsupported low-level mutations must receive a typed refusal or remain
  outside the preservation contract.

#### Acceptance criteria

- `Workbook` exposes neither `mark_dirty()` nor `replace_part()`.
- No supported example or test declares an unsupported mutation safe by adding
  coordinates to a dirty ledger.
- Dedicated planners remain able to stage owned part output, and conflicting
  planners refuse before archive assembly.
- The six Harbor image-replacement strategies that misused `replace_part()`
  have a direct supported replacement in `Worksheet.replace_image()`.

### R16 Make oracle write and evaluation semantics non-destructive

#### Confirmed failure

`oracle.recalc(in_place=True)` replaces the workbook package with LibreOffice's
conversion output. That bypasses the preservation spine and can discard macros,
extensions, or byte-preserved parts. `Workbook.evaluate()` appears to evaluate
the live workbook but actually evaluates the as-loaded retained source, so
unsaved workbook edits are invisible. Harbor agents called or inspected the
workbook method in 21 trajectories, often after editing, and several also
guessed incompatible positional signatures.

#### Required change

- Remove the `in_place=True` package-replacement path from `recalc()`.
  Recalculation may return evidence or write a separate candidate path.
- Keep `write_back()` as the only API that mutates a candidate with calculated
  caches, and route it through the preservation planner.
- Remove `Workbook.evaluate()` entirely. A workbook method cannot accurately
  represent an operation that ignores the receiver's unsaved state.
- Keep module-level source evaluation explicit: the caller supplies the source
  package and scenario arguments directly. Do not redesign the retained
  module-level API merely to compensate for the removed wrapper.
- Macro-enabled workbooks must never be silently converted to `.xlsx`.

#### Acceptance criteria

- No oracle API replaces a caller's source package with generic converter
  output.
- `Workbook` exposes no `evaluate()` method. Module-level evaluation requires
  an explicit source and never labels retained source-state results as live
  workbook evaluation.
- Separate-path recalc and certification leave the source byte-identical.
- Macro, external-link, and extension fixtures prove the write boundary.

### R17 Report safety-derived save effects

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
- `pivot_refresh_on_load_enabled`;
- `calc_chain_removed`;
- `recalculation_metadata_changed`; and
- relationship or content-type changes created by supported lifecycle edits.

Do not redesign the rest of the receipt schema.

#### Acceptance criteria

- Derived effects are deterministic, versioned, and JSON serializable.
- The receipt distinguishes direct requested edits from every cache,
  recalculation, relationship, and content-type consequence in this cut.
- Receipt generation does not trigger a second full save or duplicate archive
  decompression.

### R18 Make the core code reviewable

The main risk is concentrated in functions too large for reliable review:

- `_save_preserved`: 727 lines, about 174 branch points;
- `scan_sheet`: 350 lines, about 101 branch points;
- `apply_model_shift`: 172 lines, about 43 branch points;
- `shift_blockers`: 142 lines, about 43 branch points;
- `begin_move_range`: 138 lines, about 49 branch points;
- `plan_workbook_xml`: 143 lines, about 47 branch points.

R10 deletes the separate 144-line `lint_formula()` implementation rather than
spending review effort refactoring a feature that does not make the release
cut.

Refactor in behavior-preserving slices:

1. split save planning by sheet, workbook metadata, relationships, lifecycle,
   and delivery;
2. give each phase an explicit immutable plan result;
3. separate XML tokenization from worksheet semantic validation;
4. separate structural blocker discovery from rewrite-plan construction; and
5. keep one final commit phase after all preflight succeeds.

`Workbook.validate()` currently performs a complete preserve save into
`BytesIO`. It was called in 129 paper-full trajectories, often immediately
before the real save. Once the planner is explicit, make `validate()` execute
the same complete preflight and plan validation without copying and compressing
the archive a second time. Any failure that can currently occur only during
assembly must be moved into preflight or remain a checked final-commit I/O
failure.

Land the correctness changes above first. Then refactor in behavior-preserving
slices and compare generated packages across the existing corpus before and
after each slice.

AST comparison found five modified upstream files with no executable change
after docstrings are stripped:

- `openpyxl/formula/translate.py`;
- `openpyxl/styles/fills.py`;
- `openpyxl/worksheet/_write_only.py`;
- `openpyxl/worksheet/filters.py`; and
- `openpyxl/worksheet/header_footer.py`.

Revert those diffs to the exact upstream fork-point content. Formatting-only
changes make future upstream comparison harder and have no package value.

#### Acceptance criteria

- The save planner and scanner are split into named plan phases with explicit
  inputs and immutable results.
- `validate()` and `save()` share the exact planner; validation does not build a
  duplicate archive and every non-I/O refusal agrees between the two paths.
- A golden corpus produces the same permitted package deltas immediately before
  and after each behavior-preserving refactor slice.
- The five accidental upstream files are byte-identical to the fork point.
- No correctness fix is hidden inside a nominal refactor commit.

### R19 Classify file-like sources by content

Path inputs can use the extension to choose preserve mode. File-like inputs
may have no meaningful name or a misleading suffix. The current default
disables preservation for a file-like object named `.xls` or `.xlsb` even if
its bytes are actually OOXML.

For seekable file-like inputs, sniff the container and OOXML content type while
restoring the caller's position. Use the content result for the default. Keep
an explicit `preserve=True` or `False` authoritative.

#### Acceptance criteria

- A nameless `BytesIO` containing OOXML preserves by default.
- Misleading `.xls` and `.xlsb` names do not disable preservation for OOXML
  bytes.
- A non-OOXML container does not enter preserve mode by extension alone.
- Sniffing restores the caller's exact stream position and does not close the
  stream.

### R20 Remove Paper-defined ZIP and source-size eligibility caps

The fork added resource-policy limits that do not exist at the upstream
openpyxl fork point:

- 10,000 ZIP entries;
- 256 MiB per uncompressed part in preserve mode;
- 512 MiB aggregate uncompressed package size;
- a 500x compression-ratio refusal above 64 MiB;
- a 512 MiB compressed source-file limit; and
- a separate 2 GiB per-part limit even on the stock load path.

These thresholds are not XLSX validity rules. The original Paper proposal did
include decompression caps as a general hardening item, but it did not derive
these specific thresholds from upstream behavior, a producer constraint, or a
measured package requirement. This proposal deliberately reverses that policy:
resource-exhaustion limits belong to the caller or execution environment, not
to workbook eligibility inside the compatibility fork.

The caps can reject a structurally valid workbook solely because Paper is
installed. The stock-path checks are a direct upstream compatibility
regression. The preserve-path checks also conflate package integrity with a
caller-level resource and security policy.

The integrity validation remains valuable. Keep checks for:

- duplicate and ASCII-case-colliding OPC part names;
- overlapping ZIP entries;
- local/central header disagreement;
- truncated or invalid compressed streams;
- declared-size and actual-size disagreement;
- CRC failure;
- encryption, unsupported flags, and unsupported compression methods; and
- other conditions that make the package ambiguous or unsafe to preserve
  faithfully.

#### Required change

- Remove `_check_decompression_caps()` from the stock load path and restore
  upstream behavior when `preserve=False`.
- Remove the entry-count, per-part size, aggregate-size, and compression-ratio
  refusal policy from preserve loading, ZIP validation, package diff, receipts,
  and other package APIs.
- Do not replace the constants with public tuning knobs. A workbook does not
  become structurally invalid at a caller-configured byte threshold.
- Keep decompression and validation streaming in bounded chunks. "No arbitrary
  cap" must not mean constructing an uncompressed package in one allocation.
- Remove `MAX_SOURCE_BYTES` without coupling this compatibility fix to a new
  retained-source abstraction. Preserve mode may continue retaining one
  immutable compressed `bytes` snapshot, as originally designed.
- Measure retained-source, package-diff, and receipt memory on representative
  large workbooks after removing the caps. Do not add a new source abstraction
  as part of this release cut.
- Preserve typed refusals for malformed or ambiguous ZIP/OPC structure. Do not
  describe resource exhaustion or a removed policy threshold as file
  corruption.

#### Acceptance criteria

- `preserve=False` introduces no Paper-specific entry, byte-size, aggregate,
  or compression-ratio refusal relative to upstream.
- Preserve mode and package inspection do not reject an otherwise valid XLSX
  solely for exceeding any removed Paper threshold.
- Memory measurements for retained-source and package-comparison paths are
  recorded separately from workbook-validity behavior; they do not reintroduce
  hard package-eligibility thresholds.
- The existing malformed-ZIP corpus still produces the same typed integrity
  refusals for duplicate names, overlap, header disagreement, truncation, CRC,
  encryption, and unsupported compression.
- Tests distinguish structural invalidity from resource-policy behavior; no
  test calls an oversized but otherwise valid package a decompression bomb.

## Harbor evidence used for this cut

The Harbor export is evidence about API ergonomics and saved artifacts. It is
not a clean package ablation because the runtime and skill changed together in
the paper treatments.

### Observed API use

Counts below are complete trajectories containing an actual call or code path,
not mentions in the skill text:

| API or operation | Paper full (180) | Paper light (179) | Release decision |
| --- | ---: | ---: | --- |
| `save(..., receipt=True)` | 162 | 53 | Keep; add derived effects in R17 |
| `validate()` | 129 | 29 | Keep; route through the pure planner in R18 |
| `scan_errors()` | 51 | 1 | Keep diagnostic |
| `copy_format()` | 19 | 0 | Keep explicit operation |
| `protect_for_delivery()` / `scrub()` | 12 / 12 | 3 / 3 | Remove in R5; calls came from the synthetic task that explicitly required both operations |
| `model_map()` | 13 | 4 | Remove in R5; nearly all calls were induced by the delivery helper |
| `allowed_values()` | 7 | 1 | Keep; low use is not a defect |
| `replace_part()` | 7 | 0 | Remove in R15; all demonstrated image use needed the dedicated R8 operation |
| `mark_dirty()` | 4 | 0 | Remove in R15; ledger membership cannot validate an unsupported mutation |
| `Chart.repoint()` | 1 | 1 | Keep; fix caches and composition |
| `locate()` | 1 | 0 | Remove in R11; it guesses a value target from nearby cells |
| `set_input()` | 0 | 0 | Remove in R11; it mutates a guessed target when no defined name resolves |
| current `append_row()` | 0 | 0 | Keep the table-expansion job in R11 as public atomic `Worksheet.append_table_row()`; zero use reflects the hidden API, not a redundant capability |

The high full-skill use of receipts and validation is partly instruction-driven.
The light skill achieved comparable task outcomes with far fewer calls. This
supports keeping those APIs available but not making them mandatory internal
steps or expanding the surface on usage counts alone.

### Evidence that supports existing core behavior

- In the dedicated preservation suite, the bare runtime passed 542/730 checks
  and paper-xlsx runtime-only passed 727/730. Runtime-only package integrity was
  390/390.
- In 50 sampled runtime-only traces, agents did not call validation or receipt
  APIs. Preserve-by-default still delivered the integrity gain.
- Exact artifact replay after the formula-cache fix recovered 12 binary task
  passes and lost none. One task moved from 60/149 to 149/149 rubric checks.
- In the Harbor run, paper-full and paper-light preserved the three custom XML
  parts in all 12 contact-update trials; upstream/no-skill preserved them in 9
  of 12. The same pattern appeared on preserved drawings and charts in the
  dashboard task.
- The removed manifest produced an 868 KB trace and more than 150 seconds of
  work on a 2.39-million-cell workbook. It stays removed.

### Evidence that produced required changes

- Six `replace_part()` image attempts all missed the isolated-image package
  contract. This produced the dedicated R8 API and removal of the raw escape
  hatch in R15.
- Twenty-seven distinct trajectories hit the structural-shift/chart-edit
  same-part refusal. This produced R7.
- Twenty-two of 24 paper outputs on the revenue-row task normalized unrelated
  data-validation defaults. This produced R9.
- Default formula lint emitted false unknown-sheet warnings during Paper's own
  rename sequence, and the run showed no formula-quality benefit. This
  reinforced removal in R10.
- `locate()` appeared in one full trajectory and `set_input()` had no calls;
  their guessed-target contract remains unjustified. The internal
  `append_row()` helper also had no calls, but code review showed that it owns a
  distinct table-expansion job that generic `Worksheet.append()` cannot perform.
  Promote and harden that job rather than inferring redundancy from zero use.
- The current pivot-refresh helper had no demonstrated Harbor demand, but code
  review confirmed a real preservation job: request Excel to refresh selected
  preserved pivot caches after their source data changes. Keep that narrow job
  while replacing the untargeted raw-replacement implementation.
- Paper agents inspected or used internal `cache_writes` state in 13
  trajectories after formula-cache or oracle write-back problems. This
  reinforced the public/internal boundary and R16 rather than making the ledger
  public.

### Evaluator findings that must not become package regressions

Some Harbor score losses are contradictory grader policy, not evidence to undo
preservation behavior:

- The sheet-rename grader required dependent formula and chart rewrites, then
  marked the rewritten worksheet part as unexpected collateral.
- The regional and macro tasks accepted cleared dependent formula caches as
  safe, then also penalized the worksheet part containing those cleared caches
  as an unexpected changed part.
- A successful package save does not prove calculation, rendering, or task
  correctness, and an unchanged-input replay is not model performance.

Release validation must classify package correctness, model behavior,
treatment compliance, and evaluator validity separately.

## Pre-open-source release gate

Publication is blocked until every R1-R20 change and the following proof set is
complete.

### Required automated proof

- Preserve-mode no-op save across the frozen fixture corpus.
- Exact part-level change allowlists for every supported edit.
- Refusal atomicity for model, ledger, source identity, and destination bytes.
- Linear-scaling benchmarks for cell writes and appends.
- Differential upstream tests with `preserve=False`.
- Object-fingerprint failure and refuse-before-write tests.
- Rename plus internal-hyperlink tests.
- One-session structural shift plus chart edit tests.
- Isolated shared-media replacement tests.
- Atomic table-row append tests covering totals, declared calculated columns,
  filters, comments/hyperlinks, content-below refusal, and rollback at every
  commit boundary.
- Targeted and explicit-all pivot-refresh tests covering shared caches,
  already-enabled caches, exact XML deltas, unknown/ambiguous names, and
  package-plan conflicts.
- Lexical region-preservation tests, including omitted defaults.
- Public-surface absence tests for every removed helper and type in R5, R10,
  R11, R15, and R16, plus exact-delta tests for equivalent explicit stock
  protection, metadata, named-cell, and coordinate edits in preserve mode.
- Chart formula/cache consistency tests.
- Closed-world structural surface and XML part-editor adversarial fixtures.
- Public API snapshot tests with internal exports absent.
- Oracle source/candidate custody tests.
- Valid large-package tests above every removed Paper threshold and malformed
  ZIP/OPC integrity tests.
- Independent load in a known-good LibreOffice environment.
- Real-producer fixtures from desktop Excel, LibreOffice, and Google Sheets,
  including pivots/caches, external links, macros, rich text, drawings, x14,
  threaded review content, and 1904-date workbooks.

### Required evaluation proof

Run the fixed task identities needed to exercise the changed boundaries under
both upstream and paper-xlsx with the same skill and agent configuration. At
minimum include:

- isolated image replacement;
- table row expansion with a totals row and declared calculated column;
- targeted pivot refresh after editing source data;
- revenue row insertion;
- forecast column insertion;
- sheet rename with hyperlinks;
- formula repair with x14 validation;
- shared-formula repair;
- macro-enabled input update; and
- custom XML preservation.

Report task score, exact package delta, cache behavior, refusal correctness,
completion/tool errors, runtime, and treatment compliance separately. The
release decision is based on package assertions and regressions, not aggregate
score alone.

## Definition of done

This proposal is implemented when:

- every R1-R20 item is complete;
- every executable difference from upstream has an owner, test, and stated
  compatibility effect;
- normal value writes and appends are linear;
- no successful edit can be skipped because its fingerprint is unstable;
- rename, relationship, structural, and chart plans compose when their exact
  edits do not conflict;
- no heuristic authorizes destructive mutation while reporting uncertainty;
- chart and formula source references cannot disagree with retained caches;
- isolated image replacement needs no raw ZIP/XML rewrite;
- table-aware row append is a public atomic operation and does not infer an
  undeclared calculated formula from the preceding row;
- pivot refresh is explicitly scoped, planner-owned, and changes only the
  selected cache-definition root attributes;
- supported region edits preserve unrelated lexical XML state;
- valid workbooks are not rejected solely for exceeding a Paper-defined ZIP or
  source-size threshold;
- preserve-off behavior matches the fork-point upstream contract in the tested
  API matrix;
- internal implementation objects are not accidentally promised as stable
  public API;
- no heuristic findings, formula lint, guessed targeting, raw dirty/part
  declaration, untargeted pivot patch, or stale-source workbook-evaluation API
  remains in the release surface;
- core planners are small enough to review by phase; and
- the independent-loader, real-producer, targeted Harbor, and full unit-test
  gates are green with package correctness reported separately from agent and
  evaluator behavior.
