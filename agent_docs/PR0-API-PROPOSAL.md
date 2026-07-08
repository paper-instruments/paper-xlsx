# PR-0 — API Proposal (Phase 1.5)

**Status:** the design contract for paper-xlsx v0. Everything below is grounded in Phase-0
evidence (`ARCHITECTURE-NOTES.md`, `OPEN-QUESTIONS.md`) plus two PR-0 spikes
(`scratch/probes/pr0_composed_save.py`, `pr0_g9_chokepoints.py`). CONVENTIONS remains the law;
this document exercises the decisions CONVENTIONS delegates to PR-0 and records the sanctioned
amendments to pinned shapes (each marked **[AMENDMENT]** with its evidence). Implementation
phases follow this document verbatim; deviations discovered during implementation are amended
here in the same commit that implements them, never silently.

---

## 1. Module layout

| Module | Contents | Phase |
|---|---|---|
| `openpyxl/errors.py` | `PaperRefusal` taxonomy + `LossySaveWarning` (name pinned by CONVENTIONS §2) | 2a |
| `openpyxl/package/` | the kernel: `xml_equivalent`, `diff_package`, `diff_cells`, typed results | 2a (diff_cells 4) |
| `openpyxl/preserve/` | the spine: retention store, dirty ledger, splice writer, collateral policies, lossy-content inspector, manifest report | 2a–2d, 4 |
| `openpyxl/oracle.py` | LibreOffice driver, `recalc()`, `certify()` | 5 |

Note: `openpyxl.package` (pinned name) is distinct from upstream's `openpyxl.packaging`
(content-types/rels machinery). The near-collision is unfortunate but the kernel name is pinned.

## 2. Exception taxonomy (pinned shape, confirmed against the code)

```python
# openpyxl/errors.py
class PaperRefusal(Exception):
    """Base for all safe refusals. A refused operation left the model, the
    ledger, and every file exactly as they were (CONVENTIONS §1.3)."""

class AmbiguousTargetError(PaperRefusal): ...
class TargetNotFoundError(PaperRefusal): ...
class UnsupportedStructureError(PaperRefusal): ...
class BoundaryViolationError(PaperRefusal): ...
class RelationshipPolicyError(PaperRefusal): ...
class OracleUnavailableError(PaperRefusal): ...
class OracleTimeoutError(PaperRefusal): ...

class LossySaveWarning(UserWarning):
    """Loud, structured warning on the stock save path when content that
    cannot be preserved is about to be rebuilt lossily or dropped.
    .losses -> list of {"kind", "location", "detail"} dicts."""
```

Programmer errors stay `TypeError`/`ValueError` (CONVENTIONS §2): notably
`load_workbook(preserve=True, read_only=True)` raises **`ValueError`** (correctly-typed flags in
an invalid combination; in-repo precedent `copy_worksheet`, workbook.py:407-408), raised at the
top of `ExcelReader.__init__` before any file handle opens.

## 3. Load / save surface

```python
def load_workbook(filename, read_only=False, keep_vba=KEEP_VBA, data_only=False,
                  keep_links=True, rich_text=False, *, preserve=False): ...
# openpyxl.open is the same callable; pandas reaches this via
# pd.ExcelWriter(..., engine="openpyxl", mode="a", engine_kwargs={"preserve": True})
```

- `preserve=True` retains the source archive as one immutable `bytes` blob on the workbook
  (new attribute `wb._paper_source`; never `_archive` (owned by read_only), never
  `vba_archive` (flips mime_type)). File-like sources are read eagerly (pandas hands an open
  `r+b` handle that is later overwritten in place — Q3).
- `Workbook.preserve` — read-only property, `False` on stock/fresh workbooks.
- `preserve + read_only` → `ValueError`. `preserve + write_only` is unreachable by
  construction. `preserve + data_only` loads fine; its **save** refuses (Phase 3).
- Save dispatch: one guard at the top of `save_workbook` (writer/excel.py:279) routes
  preserve-mode workbooks to the splice save. `Workbook.save` needs no edit.
- **Save targets (G4 resolution) [AMENDMENT to CONVENTIONS §7 wording]:** path targets are
  written temp-file-then-`os.replace` (atomic; in-place truncation is the measured hazard);
  file-like targets (pandas) are built fully in memory and written with a single
  `seek(0)/write/truncate` choreography — atomic rename is impossible for handles, so the
  in-memory build is the atomicity mechanism (nothing partial ever reaches the handle).
- Deterministic zip writing: fixed `ZipInfo` date_time, uniform external_attr, entries written
  via `writestr` (payload determinism measured achievable — Q12); `properties.modified` is NOT
  auto-stamped under preserve (D3).

### Example — the pandas path, zero pandas changes

```python
with pd.ExcelWriter("model.xlsx", engine="openpyxl", mode="a",
                    engine_kwargs={"preserve": True}) as xw:
    df.to_excel(xw, sheet_name="Appendix", index=False)
# charts, sparklines, VBA, pivots in model.xlsx survive byte-identical
```

### Example — surgical edit with refusal handling

```python
from openpyxl import load_workbook
from openpyxl.errors import PaperRefusal

wb = load_workbook("client_model.xlsx", preserve=True)
wb["Model"]["B8"] = 0.12
try:
    wb.save("client_model_v2.xlsx")
except PaperRefusal as r:
    print("refused:", r)   # file on disk untouched, model still usable
```

## 4. `mark_dirty` (the documented escape hatch)

```python
Workbook.mark_dirty(target: str) -> None
```

- `target` containing `!` → sheet-qualified A1 region (pinned addressing): the ledger marks
  those cells dirty and the splice re-emits them from the model.
- otherwise → must exactly match a part name in the retained package (`"xl/media/image1.png"`);
  that part is re-serialized from the model if modeled, else `UnsupportedStructureError`
  (a raw part with no model source has nothing to regenerate from).
- unknown target → `TargetNotFoundError`. Only meaningful under preserve; stock → `ValueError`.

## 5. Package kernel (`openpyxl.package`, CONVENTIONS §7)

```python
def xml_equivalent(a, b) -> bool
    # a, b: bytes or paths to XML payloads. Clark-name comparison: prefixes
    # insignificant, attribute order insignificant, inter-element whitespace
    # insignificant, text content NEVER normalized. Explicit-schema-default
    # tolerance (e.g. showDropDown="0") is a documented, versioned table.

def diff_package(path_a, path_b) -> PackageDiff
    # part-by-part: XML parts semantic, binary parts size+sha256.
    # PackageDiff.added/.removed/.changed/.identical; .to_dict() ->
    # {"schema": "package_diff", "version": 1, "added": [...], "removed": [...],
    #  "changed": [{"part": ..., "kind": "xml|binary", "detail": [...]}], ...}

def diff_cells(path_a, path_b) -> CellsDiff   # Phase 4
    # [{"address": "Model!B7", "old_value": ..., "new_value": ...,
    #   "old_formula": ..., "new_formula": ...}]; deterministic order.
```

## 6. Frozen decisions

**D1 — Strings: inline strings for ALL operation classes.** Every string cell the splice emits
is `t="inlineStr"` via the existing `write_cell` machinery — upstream's only string mode (it has
no sharedStrings writer at all). `xl/sharedStrings.xml`, where present, is raw-copied
byte-identical; indices never renumber. Append-only sst stays a documented post-v0 fallback
gated on a real-Excel fixture. **[AMENDMENT to §3.5]:** sharedStrings drops OUT of the
sanctioned single-cell collateral set.

**D2 — Sanctioned collateral sets, per operation class [AMENDMENT to §3.5, enforced literally
by the budget test]:**

| Operation class | Parts that may change |
|---|---|
| cell value/formula edit (incl. strings) | sheet part; calcChain **removed** + its CT override + its workbook rel (formula edits only); workbook.xml (calcPr, formula edits only) |
| style-bearing cell edit | + styles.xml (append-only) |
| comment add/edit | + comments part, VML part, sheet rels, [Content_Types] |
| hyperlink add | sheet part + sheet rels (append-only) |
| sheet append (incl. pandas) | + new sheet part, workbook.xml, workbook rels, [Content_Types] |
| chart/image onto preserved sheet without drawing | + new drawing/chart/media parts + their rels, sheet part, sheet rels, [Content_Types] |
| workbook-property edit (explicit) | docProps/core.xml |

Anything outside the class's set is a test failure.

**D3 — core.xml modified-stamp (G3): preserve-mode save does NOT auto-stamp
`properties.modified`.** docProps/core.xml is raw-copied unless the user explicitly changed
`wb.properties`. This preserves the pinned no-op payload-identity invariant. Logged in PAPER.md
as sanctioned deviation (stock path keeps stock stamping). Any stamping that does occur goes
through an injectable clock.

**D4 — Performance budget: preserve save ≤ 2× stock save** on the large fixture.
**[AMENDED in Phase 2c with evidence]:** originally pinned at 1.5× from the composed prototype
(0.16×) plus the Q4 spike-scanner projection (~0.8×). The production scanner does strictly more
work than the spike (per-scope namespace tracking, exact-parent-chain matching, span collection
for every row/cell/region, shared-formula/array/cm-vm inventory, guard checks) and measures
**1.82× on 600k cells (3.99s vs 2.19s) and 1.87× on 150k cells** after optimization
(single-scan reuse, byte dispatch, fast tag-end, selective attribute parsing). 2× is the
honest evidence-based budget; expat-byte-offset or lxml span acceleration remains the
non-semantic contingency if it must shrink. Fresh-generation path: unaffected within noise
(large-fixture stock load +0.7%).

**D5 — Chokepoint inventory (frozen; Q2 + verifier + G9 spikes).**
- *Tier 1 — instrumented funnels:* `Cell._bind_value`; `Cell.hyperlink`/`comment` setters;
  `cell.data_type` (converted to a property — silent formula demotion measured);
  `Worksheet.__setitem__/__delitem__`; `append`; `merge_cells`/`unmerge_cells`;
  `insert_rows/insert_cols/delete_rows/delete_cols` + `move_range`; the four style descriptors +
  `NumberFormatDescriptor` + `NamedStyleDescriptor`; `NamedStyle.__setattr__` +
  `add_named_style`; `ConditionalFormattingList.add/__setitem__/__delitem__`;
  `add_data_validation`; `DefinedNameDict` all mutating dunders (`__setitem__/__delitem__/pop/
  popitem/clear/update/setdefault/__ior__` — the missing overrides are added);
  `TableList.add/__delitem__`; sheet lifecycle (`create_sheet/_add_sheet/remove/__delitem__/
  move_sheet/copy_worksheet`, `title` setter, `active` setter, `epoch` setter);
  `freeze_panes`; `print_area`/`print_title_rows/cols`; `add_chart/add_image/add_pivot/
  add_table`; `Comment.text`; protection password setters; `AutoFilter.ref` + filter methods;
  row/col breaks appends; `DimensionHolder.group`; `CustomPropertyList.append/__delitem__`.
- *Tier 2 — satellite snapshots* (fully-modeled elements, snapshotted at preserve-load,
  re-serialized + compared at save; catches every by-reference/bare-attribute path):
  worksheet: `sheetPr, dimension?, sheetViews, sheetFormatPr, cols, sheetProtection, scenarios,
  autoFilter, mergeCells, dataValidations, hyperlinks, printOptions, pageMargins, pageSetup,
  headerFooter, rowBreaks, colBreaks, tableParts` (+ classic `conditionalFormatting`, gated by
  D18); chartsheet: `sheetPr, sheetViews, sheetProtection, pageMargins, pageSetup,
  headerFooter`; workbook.xml: `workbookPr, workbookProtection, bookViews, sheets,
  definedNames, calcPr`; docProps/core.xml; each table part; row/column dimension state.
- *Tier 3 — named always-dirty / special:* nested StyleProxy leak → styles.xml semantic
  re-diff at save; `CellRichText` in-place → always-dirty per cell; per-sheet hyperlink-set
  hash (post-hoc `hyperlink.target=`); `wb.loaded_theme` bytes-compare; post-hoc in-session
  chart/image object or anchor mutation → that drawing/chart part re-emitted (measured:
  chart part only); `wb.code_name` → workbook.xml (Tier-2 covered), `wb.template` →
  [Content_Types] (content-types policy D13).
- *Not chokepoints:* pure reads that materialize (`row_dimensions[n]`, `ws['Z99']`,
  `iter_rows`) — the ledger keys on semantic mutation; the no-op invariant test enforces it.
  `ws._rels.append` is measured-discarded by stock save today → not a public mutation path;
  `mark_dirty` territory.
- The ledger **arms after load completes** (the reader itself fires `create_sheet`, style sets,
  and `_bind_value`-bypassing writes — measured).

**D6 — Splice guard set (Q4 + skeptic, verbatim-adopted).** The scanner refuses
(`UnsupportedStructureError` unless noted) before any write when: DOCTYPE present; encoding not
UTF-8; target cell's in-scope default namespace ≠ spreadsheetml main or the element is
prefix-qualified (unguarded failure = measured silent value deletion); rows/cells without `@r`
(all operations, v0); a dirty cell carries unexpected children (e.g. cell-level `extLst`) or
`cm`/`vm` attributes (dynamic-array metadata would go stale). Target matching is by EXACT
parent chain `worksheet→sheetData→row→c` at depth 3 — never ancestor containment (two legal
decoys measured). Replacement cells carry over verbatim every original attribute not
intentionally rewritten (`ph`, foreign-ns attrs). `TargetNotFoundError` when a dirty
coordinate's spliceable location cannot be found. Refusals never fall back to echo re-emission.

**D7 — Shared/array formulas (G1).** Per touched sheet, a pre-scan pass over the original
bytes collects `si→ref` shared-group maps, `t="array"` refs, and `cm`/`vm` cells. If any dirty
cell intersects a shared group, the WHOLE group dissolves: every member becomes an explicit
`<f>` from the model's already-expanded formulas (openpyxl expands shared groups via Translator
at load — the model knows every member's formula). Semantically identical, single splice pass,
no si bookkeeping. Dirty ∩ array-formula ref → refuse in v0. Dirty cell carrying `cm`/`vm` →
refuse in v0.

**D8 — Sheet lifecycle under preserve, v0 (G7).** `create_sheet` (append) is supported —
including through pandas. `del wb[name]`, `ws.title = ...` (rename), `move_sheet`, and
`copy_worksheet` **refuse** (`UnsupportedStructureError`): delete requires a localSheetId remap
of preserved definedName bytes plus a dependent-part cascade; rename silently strands every
formula referencing the old name (stock behavior is itself corrupting). pandas
`if_sheet_exists="replace"` therefore refuses in v0; `"overlay"` and `"new"` work.

**D9 — Mixed fresh-on-preserved (Q6). [AMENDED in Phase 2d]:** fresh charts/images under
preserve mode refuse in v0 everywhere (including new in-session sheets) — generating drawing
parts + their rels/content-types alongside the preserved package was descoped from Phase 2d to
keep the spine tight; the refusal is typed and names the option (build charts in a separate
stock-mode workbook). The anchor-merge evidence from Q6 stands for the post-v0 lift. Original
decision text (now post-v0): new chart/image onto a preserved sheet WITHOUT an
existing drawing: supported — part names allocated `1 + max existing number per family` from
the retained namelist (never per-save session counters), one appended sheet-rels entry with
`rId = max numeric + 1`, `<drawing r:id>` spliced at its CT_Worksheet slot, targeted
[Content_Types] appends. Onto a sheet WITH a preserved drawing: **refuse** at
`add_chart`/`add_image` time, naming the drawing part and the options (anchor-merge is the
documented v0.5 lift; empirically tractable). Charts on new in-session sheets: supported.
Preserved drawings/images/charts are never re-serialized and never routed through
`Image._data()` (double-save crash).

**D10 — Raw compressed-stream copy guards (G8).** The fast path requires: no data-descriptor
(GP flag bit 3), no Zip64 local header, method DEFLATE or STORED, CPython-private `ZipFile`
attrs present (probed at import). Any miss → documented recompression fallback (0.35s vs
0.0015s on the large fixture — correctness identical).

**D11 — Rels and part resolution (G10).** Under preserve: all rels parts are append-only with
`rId = max numeric existing + 1`; never rebuilt or renumbered (stock renumbering measured to
dangle preserved r:ids). Touched parts are located via [Content_Types] + rels targets, never by
pattern-matching canonical paths.

**D12 — Content types come from the retained `[Content_Types].xml`**, edited by targeted
append/remove only — never regenerated from `wb.mime_type` (keys off `vba_archive` truthiness;
a preserve-loaded .xlsm without keep_vba would self-report XLSX).

**D13 — calcChain (Q12): active deletion cascade.** On any formula-affecting edit the splice
save removes `xl/calcChain.xml` AND its [Content_Types] override AND its workbook rel (all
three registrations proven load-bearing). No formula edits → calcChain raw-copies through.

**D14 — Lossy-save warning (Phase 2a stock + preserve fallback paths).** Content-level, not
part-list-level (zero parts are removed while sparklines are gutted — measured):
`LossySaveWarning` enumerates worksheet-extLst families by URI name, drawings carrying
non-chart/image anchors or mc:AlternateContent, chart auxiliary parts, VBA-without-keep_vba,
docProps/app.xml reset, unmodeled part families (pivots etc.). Fires on both save paths where
applicable; under preserve, stock's load-time "will be removed" extension warning is suppressed
(it becomes false — the splice preserves extLst).

**D15 — Region write-policy matrix (G5: merges Q2 Tier-2 with Q5's gates).** For a
Tier-2-snapshot region that the save finds changed:

| Region | Write mechanism | Refusal gate |
|---|---|---|
| mergeCells, dataValidations, hyperlinks, autoFilter, sheetProtection, printOptions, pageMargins, rowBreaks/colBreaks, sheetViews, cols, sheetFormatPr, sheetPr, tableParts | splice-replace whole element from model | element (or sheetViews/autoFilter level) carries extLst → refuse; DV range intersecting x14 DV `xm:sqref` → refuse |
| conditionalFormatting (classic) | splice-replace the touched block only, loud warning on xr:uid drop | touched block has cfRule-level extLst (x14 twin pointer) or intersects x14 CF `xm:sqref` → refuse |
| pageSetup | splice-replace with original `r:id` carried over | — |
| headerFooter | splice-replace, empty-children emission suppressed | — |
| drawing, legacyDrawing, anything in extLst, protectedRanges, phoneticPr, customSheetViews, cellWatches, ignoredErrors, smartTags, picture, oleObjects, controls, mc:AlternateContent | passthrough only | any edit → refuse |

A read-only extLst perception pass (URI inventory, `xm:sqref`/`x14:id` extraction) powers the
gates and the Phase-4 confession block.

**D16 — Volatile table (§3.7, restated verbatim, no deviations):** `NOW`, `TODAY`, `RAND`,
`RANDBETWEEN` (and their downstream cells) excluded from certification comparison, reported
separately; `INDIRECT`/`OFFSET` stay in. Tolerance: relative 1e-9, absolute floor 1e-11; text
and error values exact.

**D17 — Oracle driver rules (Q10, all measured):** per-invocation unique
`-env:UserInstallation` profile; success = rc 0 AND output exists; never parse stderr; timeout
kills the process group and raises `OracleTimeoutError`; detection `which("soffice")` →
`which("libreoffice")` → macOS app-bundle fallback → `OracleUnavailableError`; temp copies
only (tested invariant); "cached value absent" includes EMPTY `<v></v>`; LO output is an
answer key, never bytes to splice.

## 7. Oracle API (Phase 5)

```python
from openpyxl import oracle

oracle.available() -> bool
oracle.recalc(path, *, output_path=None, in_place=False, timeout=120.0) -> RecalcResult
    # exactly one of output_path/in_place may direct the result; neither ->
    # recalc scan only, nothing written. RecalcResult.to_dict() ->
    # {"schema": "oracle_recalc", "version": 1, "status": "ok|error",
    #  "cells_scanned": n, "formula_cells": n, "error_cells": n,
    #  "errors": [{"sheet": ..., "cell": ..., "value": ...}]}   # skill-compatible shape
oracle.certify(path, *, timeout=120.0) -> CertificationResult
    # .status in {"CERTIFIED", "DIVERGED", "BASELINE_UNVERIFIABLE"}
    # .divergences [{"address", "cached", "computed"}], .volatile_excluded [...]
    # measurements, never judgments; pre-flights on an untouched TEMP copy
```

## 8. Structural edits (Phase 6)

- **6a:** under preserve, `insert_rows/insert_cols/delete_rows/delete_cols` refuse
  (`UnsupportedStructureError`, message naming every stranded reference: formulas incl.
  cross-sheet, defined names, CF/DV ranges, merges, tables, chart series over preserved
  charts). Stock path: stock behavior + `LossySaveWarning`-class warning. Justification
  artifact: the measured 1100/6399/5400-vs-7499/6500 wrong-number set.
- **6b:** a new position-aware, `$`-insensitive, range-expanding rewriter (Excel INSERT
  semantics — Translator is fill-semantics and stays untouched) over tokenizer operands,
  applied as ledger entries. Under preserve, a successful structural edit returns an
  `AddressRemap` (`.map("Model!B12") -> "Model!B13"`; pre-edit addresses must be re-resolved,
  CONVENTIONS §2); the stock path keeps returning `None`.
- **6c:** chart series-range rewriting as a targeted patch inside preserved chart parts,
  lifting the 6a chart refusal. Scoped honestly; if it slips, the refusal stands.

## 9. G9 measurements (recorded)

1. Post-hoc in-session chart mutation → `xl/charts/chartN.xml` only (measured).
2. `add_pivot` exists (`ws._pivots`); pivot creation from scratch is impossible in openpyxl —
   under preserve, loaded pivot parts raw-copy through untouched.
3. `ws._rels.append` + stock save → appended rel silently DISCARDED (rels rebuilt wholesale).
   Not a functioning public mutation path today; `mark_dirty` territory under preserve.
4. `wb.code_name` → workbook.xml (Tier-2 covered); `wb.template=True` → [Content_Types]
   (D12 policy handles it).

## 10. Amendments to pinned shapes (consolidated register)

1. §3.5 collateral set → per-operation-class table (D2); sharedStrings removed (D1);
   calcChain cascade spelled out (D13); styles.xml added for style-bearing edits.
2. §3.2 wording → retention is the compressed whole-file blob; "≈ file size" holds for it
   (decompressed retention measures 7×).
3. §7 atomic-rename → dual-mode save targets (D3/G4); no-op invariant preserved via D3's
   no-auto-stamp rule.
4. PLAN §C "echo events verbatim" → byte-span splice (the only literally-verbatim mechanism;
   §3.4's invariant unchanged).
5. PLAN damage-table row 1 / battery jobs 1–2 wording → coverage-gated loss claims
   (already encoded in the Phase-1 battery).
6. D7 shared-formula default, settled by the Batch-0 item-zero probe (PLAN-v0.1 §0.1):
   the default is UN-SHARE (dissolve-on-touch — any edit intersecting a group's observed
   `si=` members re-emits the whole group as model-translated plain formulas), REFUSE for
   orphan followers and ref-less hosts (group extent named, atomic). Probed adversarially:
   master/follower/literal/delete edits, two-group isolation (untouched group byte-verbatim),
   gap cells in stale refs, followers beyond stale refs. Known lossy side effect, enumerated
   not hidden: dissolved untouched members lose cached `<v>` values; mitigations are the
   auto-set fullCalcOnLoad flag (verified), honest certify() unverifiable reporting, and
   recalc(); write-back (PLAN-v0.1 §5.3) is the cure. Battery job 6: correct.
7. Batch-1 surface widenings (PLAN-v0.1 1.2/1.6/1.7): (a) D2's cell-edit collateral
   set gains xl/workbook.xml (calcPr) when a VALUE edit intersects the dependency
   sketch of any formula — the recalc-on-load honesty flag now covers the most
   common agent edit, not just formula-text changes; (b) the workbook_manifest
   schema gains per-sheet "protection" and workbook-level "workbook_protection"
   booleans, and a new pinned warning ProtectedWriteWarning + wb.strict_protection
   flag implement protection awareness (reported, never enforced or bypassed);
   (c) the oracle_certification schema gains "external_excluded" and
   "unsupported_excluded" (excluded-with-reason, like volatiles, with downstream
   taint inheritance) so DIVERGED keeps meaning genuine disagreement; the
   ORACLE_UNSUPPORTED_FUNCS catalog is pinned data in oracle.py.
