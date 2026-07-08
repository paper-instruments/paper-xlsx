# ARCHITECTURE-NOTES — paper-xlsx Phase 0 source tour

**Status:** Phase 0 deliverable, 2026-07-07. Answers PLAN Phase 0 items 1–9 against the actual
tree (upstream 3.1.5 == `paper-base` == `c4986390b`). Every claim below was verified by reading
this tree or by a probe under `scratch/probes/` (gitignored; spikes never merge). Full per-topic
reports with verbatim probe output live in `scratch/results/*.md`. Companion document:
`OPEN-QUESTIONS.md` (the ten Phase-0 open questions, answered with evidence).

**How to read file:line citations:** all paths relative to repo root; lines verified on
`paper-base + 1` (the bootstrap commit does not touch package code).

---

## 1. Where the archive is opened at load, and where its bytes are discarded

- Open: `load_workbook` (`openpyxl/reader/excel.py:316-349`) → `ExcelReader.__init__`
  (`excel.py:121-130`) → `_validate_archive(fn)` (`excel.py:123`) → `ZipFile(filename, 'r')` at
  `excel.py:95`. File-like inputs already work (`excel.py:75`) — pandas relies on this.
- Discard: `ExcelReader.read()` ends with `if not self.read_only: self.archive.close()` at
  `excel.py:306-307`; the reader (sole holder of the archive) is garbage after `load_workbook`
  returns `reader.wb` (`excel.py:349`). **This is the retention hook point** — stash retained
  bytes in `read_workbook()` (`excel.py:150-170`) or just before the close.
- Precedents already in-tree:
  - `keep_vba=True` copies **the entire source archive** into an in-memory zip
    `wb.vba_archive` (`excel.py:162-165`); save re-emits only the VBA-family subset via regex in
    `ExcelWriter._merge_vba` (`openpyxl/writer/excel.py:96-110`). The fork's retention mechanism
    exists in embryo. Caveat: `vba_archive` truthiness flips `Workbook.mime_type` to
    macro-enabled (`openpyxl/workbook/workbook.py:360-370`) — preserve mode needs a NEW attribute.
  - `read_only=True` parks the **open** ZipFile as `wb._archive` (`excel.py:167-168`) — never
    reuse that name; `Workbook.close()` closes anything so named (`workbook.py:417-422`).
  - `wb.loaded_theme` is raw retained bytes written back verbatim (`excel.py:185-187`,
    `writer/excel.py:63-64`) — a per-part raw-retention precedent.
- Save side: `Workbook.save` → `save_workbook` → `ZipFile(filename,'w')` at
  `writer/excel.py:291` **truncates the target in place** (same inode, measured) before writing.
  Nothing on the save path reads the original file. Measured hazards of retaining a handle or a
  path instead of bytes: loud `BadZipFile` on large files, **silently stale bytes** on small
  files (8 KB buffer cache), silently-different-file on path re-read. Bytes-not-path
  (CONVENTIONS §3.2) is confirmed correct. Costs measured on `large.xlsx` (3.39 MB, 600k cells):
  whole-file-bytes retention +0.7 ms at load, memory = file size; see OPEN-QUESTIONS Q1.

## 2. The full path of a cell write; the ledger chokepoint inventory

- Canonical funnel: `ws['A1']=v` → `Worksheet.__setitem__` (`openpyxl/worksheet/worksheet.py:316-317`)
  → `Cell.value` setter (`openpyxl/cell/cell.py:215-218`) → `_bind_value` (`cell.py:176-203`).
  `ws.cell(row=, column=, value=)` and `ws.append` also route through it — but `append` inserts
  cells into `ws._cells` directly (`worksheet.py:674,681`), and `None` entries skip the setter.
- Styles: `cell.font = ...` → `StyleDescriptor.__set__` (`openpyxl/styles/styleable.py:22-26`),
  which interns the object in workbook-level `IndexedList`s and stores an index in `cell._style`.
  The getter returns a `StyleProxy` (`styleable.py:29-34`) that blocks top-level assignment but
  **leaks nested objects unproxied** (`styles/proxy.py:23-24`): `c.font.color.rgb=...` mutates
  the SHARED interned Font — restyles every aliased cell, corrupts the `IndexedList` hash
  registry, and mutates the process-wide `DEFAULT_FONT` singleton (`workbook.py:95-96`).
  Upstream bug class, and a ledger bypass that method instrumentation cannot see.
- NamedStyle aliasing: `cell.style="name"` **copies** the StyleArray (`styleable.py:87`).
  Post-hoc `ns.font = Font(...)` fires `NamedStyle.__setattr__._recalculate`
  (`styles/named_styles.py:74-79,97-109`) and changes **styles.xml only** — already-styled cells
  keep their old fontId in openpyxl's own model. Post-hoc in-place mutation (`ns.font.i=True`)
  bypasses the hook and restyles every aliased cell via the shared interned object.
- Row/column dimensions: `ws.row_dimensions[n]` / `ws.column_dimensions['A']` hand out mutable
  objects by reference (holder `worksheet.py:110-113`, descriptors
  `worksheet/dimensions.py:195-231`); no chokepoint fires on `.width=`. **Reads materialize
  entries that serialize** — `ws.row_dimensions[5]`, `ws['Z99']`, and `iter_rows` over empty
  regions all change stock saved output (measured). The ledger must key on semantic mutation,
  never on materialization.
- The full inventory (~35 clean Tier-1 funnels, ~14 by-reference bypass families, part→region
  mapping for each, verified by a 60-case part-payload diff probe with zero control noise) is in
  OPEN-QUESTIONS Q2 and `scratch/results/q2-chokepoints.md` +
  `scratch/results/v1-chokepoint-critic.md` (adversarial verification found five additional
  paths: `ws.sheet_properties` family, `cell.data_type` direct set, `wb.loaded_theme`,
  chartsheet satellites, post-hoc image/chart anchor mutation).

## 3. How the worksheet writer emits sheet XML; the event machinery

- `WorksheetWriter` (`openpyxl/worksheet/_writer.py:44-390`): temp-file backed (BytesIO
  supported), `write()` = `write_top()` (:90-103) + `write_rows()` (:120-127) + `write_tail()`
  (:303-351). `rows()` (:106-117) sorts `ws._cells` by (row, col) — already coordinate-ordered,
  matching the splice's two-sorted-streams design.
- Worksheet is NOT `Serialisable` — there is no `__elements__`; child-element order is
  procedural in `_writer.py`. Emitted order: sheetPr, dimension, sheetViews, sheetFormatPr, cols,
  sheetData, sheetProtection, scenarios, autoFilter, mergeCells, conditionalFormatting,
  dataValidations, hyperlinks, printOptions, pageMargins, pageSetup, headerFooter, rowBreaks,
  colBreaks, drawing, legacyDrawing, tableParts. Never emitted: sheetCalcPr, protectedRanges,
  sortState (deliberate no-op :168-173), customSheetViews, phoneticPr, cellWatches,
  ignoredErrors, smartTags, picture, oleObjects, controls, webPublishItems — **and extLst** (no
  writer code path exists). The splice writer must pin the CT_Worksheet child sequence as a
  constant to place newly-introduced elements correctly.
- Cell emission: `openpyxl/cell/_writer.py` — `_set_attributes` (:12-42), `etree_write_cell`
  (:45-86), `lxml_write_cell` (:89-130); fork point `write_cell = lxml_write_cell if LXML else
  etree_write_cell` (:133-136). LXML flag: `openpyxl/xml/__init__.py:26` (env `OPENPYXL_LXML`).
  Incremental writer: lxml's `xmlfile` or `et_xmlfile.xmlfile` 2.0.0 (`openpyxl/xml/functions.py:13-37`).
- **Spiked: both backends emit a single bare `<c>` fragment with no declaration and no prefix**
  (e.g. `b'<c r="A1" t="n"><v>42</v></c>'`) — directly spliceable into a default-namespace sheet
  stream. Caveats the splice wrapper must own: `write_cell` side-effects (`ws._hyperlinks`
  append at `cell/_writer.py:39-40`; `cell.style_id` mutates `wb._cell_styles`,
  `styleable.py:139-143`); lxml refuses a second bare root per stream; backend byte drift
  (`<v></v>` vs `<v/>`). See OPEN-QUESTIONS Q4 for why the splice is byte-range, not echo.

## 4. sharedStrings in memory and on disk; style index assignment

- **This tree never writes `xl/sharedStrings.xml`.** Every string cell is emitted
  `t="inlineStr"` (`cell/_writer.py:21-22`, `:70-79`, `:113-125`); there is no sharedStrings
  writer anywhere (`openpyxl/writer/` = excel.py + theme.py; `ARC_SHARED_STRINGS` is only a
  constant). `wb.shared_strings` is reader-only: `read_string_table`
  (`openpyxl/reader/strings.py:10-24`) builds a plain list in `<si>` order (count/uniqueCount
  never read); `t="s"` cells resolve positionally at `worksheet/_reader.py:227`. Resaving an
  sst-bearing file strips the part, its content-type override, and its workbook rel (measured).
  Consequence: the CONVENTIONS §3.5 inline-vs-shared decision is nearly self-answering — inline
  IS the stock path. Details and measurements: OPEN-QUESTIONS Q7.
- Styles: load via `apply_stylesheet` (`openpyxl/styles/stylesheet.py:199-240`);
  `wb._cell_styles` is an `IndexedList` of StyleArrays in file order (duplicates allowed).
  `_normalise_numbers` **rewrites numFmtIds at load** (:165-190). Save: `write_stylesheet`
  (:243-274) iterates `wb._cell_styles` in order; xf positions of loaded entries are stable and
  new xfs append, but numFmtId values renumber and cellXfs can GROW on a no-edit round-trip
  (measured 1→3). Append-only for preserve mode therefore means byte-splicing new `<xf>`/`<numFmt>`
  elements into the original styles.xml with count bumps — `write_stylesheet` cannot be the
  preserve path. Hazard: `IndexedList._rebuild_dict` returns wrong indices after `__contains__`
  on duplicate-seeded lists (`openpyxl/utils/indexed_list.py:23-30`, measured).

## 5. workbook.xml, calcPr, calcChain

- `WorkbookWriter` (`openpyxl/workbook/_writer.py:49-197`). calcPr emitted from
  `wb.calculation` (`CalcProperties`, `openpyxl/workbook/properties.py:82-131`; defaults
  `calcId=124519, fullCalcOnLoad=True` — a loaded calcPr lacking those attrs silently gains them
  on save, measured). Sheet entries: `sheetId` and `r:id` are **renumbered to 1..N on every
  save** (`_writer.py:70-79`); hidden state written at :75-78. Defined names are regenerated
  wholesale including synthesized `_FilterDatabase`/`Print_Titles`/`Print_Area` (:92-120).
- **calcChain: zero references in package source** — no reader, no writer; the part, its
  content-type override, and its rel vanish today only because the archive is rebuilt from
  scratch (proven by zip surgery). Preserve-mode raw copy inverts that default: the fork must
  delete calcChain **actively** on formula change — part + `[Content_Types].xml` override +
  workbook rel, all three (the collateral set needs this cascade spelled out; see
  OPEN-QUESTIONS "Cross-cutting gaps").
- Manifest/rels: `[Content_Types].xml` regenerated from scratch each save
  (`openpyxl/packaging/manifest.py:159-179`), written LAST so it can scan the final namelist.
  All rels parts are regenerated with positional insertion-order rIds
  (`openpyxl/packaging/relationship.py:53-62`) — original rId numbering is not preserved;
  adding one hyperlink renumbered an existing drawing rel (measured). Under preserve, sheet rels
  must become append-only with fresh unused rIds.

## 6. How charts/images added in-session are tracked and written

- Storage: `ws._charts`/`ws._images` lists (`worksheet.py:117-118`); `add_chart`/`add_image`
  only set `.anchor` and append (`worksheet.py:552-569`). All part names and rIds are minted at
  save: `drawing._id = len(self._drawings)` (`writer/excel.py:132-133`) → `drawing{N}.xml`;
  same pattern for charts and images. Counters restart at 1 on every save → guaranteed
  collisions with preserved parts of the same names; rIds are `rId{len+1}`
  (`relationship.py:59-62`) → collides with non-contiguous preserved ids (measured).
- **The premise "stock drops charts at load" is stale for 3.1.5**: `reader/excel.py:265-271`
  reads drawing rels into `ws._charts`/`ws._images` via `openpyxl/reader/drawings.py:21-71`,
  and the save path regenerates chart and drawing parts. On the openpyxl-authored trap fixture
  the chart part round-tripped byte-identical. What still dies, deterministically: shapes and
  textboxes (whole drawing dropped on TypeError, `drawings.py:30-34`), `mc:AlternateContent`
  inside drawings, chart-internal `extLst`, chart auxiliary parts (colors/style/embedded
  workbook — read into `.deps`, never written), worksheet `extLst` (sparklines/x14), and images
  without Pillow. The damage model is coverage-gated, not absolute; real-Excel fixtures are
  load-bearing for the "charts gone" battery claims (see FIXTURE-REQUESTS.md).
- One drawing per sheet is a hard invariant: the writer emits at most one `<drawing>` element
  (`worksheet/_writer.py:244-250`), and **LibreOffice silently drops a second one** (measured).
  Adding a chart to a preserved sheet that already has a drawing requires merging anchors INTO
  the preserved drawing part — empirically tractable (4-edit splice verified against both
  openpyxl and an `xdr:`-prefixed LibreOffice-authored drawing) but proposed as post-v0; v0
  refuses. Upstream bug found on the way: saving a workbook loaded with images TWICE crashes
  (`Image._data()` closes its BytesIO, `drawing/image.py:59`).

## 7. The Serialisable descriptor framework (house style, ten lines)

1. XML classes subclass `Serialisable` (`openpyxl/descriptors/serialisable.py:24`); the
   metaclass (`descriptors/__init__.py:21-58`) sorts class-level descriptors into `__attrs__`
   (XML attributes), `__nested__`, and `__elements__` (children).
2. Plain `Typed`/`Convertible` with non-Serialisable `expected_type` → attribute; `nested=True`,
   `Sequence`, or Serialisable `expected_type` → child element.
3. `Typed.__set__` enforces `isinstance` → `TypeError`; `allow_none=True` permits None
   (`descriptors/base.py:28-47`); `Convertible` coerces.
4. `tagname` names the element; `namespace` attaches Clark-style namespaces
   (`descriptors/namespace.py:4-12`); namespaced attributes (e.g. `r:id` via `Relation`) live in
   `__namespaced__`.
5. `from_tree` builds kwargs from `node.attrib` plus recursive child conversion, ends
   `cls(**attrib)` (`serialisable.py:47-103`).
6. `to_tree` emits attrs (None skipped) and children in `__elements__` order (:106-157).
7. `Sequence` collects repeated same-tag children; `NestedSequence` adds a container with
   optional `count`.
8. **Unknown children are silently skipped** (:77-79); **foreign-namespace attributes are
   silently deleted** (:58-61); **unknown plain attributes CRASH** with TypeError at
   `cls(**attrib)` (:103) — sometimes converted to warn-and-discard by callers
   (`worksheet/_reader.py:307-313`). This two-mode behavior (silent drop + hard crash) is the
   mechanistic root of the entire damage model.
9. `ExtensionList` models only `ext/@uri` (`descriptors/excel.py:58-75`): declared extLst that
   does round-trip is re-emitted as a **hollow shell** (`<ext uri="..."/>` with the payload
   destroyed, probe-proven) — worse than dropping it. Nothing in the framework retains raw XML.
10. Worked example — `Hyperlink` (`worksheet/hyperlink.py:10-36`): `ref = String()`,
    `tooltip = String(allow_none=True)`, `id = Relation()`, explicit `__attrs__` to exclude a
    non-persisted field. New paper-xlsx vocabulary follows exactly this pattern.

## 8. Formula tooling: Tokenizer and Translator

- Tokenizer (`openpyxl/formula/tokenizer.py`): operand classification is a fallback — plain,
  absolute, whole-column, sheet-qualified, external, defined-name, and structured/table
  references ALL tokenize as single OPERAND/RANGE tokens (`make_operand`, :372-387;
  bracket handling :123-146). It isolates references reliably; it does not distinguish kinds.
- Translator (`openpyxl/formula/translate.py:33-166`): fill/copy semantics — `$`-anchored parts
  are pinned (:65-66, :78-79), unmatched operands are presumed defined names and left alone
  (:130-132). Spiked +1 row: `=SUM(B2:B11)`→`=SUM(B3:B12)`; `=SUM($B$2:$B$11)`→unchanged;
  `=SUM(B:B)`→unchanged; `=Table1[Amount]`→unchanged; cross-sheet shifts work. It is
  production-load-bearing for shared-formula expansion at load (`worksheet/_reader.py:249-272`).
- **Semantic gap for Phase 6b:** Excel INSERT semantics shift `$`-anchored refs too, shift only
  refs at/below the insertion point, and EXPAND ranges spanning it. Naive Translator use would
  wrongly pin every `$`-anchored assumption and shift unconditionally. 6b reuses the Tokenizer's
  operand isolation + `strip_ws_name` and builds its own position-aware, `$`-insensitive,
  range-expanding rewriter beside Translator (which stays untouched).
- Addressing toolbox (`openpyxl/utils/cell.py`, probe-verified): `quote_sheetname` (:232-240,
  always quotes despite its docstring), `absolute_coordinate` (:58-73), `range_boundaries`
  (:139-179, rejects sheet-qualified), `range_to_tuple` (:218-229, requires `!`, does not
  un-escape doubled quotes), `coordinate_to_tuple` (:206-215, rejects `$` — its sibling
  `coordinate_from_string` accepts it). One layer up, `worksheet/cell_range.py` `CellRange` /
  `MultiCellRange` already parse sheet-qualified A1 and implement
  shift/expand/union/intersection/containment — the natural typed carrier for the pinned
  addressing. The PR-0 wrapper adds exactly two fixes: quote un-escaping and `$` tolerance.

## 9. Baseline vitals (recorded before any surgery)

- Upstream pytest suite, this machine, `.venv` Python 3.13.3 + lxml 5.4.0 + pandas 3.0.3:
  **2592 passed, 6 skipped, 7 xfailed in 2.77s** — matches PAPER.md's recorded fork-point
  baseline exactly (`scratch/results/baseline_pytest.txt`).
- Provenance verified: full hg-converted history (9,142 commits back to the 2010-04-09 initial
  import, 123 tags); `paper-base` == tag `3.1.5` == `c4986390b`; HEAD is exactly one commit
  ahead (fork-identity bootstrap: `__paper_version__`, packaging, CI, PAPER.md). PyPI's latest
  openpyxl is still 3.1.5 (2024-06-28) as of 2026-07-07, so the fork base is the current
  upstream stable release. The wheel ships `openpyxl/` as the import package (name rule holds).
- Performance seeds for the Phase-2 guardrail (large.xlsx, 3.39 MB, 600k cells, medians of 3):
  stock load 2.505 s / 301.6 MB max RSS; stock save 2.174 s; LibreOffice warm convert 2.09 s.

---

## Stale-premise register (where the plan's damage model needed correction)

Recorded here so nobody re-imports training-era assumptions; details in OPEN-QUESTIONS.md.

1. **"Charts/images deleted on load+save" is not literally true in 3.1.5** for
   openpyxl-parseable charts — they are read back and regenerated (lossily). The deterministic
   kills are: shapes/textboxes, drawing `mc:AlternateContent`, chart-internal extLst, chart
   auxiliary parts, worksheet extLst (sparklines/x14 — with a load-time warning), VBA without
   flag, docProps/app.xml, and unmodeled real-Excel content. Battery jobs 1–2 must claim chart
   loss on real-Excel/shape-bearing fixtures, and the Phase-2a lossy-save warning must be
   content-level, not part-list-level (zero parts are removed while sparklines are gutted).
2. **Stock never writes sharedStrings.xml** — inline strings are upstream's only string mode.
3. **Stock already relies on Excel rebuilding calcChain** (drop-by-omission today).
4. **Extension drop is warned at LOAD** for worksheet extLst (not silent, not save-time);
   under preserve mode the stock warning text becomes false and must be suppressed/rephrased.
5. **openpyxl's reader mutates the model at load** (`_clean_merge_range` writes borders/
   protection through chokepoints; `create_sheet` fires during load) — the ledger arms only
   after load completes, and load-time dirt is whitelisted.
6. Upstream bugs found by probes (candidates for fixture-backed fork patches, not silent fixes):
   StyleProxy nested-mutation registry corruption + process-wide `DEFAULT_FONT` contamination;
   double-save crash on image-bearing workbooks; `<col phonetic="1"/>` load crash;
   `mc:Fallback` colBreaks resurrected as real by depth-blind dispatch; `IndexedList`
   wrong-index bug; docProps/app.xml silently reset; write_only post-save exceptions contradict
   docstring (`StopIteration`/`FileNotFoundError`, not `WorkbookAlreadySaved`).
