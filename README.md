<div align="center">
  <a href="https://github.com/paper-instruments/paper-xlsx">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/paper-instruments/paper-xlsx/main/.github/assets/logo-dark.svg">
      <img alt="paper-xlsx logo" src="https://raw.githubusercontent.com/paper-instruments/paper-xlsx/main/.github/assets/logo-light.svg" height="128">
    </picture>
  </a>
  <h1>paper-xlsx</h1>

[![PyPI](https://img.shields.io/pypi/v/paper-xlsx.svg)](https://pypi.org/project/paper-xlsx/)
[![Python versions](https://img.shields.io/pypi/pyversions/paper-xlsx.svg)](https://pypi.org/project/paper-xlsx/)
[![Test](https://github.com/paper-instruments/paper-xlsx/actions/workflows/test.yml/badge.svg)](https://github.com/paper-instruments/paper-xlsx/actions/workflows/test.yml)

</div>

**An import-compatible, agent-safe fork of openpyxl designed to prevent silent loss during supported edits to existing Excel files.**

`paper-xlsx` is an import-compatible hard fork of [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl) 3.1.5 for safely inspecting, editing, and verifying existing Excel workbooks. It keeps everything that makes openpyxl excellent — the reader, the object model, the formula tokenizer, fifteen years of absorbed producer quirks — and adds a preserve-mode save path that retains package content openpyxl does not model.

```python
import openpyxl   # the import name is unchanged — see "Drop-in by design"
```

Under the default **preserve mode**, the original file's bytes are the source of truth. Edits are spliced into those bytes surgically; everything untouched survives byte-identical *by construction*, not by coverage. Within the supported preserve-mode surface, an operation either completes correctly or refuses with a typed error before delivery. Protection writes can also emit an explicit advisory warning.

---

## Why paper-xlsx exists

openpyxl is the de facto standard Python Excel library — it is what pandas uses under `read_excel` and `ExcelWriter` — and its object model is genuinely good. The problem is its persistence core: on save it regenerates the entire file from its in-memory model, so anything it does not fully model is degraded or dropped. Its own documentation says so:

> openpyxl does currently not read all possible items in an Excel file so shapes will be lost from existing files if they are opened and saved with the same name.
>
> — openpyxl tutorial (`doc/tutorial.rst`)

openpyxl's historic failure mode is **the file that opens fine and is quietly wrong**. The failures this fork was built to kill, each verified against openpyxl 3.1.5 before the fork was designed:

- **`data_only=True` + save permanently destroys every formula.** The load gives you cached values instead of formulas; saving writes those values back as the file. Measured: a sheet with 3 formulas round-trips to 0 — only literals remain.
- **`insert_rows` / `delete_rows` move cells and update nothing.** Not formulas, not defined names, not chart ranges. One inserted row leaves every downstream `SUM`, name, and cross-sheet reference pointing at the wrong cells — and the recalculated numbers *look plausible*, which is what makes this the most dangerous failure in the list.
- **Written formulas carry stale or empty cached values.** openpyxl never calculates, so a formula it writes has no result value, and a formula whose inputs it changed keeps the old one. Any pipeline (or human) reading the file trusts a number the file no longer justifies.
- **VBA is stripped from `.xlsm` files** unless you remembered `keep_vba=True`.
- **Unmodeled and half-modeled drawing content dies deterministically.** Shapes and textboxes, drawing `mc:AlternateContent`, chart-internal `extLst`, chart auxiliary parts, and worksheet extension lists (sparklines, x14 conditional formatting and validations) are lost on a plain load+save. Charts that openpyxl *can* parse are re-read and regenerated — lossily — rather than deleted outright; the fully-modeled basics (merges, classic conditional formatting, data validations, comments, hyperlinks) do survive. Stock openpyxl is genuinely good at everything it fully models. The carnage is exactly the unmodeled and half-modeled set — which is what real Excel files are full of.

Humans catch these failures by eyeballing the file. An agent editing a workbook programmatically cannot. It needs edits to either work, refuse loudly, or warn precisely — as typed, machine-readable outcomes. That is what this fork provides.

## Quick start

```bash
pip uninstall -y openpyxl paper-xlsx   # required: see "Drop-in by design"
pip install paper-xlsx
paper-xlsx-doctor                       # verify the install is coherent
```

> [!IMPORTANT]
> The PyPI distribution is `paper-xlsx`, but you still write `import openpyxl` — never `import paper_xlsx`. **Do not install `openpyxl` and `paper-xlsx` in the same environment**: both distributions own the same import tree, and package managers cannot safely arbitrate that. Always uninstall both first, then install `paper-xlsx`.

Build a small model, reload it, and make a safe edit with a machine-readable receipt:

```python
from openpyxl import Workbook, load_workbook

wb = Workbook()
ws = wb.active
ws["A1"], ws["B1"] = "Growth rate", 0.05
ws["A2"], ws["B2"] = "Revenue", 1000
ws["B3"] = "=B2 * (1 + B1)"
wb.save("model.xlsx")

wb = load_workbook("model.xlsx")        # preserve mode: on by default

wb.sheetnames                           # inspect structure directly
wb.active["B1"] = 0.07                  # ordinary openpyxl cell API
receipt = wb.save("model_v2.xlsx", receipt=True)
receipt.to_dict()["cells_changed"]      # {'xl/worksheets/sheet1.xml': {'B1': 'changed'}}
```

When an edit cannot be made safely, you get a typed refusal instead of a corrupted file — and a refused operation leaves the model, the ledger, and the disk exactly as they were:

```python
from openpyxl.errors import PaperRefusal

try:
    wb.active.insert_rows(1)             # may strand an unmodeled reference
except PaperRefusal as err:
    err.kind, err.anchor, err.options   # machine-readable: what, where, remedies
```

## What we changed from openpyxl — and why

Every claim below is traceable to a commit in this repository. The fork point is [`021192cf`](https://github.com/paper-instruments/paper-xlsx/commit/021192cf264012d7b5dba537f9994ee3f59ff223), on top of upstream openpyxl 3.1.5 (tagged `paper-base`). CI runs the upstream test suite alongside Paper's contract tests to catch compatibility regressions. Preserve mode becoming the default remains a deliberate behavioral change.

### Added

- **Preserve mode** ([`021192cf`](https://github.com/paper-instruments/paper-xlsx/commit/021192cf264012d7b5dba537f9994ee3f59ff223), the ~31,000-line bootstrap): `load_workbook(path)` retains the original archive bytes, a **dirty ledger** wired into openpyxl's setters records supported semantic mutations, and save **splices** only the dirty byte ranges into the original parts instead of regenerating files. Untouched parts are raw-copied byte-identical; a no-op save produces a byte-identical file. The machinery lives in the new [`openpyxl/preserve/`](https://github.com/paper-instruments/paper-xlsx/tree/main/openpyxl/preserve) tree (30 modules: ledger, splice writer, cross-part discipline, deterministic atomic zip I/O, and friends).
- **A typed refusal taxonomy** ([`openpyxl/errors.py`](https://github.com/paper-instruments/paper-xlsx/blob/main/openpyxl/errors.py)): `PaperRefusal` and its subclasses — `AmbiguousTargetError`, `TargetNotFoundError`, `UnsupportedStructureError`, `BoundaryViolationError`, `RelationshipPolicyError`, `OracleUnavailableError`, `OracleTimeoutError`. Guarded operations validate before committing mutations; a refused operation changes nothing in memory or on disk. Each refusal carries machine-readable `kind`, `anchor`, and `options` fields.
- **The oracle** ([`openpyxl/oracle.py`](https://github.com/paper-instruments/paper-xlsx/blob/main/openpyxl/oracle.py)): openpyxl never calculates, and this fork deliberately ships **no formula engine**. `oracle.recalc()`, `oracle.certify()`, `oracle.evaluate()`, and `oracle.evaluate_many()` use a headless, profile-isolated LibreOffice process working on temporary copies. `recalc(source)` returns narrow calculation/error evidence without writing. `recalc(source, output_path=...)` writes only a separate Paper-preserved candidate: eligible LibreOffice-calculated caches are spliced into the original package structure, the source stays untouched, and full recalculation remains requested. Recalculation and evaluation status say only whether recognized formula errors were detected; they do not claim Excel equivalence or financial correctness.
- **Targeted inspection helpers**: `wb.search(...)`, `ws.allowed_values(cell)`, `openpyxl.preserve.scan_errors()`, and `diff_workbooks()`. `allowed_values()` reports literal lists and deterministic static one-dimensional ranges, or raises a typed refusal when the source cannot be represented exactly. `scan_errors()` reports cached error values and actual formula error operands, not matching text inside string literals. These helpers supplement ordinary workbook objects without guessing workbook roles or mutation targets.
- **Guarded structural edits**: row and column insertion and deletion on loaded preserve-mode sheets rewrite supported dependent formulas, defined names, print areas, table ranges, and chart references — or refuse before mutation — and return an `AddressRemap`. Sheet renames rewrite supported dependencies through normal title assignment and return no remap. `move_range()` tracks the moved cells and refuses intersections or outside references it cannot keep coherent; like upstream, it returns `None` (shipped across 0.1.2, see [`doc/changes.rst`](https://github.com/paper-instruments/paper-xlsx/blob/main/doc/changes.rst)).
- **Narrow mutation helpers**: `ws.append_table_row(...)` expands a supported named worksheet table atomically. It preflights the retained table XML, relationships, geometry, formulas, totals, filters, styles, formats, merged/spill regions, and protection state. A loaded table without retained preserve-mode source, or a connected, extended, sorted, or otherwise unsupported table, refuses before mutation. A refusal is final: do not bypass it with generic row insertion, `preserve=False`, or raw package editing. `openpyxl.preserve.copy_format(...)` applies one complete cell style through a range-local transaction, and `Chart.repoint(...)` validates the complete chart patch before changing the model. `ws.replace_image(...)` retargets one loaded image relationship without changing its anchor, and `wb.set_pivot_refresh_on_load(pivots=[...])` changes only selected pivot-cache refresh metadata. Path saves use fsync-before-rename and ZIP integrity validation.
- **`paper-xlsx-doctor`**: a console script that verifies the installed distribution actually owns the `openpyxl` import tree.

### Changed

- **Preserve mode became the default** ([`6c2b99f7`](https://github.com/paper-instruments/paper-xlsx/commit/6c2b99f7), 0.1.3 — a deliberate breaking change: *"editable OOXML loads now use preserve mode unless preserve=False is passed explicitly"*). Filesystem paths are classified by supported OOXML suffix. Seekable file-like sources are classified from their ZIP container and workbook content type, regardless of a missing or misleading filename. `read_only=True` loads keep stock behavior. `preserve=False` is the explicit compatibility path.
- **Formula caches are invalidated instead of trusted** ([`a0b89793`](https://github.com/paper-instruments/paper-xlsx/commit/a0b89793) through [`1a16aa07`](https://github.com/paper-instruments/paper-xlsx/commit/1a16aa07), 0.1.3): when you edit a formula, or write a value into a cell that formulas read, the save strips the now-stale cached results from the file and sets the workbook to fully recalculate on open. The risk this closes: a human opens the edited file in Excel and silently trusts a stale number. The commit stream handles the ugly realities — array/spill formula followers, namespace-prefixed formula elements, whole-column array references — and style-only edits keep their caches untouched.
- **The stock path stays stock.** `preserve=False` does not run Paper ledgers, scanners, warnings, structural guards, or ZIP eligibility policies. It is the compatibility escape hatch for callers that deliberately want upstream openpyxl behavior.

### Archive validation

- **Integrity without package-defined eligibility caps.** Paper validates ZIP integrity but does not impose fixed entry-count, byte-size, or compression-ratio limits. Resource limits belong to the caller or execution environment.

## How it works

The architecture is a **spine transplant**. Stock openpyxl's object model holds two jobs: an in-memory representation of the grid (excellent — kept forever) and the source from which the entire file is regenerated at save (where losslessness dies — terminated). Under preserve mode, the original archive is the source of truth and the object model becomes a source of *edits to it*:

1. **Byte retention** — the load keeps every part of the original archive. Parts openpyxl never parses (drawings, VBA, pivot caches, media, custom XML) exist only as retained bytes.
2. **The dirty ledger** — instrumented chokepoints in openpyxl's own setters record every semantic mutation. A compare-based diff-save is impossible here, because openpyxl cannot serialize a faithful candidate to compare against — serialization *is* the lossy act. The ledger is the only honest record of what changed.
3. **The splice writer** — touched sheets are stream-patched: untouched byte ranges are copied verbatim, dirty cells are replaced at their exact scanned spans. Unmodeled XML passes through untouched *because it is never interpreted*. Untouched parts are raw-copied without recompression.
4. **Cross-part discipline** — every operation class has a sanctioned set of parts it may touch, enforced in tests by an exact changed-part budget: the package diff must show exactly the expected parts changed and every other part byte-identical.

| A real workbook, edited and saved | stock openpyxl 3.1.5 | paper-xlsx preserve mode |
|---|---|---|
| Shapes, textboxes, `mc:AlternateContent` | silently dropped | survive byte-identical |
| Charts | regenerated from the model; chart `extLst` and auxiliary parts lost | untouched charts survive byte-identical; title and series-range edits are spliced |
| Sparklines, x14 validations/formatting | dropped (load-time warning) | survive byte-identical |
| VBA project in `.xlsm` | stripped unless `keep_vba=True` | survives |
| `data_only=True`, then save | every formula replaced by its cached value, silently | refused unless explicitly allowed |
| `insert_rows` / `delete_rows` | cells move, references don't — silent corruption | references rewritten (or edit refused); `AddressRemap` returned |
| Formula caches after an edit | stale values left in the file | invalidated; full recalc forced on open |
| Unsafe or ambiguous operation | best guess, silently | typed `PaperRefusal`, atomic |

The full preserve-mode guide — including the refusal taxonomy, receipts, the oracle, and delivery — is in [`doc/paper.rst`](https://github.com/paper-instruments/paper-xlsx/blob/main/doc/paper.rst).

## Evaluation scenarios and regression coverage

The fork is developed against **15 realistic spreadsheet-editing scenarios**. The recorded agent evaluation used four package-plus-guidance treatments: two with stock `openpyxl==3.1.5` and two with `paper-xlsx==0.1.3`. Because the package and agent guidance changed together, those results are workflow evidence, not a package-only causal comparison. Hidden graders reopen each produced package independently and verify not just the requested edit but the collateral: formulas, names, tables, charts, drawings, review content, caches, and relationships. The scenarios continue to inform the 0.2.0 regression suite and include:

- repairing one member of a shared-formula block without disturbing its siblings or a neighboring array formula,
- fixing a formula without dropping the workbook's x14 validation dropdowns,
- renaming a sheet with every dependent formula, defined name, print area, and chart reference rewritten,
- updating an input in a macro-enabled `.xlsm` and delivering it with the VBA project intact,
- retargeting a chart series without disturbing sibling images or drawing anchors,
- expanding a named table and editing review content while preserving formulas, drawings, and relationships.

We don't publish aggregate pass rates here: results are tracked in the internal evaluation harness, not in this repository, and this README doesn't quote numbers you can't check.

The library's test discipline is documented in [CONTRIBUTING.md](https://github.com/paper-instruments/paper-xlsx/blob/main/CONTRIBUTING.md): the upstream pytest suite provides compatibility coverage, while Paper's persistence tests use saved-and-reopened assertions, exact changed-part budgets, refusal-atomicity checks, a provenance-labelled fixture corpus, and a headless LibreOffice load smoke where those checks apply.

## Drop-in by design

The import name `openpyxl` is **frozen forever**. `import openpyxl` appears in millions of scripts, in pandas itself, and in every model's training prior — so the fork keeps the import and renames only the distribution, the same split as Pillow (`pip install pillow`, `import PIL`).

- PyPI distribution / GitHub repository: **`paper-xlsx`**
- Python import: **`openpyxl`** — never `import paper_xlsx`, anywhere
- Fork sentinel: `openpyxl.__paper_version__` (`"0.2.0"`)
- Upstream base: openpyxl **3.1.5** (tag `paper-base`; upstream releases are merged, not rebased)

Note the two version numbers: `paper-xlsx` is versioned independently (currently **0.2.0**, early and pre-1.0) while `openpyxl.__version__` reports the upstream base (**3.1.5**) it wraps. pandas flows through this fork automatically — preserve-by-default covers files pandas opens for editing, and fresh `ExcelWriter` output is untouched stock behavior. Python **3.9–3.13** are supported and tested in CI, on Linux and Windows, with and without lxml.

Everything upstream openpyxl documents remains available. Preserve-by-default is the deliberate behavioral change; `preserve=False` restores the upstream-compatible load/edit/save path.

## Documentation

- [`doc/paper.rst`](https://github.com/paper-instruments/paper-xlsx/blob/main/doc/paper.rst) — the preserve-mode guide: loading and saving, formula-cache freshness, perception, editing, the oracle, delivery, the refusal taxonomy, and the compatibility opt-out. Ships inside the sdist.
- The remaining Sphinx docs are upstream openpyxl's, and everything they document still applies.

## Roadmap

paper-xlsx is pre-1.0 and its surface grows only as fast as the contract harness can prove it. Direction, clearly distinguished from what is shipped today:

- **Shrinking the refusal set**: preserve mode still refuses operations whose splice coverage isn't proven — for example chartsheet edits, generating table and pivot parts on newly added sheets, and comment changes on sheets that already carry comment parts (see the refusal sites in [`openpyxl/preserve/saver.py`](https://github.com/paper-instruments/paper-xlsx/blob/main/openpyxl/preserve/saver.py)). Each becomes supported as coverage lands.

Nothing on this list is presented as a current capability; when it ships, it appears in [`doc/changes.rst`](https://github.com/paper-instruments/paper-xlsx/blob/main/doc/changes.rst).

## Contributing

Contributions are welcome — see [CONTRIBUTING.md](https://github.com/paper-instruments/paper-xlsx/blob/main/CONTRIBUTING.md) for the engineering discipline this fork runs on. The short version: the upstream suite must remain green; persistence changes need saved-and-reopened assertions and exact package-delta checks; package-format regressions should use representative frozen fixtures; guarded refusals must be atomic; and new XML handling goes through openpyxl's `Serialisable` descriptor framework rather than string-formatted XML.

The most valuable non-code contribution right now is real-world fixtures: workbooks authored by desktop Excel or Google Sheets under the provenance rules in [CONTRIBUTING.md](https://github.com/paper-instruments/paper-xlsx/blob/main/CONTRIBUTING.md).

## Community

- **Bugs and feature requests**: [GitHub Issues](https://github.com/paper-instruments/paper-xlsx/issues)
- **Questions and ideas**: [GitHub Discussions](https://github.com/paper-instruments/paper-xlsx/discussions)

There is no Discord, Slack, or forum.

## Acknowledgments

paper-xlsx exists because openpyxl's object model and format coverage are excellent — we forked its persistence layer, not its judgment. Deep thanks to Eric Gazoni, Charlie Clark, and the openpyxl contributors (see [AUTHORS.rst](https://github.com/paper-instruments/paper-xlsx/blob/main/AUTHORS.rst)) for fifteen years of careful work this project stands on. Upstream openpyxl lives at [foss.heptapod.net/openpyxl/openpyxl](https://foss.heptapod.net/openpyxl/openpyxl).

If you reference this project in writing, cite it as *paper-xlsx* (Paper Instruments, Inc.), a fork of *openpyxl* by Eric Gazoni, Charlie Clark, and contributors, and link to this repository.
