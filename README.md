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

`paper-xlsx` is an import-compatible hard fork of [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl) 3.1.5 for safely inspecting, editing, and verifying existing Excel workbooks. It keeps openpyxl's reader, object model, formula tokenizer, and broad file-format support. It adds a preserve-mode save path that retains package content openpyxl does not model.

```python
import openpyxl   # the import name is unchanged; see "Drop-in by design"
```

Under the default **preserve mode**, the original file's bytes are the source of truth. paper-xlsx writes supported edits into the retained package and copies untouched parts byte-for-byte. If it cannot express an edit safely, it raises a typed error before saving. Writes to protected cells can also emit an advisory warning.

---

## Why paper-xlsx exists

openpyxl is a widely used Python Excel library, and pandas can use it for `.xlsx` files through `read_excel` and `ExcelWriter`. Its object model is strong, but its save path regenerates the entire file from that model. Content that openpyxl does not fully model can be degraded or removed. Its documentation says:

> openpyxl does currently not read all possible items in an Excel file so shapes will be lost from existing files if they are opened and saved with the same name.
>
> openpyxl tutorial (`doc/tutorial.rst`)

An `.xlsx` file can open normally even when an edit has damaged its contents. The results can still look reasonable, so automated checks may not detect the problem. These problems occur with openpyxl 3.1.5:

- **Saving after `data_only=True` can replace formulas with values.** `data_only=True` loads cached results instead of formulas. If you save that workbook, those results are written back as values. In a test with three formulas, none remained after the save.
- **`insert_rows()` and `delete_rows()` do not update related references.** Cells move, but formulas, defined names, and chart ranges can still point to the old locations. The resulting values can look plausible even when the references are wrong.
- **Formula results can be missing.** openpyxl does not calculate formulas. New or saved formulas can have empty cached results until a spreadsheet application recalculates them.
- **VBA can be removed from `.xlsm` files.** Callers must pass `keep_vba=True` to preserve it.
- **Unsupported drawing content can be lost during save.** Shapes, text boxes, sparklines, and newer validation or conditional-formatting extensions can be removed. Supported charts are regenerated, which can remove chart extensions and related parts. Content that openpyxl fully supports, such as merged cells, standard validations, comments, and hyperlinks, is preserved.

People can inspect a workbook after an edit. An automated agent cannot rely on visual review. paper-xlsx preserves supported content and returns a typed refusal when it cannot make an edit safely.

## Quick start

```bash
pip uninstall -y openpyxl paper-xlsx   # required: see "Drop-in by design"
pip install paper-xlsx
paper-xlsx-doctor                       # verify the install is coherent
```

> [!IMPORTANT]
> The PyPI distribution is `paper-xlsx`, but you still write `import openpyxl`, never `import paper_xlsx`. **Do not install `openpyxl` and `paper-xlsx` in the same environment**: both distributions own the same import tree, and package managers cannot safely arbitrate that. Always uninstall both first, then install `paper-xlsx`.

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
receipt.to_dict()["cells_changed"]
# {'xl/worksheets/sheet1.xml': {'B1': 'changed', 'B3': 'changed'}}
```

When an edit cannot be made safely, paper-xlsx raises a typed refusal instead of writing a corrupted file. A refused operation leaves the model, the ledger, and the disk exactly as they were:

```python
from openpyxl import load_workbook
from openpyxl.errors import PaperRefusal

values_only = load_workbook("model.xlsx", data_only=True)
try:
    values_only.save("model_values_only.xlsx")
except PaperRefusal as err:
    print(err)
    if err.kind is not None:
        print(err.kind, err.anchor, err.options)  # optional structured context
```

## What paper-xlsx adds

paper-xlsx is based on openpyxl 3.1.5. CI runs the upstream test suite alongside Paper's contract tests to catch compatibility regressions. Preserve mode is the default for editable supported OOXML workbooks; pass `preserve=False` to use upstream-compatible behavior.

### Preservation and guarded editing

- **Preserve mode**: `load_workbook(path)` retains the original archive bytes, a **dirty ledger** wired into openpyxl's setters records supported semantic mutations, and save **splices** only the dirty byte ranges into the original parts instead of regenerating files. Untouched parts are raw-copied byte-identical; a no-op save produces a byte-identical file. The machinery lives in [`openpyxl/preserve/`](https://github.com/paper-instruments/paper-xlsx/tree/main/openpyxl/preserve): ledger, splice writer, cross-part discipline, deterministic atomic ZIP I/O, and related preservation code.
- **A typed refusal taxonomy** ([`openpyxl/errors.py`](https://github.com/paper-instruments/paper-xlsx/blob/main/openpyxl/errors.py)): `PaperRefusal` and its subclasses (`AmbiguousTargetError`, `TargetNotFoundError`, `UnsupportedStructureError`, `BoundaryViolationError`, `RelationshipPolicyError`, `OracleUnavailableError`, `OracleTimeoutError`). Guarded operations validate before committing mutations; a refused operation changes nothing in memory or on disk. Refusals expose optional `kind`, `anchor`, and `options` fields. The exception message is always the complete explanation.
- **The oracle** ([`openpyxl/oracle.py`](https://github.com/paper-instruments/paper-xlsx/blob/main/openpyxl/oracle.py)): openpyxl never calculates, and this fork deliberately ships **no formula engine**. `oracle.recalc()`, `oracle.certify()`, `oracle.evaluate()`, and `oracle.evaluate_many()` use a headless, profile-isolated LibreOffice process working on temporary copies. `recalc(source)` returns narrow calculation/error evidence without writing. `recalc(source, output_path=...)` writes only a separate Paper-preserved candidate: eligible LibreOffice-calculated caches are spliced into the original package structure, the source stays untouched, and full recalculation remains requested. If those writes or that recalculation can affect a local pivot source, the candidate requests cache refresh-on-open and `RecalcResult.pivot_refreshes` reports that Excel requirement. Recalculation and evaluation status say only whether recognized formula errors were detected; they do not claim Excel equivalence or financial correctness.
- **Targeted inspection helpers**: `wb.search(...)`, `ws.allowed_values(cell)`, `openpyxl.preserve.scan_errors()`, and `diff_workbooks()`. `allowed_values()` reports literal lists and deterministic static one-dimensional ranges, or raises a typed refusal when the source cannot be represented exactly. `scan_errors()` reports cached error values and actual formula error operands, not matching text inside string literals. These helpers supplement ordinary workbook objects without guessing workbook roles or mutation targets.
- **Guarded structural edits**: row and column insertion and deletion on loaded preserve-mode sheets rewrite supported dependent formulas, defined names, print areas, table ranges, and chart references, or refuse before mutation, and return an `AddressRemap`. Sheet renames rewrite supported dependencies through normal title assignment and return no remap. `move_range()` tracks the moved cells and refuses intersections or outside references it cannot keep coherent; like upstream, it returns `None`.
- **Narrow mutation helpers**: `ws.append_table_row(...)` expands a supported named worksheet table atomically. It preflights the retained table XML, relationships, geometry, formulas, totals, filters, styles, formats, merged/spill regions, and protection state. A loaded table without retained preserve-mode source, or a connected, extended, sorted, or otherwise unsupported table, refuses before mutation. A refusal is final: do not bypass it with generic row insertion, `preserve=False`, or raw package editing. `openpyxl.preserve.copy_format(...)` applies one complete cell style through a range-local transaction, and `Chart.repoint(...)` validates the complete chart patch before changing the model. `ws.replace_image(...)` retargets one loaded image relationship without changing its anchor. `Worksheet.pivots.create(...)` builds a Paper-managed classic worksheet PivotTable with materialized output; `PivotTable.refresh()`, `repoint_source()`, `move()`, `update()`, `rename()`, and `delete()` rebuild or remove that dedicated-cache graph. Shared caches disable layout/update and the other isolation-sensitive verbs. Foreign pivots remain preserved and may only receive `wb.set_pivot_refresh_on_load(pivots=[...])` when qualified. `PivotTable.qualify_adoption()` is a read-only eligibility report. `PivotTable.adopt()` converts a qualified foreign pivot into the Paper-managed lifecycle: dedicated caches are replaced in place, and shared caches are isolated when safe. Eligibility stays gated on desktop Excel evidence; this is not edit-any-PivotTable support. Formula-backed pivot sources may use stock LibreOffice for values; LibreOffice-rewritten workbooks are never published. Desktop Excel is not a runtime dependency. Edits that change an existing pivot's local source, touch that source, transitively affect its formulas (including calculation-relevant cell formatting, row/column display state, and filtering), or accompany a known volatile built-in in the source refuse at validation/save unless that pivot is explicitly refreshed or selected for refresh. Exact direct ranges, static defined names, and named-table sources are recognized; unresolved local sources refuse conservatively. Runtime volatility of user-defined functions is not declared in OOXML and is not inferred; callers using a UDF in a pivot source must explicitly request refresh. The receipt reports that Excel must refresh the cache on open and that headless readers may still see its old results. Path saves are written to a temporary file, fsynced, and atomically moved into place.

```python
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table

wb = Workbook()
data = wb.active
data.title = "Data"
data.append(("Region", "Amount"))
data.append(("East", 10))
data.append(("West", 20))
data.add_table(Table(displayName="RegionTable", ref="A1:B3"))
wb.create_sheet("Summary")
wb.save("sales.xlsx")

wb = load_workbook("sales.xlsx", preserve=True)
pivot = wb["Summary"].pivots.create(
    name="ByRegion",
    source="RegionTable",
    destination="A1",
    rows=["Region"],
    values=["Amount"],
)
wb["Data"]["B2"] = 99
pivot.refresh()
wb.save("sales-out.xlsx")
```
- **`paper-xlsx-doctor`**: a console script that verifies the installed distribution actually owns the `openpyxl` import tree.

### Default behavior and compatibility

- **Preserve mode is the default.** Editable supported OOXML workbooks load in preserve mode unless `preserve=False` is passed. Filesystem paths are classified by supported OOXML suffix. Seekable file-like sources are classified from their ZIP container and workbook content type, regardless of a missing or misleading filename. `read_only=True` loads retain upstream behavior, and `preserve=False` selects the upstream-compatible load/edit/save path.
- **Formula caches are invalidated instead of trusted.** When you edit a formula, or write a value into a cell that formulas read, save strips the now-stale cached results from the file and sets the workbook to fully recalculate on open. This prevents a human from opening the edited file in Excel and silently trusting a stale number. The implementation handles array/spill formula followers, namespace-prefixed formula elements, and whole-column array references; style-only edits keep their caches untouched.
- **The stock path stays stock.** `preserve=False` does not run Paper ledgers, scanners, warnings, structural guards, or ZIP eligibility policies. It is the compatibility escape hatch for callers that deliberately want upstream openpyxl behavior.

### Archive validation

- **Integrity without package-defined eligibility caps.** Paper validates ZIP integrity but does not impose fixed entry-count, byte-size, or compression-ratio limits. Resource limits belong to the caller or execution environment.

## How it works

Preserve mode separates the in-memory workbook model from the saved package. Stock openpyxl uses its object model both to represent the workbook and to regenerate the file during save. paper-xlsx keeps the original archive as the source of truth and uses the object model to describe edits:

1. **Byte retention**: the load keeps every part of the original archive. Content that openpyxl does not parse or fully represent remains available in the retained bytes. This includes content in drawings, VBA projects, pivot caches, media, and custom XML.
2. **The dirty ledger**: instrumented chokepoints in openpyxl's own setters record every semantic mutation. A compare-based diff-save is impossible here, because openpyxl cannot serialize a faithful candidate to compare against; serialization is the lossy act. The ledger records what changed.
3. **The splice writer**: touched sheets are stream-patched. Untouched byte ranges are copied verbatim; dirty cells are replaced at their exact scanned spans. Unmodeled XML passes through untouched because it is never interpreted. Untouched parts are raw-copied without recompression.
4. **Cross-part discipline**: every operation class has a sanctioned set of parts it may touch, enforced in tests by an exact changed-part budget. The package diff must show exactly the expected parts changed and every other part byte-identical.

| A real workbook, edited and saved | stock openpyxl 3.1.5 | paper-xlsx preserve mode |
|---|---|---|
| Shapes, textboxes, `mc:AlternateContent` | silently dropped | survive byte-identical |
| Charts | regenerated from the model; chart `extLst` and auxiliary parts lost | untouched charts survive byte-identical; title and series-range edits are spliced |
| Sparklines, x14 validations/formatting | dropped (load-time warning) | survive byte-identical |
| VBA project in `.xlsm` | stripped unless `keep_vba=True` | survives |
| `data_only=True`, then save | every formula replaced by its cached value, silently | refused unless explicitly allowed |
| `insert_rows` / `delete_rows` | cells move, references don't; silent corruption | references rewritten (or edit refused); `AddressRemap` returned |
| Formula caches after an edit | cached results can be cleared; formulas are not calculated | affected caches invalidated; full recalc requested on open |
| Unsafe or ambiguous operation | best guess, silently | typed `PaperRefusal`, atomic |

The full preserve-mode guide, including the refusal taxonomy, receipts, the oracle, and delivery, is in [`doc/paper.rst`](https://github.com/paper-instruments/paper-xlsx/blob/main/doc/paper.rst).

## Testing

The test suite covers realistic spreadsheet edits, including:

- repairing one member of a shared-formula block without disturbing its siblings or a neighboring array formula,
- fixing a formula without dropping the workbook's x14 validation dropdowns,
- renaming a sheet with every dependent formula, defined name, print area, and chart reference rewritten,
- updating an input in a macro-enabled `.xlsm` and delivering it with the VBA project intact,
- retargeting a chart series without disturbing sibling images or drawing anchors,
- expanding a named table and editing review content while preserving formulas, drawings, and relationships.

The library's test discipline is documented in [CONTRIBUTING.md](https://github.com/paper-instruments/paper-xlsx/blob/main/CONTRIBUTING.md): the upstream pytest suite provides compatibility coverage, while Paper's persistence tests use saved-and-reopened assertions, exact changed-part budgets, refusal-atomicity checks, a provenance-labelled fixture corpus, and a headless LibreOffice load smoke where those checks apply.

## Drop-in by design

The fork keeps the `openpyxl` import name so existing code does not need new imports. Only the distribution name changes, similar to Pillow (`pip install pillow`, `import PIL`).

- PyPI distribution / GitHub repository: **`paper-xlsx`**
- Python import: **`openpyxl`**, never `import paper_xlsx`, anywhere
- Fork sentinel: `openpyxl.__paper_version__`
- Upstream base: openpyxl **3.1.5**

`paper-xlsx` is versioned independently from its upstream base. `openpyxl.__paper_version__` reports the installed paper-xlsx distribution version, while `openpyxl.__version__` reports the upstream base version. pandas workflows that use the openpyxl engine use this fork automatically. Preserve mode applies when pandas opens an existing file for editing; a new `ExcelWriter` workbook uses the standard creation path. Python **3.9–3.13** are supported and tested on Linux. CI also tests Python 3.13 on Windows and with and without lxml.

The upstream openpyxl APIs remain available. Preserve-by-default changes save behavior for existing editable OOXML workbooks; `preserve=False` restores the upstream-compatible load/edit/save path.

## Documentation

- [`doc/paper.rst`](https://github.com/paper-instruments/paper-xlsx/blob/main/doc/paper.rst): the preserve-mode guide, covering loading and saving, formula-cache freshness, perception, editing, the oracle, delivery, the refusal taxonomy, and the compatibility opt-out. Ships inside the sdist.
- The remaining Sphinx docs cover the upstream openpyxl APIs. Use `preserve=False` when you need upstream save behavior.

## Current limitations

Preserve mode refuses operations that it cannot save safely. Current examples include chartsheet edits, table or pivot creation on newly added sheets, and comment changes on sheets that already contain comment parts. See the refusal sites in [`openpyxl/preserve/saver.py`](https://github.com/paper-instruments/paper-xlsx/blob/main/openpyxl/preserve/saver.py).

## Contributing

See [CONTRIBUTING.md](https://github.com/paper-instruments/paper-xlsx/blob/main/CONTRIBUTING.md) for the engineering discipline this fork runs on. The short version: the upstream suite must remain green; persistence changes need saved-and-reopened assertions and exact package-delta checks; package-format regressions should use representative frozen fixtures; guarded refusals must be atomic; and new XML handling goes through openpyxl's `Serialisable` descriptor framework rather than string-formatted XML.

Useful non-code contributions include real-world fixtures authored by desktop Excel or Google Sheets under the provenance rules in [CONTRIBUTING.md](https://github.com/paper-instruments/paper-xlsx/blob/main/CONTRIBUTING.md).

## Community

- **Bugs and feature requests**: [GitHub Issues](https://github.com/paper-instruments/paper-xlsx/issues)
- **Questions and ideas**: [GitHub Discussions](https://github.com/paper-instruments/paper-xlsx/discussions)

## Acknowledgments

paper-xlsx exists because openpyxl's object model and format coverage are excellent. Thanks to Eric Gazoni, Charlie Clark, and the openpyxl contributors (see [AUTHORS.rst](https://github.com/paper-instruments/paper-xlsx/blob/main/AUTHORS.rst)) for the work this project builds on. Upstream openpyxl lives at [foss.heptapod.net/openpyxl/openpyxl](https://foss.heptapod.net/openpyxl/openpyxl).

## Citation

If you reference paper-xlsx in research or writing:

```bibtex
@software{paper_xlsx,
  title   = {paper-xlsx: an agent-first structure editor for Excel documents},
  author  = {{Paper Instruments, Inc.}},
  year    = {2026},
  url     = {https://github.com/paper-instruments/paper-xlsx}
}
```

paper-xlsx is a fork of *openpyxl* by Eric Gazoni, Charlie Clark, and contributors.

## License

MIT, inherited from openpyxl. Original work © 2010 openpyxl; fork additions © 2026 Paper Instruments, Inc. This fork preserves the upstream license and attribution. See [LICENCE.rst](https://github.com/paper-instruments/paper-xlsx/blob/main/LICENCE.rst).
