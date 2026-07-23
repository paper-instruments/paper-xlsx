# paper-xlsx

**A drop-in, agent-safe fork of openpyxl that will not silently corrupt a real Excel file.**

[![PyPI](https://img.shields.io/pypi/v/paper-xlsx)](https://pypi.org/project/paper-xlsx/)
[![Python versions](https://img.shields.io/pypi/pyversions/paper-xlsx)](https://pypi.org/project/paper-xlsx/)
[![Test](https://github.com/paper-instruments/paper-xlsx/actions/workflows/test.yml/badge.svg)](https://github.com/paper-instruments/paper-xlsx/actions/workflows/test.yml)

`paper-xlsx` is a strict-superset hard fork of [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl) 3.1.5 for safely inspecting, editing, and verifying existing Excel workbooks. It keeps everything that makes openpyxl excellent — the reader, the object model, the formula tokenizer, fifteen years of absorbed producer quirks — and replaces the one thing that isn't: a save path that regenerates the whole file from memory and silently destroys whatever it doesn't model.

```python
import openpyxl   # the import name is unchanged — see "Drop-in by design"
```

Under the default **preserve mode**, the original file's bytes are the source of truth. Edits are spliced into those bytes surgically; everything untouched survives byte-identical *by construction*, not by coverage. Every operation has exactly three legal outcomes: **done correctly**, **refused with a typed error** that says what was found and why it was unsafe, or **done with a loud warning** enumerating exactly what could not be preserved. There is no silent fourth option.

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
wb.active.locate("Growth rate")         # find a value cell by its label

wb.set_input("Growth rate", 0.07)       # refuses to overwrite a formula
receipt = wb.save("model_v2.xlsx", receipt=True)
receipt.to_dict()["cells_changed"]      # {'xl/worksheets/sheet1.xml': {'B1': 'changed'}}
```

When an edit cannot be made safely, you get a typed refusal instead of a corrupted file — and a refused operation leaves the model, the ledger, and the disk exactly as they were:

```python
from openpyxl.errors import PaperRefusal

try:
    wb.set_input("Growth", 0.07)        # ambiguous or missing label
except PaperRefusal as err:
    err.kind, err.anchor, err.options   # machine-readable: what, where, remedies
```

## What we changed from openpyxl — and why

Every claim below is traceable to a commit in this repository. The fork point is [`021192cf`](https://github.com/paper-instruments/paper-xlsx/commit/021192cf264012d7b5dba537f9994ee3f59ff223), on top of upstream openpyxl 3.1.5 (tagged `paper-base`). Upstream openpyxl's own test suite runs green on every change — that is the mechanical proof that existing callers keep working.

### Added

- **Preserve mode** ([`021192cf`](https://github.com/paper-instruments/paper-xlsx/commit/021192cf264012d7b5dba537f9994ee3f59ff223), the ~31,000-line bootstrap): `load_workbook(path)` retains the original archive bytes, a **dirty ledger** wired into openpyxl's own setters records every semantic mutation, and save **splices** only the dirty byte ranges into the original parts instead of regenerating files. Untouched parts are raw-copied byte-identical; a no-op save produces a byte-identical file. The machinery lives in the new [`openpyxl/preserve/`](openpyxl/preserve/) tree (30 modules: ledger, splice writer, cross-part discipline, deterministic atomic zip I/O, and friends).
- **A typed refusal taxonomy** ([`openpyxl/errors.py`](openpyxl/errors.py)): `PaperRefusal` and its subclasses — `AmbiguousTargetError`, `TargetNotFoundError`, `UnsupportedStructureError`, `BoundaryViolationError`, `RelationshipPolicyError`, `OracleUnavailableError`, `OracleTimeoutError`. Refusals are **atomic**: validate fully, then mutate; a refused operation changes nothing in memory or on disk. Each carries machine-readable `kind`, `anchor`, and `options` fields.
- **The oracle** ([`openpyxl/oracle.py`](openpyxl/oracle.py)): openpyxl never calculates, and this fork deliberately ships **no formula engine** — a partial engine is a silent-wrongness machine. Instead, `oracle.recalc()`, `oracle.certify()`, `wb.evaluate()`, and `oracle.write_back()` delegate to a headless, profile-isolated LibreOffice process working on temp copies, and report measurements (`CERTIFIED` / `DIVERGED` / `BASELINE_UNVERIFIABLE`), never judgments. All preservation guarantees hold with no LibreOffice installed.
- **Perception helpers**: `ws.locate(label)`, `wb.search(...)`, `ws.allowed_values(cell)`, `wb.model_map()` (inputs/calculations/outputs classification), `openpyxl.preserve.scan_errors()`, `findings()`, and `diff_workbooks()` — structured, versioned JSON-compatible payloads throughout, built so an agent asks targeted questions instead of parsing repr output.
- **Guarded structural edits**: `insert_rows`, `delete_rows`, sheet renames, and range moves rewrite dependent formulas, defined names, print areas, table ranges, and chart references before mutating — or refuse if a reference cannot be rewritten — and return an `AddressRemap` for every shifted address (shipped across 0.1.2, see [`doc/changes.rst`](doc/changes.rst)). Stock openpyxl's silent reference corruption ceases to exist as an outcome.
- **Delivery helpers**: `wb.scrub(remove=("comments", "metadata", "personal", "hidden-sheets"))` and `wb.protect_for_delivery()`, plus hardened path saves (fsync-before-rename, decompression caps, ZIP consistency checks).
- **`paper-xlsx-doctor`**: a console script that verifies the installed distribution actually owns the `openpyxl` import tree.

### Changed

- **Preserve mode became the default** ([`6c2b99f7`](https://github.com/paper-instruments/paper-xlsx/commit/6c2b99f7), 0.1.3 — a deliberate breaking change: *"editable OOXML loads now use preserve mode unless preserve=False is passed explicitly"*). Preserve started as pure opt-in; once the contract harness and fixture corpus proved the spine, the safe path became the default path. The default is classified by source type ([`b1773f3c`](https://github.com/paper-instruments/paper-xlsx/commit/b1773f3c)): filesystem paths get preserve only with an OOXML suffix (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`); file-like streams are validated from their bytes and stay preserved unless named as legacy `.xls`/`.xlsb`; `read_only=True` loads keep stock behavior. `preserve=False` remains a first-class opt-out to the stock, potentially lossy round trip.
- **Formula caches are invalidated instead of trusted** ([`a0b89793`](https://github.com/paper-instruments/paper-xlsx/commit/a0b89793) through [`1a16aa07`](https://github.com/paper-instruments/paper-xlsx/commit/1a16aa07), 0.1.3): when you edit a formula, or write a value into a cell that formulas read, the save strips the now-stale cached results from the file and sets the workbook to fully recalculate on open. The risk this closes: a human opens the edited file in Excel and silently trusts a stale number. The commit stream handles the ugly realities — array/spill formula followers, namespace-prefixed formula elements, whole-column array references — and style-only edits keep their caches untouched.
- **The stock path now warns loudly.** With `preserve=False`, a save that is about to drop content it cannot preserve emits a structured `LossySaveWarning` enumerating what dies, instead of upstream's silence. This is one of the fork's two sanctioned deviations from stock behavior (the other is preserve mode itself).

### Removed

- **The package-wide `wb.manifest()` preflight API** ([`ea327e23`](https://github.com/paper-instruments/paper-xlsx/commit/ea327e23), 0.1.3). The bootstrap shipped a full-workbook inventory — every sheet, formula address, chart, extension, and an honesty "confession" block — intended as an agent's preflight. In practice it was the wrong shape: a single call walked every cell, built a full dependency map, and re-scanned the archive several times, to answer questions agents ask one at a time. It was deleted in favor of the targeted primitives above, and the preflight *guarantee* moved into the machinery itself — preservation checks now run automatically during load, mutation, validation, and save. From the changelog: *"Remove the package-wide workbook manifest API in favor of targeted inspection through standard workbook objects, search, locate, validation, findings, and optional model-map analysis."* We built it, learned it was the wrong shape, and cut it before 1.0.

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

The full preserve-mode guide — including the refusal taxonomy, receipts, the oracle, and delivery — is in [`doc/paper.rst`](doc/paper.rst).

## Battle-tested against real failure modes

The fork is developed against an evaluation suite of **15 realistic spreadsheet-editing tasks**, each run across four treatments comparing stock `openpyxl==3.1.5` (with and without Anthropic's published xlsx agent skill) against `paper-xlsx==0.1.3`. Hidden graders reopen each produced package independently and verify not just the requested edit but the collateral: formulas, names, tables, charts, drawings, review content, caches, and relationships. Tasks include:

- repairing one member of a shared-formula block without disturbing its siblings or a neighboring array formula,
- fixing a formula without dropping the workbook's x14 validation dropdowns,
- renaming a sheet with every dependent formula, defined name, print area, and chart reference rewritten,
- updating an input in a macro-enabled `.xlsm` and delivering it with the VBA project intact,
- retargeting a chart series without disturbing sibling images or drawing anchors,
- scrubbing a workbook for external delivery — protection applied, private residue removed, formulas and review content preserved.

We don't publish aggregate pass rates here: results are tracked in the internal evaluation harness, not in this repository, and this README doesn't quote numbers you can't check.

The library's own test discipline is documented in [CONTRIBUTING.md](CONTRIBUTING.md): upstream's full pytest suite green on every change, a frozen hash-pinned fixture corpus with honest provenance labels, save→reopen→assert everywhere, exact changed-part budgets, refusal-atomicity checks (output bytes equal input bytes), and a headless LibreOffice load smoke. Fixture buckets we still need — genuinely Excel-authored and Google-Sheets-authored files we refuse to synthesize — are listed in [FIXTURE-REQUESTS.md](FIXTURE-REQUESTS.md).

## Drop-in by design

The import name `openpyxl` is **frozen forever**. `import openpyxl` appears in millions of scripts, in pandas itself, and in every model's training prior — so the fork keeps the import and renames only the distribution, the same split as Pillow (`pip install pillow`, `import PIL`).

- PyPI distribution / GitHub repository: **`paper-xlsx`**
- Python import: **`openpyxl`** — never `import paper_xlsx`, anywhere
- Fork sentinel: `openpyxl.__paper_version__` (`"0.1.3"`)
- Upstream base: openpyxl **3.1.5** (tag `paper-base`; upstream releases are merged, not rebased)

Note the two version numbers: `paper-xlsx` is versioned independently (currently **0.1.3**, early and pre-1.0) while `openpyxl.__version__` reports the upstream base (**3.1.5**) it wraps. pandas flows through this fork automatically — preserve-by-default covers files pandas opens for editing, and fresh `ExcelWriter` output is untouched stock behavior. Python **3.9–3.13** are supported and tested in CI, on Linux and Windows, with and without lxml.

Everything upstream openpyxl documents still works, unchanged. The additions are a strict superset; the only behavioral deltas are the two sanctioned ones above (preserve-by-default and loud lossy-save warnings), both reversible per load with `preserve=False`.

## Documentation

- [`doc/paper.rst`](doc/paper.rst) — the preserve-mode guide: loading and saving, formula-cache freshness, perception, editing, the oracle, delivery, the refusal taxonomy, and the compatibility opt-out. Ships inside the sdist.
- The remaining Sphinx docs are upstream openpyxl's, and everything they document still applies.

## Roadmap

paper-xlsx is pre-1.0 and its surface grows only as fast as the contract harness can prove it. Direction, clearly distinguished from what is shipped today:

- **Filling the fixture corpus** with genuinely Excel-authored and Google-Sheets-authored workbooks ([FIXTURE-REQUESTS.md](FIXTURE-REQUESTS.md)) — the load-bearing bucket for everything below.
- **Shrinking the refusal set**: preserve mode still refuses operations whose splice coverage isn't proven — for example chartsheet edits, generating table and pivot parts on newly added sheets, and comment changes on sheets that already carry comment parts (see the refusal sites in [`openpyxl/preserve/saver.py`](openpyxl/preserve/saver.py)). Each becomes supported as coverage lands.
- **A deeper computation layer** on the oracle: broader evaluation workflows and formula pre-flight linting.

Nothing on this list is presented as a current capability; when it ships, it appears in [`doc/changes.rst`](doc/changes.rst).

## Contributing

Contributions are welcome — see [CONTRIBUTING.md](CONTRIBUTING.md) for the engineering discipline this fork runs on. The short version: upstream's test suite stays green on every PR; every test asserts on a saved-and-reopened file, never an in-memory object; every bug fix lands with a frozen fixture reproducing it (*no fix without a fixture*); refusals must be atomic down to the byte; and new XML handling goes through openpyxl's `Serialisable` descriptor framework — never string-formatted XML.

The most valuable non-code contribution right now is real-world fixtures: workbooks authored by desktop Excel or Google Sheets, per [FIXTURE-REQUESTS.md](FIXTURE-REQUESTS.md).

## Community

- **Bugs and feature requests**: [GitHub Issues](https://github.com/paper-instruments/paper-xlsx/issues)
- **Questions and ideas**: [GitHub Discussions](https://github.com/paper-instruments/paper-xlsx/discussions)

There is no Discord, Slack, or forum.

## Acknowledgments

paper-xlsx exists because openpyxl's object model and format coverage are excellent — we forked its persistence layer, not its judgment. Deep thanks to Eric Gazoni, Charlie Clark, and the openpyxl contributors (see [AUTHORS.rst](AUTHORS.rst)) for fifteen years of careful work this project stands on. Upstream openpyxl lives at [foss.heptapod.net/openpyxl/openpyxl](https://foss.heptapod.net/openpyxl/openpyxl).

If you reference this project in writing, cite it as *paper-xlsx* (Paper Instruments, Inc.), a fork of *openpyxl* by Eric Gazoni, Charlie Clark, and contributors, and link to this repository.
