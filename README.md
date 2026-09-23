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

**An import-compatible, agent-safe fork of openpyxl for creating and editing Excel workbooks without silent data loss.**

`paper-xlsx` is an import-compatible hard fork of [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl) for creating and editing Excel (`.xlsx`) workbooks. It keeps the `openpyxl` import name and upstream object model. When editing an existing file, preserve mode retains package content that openpyxl does not model.

## Installation

Use an [activated virtual environment](https://docs.python.org/3/library/venv.html).

```bash
python -m pip uninstall -y openpyxl paper-xlsx
python -m pip install paper-xlsx
python -m paper_xlsx_doctor
```

Both distributions provide the `openpyxl` import package. Do not install `openpyxl` and `paper-xlsx` in the same environment.

To return to `openpyxl`, run the uninstall line above again, then `python -m pip install openpyxl`.

## Quick start

Create a workbook, reopen it in preserve mode, and save an edit with a machine-readable receipt:

```python
from openpyxl import Workbook, load_workbook

wb = Workbook()
ws = wb.active
ws["A1"], ws["B1"] = "Growth rate", 0.05
ws["A2"], ws["B2"] = "Revenue", 1000
ws["B3"] = "=B2 * (1 + B1)"
wb.save("model.xlsx")

wb = load_workbook("model.xlsx")
wb.active["B1"] = 0.07
receipt = wb.save("model-v2.xlsx", receipt=True)
print(receipt.to_dict()["cells_changed"])
```

Preserve mode keeps untouched package content and refuses edits it cannot save safely.

## Documentation

Read the [paper-xlsx documentation](https://docs.paperinstruments.com/).

## Contributing

Contributions are welcome. See [CONTRIBUTING.md](https://github.com/paper-instruments/paper-xlsx/blob/main/CONTRIBUTING.md).

## Acknowledgments

paper-xlsx builds on openpyxl by Eric Gazoni, Charlie Clark, and contributors. This fork preserves their API, license, and attribution.

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

Cite it as a fork of *openpyxl* by Eric Gazoni, Charlie Clark, and contributors.

## License

MIT, inherited from openpyxl. Original work © 2010 openpyxl; fork additions © 2026 Paper Instruments, Inc. See [LICENCE.rst](https://github.com/paper-instruments/paper-xlsx/blob/main/LICENCE.rst).
