import io
import warnings

import pytest

from openpyxl import Workbook, load_workbook
from openpyxl.errors import UnsupportedStructureError


def _preserved_workbook():
    workbook = Workbook()
    workbook.active.title = "Data"
    workbook.active["A1"] = 1
    workbook.active["B1"] = 2
    workbook.create_sheet("Other")
    target = io.BytesIO()
    workbook.save(target)
    return load_workbook(io.BytesIO(target.getvalue()), preserve=True)


def _ledger_cells(workbook):
    return {
        sheet: set(values)
        for sheet, values in workbook._paper_ledger.cells.items()}


def _ledger_state(workbook):
    ledger = workbook._paper_ledger
    return {slot: repr(getattr(ledger, slot)) for slot in ledger.__slots__}


def test_interrupt_during_cell_ledger_mark_rolls_back_value(monkeypatch):
    from openpyxl.cell import cell as cell_module

    workbook = _preserved_workbook()
    cell = workbook["Data"]["A1"]
    before = (cell.value, cell.data_type, _ledger_cells(workbook))
    real_mark = cell_module._mark_cell_dirty

    def interrupt(*args, **kwargs):
        real_mark(*args, **kwargs)
        raise KeyboardInterrupt("injected")

    monkeypatch.setattr(cell_module, "_mark_cell_dirty", interrupt)
    with pytest.raises(KeyboardInterrupt):
        cell.value = "changed"

    assert (cell.value, cell.data_type, _ledger_cells(workbook)) == before


@pytest.mark.parametrize("mutation", ["data_type", "hyperlink", "comment"])
def test_interrupt_during_cell_metadata_mark_rolls_back(
        monkeypatch, mutation):
    from openpyxl.cell import cell as cell_module
    from openpyxl.comments import Comment

    workbook = _preserved_workbook()
    cell = workbook["Data"]["A1"]
    before = (cell.data_type, cell.hyperlink, cell.comment,
              _ledger_cells(workbook))
    real_mark = cell_module._mark_cell_dirty

    def interrupt(*args, **kwargs):
        real_mark(*args, **kwargs)
        raise KeyboardInterrupt("injected")

    monkeypatch.setattr(cell_module, "_mark_cell_dirty", interrupt)
    with pytest.raises(KeyboardInterrupt):
        if mutation == "data_type":
            cell.data_type = "s"
        elif mutation == "hyperlink":
            cell.hyperlink = "https://example.com"
        else:
            cell.comment = Comment("note", "author")

    assert (cell.data_type, cell.hyperlink, cell.comment,
            _ledger_cells(workbook)) == before


def test_noop_data_type_assignment_does_not_mark_value_overwrite():
    workbook = _preserved_workbook()
    cell = workbook["Data"]["A1"]

    cell.data_type = cell.data_type

    assert workbook._paper_ledger.cells.get(cell.parent, set()) == set()
    assert workbook._paper_ledger.value_overwrites.get(
        cell.parent, set()) == set()


def test_strict_protection_refuses_data_type_change_atomically():
    workbook = _preserved_workbook()
    sheet = workbook["Data"]
    sheet["A1"] = "=1+1"
    sheet.protection.sheet = True
    workbook.strict_protection = True
    cell = sheet["A1"]
    before = (cell.value, cell.data_type, _ledger_state(workbook))

    with pytest.raises(UnsupportedStructureError):
        cell.data_type = "s"

    assert (cell.value, cell.data_type, _ledger_state(workbook)) == before


def test_append_failure_restores_partial_row_and_ledger():
    workbook = _preserved_workbook()
    sheet = workbook["Data"]
    before_package = io.BytesIO()
    workbook.save(before_package)
    before_cells = dict(sheet._cells)
    before_row = sheet._current_row
    before_ledger = _ledger_cells(workbook)

    with pytest.raises(ValueError):
        sheet.append(["=1+1", object()])

    assert sheet._cells == before_cells
    assert sheet._current_row == before_row
    assert _ledger_cells(workbook) == before_ledger
    assert not workbook._paper_ledger.formulas_changed
    after_package = io.BytesIO()
    workbook.save(after_package)
    assert after_package.getvalue() == before_package.getvalue()


def test_append_failure_restores_prebuilt_cell_binding():
    from openpyxl.cell import Cell

    workbook = _preserved_workbook()
    sheet = workbook["Data"]
    cell = Cell(sheet, value="prebuilt")
    before = (cell.parent, cell.row, cell.column)

    with pytest.raises(ValueError):
        sheet.append([cell, object()])

    assert (cell.parent, cell.row, cell.column) == before


def test_structural_interrupt_restores_cells_and_ledger(monkeypatch):
    from openpyxl.preserve import structural

    workbook = _preserved_workbook()
    sheet = workbook["Data"]
    before_cells = dict(sheet._cells)
    before_ledger = _ledger_cells(workbook)

    def interrupt(*_args, **_kwargs):
        raise KeyboardInterrupt("injected")

    monkeypatch.setattr(structural, "apply_model_shift", interrupt)
    with pytest.raises(KeyboardInterrupt):
        sheet.insert_rows(1)

    assert sheet._cells == before_cells
    assert _ledger_cells(workbook) == before_ledger


def test_sheet_create_and_remove_failures_restore_workbook(monkeypatch):
    from openpyxl.preserve import ledger

    workbook = _preserved_workbook()
    sheets = workbook._sheets
    before = list(sheets)
    before_ledger = _ledger_state(workbook)
    real_added = ledger.mark_sheet_added

    def fail_added(*args, **kwargs):
        real_added(*args, **kwargs)
        raise RuntimeError("injected create failure")

    monkeypatch.setattr(ledger, "mark_sheet_added", fail_added)
    with pytest.raises(RuntimeError, match="create failure"):
        workbook.create_sheet("Failed")
    assert workbook._sheets is sheets
    assert list(sheets) == before
    assert _ledger_state(workbook) == before_ledger

    monkeypatch.setattr(ledger, "mark_sheet_added", real_added)
    real_removed = ledger.record_sheet_removal

    def fail_removed(*args, **kwargs):
        real_removed(*args, **kwargs)
        raise RuntimeError("injected remove failure")

    monkeypatch.setattr(ledger, "record_sheet_removal", fail_removed)
    with pytest.raises(RuntimeError, match="remove failure"):
        workbook.remove(workbook["Other"])
    assert workbook._sheets is sheets
    assert list(sheets) == before
    assert _ledger_state(workbook) == before_ledger


def test_mark_dirty_failure_restores_existing_claims(monkeypatch):
    workbook = _preserved_workbook()
    ledger = workbook._paper_ledger
    before = _ledger_cells(workbook)
    calls = 0
    real_mark = type(ledger).mark_cell

    def fail_second(self, *args, **kwargs):
        nonlocal calls
        calls += 1
        real_mark(self, *args, **kwargs)
        if calls == 2:
            raise RuntimeError("injected claim failure")

    monkeypatch.setattr(type(ledger), "mark_cell", fail_second)
    with pytest.raises(RuntimeError, match="claim failure"):
        workbook.mark_dirty("Data!A1:B1")

    assert _ledger_cells(workbook) == before


def test_table_append_lint_failure_happens_before_commit(tmp_path):
    from openpyxl.errors import LintWarning
    from openpyxl.preserve.tables import append_row
    from openpyxl.worksheet.table import Table

    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Amount", "Formula"])
    sheet.append([1, "=SUM(A2)"])
    sheet.add_table(Table(displayName="Inputs", ref="A1:B2"))
    source = tmp_path / "table.xlsx"
    workbook.save(source)
    workbook = load_workbook(source, preserve=True)
    before = io.BytesIO()
    workbook.save(before)

    with warnings.catch_warnings():
        warnings.simplefilter("error", LintWarning)
        with pytest.raises(LintWarning):
            append_row(workbook.active, "Inputs", [3, "=SUMM(A1)"])

    after = io.BytesIO()
    workbook.save(after)
    assert after.getvalue() == before.getvalue()
