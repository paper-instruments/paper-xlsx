"""PR 7: dedicated-cache deletion clearing and graph removal."""
from __future__ import annotations

from openpyxl import load_workbook

from .support.harness import save_and_reopen
from .test_pivot_create_package import _create_by_region


_TABLE = "features/tables.xlsx"


def test_delete_clears_owned_output_only(fixture_copy, tmp_path):
    src = fixture_copy(_TABLE)
    wb = load_workbook(src, preserve=True)
    wb["Data"]["Z1"] = "keep"
    _create_by_region(wb["Data"])
    wb = save_and_reopen(wb, str(tmp_path / "created.xlsx"), preserve=True)
    wb["Data"].pivots["ByRegion"].delete()
    wb = save_and_reopen(wb, str(tmp_path / "deleted.xlsx"), preserve=True)
    assert wb["Data"]["Z1"].value == "keep"
    assert wb["Data"]["E3"].value is None
    assert wb["Data"]["F4"].value is None
