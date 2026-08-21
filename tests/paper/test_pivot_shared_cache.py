"""PR 7: shared-cache isolation for managed lifecycle verbs."""
from __future__ import annotations

import pytest

from openpyxl import load_workbook
from openpyxl.errors import (
    RelationshipPolicyError,
    UnsupportedStructureError,
)

from .test_pivot_graph import _basic_package, _write_package


def test_shared_or_foreign_cache_refuses_lifecycle(tmp_path):
    src = _write_package(tmp_path, "shared.xlsx", _basic_package())
    wb = load_workbook(src, preserve=True)
    pivot = wb["Summary"].pivots["SalesByRegion"]
    assert pivot.capabilities.can_headless_refresh is False
    assert pivot.capabilities.can_delete is False
    with pytest.raises((UnsupportedStructureError, RelationshipPolicyError)):
        pivot.refresh()
    with pytest.raises((UnsupportedStructureError, RelationshipPolicyError)):
        pivot.repoint_source("A1:B5")
    with pytest.raises((UnsupportedStructureError, RelationshipPolicyError)):
        pivot.delete()
