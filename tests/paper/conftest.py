"""Shared conftest for the paper-xlsx contract harness (CONVENTIONS §4)."""
from __future__ import annotations

import os
import shutil

import pytest

# every preserve-mode save in this suite runs the ledger cross-check: a
# splice-changed cell the ledger never recorded fails the test run hard
os.environ.setdefault("PAPER_LEDGER_CROSSCHECK", "1")

FIXTURES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures")


def fixture_path(rel):
    """Absolute path of a corpus fixture (read-only — never mutate these)."""
    path = os.path.join(FIXTURES_DIR, *rel.split("/"))
    if not os.path.exists(path):
        raise FileNotFoundError("no such fixture: {0}".format(rel))
    return path


@pytest.fixture
def fixtures_dir():
    return FIXTURES_DIR


@pytest.fixture
def fixture_copy(tmp_path):
    """Copy a corpus fixture into tmp_path and return the copy's path.

    All mutation happens on copies; the corpus stays byte-frozen
    (test_manifest.py enforces it independently).
    """
    def _copy(rel, name=None):
        src = fixture_path(rel)
        dst = os.path.join(str(tmp_path), name or os.path.basename(src))
        shutil.copyfile(src, dst)
        return dst

    return _copy


@pytest.fixture
def lo():
    """The LibreOffice test driver module, skipping (loudly) when absent."""
    from .support import lo as lo_mod

    lo_mod.require_lo()
    return lo_mod
