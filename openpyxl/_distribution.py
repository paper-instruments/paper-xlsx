"""Runtime guard for the shared ``openpyxl`` import namespace."""

from __future__ import annotations

import re
from importlib import metadata


_GUARDED_DISTRIBUTIONS = ("openpyxl", "paper-xlsx")


def _canonical_name(name):
    return re.sub(r"[-_.]+", "-", name).lower()


def _installed_distribution_names():
    """The guarded distribution names that are actually installed.

    Two targeted lookups keep ``import openpyxl`` overhead independent of
    the environment's size; enumerating every installed distribution's
    metadata costs time linear in the number of packages, on every import.
    A lookup that fails for any reason other than a clean hit counts the
    distribution as absent — this guard must never be the thing that
    breaks ``import openpyxl``, and ``paper-xlsx-doctor`` performs the
    deep verification.
    """
    names = set()
    for requested in _GUARDED_DISTRIBUTIONS:
        try:
            metadata.distribution(requested)
        except Exception:
            continue
        names.add(requested)
    return names


def assert_single_openpyxl_distribution(distribution_names=None):
    """Fail when both distributions claim the same import package.

    Python package metadata has no supported way for ``paper-xlsx`` to
    satisfy another project's dependency on the separate ``openpyxl``
    distribution. Installing both makes their files overwrite each other in
    an order-dependent way, so continuing would make the fork identity
    unknowable.
    """
    names = (_installed_distribution_names() if distribution_names is None
             else {_canonical_name(name) for name in distribution_names})
    if {"openpyxl", "paper-xlsx"}.issubset(names):
        raise ImportError(
            "paper-xlsx and openpyxl are both installed, but both provide "
            "the 'openpyxl' import package. Uninstall both distributions, "
            "then install only paper-xlsx in this environment."
        )
