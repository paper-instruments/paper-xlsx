"""Runtime guard for the shared ``openpyxl`` import namespace."""

from __future__ import annotations

import re
from importlib import metadata


def _canonical_name(name):
    return re.sub(r"[-_.]+", "-", name).lower()


def _installed_distribution_names():
    names = set()
    try:
        distributions = metadata.distributions()
    except Exception:
        return names
    for distribution in distributions:
        name = distribution.metadata.get("Name")
        if name:
            names.add(_canonical_name(name))
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
