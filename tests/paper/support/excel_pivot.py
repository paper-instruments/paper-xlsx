# paper-xlsx: optional desktop Excel PivotTable acceptance runner

"""Release-gate helper. Excel is never an installation or runtime dependency.

This module lives under ``tests/paper/support/`` and must not be imported by
``openpyxl`` package code. Set ``PAPER_EXCEL_PIVOT=1`` and provide a pinned
desktop Excel automation host to run a transcript. Ordinary CI skips.

Kinds:

- ``managed``: Paper-created dedicated-cache pivots
- ``foreign-dedicated``: Excel-authored dedicated-cache adoption
- ``foreign-shared``: Excel-authored shared-cache isolation

A ``foreign-shared`` transcript must compare unselected siblings to a
refreshed control copy of the original workbook, not to pre-refresh
display. Repair prompts are failures. Excel-normalized bytes after Excel
saves are not part of Paper's preservation budget.

LibreOffice rewritten workbooks are never published through this runner.
Committed machine-readable transcripts, not this stub, are the release
evidence. Foreign adoption stays ineligible until those transcripts exist.
"""

from __future__ import annotations

import os


TRANSCRIPT_KINDS = ("managed", "foreign-dedicated", "foreign-shared")


def excel_available():
    return os.environ.get("PAPER_EXCEL_PIVOT") == "1"


def run_transcript(path, expected, *, kind="managed"):
    """Open ``path`` in Excel and compare PivotTable semantics to ``expected``.

    Raises ``NotImplementedError`` until a pinned desktop host is configured
    for the release matrix. Callers must treat repair prompts as failures.
    Committed Excel transcripts, not this stub, are the release evidence.
    """
    if kind not in TRANSCRIPT_KINDS:
        raise ValueError("unknown Excel transcript kind %r" % kind)
    if not excel_available():
        raise RuntimeError("desktop Excel acceptance is not enabled")
    raise NotImplementedError(
        "desktop Excel PivotTable transcripts require a pinned host; "
        "commit machine-readable transcripts with producer/version metadata"
    )
