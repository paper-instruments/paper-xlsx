# paper-xlsx: optional desktop Excel PivotTable acceptance runner

"""Release-gate helper. Excel is never an installation or runtime dependency.

This module lives under ``tests/paper/support/`` and must not be imported by
``openpyxl`` package code. Set ``PAPER_EXCEL_PIVOT=1`` and provide a pinned
desktop Excel automation host to run a transcript. Ordinary CI skips.

LibreOffice rewritten workbooks are never published through this runner.
"""

from __future__ import annotations

import os


def excel_available():
    return os.environ.get("PAPER_EXCEL_PIVOT") == "1"


def run_transcript(path, expected):
    """Open ``path`` in Excel and compare PivotTable semantics to ``expected``.

    Raises ``NotImplementedError`` until a pinned desktop host is configured
    for the release matrix. Callers must treat repair prompts as failures.
    Committed Excel transcripts, not this stub, are the release evidence.
    """
    if not excel_available():
        raise RuntimeError("desktop Excel acceptance is not enabled")
    raise NotImplementedError(
        "desktop Excel PivotTable transcripts require a pinned host; "
        "commit machine-readable transcripts with producer/version metadata"
    )
