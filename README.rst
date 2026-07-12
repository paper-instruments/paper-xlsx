paper-xlsx
==========

``paper-xlsx`` is an agent-first Python library for safely inspecting and
editing existing Excel (``.xlsx``) files. It is a strict-superset hard fork of
``openpyxl`` 3.1.5 and a drop-in replacement. The distribution is renamed; the
import name stays ``openpyxl``, so existing code keeps working unchanged.

Why it exists
-------------

openpyxl is excellent at *creating* a workbook from scratch. The harder problem
is changing a real workbook without dropping charts or leaving formulas pointed
at the wrong cells. That is **silent corruption**: a file that opens fine and is
quietly wrong, often with numbers that still look plausible.

An agent cannot eyeball the result. It needs the workbook's structure and every
edit outcome as typed, machine-readable data, and it needs the library to refuse
rather than guess.

Safety contract
---------------

Preserve mode applies the shared rule: every operation either does exactly what
it claims or refuses atomically. ``load_workbook(path, preserve=True)`` keeps the
original package bytes as the source of truth. Every session has one of three
explicit outcomes:

* a **correct save**: your edits are spliced into the original bytes, and
  everything untouched survives byte-identical;
* a **typed refusal**: an unsafe edit changes nothing on disk or in memory and
  the exception names the remedy;
* a **loud warning**: a stock-mode path reports that an operation may be lossy.

``manifest``, ``model_map``, and ``locate`` expose workbook structure through
versioned, JSON-compatible payloads.

Quick start
-----------

.. code-block:: python

    from openpyxl import Workbook, load_workbook

    # create a workbook with an input and a formula
    wb = Workbook()
    ws = wb.active
    ws["A1"], ws["B1"] = "Growth rate", 0.05
    ws["A2"], ws["B2"] = "Revenue", 1000
    ws["B3"] = "=B2 * (1 + B1)"
    wb.save("model.xlsx")

    # reopen it in preserve mode: the original bytes are the source of truth
    wb = load_workbook("model.xlsx", preserve=True)

    wb.manifest().to_dict()             # what's in the file, what survives a save
    wb.model_map().to_dict()            # inputs / calculations / outputs
    wb.active.locate("Growth rate")     # find a value cell by its label

    wb.set_input("Growth rate", 0.07)   # set an input; does not overwrite formulas
    receipt = wb.save("model_v2.xlsx", receipt=True)
    receipt.to_dict()["cells_changed"]  # {'xl/worksheets/sheet1.xml': {'B1': 'changed'}}

Structural edits follow the same contract. ``ws.insert_rows(5)`` rewrites the
formulas, defined names, and chart series that point into the shifted range and
returns an ``AddressRemap``. If it cannot rewrite a reference safely, it refuses
and lists every reference that would have broken.

Computing values
----------------

``paper-xlsx`` does not implement formula calculation. When LibreOffice is
installed, the library delegates calculation to it as the oracle:

.. code-block:: python

    import openpyxl.oracle as oracle

    ev = wb.evaluate(set={"Sheet!B1": 0.10}, read=["Sheet!B3"])
    ev.outputs                          # {'Sheet!B3': 1100}
    ev.certification.status             # did LibreOffice reproduce the file's caches?
    check = oracle.recalc("model.xlsx", output_path="model_recalculated.xlsx")
    check.errors                        # formula-error cells found after recalculation

``oracle.write_back(path)`` is a separate, certification-gated operation. It
splices computed caches into the original package only when the existing
caches provide a trustworthy baseline, unless the caller explicitly accepts an
uncertified write.

Results expose versioned, JSON-compatible payload contracts with stable
``schema`` and ``version`` fields. The **Preserve mode** guide
(``doc/paper.rst``) provides the complete API overview and refusal taxonomy.

Preserve mode is opt-in today: pass ``preserve=True`` per call, or set
``PAPER_PRESERVE_DEFAULT=1``. Preserve-by-default for the public/pandas API
is release-gated.

Drop-in and name map
--------------------

Only the distribution and repository are renamed. The importable package stays
``openpyxl``. This is the same distribution/import split as Pillow
(``pip install pillow``, ``import PIL``), and it preserves existing code,
snippets, and model priors. Existing code that says ``import openpyxl`` keeps
working unchanged, and every upstream feature is still available. Preserve mode
and the new API are purely additive.

* GitHub repository: ``paper-xlsx``
* PyPI distribution: ``paper-xlsx``
* Built wheel/sdist names: ``paper_xlsx-*``
* Python import: ``openpyxl``
* Fork sentinel: ``openpyxl.__paper_version__ = "0.1.2"``
* Upstream base: openpyxl **3.1.5** (marker tag ``paper-base``)

Upstream releases are merged rather than rebased. The ``paper-base`` tag records
the current openpyxl fork point.

Installation
------------

``paper-xlsx`` is published on PyPI (this source tree carries the version in
``openpyxl/_paper_version.py``)::

    python -m pip install paper-xlsx

Install it in an environment that does not also contain the ``openpyxl``
distribution. Both distributions own the same ``openpyxl`` import tree, and
Python package metadata cannot make one satisfy a dependency on the other. If
both are present, file ownership depends on installation order; the fork raises
on import when it can detect that state. Verify any installed environment with::

    paper-xlsx-doctor

The source repository remains private. Users with repository access can install
the current branch from Git::

    pip install "paper-xlsx @ git+https://github.com/The-LLM-Data-Company/paper-xlsx.git@main"

Verification
------------

::

    python -c "import openpyxl; print(openpyxl.__paper_version__)"

Expected output::

    0.1.2

How it's tested
---------------

* Upstream openpyxl's test suite runs on every change to check compatibility
  with existing behavior.
* A frozen, hash-pinned fixture corpus under ``tests/paper`` records exact
  provenance. Its current files come from openpyxl, LibreOffice, or documented
  package surgery; the still-missing real-Excel and Google Sheets fixture
  buckets are listed in the corpus README.
* The contract harness saves and reopens before asserting, enforces exact
  changed-part budgets, exercises refusal atomicity, and runs a headless
  LibreOffice load smoke.

License
-------

MIT, inherited from openpyxl. Original work © the openpyxl authors; fork
additions © Paper Instruments, Inc. The fork preserves the upstream license and
attribution. See ``LICENCE.rst``.
