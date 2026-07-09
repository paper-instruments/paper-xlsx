paper-xlsx
==========

``paper-xlsx`` is Paper Instruments' hard fork of ``openpyxl``, the standard
Python library for reading and writing Excel ``.xlsx``, ``.xlsm``, ``.xltx``,
and ``.xltm`` files.

This fork starts from upstream ``openpyxl`` tag ``3.1.5``. The package is
renamed for distribution, but the Python import name remains ``openpyxl``.
That mismatch is intentional and must not be "fixed": this package is meant to
be a drop-in replacement for existing code that says ``import openpyxl``.

Name map
--------

* GitHub repository: ``paper-xlsx``
* PyPI distribution: ``paper-xlsx``
* Built wheel/sdist names: ``paper_xlsx-*``
* Python import: ``openpyxl``
* Fork sentinel: ``openpyxl.__paper_version__``

The paper API in 90 seconds
---------------------------

Everything upstream openpyxl does still works. What the fork adds is a
**preserve mode** whose contract is: every session ends in a correct
save, a typed refusal, or a loud warning — never silent loss.

.. code-block:: python

    from openpyxl import load_workbook, oracle

    wb = load_workbook("model.xlsx", preserve=True)   # bytes = truth

    # perceive
    wb.manifest().to_dict()          # what is in here, what survives
    wb.model_map().to_dict()         # inputs / calculations / outputs
    ws = wb["Model"]
    cell = ws.locate("Growth rate")  # value cell by LABEL (typed
                                     # AmbiguousTargetError when unsure)

    # edit (all guarded: refusals are atomic, typed, and name remedies)
    wb.set_input("Growth rate", 0.07)     # never overwrites a formula
    remap = ws.insert_rows(5)             # references rewritten, or a
                                          # refusal listing every victim
    receipt = wb.save("out.xlsx", receipt=True)   # what ACTUALLY changed

    # compute (LibreOffice as the calculation oracle)
    ev = wb.evaluate(set={"Model!B2": 1000}, read=["Model!B12"])
    ev.outputs, ev.certification.status          # values + trust story
    oracle.write_back("out.xlsx")                # cache real values,
                                                 # certification-gated

    # deliver
    wb.protect_for_delivery()        # lock all but classified inputs
    wb.scrub()                       # comments/metadata/personal, with
                                     # a report of everything touched

Five pinned JSON schemas (``workbook_manifest``, ``model_map``,
``evaluation``, ``oracle_write_back``, ``edit_receipt``,
``workbook_diff``) make every result machine-consumable. See
``doc/paper.rst`` for the full tour and the refusal taxonomy.

Installation
------------

From PyPI::

    pip install paper-xlsx

From the repository::

    pip install "paper-xlsx @ git+https://github.com/The-LLM-Data-Company/paper-xlsx.git@main"

Verification
------------

::

    python -c "import openpyxl; print(openpyxl.__paper_version__)"

Expected output::

    0.1.0

Fork ledger
-----------

``PAPER.md`` records the upstream base tag, conversion notes, baseline test
results, sanctioned deviations, release safety notes, and future upstream merge
policy.
