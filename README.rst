paper-xlsx
==========

``paper-xlsx`` is Paper Instruments' hard fork of ``openpyxl`` for **lossless,
safe editing of existing Excel files**.

Plain openpyxl is excellent at *writing* new spreadsheets, but *editing* one it
did not create is lossy: open a real ``.xlsx`` and save it, and the charts,
pivots, VBA and formatting it does not model are silently dropped — and a
structural edit such as ``insert_rows`` leaves formulas pointing at the wrong
cells, producing numbers that look plausible and are wrong. ``paper-xlsx`` keeps
everything else the same and adds a **preserve mode** that closes that gap.

``load_workbook(path, preserve=True)`` keeps the original package bytes as the
source of truth. Every session then ends in exactly one of three outcomes, never
a silent fourth:

* a **correct save** — your edits spliced into the original bytes; everything
  you did not touch survives byte-identical;
* a **typed refusal** — when an edit cannot be made safely, an exception that
  changes nothing (on disk *and* in memory) and names the remedy;
* a **loud warning** — never a silent, plausible-looking wrong result.

Quick start
-----------

.. code-block:: python

    from openpyxl import Workbook, load_workbook

    # a workbook a colleague sent you (imagine it also had charts and styling)
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

    wb.set_input("Growth rate", 0.07)   # set an input — never overwrites a formula
    receipt = wb.save("model_v2.xlsx", receipt=True)
    receipt.to_dict()["cells_changed"]  # {'xl/worksheets/sheet1.xml': {'B1': 'changed'}}

Structural edits are guarded the same way: ``ws.insert_rows(5)`` rewrites the
formulas, defined names and chart series that point into the shifted range and
hands back an ``AddressRemap``; if it cannot rewrite a reference safely it
refuses and lists every reference that would have broken.

With LibreOffice installed, ``paper-xlsx`` can also *compute* — it is the
calculation oracle, so values are never computed in-process:

.. code-block:: python

    import openpyxl.oracle as oracle

    ev = wb.evaluate(set={"Sheet!B1": 0.10}, read=["Sheet!B3"])
    ev.outputs                          # {'Sheet!B3': 1100}
    ev.certification.status             # did LibreOffice reproduce the file's caches?
    oracle.write_back("model.xlsx")     # write real cached values back, gated on that

Six pinned JSON schemas (``workbook_manifest``, ``model_map``, ``evaluation``,
``oracle_write_back``, ``edit_receipt``, ``workbook_diff``) make every result
machine-consumable. The **Preserve mode** guide (``doc/paper.rst``) is the full
tour and the refusal taxonomy.

What it does not do
-------------------

``paper-xlsx`` guards the file format; it is not a spreadsheet engine. It
deliberately does **not**:

* **calculate** in-process — LibreOffice is the oracle (``wb.evaluate``,
  ``oracle.write_back``, ``oracle.certify``);
* **render** — no drawing or layout of charts and images;
* **create pivot tables or edit VBA** — existing ones are preserved verbatim,
  not authored;
* **validate inputs semantically** — it guards structure, not business meaning.

Preserve mode is opt-in today: pass ``preserve=True`` per call, or set
``PAPER_PRESERVE_DEFAULT=1``. Preserve-by-default for the public/pandas surface
is release-gated.

Drop-in and name map
--------------------

The Python import name stays ``openpyxl``, so existing code that says
``import openpyxl`` keeps working unchanged and every upstream feature is still
available. Only the distribution is renamed; preserve mode and the new API are
additive.

* GitHub repository: ``paper-xlsx``
* PyPI distribution: ``paper-xlsx``
* Built wheel/sdist names: ``paper_xlsx-*``
* Python import: ``openpyxl``
* Fork sentinel: ``openpyxl.__paper_version__``

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
