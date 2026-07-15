paper-xlsx
==========

``paper-xlsx`` is an agent-first Python library for safely inspecting, editing,
and verifying existing Excel (``.xlsx``) workbooks. It is a strict-superset hard
fork of ``openpyxl`` 3.1.5 and a drop-in replacement. The distribution is
renamed; the import name stays ``openpyxl``, so existing imports do not
change.

.. code-block:: python

    import openpyxl                    # the import name is unchanged

Why it exists
-------------

openpyxl is excellent at *creating* workbooks. Its object model, file-format
coverage, and years of absorbed edge cases are why this fork builds on it.

The harder problem is changing a real workbook without dropping charts or
leaving formulas pointed at the wrong cells. That is **silent corruption**: a
file that opens fine and is quietly wrong, often with numbers that still look
plausible. An agent cannot eyeball the result, so it needs the workbook's
structure and every edit outcome as typed, machine-readable data. It also needs
the library to refuse rather than guess.

Safety contract
---------------

Every added operation either does exactly what it claims or refuses atomically.
``load_workbook(path)`` keeps the original package bytes as the source of
truth by default. Every editing session has one of three explicit outcomes:

* a **correct save**: edits are spliced into the original bytes, and unrelated
  package content survives byte-identical; formula-affecting edits may
  intentionally invalidate caches and update calculation metadata;
* a **typed refusal**: an unsafe edit changes nothing on disk or in memory and
  the exception identifies the remedy;
* a **loud warning**: a stock-mode path reports that an operation may be lossy.

Preserve mode is the default for editable OOXML workbooks, including files
opened indirectly by pandas append mode. Pass ``preserve=False`` explicitly
only when you intend to use openpyxl's stock, potentially lossy round trip.
Read-only and unsupported-format loads retain stock behavior.

A short example
---------------

Inspect a workbook, change a labeled input, and save with a machine-readable
receipt:

.. code-block:: python

    from openpyxl import Workbook, load_workbook

    wb = Workbook()
    ws = wb.active
    ws["A1"], ws["B1"] = "Growth rate", 0.05
    ws["A2"], ws["B2"] = "Revenue", 1000
    ws["B3"] = "=B2 * (1 + B1)"
    wb.save("model.xlsx")

    wb = load_workbook("model.xlsx")

    wb.sheetnames                       # inspect workbook structure directly
    wb.active.locate("Growth rate")     # find a value cell by its label

    wb.set_input("Growth rate", 0.07)   # does not overwrite formulas
    receipt = wb.save("model_v2.xlsx", receipt=True)
    receipt.to_dict()["cells_changed"]  # {'xl/worksheets/sheet1.xml': {'B1': 'changed'}}

What it adds
------------

Inspecting a workbook
+++++++++++++++++++++

* **``wb.model_map()``** classifies populated cells as inputs, calculations,
  outputs, or constants through a dependency sketch when that analysis is
  explicitly useful. It returns the versioned ``model_map`` payload.
* **``ws.locate()`` / ``wb.search()``** find values by label or search text and
  refuse when a target is ambiguous rather than selecting one.
* **``ws.allowed_values()`` / ``openpyxl.preserve.scan_errors()`` /
  ``findings()``** expose validation choices, formula-error cells, and advisory
  workbook hygiene findings as structured data.
* **``openpyxl.preserve.diff_workbooks()``** distinguishes content changes from
  addresses shifted by structural edits. It returns the versioned
  ``workbook_diff`` payload.

Preservation checks run automatically during load, mutation, validation, and
save; they do not require a package-wide preflight inventory call.

Editing one workbook
++++++++++++++++++++

* **``load_workbook(...)`` / ``wb.save(..., receipt=True)`` /
  ``wb.validate()``** retain original package bytes by default, return an
  ``EditReceipt``, and run save validation without writing. Pass
  ``preserve=False`` to request stock openpyxl behavior explicitly.
* **``wb.set_input()``** resolves a defined name or label and changes the input
  only if the target is not a formula.
* **Row, column, and range operations** rewrite formulas, defined names, and
  chart references before mutation. They return an ``AddressRemap`` for every
  shifted pre-edit address and refuse if a reference cannot be rewritten.
* **Cell, style, comment, table, chart, image, and worksheet operations** work
  under preserve mode while guarding loaded package structures. Supported chart
  edits include titles and series ranges.
* **``openpyxl.preserve.copy_format()`` / ``apply_profile()``** apply formatting
  as data without leaving preserve mode.
* **``wb.mark_dirty()`` / ``wb.replace_part()``** register mutations made
  outside preserve APIs and replace unmanaged package parts such as media.
* **``wb.formula_lint``** controls preflight checks for malformed or unresolved
  formula references with ``"off"``, ``"warn"``, or ``"refuse"`` behavior.

Computing workbook values
+++++++++++++++++++++++++

Preserve-mode saves automatically invalidate retained formula caches after
formula changes or value edits that may feed formulas, then request an
automatic full recalculation on open. Style-only and unrelated value edits keep
their caches. Until a calculation engine runs, ``data_only=True`` may therefore
return ``None`` for invalidated formulas.

* **``oracle.recalc()``** asks a profile-isolated LibreOffice process to
  recalculate a temporary copy and scan the result for formula errors.
* **``oracle.certify()``** reports whether LibreOffice reproduces the workbook's
  existing cached values as ``CERTIFIED``, ``DIVERGED``, or
  ``BASELINE_UNVERIFIABLE``.
* **``wb.evaluate()`` / ``oracle.evaluate_many()``** apply temporary inputs and
  return requested outputs plus certification. Batch evaluation reuses a warm
  LibreOffice profile pool.
* **``oracle.write_back()``** splices computed caches into the original package
  only when certification permits it, unless the caller explicitly accepts an
  uncertified write.

Preparing and verifying delivery
++++++++++++++++++++++++++++++++

* **``wb.protect_for_delivery()``** locks everything except classified inputs
  and reports the result; file-format protection remains advisory.
* **``wb.scrub()``** removes selected comments, metadata, personal information,
  or hidden sheets and reports removals and refusals.
* **``wb.set_pivot_refresh_on_load()``** preserves pivots verbatim while asking
  Excel to refresh their caches when the workbook opens.
* **Path-based saves** build the archive on disk, enforce decompression and ZIP
  consistency limits, fsync before rename, and fsync the containing directory.

``paper-xlsx`` guards workbook structure but does not calculate formulas itself.
The oracle APIs delegate calculation to LibreOffice. Their results, like the
inspection, diff, receipt, and refusal surfaces, use versioned JSON-compatible
payloads with stable ``schema`` and ``version`` fields.

Drop-in and name map
--------------------

Only the distribution and repository are renamed. The importable package stays
``openpyxl``. This is the same distribution/import split as Pillow
(``pip install pillow``, ``import PIL``), and it preserves existing code,
snippets, and model priors. Every upstream feature remains available; preserve
mode and the added APIs are additive.

* GitHub repository / PyPI distribution: **``paper-xlsx``**
* Built wheel and sdist names: ``paper_xlsx-*``
* Python import: **``openpyxl``**
* Fork sentinel: ``openpyxl.__paper_version__ = "0.1.2"``
* Upstream base: openpyxl **3.1.5** (marker tag ``paper-base``)

Upstream releases are merged rather than rebased. The ``paper-base`` tag records
the current openpyxl fork point.

Installation
------------

Install from PyPI::

    python -m pip uninstall -y openpyxl paper-xlsx
    python -m pip install paper-xlsx

The clean uninstall is required when migrating from openpyxl. Both
distributions own the frozen ``openpyxl`` import tree, and Python package
metadata cannot make one satisfy a dependency on the other. If both are
present, file ownership depends on installation order and the fork raises on
import when it can detect that state.

Confirm the install::

    paper-xlsx-doctor

Install the current branch from Git::

    pip install "paper-xlsx @ git+https://github.com/paper-instruments/paper-xlsx.git@main"

Documentation
-------------

The Sphinx docs extend the upstream openpyxl documentation to cover the fork's
additions. Start with ``doc/paper.rst`` for preserve mode, the added API surface,
and the refusal taxonomy. Everything inherited from openpyxl works as documented
in the remaining upstream documentation.

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
