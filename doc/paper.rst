Preserve mode
=============

``paper-xlsx`` is a hard fork of openpyxl. The distribution name changes, but
the import stays ``openpyxl``.

Safety contract
---------------

Editable OOXML workbooks load in preserve mode by default. The source package
bytes remain the source of truth. A save patches the parts owned by supported
edits and copies everything else unchanged.

A supported operation completes correctly or raises a
:class:`openpyxl.errors.PaperRefusal` before delivery. Refusals provide
``kind``, ``anchor``, and ``options`` fields when applicable. Writes to locked
cells can emit :class:`openpyxl.errors.ProtectedWriteWarning`; set
``wb.strict_protection = True`` to refuse them instead.

Loading and saving
------------------

.. code-block:: python

    from openpyxl import load_workbook

    wb = load_workbook("input.xlsx")
    wb["Model"]["B2"] = 125
    wb.validate()                         # plan only; no archive is built
    receipt = wb.save("output.xlsx", receipt=True)

Filesystem paths use their OOXML suffix. Seekable file-like inputs are
classified by their ZIP container and workbook content type, and the sniff
restores the caller's stream position. ``read_only=True`` uses the stock path.
Pass ``preserve=False`` to request upstream-compatible openpyxl behavior.

``validate()`` and ``save()`` use the same complete planner. Validation does
not compress a temporary workbook. Preserve-mode path saves validate first,
write to a temporary file, fsync, and atomically replace the destination.

Formula freshness
-----------------

When formula text changes, or a changed value feeds a formula, save removes
affected cached formula results and requests a full recalculation on open.
Style-only edits do not invalidate formula caches.

A preserve-mode workbook loaded with ``data_only=True`` refuses a save unless
``allow_formula_loss=True`` is explicit. Only edited cells can lose formulas;
untouched source cells remain copied from the original package.

Supported Paper APIs
--------------------

The fork keeps the standard openpyxl object model and adds a small explicit
surface:

``wb.validate()``
    Run the exact preserve save plan without building or delivering a ZIP.

``wb.save(path, receipt=True)``
    Save once and return an :class:`openpyxl.preserve.receipts.EditReceipt`.
    Version 2 receipts report direct package changes and versioned derived
    effects such as formula-cache removal, chart-cache removal, calculation
    metadata, relationships, content types, and pivot refresh flags.

``wb.search(...)`` and ``ws.allowed_values(cell)``
    Search modeled values/formulas and inspect list validation vocabulary.

``openpyxl.preserve.scan_errors(wb)``
    Report modeled formula and cached error tokens without calculating.

``openpyxl.preserve.diff_workbooks(before, after, remaps=())``
    Classify workbook cell changes and structural moves.

``openpyxl.preserve.copy_format(ws, source, destination)``
    Copy a cell format to an explicit destination range.

``ws.append_table_row(table_name, values)``
    Atomically add a row and expand a named table. Calculated columns are
    populated only from the table's declared ``calculatedColumnFormula``.
    The helper never guesses formulas from a neighboring ordinary row.

``ws.replace_image(target, replacement, name=None)``
    Replace one loaded image selected by anchor or object. Save adds a fresh
    media part and retargets only the selected drawing relationship. The old
    media part and drawing XML remain unchanged.

``wb.set_pivot_refresh_on_load(pivots=[...])``
    Set refresh metadata for named pivots. Use a sheet-qualified name when the
    name is ambiguous. ``all=True`` is an explicit package-wide alternative;
    omitting both scopes is an error.

Structural edits
----------------

On loaded preserve-mode sheets, supported row and column insertions/deletions
rewrite modeled formulas, names, print settings, tables, filters, and chart
references. They return an :class:`openpyxl.preserve.AddressRemap`. An edit
refuses if a dependent structure cannot be kept coherent.

Cell writes and ``Worksheet.append()`` use local undo journals. A failed bind
restores the target cell, style, hyperlink/comment state, ledger membership,
staged cache entry, formula flag, protection-warning state, and any exact
number-format registry delta. ``append()`` journals only its target row and
caller-supplied cell bindings; it does not snapshot the sheet or workbook.

Computation oracle
------------------

The optional :mod:`openpyxl.oracle` module uses headless LibreOffice. It does
not implement a partial formula engine.

``oracle.recalc(source, output_path=None)``
    Recalculate a temporary copy. An output path must be separate from the
    source. There is no in-place conversion mode.

``oracle.certify(source)``
    Compare source caches with recalculated values and report ``CERTIFIED``,
    ``DIVERGED``, or ``BASELINE_UNVERIFIABLE`` with exclusions.

``oracle.evaluate(source, set=..., read=...)``
    Evaluate an explicit source package and scenario. Workbook objects do not
    expose ``evaluate()`` because unsaved workbook state cannot be represented
    honestly by a retained-source evaluation.

``oracle.write_back(path)``
    Splice computed caches through the preserve machinery. This is the only
    oracle API that mutates a candidate package.

ZIP policy
----------

Paper validates package integrity: duplicate or case-colliding part names,
overlapping entries, local/central header disagreement, invalid streams,
declared-size or CRC disagreement, encryption, and unsupported compression
remain typed refusals.

The library does not define workbook eligibility caps for entry count,
uncompressed bytes, source bytes, or compression ratio. Resource limits belong
to the caller or execution environment. Reads and validation remain chunked.

Refusal taxonomy
----------------

``PaperRefusal`` is the base for ``AmbiguousTargetError``,
``TargetNotFoundError``, ``UnsupportedStructureError``,
``BoundaryViolationError``, ``RelationshipPolicyError``,
``OracleUnavailableError``, and ``OracleTimeoutError``. Invalid Python
arguments continue to use ``TypeError`` and ``ValueError``.
