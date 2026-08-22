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
    ``allowed_values()`` returns ``None`` only when no list validation covers
    the cell. It supports literal lists and deterministic static,
    one-dimensional ranges; unsupported or ambiguous sources raise a typed
    refusal rather than returning a partial answer.

``openpyxl.preserve.scan_errors(wb)``
    Report actual formula error operands and cached error values without
    calculating. Error-like text inside formula string literals is ignored.

``openpyxl.preserve.diff_workbooks(before, after, remaps=())``
    Classify workbook cell changes and structural moves.

``openpyxl.preserve.copy_format(ws, source, destination)``
    Atomically copy a cell's font, fill, border, alignment, number format, and
    protection to an explicit finite range. Values, formulas, comments,
    hyperlinks, validation, row heights, and column widths are not copied.
    Merged-cell interiors and strictly protected targets refuse before any
    destination changes.

``chart.repoint(series_index, range)``
    Point one chart value series at a sheet-qualified range. Loaded
    preserve-mode charts run the complete chart patch planner before the
    in-memory formula changes; the saved patch removes the corresponding
    stale chart cache.

``ws.append_table_row(table_name, values)``
    Atomically add a row and expand a named table. Calculated columns are
    populated only from the table's declared ``calculatedColumnFormula``.
    The helper never guesses formulas from a neighboring ordinary row.

    Success is limited to a closed worksheet-table subset. Before changing
    cells, the helper checks the retained table XML and relationships, the
    original and current table geometry, headers and columns, totals and
    filter ranges, declared formulas, destination cells, merged and array or
    spill regions, inherited styles and number formats, and protection state.
    It preserves the producer's supported auto-filter convention and keeps a
    single totals row last. New in-memory tables are supported once they have
    one text header row and at least one data row.

    Loaded tables require preserve mode because stock loading does not retain
    the source XML needed for that proof. Query or externally connected
    tables, table extensions, active sort metadata, array table formulas,
    ambiguous headers, destination conflicts, and other unknown table
    structures raise a typed refusal without mutation. Treat that refusal as
    final. Do not retry through generic row insertion, ``preserve=False``, or
    direct ZIP/XML editing.

``ws.replace_image(target, replacement, name=None)``
    Replace one loaded image selected by anchor or object. Save adds a fresh
    media part and retargets only the selected drawing relationship. The old
    media part and drawing XML remain unchanged.

``wb.set_pivot_refresh_on_load(pivots=[...])``
    Permit refresh and request refresh-on-open for named pivots. Use a
    sheet-qualified name when the name is ambiguous. ``all=True`` is an
    explicit package-wide alternative; omitting both scopes is an error.
    Validation and save refuse an edit that
    changes an existing pivot's local source, directly intersects it, or
    transitively affects a formula inside it unless its cache is selected
    through this method. A value-writing save also treats known volatile
    built-ins such as ``NOW``, ``TODAY``, ``RAND``, ``CELL``, and ``INFO`` in a
    pivot source as changed. Formula dependencies also include
    calculation-relevant cell formatting, row and column display state, and
    filtering. Exact direct ranges, static defined names, and named tables are
    recognized; dynamic or otherwise unresolved local sources refuse
    conservatively. OOXML does not declare runtime volatility for user-defined
    functions, so callers using a UDF in a pivot source must select that pivot
    explicitly. The explicit request accepts that the saved cache
    remains stale until Excel refreshes it; the edit receipt reports this
    requirement. Headless readers can observe the old cached result before
    that refresh. External pivot sources do not make unrelated local cell
    edits unsafe.

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

``ws.pivots.create(...)`` / ``PivotTable.refresh()`` / ``update()`` / ``delete()``
    Create and edit Paper-managed classic worksheet PivotTables in preserve
    mode. See :ref:`paper-pivottables`.

Computation oracle
------------------

The optional :mod:`openpyxl.oracle` module uses headless LibreOffice. It does
not implement a partial formula engine.

``oracle.recalc(source, output_path=None)``
    Recalculate a temporary copy. With no output path, return calculation and
    error-scan evidence without writing. With a separate output path, create a
    Paper-preserved candidate by splicing eligible LibreOffice-calculated
    caches into the original package structure. The source is never modified,
    LibreOffice's rewritten package is never delivered, and full recalculation
    remains requested. If cache writes or that recalculation can affect a local
    pivot source, the candidate requests pivot refresh-on-open and reports the
    cache, pivots, source, and ``excel_refresh_on_open`` requirement in
    ``result.pivot_refreshes``. Status reports only whether recognized formula
    errors were detected; it does not claim Excel equivalence or financial
    accuracy.

``oracle.certify(source)``
    Compare source caches with recalculated values and report ``CERTIFIED``,
    ``DIVERGED``, or ``BASELINE_UNVERIFIABLE`` with exclusions. Finite numeric
    divergences include absolute, relative, and binary64 ULP-scale deltas.
    Call
    ``result.classify_tolerance(abs_tol=..., rel_tol=...)`` to classify recorded
    strict divergences under an explicit caller policy. Either threshold may
    admit a numeric divergence. Formula errors, nonnumeric mismatches, and
    non-finite values remain outside. The result reports certification coverage
    separately and makes no aggregate claim when there are no strict
    divergences. It cannot apply a policy narrower than Paper's strict
    comparator because strict matches are not retained. The strict
    certification status does not change.

``oracle.evaluate(source, set=..., read=...)``
    Evaluate an explicit source package and scenario. Evaluation status, like
    recalculation status, reports only whether recognized formula errors were
    detected.

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


.. _paper-pivottables:

PivotTables
-----------

Preserve-mode ``Worksheet.pivots`` inspects relationship-resolved PivotTables
and creates Paper-managed classic worksheet pivots. Created pivots have
current materialized output, ``refreshOnLoad=False``, ``enableRefresh=True``,
and ``saveData=True``. Full lifecycle (create, inspect, refresh, repoint,
move, update, rename, delete) applies to Paper-managed dedicated-cache
pivots. A shared cache disables ``can_edit_layout`` and the other
isolation-sensitive capabilities; ``update()``, headless ``refresh()``,
``repoint_source()``, ``move()``, and ``delete()`` refuse rather than
touching siblings. Layout-only shared-cache edits are not in v1.
Foreign Excel-authored pivots stay inspectable and byte-preserved;
v1 grants them at most ``can_refresh_on_open`` through
``Workbook.set_pivot_refresh_on_load`` until ``PivotTable.adopt()``
converts a qualified dedicated-cache pivot. ``qualify_adoption()`` is
the read-only eligibility analysis; it does not mutate the workbook and
does not invoke LibreOffice. ``eligible=True`` and a successful
``adopt()`` remain withheld until desktop Excel evidence proves the
managed serializer and provenance channel. Shared-cache isolation is
not in this layer; ``adopt()`` refuses those pivots. Do not treat
foreign inspection as a license to edit any PivotTable.

Pivot creation accepts worksheet tables and sheet-qualified ranges. Inspection
also resolves supported existing pivots whose source is a static defined name.
Formula columns may require stock LibreOffice via :mod:`openpyxl.oracle`; no
LibreOffice fork or commercial backend is required, and LibreOffice never
authors the published package. Literal sources do not require LibreOffice.
Data Model/OLAP, grouping, calculated fields, slicers, PivotCharts,
``showDataAs``, Strict mutation, templates, and in-session creation on a
newly added sheet refuse.

Desktop Excel is not an installation or runtime dependency. Human-run Excel
transcripts with pinned producer/version metadata are a release gate;
``tests/paper/fixtures/pivots/RELEASE_MATRIX.json`` currently records that
those transcripts are not yet committed.

.. code-block:: python

    from openpyxl import Workbook, load_workbook
    from openpyxl.worksheet.table import Table

    wb = Workbook()
    data = wb.active
    data.title = "Data"
    data.append(("Region", "Amount"))
    data.append(("East", 10))
    data.append(("West", 20))
    data.add_table(Table(displayName="RegionTable", ref="A1:B3"))
    wb.create_sheet("Summary")
    wb.save("sales.xlsx")

    wb = load_workbook("sales.xlsx", preserve=True)
    pivot = wb["Summary"].pivots.create(
        name="ByRegion",
        source="RegionTable",
        destination="A1",
        rows=["Region"],
        values=["Amount"],
    )
    wb["Data"]["B2"] = 99
    pivot.refresh()
    wb.save("sales-out.xlsx")

Direct mutation of inherited ``Worksheet._pivots`` or low-level cache objects
is not the safe Paper API. Classic worksheet pivots are not Data Model
support.
