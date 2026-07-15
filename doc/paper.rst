Preserve mode
=============

``paper-xlsx`` is an agent-first, strict-superset hard fork of openpyxl. The
distribution is renamed; the import name stays ``openpyxl``. This guide
summarizes **preserve mode**, the added APIs for safely inspecting and editing
existing workbooks.

Safety contract
---------------

The fork exists to prevent **silent corruption**: a workbook that opens fine
and is quietly wrong. Every added operation either does exactly what it claims
or refuses atomically. Loading an editable OOXML workbook retains the original
package bytes as the source of truth by default. Every session then has one of
three outcomes:

1. **A correct save.** Your edits are spliced into the original bytes;
   unrelated package content survives byte-identical. Formula-affecting edits
   may intentionally invalidate cached results and update calculation metadata.
2. **A typed refusal.** A :class:`openpyxl.errors.PaperRefusal`
   subclass is raised *atomically*: the file on disk, the model, and the
   ledger stay exactly as they were. Refusals carry ``.kind``,
   ``.anchor`` and ``.options`` for machine handling; the message
   always names a remedy.
3. **A loud warning.** Stock-mode paths that are about to do
   something lossy (:class:`~openpyxl.errors.LossySaveWarning`,
   :class:`~openpyxl.errors.StructuralShiftWarning`,
   :class:`~openpyxl.errors.ProtectedWriteWarning`,
   :class:`~openpyxl.errors.LintWarning`).

Loading and saving
------------------

.. code-block:: python

    wb = load_workbook(path)
    wb.save(out_path)                    # the splice save
    receipt = wb.save(out_path, receipt=True)   # + an EditReceipt
    wb.validate()                        # run save validation without writing

Preserve-mode file-like saves accept an open, exact ``io.BytesIO`` with no
exported buffer views or the verified path-backed ``io.BufferedRandom`` used by
pandas append mode. Use a filesystem path for any other stream type. Stock-mode
saves retain openpyxl's file-like behavior.

Preserve mode is the default for editable OOXML workbooks, including files
opened indirectly by pandas append mode. Pass ``preserve=False`` explicitly to
request openpyxl's stock, potentially lossy round trip. Read-only and
unsupported-format loads retain stock behavior.

Under preserve, ``docProps/core.xml`` is raw-copied and the ``modified``
timestamp is **not** auto-stamped unless you explicitly change
``wb.properties`` — so a no-op save stays byte-identical.

Formula cache freshness
-----------------------

When formula text changes, or a value edit may feed a formula, preserve-mode
saves remove retained cached formula results from loaded worksheets and set
Excel to recalculate the workbook automatically and fully on open. This avoids
shipping plausible but stale values. Style-only edits and unrelated value edits
retain existing caches.

Until Excel, LibreOffice, or another calculation engine recalculates the saved
file, ``data_only=True`` may return ``None`` for invalidated formulas. Use the
oracle APIs when the task requires calculated outputs before delivery. Certified
``oracle.write_back()`` results remain authoritative and are not invalidated.

The PyPI distribution is ``paper-xlsx``, but the import package remains
``openpyxl``. Do not install the separate ``openpyxl`` distribution in the same
environment. The two distributions own the same files, and package-manager
dependency metadata does not provide a safe replacement mechanism.

Perception
----------

Start with the smallest inspection surface that answers the task:
``wb.sheetnames``, bounded worksheet ranges, ``wb.defined_names``,
``ws.calculate_dimension()``, and the workbook's standard chart, validation,
protection, and relationship collections. Preserve checks run automatically;
there is no package-wide preflight inventory API.

Use the following targeted helpers when the task calls for them:

* ``wb.model_map()``: every populated cell classified as input /
  calculation / output / constant via the dependency sketch. Schema
  ``model_map`` v1. This can be expensive on large workbooks, so use it only
  when role or dependency classification is useful.
* ``ws.locate(label, prefer="right"|"below")``: the value cell for a
  text label; raises a typed refusal on ambiguity.
* ``wb.search(text_or_regex, ...)``, ``ws.allowed_values(cell)``,
  ``openpyxl.preserve.scan_errors(wb)`` (LibreOffice-free error scan),
  ``openpyxl.preserve.findings(wb)`` (a ten-kind advisory hygiene
  taxonomy with supporting evidence).
* ``openpyxl.preserve.diff_workbooks(a, b, remaps=())``: cell diffs
  classified content-changed vs shifted-by-structural-edit. Schema
  ``workbook_diff`` v1.

Editing
-------

Cell writes, styles, comments (creation), tables, charts/images
(addition; title and series-range edits on loaded charts), sheet
lifecycle (rename/copy/delete/reorder), row/column shifts and
``move_range`` all work under preserve. Each validates references before
mutation, and refusals enumerate what would have broken. Structural edits
return an :class:`~openpyxl.preserve.AddressRemap`; every pre-edit address must
be remapped through it.

* ``wb.set_input(name_or_label, value)``: resolves defined names, then
  labels, and does not overwrite formulas.
* ``openpyxl.preserve.copy_format(ws, src, dst_range)`` and
  ``apply_profile(ws, profile)``: formatting as data, preserve-safe.
* ``wb.mark_dirty(sheet_range)``: register cell mutations made outside the
  preserve APIs.
* ``wb.replace_part(name, payload)``: raw byte swap of unmanaged
  parts (media).
* ``wb.formula_lint``: ``"off" | "warn" | "refuse"`` pre-flight
  linting of every formula bind (typos, phantom sheets/names/columns,
  and locale-specific ``;`` separators).

Computation (the oracle)
------------------------

Headless, profile-isolated LibreOffice is the calculation oracle:

* ``oracle.recalc(source)``: recompute a temporary copy and scan for errors.
* ``oracle.certify(source)``: report whether LibreOffice reproduces the file's
  cached values: ``CERTIFIED`` / ``DIVERGED`` /
  ``BASELINE_UNVERIFIABLE``, with named exclusion classes.
* ``wb.evaluate(set={...}, read=[...])``: apply inputs to a temporary copy. One
  LibreOffice run provides
  outputs and certification. Schema ``evaluation`` v1;
  ``oracle.evaluate_many`` batches with a warm profile pool.
* ``oracle.write_back(path)``: splice computed values into the
  original as caches, **certification-gated**; uncertified writes retain
  the recalc-on-load flag. Schema
  ``oracle_write_back`` v1.

Delivery
--------

* ``wb.protect_for_delivery(password=None)``: lock everything except
  classified inputs (reported, since file-format protection is
  advisory).
* ``wb.scrub(remove=("comments", "metadata", "personal",
  "hidden-sheets"))``: report removed items and refusals. Removing a hidden
  sheet that would strand references produces a refusal.
* ``wb.set_pivot_refresh_on_load()``: pivots are preserved verbatim;
  this flags their caches to refresh in Excel.
* Path saves use fsync-before-rename plus directory fsync,
  spool-to-disk archive builds for path targets, decompression caps on
  load, and central-vs-local zip header agreement on the raw-copy path.

The refusal taxonomy
--------------------

``PaperRefusal`` is the base; the pinned subclasses are
``AmbiguousTargetError``, ``TargetNotFoundError``,
``UnsupportedStructureError``, ``BoundaryViolationError``,
``RelationshipPolicyError``, ``OracleUnavailableError``, and
``OracleTimeoutError``. Every pinned class is produced and tested (a CI
check enforces it), and every refusal message states what happened,
what was not changed, and what to do instead.

Compatibility opt-out
---------------------

Code that intentionally depends on openpyxl's stock package regeneration can
pass ``preserve=False`` to ``load_workbook`` or through a caller's engine
arguments. This opt-out is explicit and local to the load; process-wide
environment configuration does not change the package's safety contract.
