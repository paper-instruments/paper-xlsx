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
or refuses atomically. Loading with ``preserve=True`` retains the original
package bytes as the source of truth. Every session then has one of three
outcomes:

1. **A correct save.** Your edits are spliced into the original bytes;
   everything untouched survives byte-identical.
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

    wb = load_workbook(path, preserve=True)
    wb.save(out_path)                    # the splice save
    receipt = wb.save(out_path, receipt=True)   # + an EditReceipt
    wb.validate()                        # run save validation without writing

Setting ``PAPER_PRESERVE_DEFAULT=1`` in the environment makes ``preserve=True``
the default for regular loads. Read-only and legacy formats fall back.

Under preserve, ``docProps/core.xml`` is raw-copied and the ``modified``
timestamp is **not** auto-stamped unless you explicitly change
``wb.properties`` — so a no-op save stays byte-identical.

Perception
----------

* ``wb.manifest()``: sheets, formulas, defined names, volatile
  functions, protection, and a package inventory of charts, pivots,
  VBA and extensions enumerated from the actual package bytes. Schema
  ``workbook_manifest`` v1.
* ``wb.model_map()``: every populated cell classified as input /
  calculation / output / constant via the dependency sketch. Schema
  ``model_map`` v1.
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
* ``wb.mark_dirty(target)``: register a mutation made outside the preserve APIs.
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

The release gate
----------------

Preserve mode remains opt-in for the public/pandas API.
``PAPER_PRESERVE_DEFAULT`` enables it by environment setting. The public
default changes after the release conditions, including real-Excel open checks,
are met.
