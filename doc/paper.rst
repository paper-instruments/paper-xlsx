The paper API
=============

``paper-xlsx`` extends openpyxl with a **preserve mode** built for
agents and pipelines that must never silently corrupt a workbook. This
document tours the public surface; ``PAPER.md`` in the repository root
is the engineering ledger behind it.

The contract
------------

Loading with ``preserve=True`` retains the original package bytes as
the source of truth. Every session then ends in exactly one of three
legal outcomes:

1. **A correct save** — your edits spliced into the original bytes;
   everything you did not touch survives byte-identical.
2. **A typed refusal** — a :class:`openpyxl.errors.PaperRefusal`
   subclass raised *atomically*: the file on disk, the model, and the
   ledger are exactly as they were. Refusals carry ``.kind``,
   ``.anchor`` and ``.options`` for machine handling; the message
   always names a remedy.
3. **A loud warning** — for stock-mode paths that are about to do
   something lossy (:class:`~openpyxl.errors.LossySaveWarning`,
   :class:`~openpyxl.errors.StructuralShiftWarning`,
   :class:`~openpyxl.errors.ProtectedWriteWarning`,
   :class:`~openpyxl.errors.LintWarning`).

The forbidden outcome — an edit accepted but missing or wrong in the
saved file, or collateral damage to preserved content — is what the
whole architecture exists to prevent.

Loading and saving
------------------

.. code-block:: python

    wb = load_workbook(path, preserve=True)
    wb.save(out_path)                    # the splice save
    receipt = wb.save(out_path, receipt=True)   # + an EditReceipt
    wb.validate()                        # every refusal a save WOULD
                                         # raise, raised now; writes
                                         # nothing

``PAPER_PRESERVE_DEFAULT=1`` in the environment makes ``preserve=True``
the default for regular loads (read-only and legacy formats fall back).

Perception
----------

* ``wb.manifest()`` — sheets, formulas, defined names, volatile
  functions, protection, and the **confession block**: charts, pivots,
  VBA and extensions enumerated from the actual package bytes. Schema
  ``workbook_manifest`` v1.
* ``wb.model_map()`` — every populated cell classified as input /
  calculation / output / constant via the dependency sketch. Schema
  ``model_map`` v1.
* ``ws.locate(label, prefer="right"|"below")`` — the value cell for a
  text label; refuses typed (never guesses) on ambiguity.
* ``wb.search(text_or_regex, ...)``, ``ws.allowed_values(cell)``,
  ``openpyxl.preserve.scan_errors(wb)`` (LibreOffice-free error scan),
  ``openpyxl.preserve.findings(wb)`` (a ten-kind advisory hygiene
  taxonomy — measurements with evidence, never judgments).
* ``openpyxl.preserve.diff_workbooks(a, b, remaps=())`` — cell diffs
  classified content-changed vs shifted-by-structural-edit. Schema
  ``workbook_diff`` v1.

Editing
-------

Cell writes, styles, comments (creation), tables, charts/images
(addition; title and series-range edits on loaded charts), sheet
lifecycle (rename/copy/delete/reorder), row/column shifts and
``move_range`` all work under preserve — each guarded, with refusals
that enumerate exactly what would have broken. Structural edits return
an :class:`~openpyxl.preserve.AddressRemap`; every pre-edit address
must be remapped through it.

* ``wb.set_input(name_or_label, value)`` — resolves defined names then
  labels; never overwrites a formula.
* ``openpyxl.preserve.copy_format(ws, src, dst_range)`` and
  ``apply_profile(ws, profile)`` — formatting as data, preserve-safe.
* ``wb.mark_dirty(target)`` — the escape hatch for below-API mutations.
* ``wb.replace_part(name, payload)`` — raw byte swap of unmanaged
  parts (media).
* ``wb.formula_lint`` — ``"off" | "warn" | "refuse"`` pre-flight
  linting of every formula bind (typos, phantom sheets/names/columns,
  the ``;`` locale trap).

Computation (the oracle)
------------------------

LibreOffice, headless and profile-isolated, is the calculation oracle:

* ``oracle.recalc(source)`` — recompute a temp copy, scan for errors.
* ``oracle.certify(source)`` — does LibreOffice reproduce the file's
  own cached values? ``CERTIFIED`` / ``DIVERGED`` /
  ``BASELINE_UNVERIFIABLE``, with named exclusion classes.
* ``wb.evaluate(set={...}, read=[...])`` — the scenario runner: inputs
  applied through the spine to a temp copy, one LibreOffice run serves
  outputs and certification. Schema ``evaluation`` v1;
  ``oracle.evaluate_many`` batches with a warm profile pool.
* ``oracle.write_back(path)`` — splice computed values into the
  original as caches, **certification-gated**; never clears the
  recalc-on-load flag for uncertified writes. Schema
  ``oracle_write_back`` v1.

Delivery
--------

* ``wb.protect_for_delivery(password=None)`` — lock everything except
  classified inputs (reported, since file-format protection is
  advisory).
* ``wb.scrub(remove=("comments", "metadata", "personal",
  "hidden-sheets"))`` — returns a report of everything removed AND
  everything that could not be (a hidden sheet whose removal would
  strand references refuses and is reported).
* ``wb.set_pivot_refresh_on_load()`` — pivots are preserved verbatim;
  this flags their caches to refresh in Excel.
* Saves are hardened: fsync-before-rename plus directory fsync,
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
what was NOT changed, and what to do instead.

The release gate
----------------

Preserve-by-default for the public/pandas surface is release-gated: the
mechanism ships (the ``PAPER_PRESERVE_DEFAULT`` switch), internal
harness images flip it, and the public default flips only when the
gate's conditions (real-Excel open checks over the
``agent_docs/FIXTURE-REQUESTS.md`` queue among them) are met.
