# Practical hardening proposal

This proposal retains only hardening work that materially improves correctness
and safety when editing real spreadsheets. It is based on failures reproduced
during the post-0.1.1 audit. Performance tuning, malformed-internal-object
defenses, diagnostic polish, and hypothetical edge cases are intentionally out
of scope.

## Release objective

Preserve mode should remain trustworthy when a professional workflow opens an
existing workbook, makes a bounded edit, and saves or evaluates it. A successful
operation must not silently lose formulas, redirect references, certify the
wrong value, overwrite concurrent work, or partially mutate the workbook before
refusing.

## 1. Protect delivery from concurrent replacement

### Finding

A path-backed preserve session could lose custody of its source or destination
between validation and commit. Some save paths also behaved incorrectly when
ledger cross-checking was disabled, even though that is a supported mode.

### Practical impact

A user could overwrite a workbook changed by another process, or receive a
successful save result for a destination that no longer contains the validated
artifact. This is unacceptable for shared deal files, financial models, and
other controlled documents.

### Required fix

- Bind a preserve session to the content and filesystem identity it loaded.
- Refuse if the source or destination changes before commit.
- Keep the staged candidate bound through final validation and atomic replace.
- Make path and file-handle saves work with ledger cross-checking both enabled
  and disabled.
- Preserve normal filesystem permission behavior and refuse non-regular targets
  without blocking.

### Acceptance criteria

Reproductions covering concurrent replacement, aliases, supported file handles,
cross-check-disabled saves, and failed commits must either save the validated
artifact or refuse without changing the destination.

## 2. Prevent formula loss in data-only workbooks

### Finding

With `data_only=True`, the in-memory model does not retain enough information to
distinguish some formulas from cached values or blanks. Operations such as input
assignment and worksheet copying could therefore erase or misrepresent formulas
that still exist in the source package.

### Practical impact

A workbook can open normally while calculations have silently disappeared. In
a financial model, that can turn a live assumption or output into a plausible
but permanently stale value.

### Required fix

- Prove mutation targets against retained worksheet XML when formula status is
  absent from the data-only object model.
- Refuse edits or copies that cannot preserve source formulas.
- Continue to allow operations on sheets created during the current session,
  where the model is authoritative.

### Acceptance criteria

No supported data-only operation may delete, replace, or copy a source formula
as a blank or cached scalar. Unprovable operations must refuse atomically.

## 3. Rewrite references completely during structural edits

### Finding

Row, column, table, and worksheet changes did not consistently update every
dependent reference. Risk areas included formulas, defined names, chart ranges,
table ranges, and repeated edits in one session.

### Practical impact

The resulting workbook can look intact and calculate without an obvious error
while formulas or charts point to the wrong cells. This is the central silent
corruption risk for model restructuring.

### Required fix

- Build one operation-scoped reference inventory before mutation.
- Validate all affected reference classes before changing the workbook.
- Rewrite formulas, defined names, table ranges, and chart ranges consistently.
- Return an accurate `AddressRemap` for the committed operation.
- Refuse unsupported dynamic or ambiguous references instead of guessing.

### Acceptance criteria

Tests must cover repeated insert/delete operations, cross-sheet references,
defined names, tables, and chart series. Every supported reference must resolve
to the intended cells after the edit; unsupported references must leave the
workbook unchanged.

## 4. Make oracle certification and write-back type-correct

### Finding

Certification and write-back could treat values as equivalent when their Excel
types differed, undercount error cells, or certify workbooks containing formulas
that were excluded or not actually checkable.

### Practical impact

A plausible-looking number, text value, boolean, or error can be materially
different in a financial model. A false `CERTIFIED` result gives downstream
automation stronger assurance than the evidence supports.

### Required fix

- Compare cached and computed values with Excel-aware types.
- Include formula errors and formula-cache errors in certification evidence.
- Treat volatile, external, unsupported, input-dependent, and cacheless formulas
  as unverifiable unless the contract explicitly proves them.
- Permit cache write-back only when the certification evidence supports it.

### Acceptance criteria

Certification must not report `CERTIFIED` when any formula is unchecked,
excluded, type-mismatched, or erroneous. Write-back must remain gated by the
same evidence and must never replace a formula with an incorrectly typed cache.

## 5. Preserve refusal atomicity across workbook mutations

### Finding

Some guarded operations changed cells, workbook collections, calculation flags,
or warning state before all validation completed. Interrupt-class failures also
exposed rollback gaps.

### Practical impact

A caller may catch a typed refusal and continue using the workbook under the
documented assumption that nothing changed. Partial in-memory mutation violates
that assumption and can contaminate a later successful save.

### Required fix

- Validate the complete operation before mutation wherever possible.
- Snapshot and restore every touched model field when mutation must precede a
  later validation step.
- Roll back on all failure paths that can escape the operation, including
  interrupt-class exceptions.
- Do not consume write-only or one-shot workbook state before commit is certain.

### Acceptance criteria

For each guarded public mutation, a forced failure at every post-validation
boundary must leave serialized output and relevant in-memory state unchanged.

## 6. Keep XML splicing semantically exact

### Finding

Cell-level splicing could mishandle namespace declarations, quoted foreign
attributes, whitespace-sensitive cached values, and formula metadata. These
cases can alter XML that the operation did not intend to own.

### Practical impact

An edit to one cell can damage extension metadata or change a formula cache in a
way that is invisible in the Python object model. Real workbooks produced by
Excel and third-party tools commonly contain such package details.

### Required fix

- Preserve unowned attributes and namespace declarations exactly.
- Parse attributes structurally rather than removing them with broad regular
  expressions.
- Preserve whitespace-sensitive values and formula metadata.
- Validate the emitted cell fragment before replacing source bytes.

### Acceptance criteria

Golden-package tests must show that a supported cell edit changes only the
owned cell content and required package metadata. Unowned XML must remain
byte-identical; ambiguous fragments must refuse.

## Delivery plan

Implement these areas as separate, reviewable pull requests. Each PR should
start from a minimal failing reproduction, contain the smallest production
change that fixes it, and pass the default and standard-library XML backends.
Delivery, oracle, and XML-splice changes should also include end-to-end package
round trips and LibreOffice load or calculation checks where applicable.

Do not transplant the expansive audit patch wholesale. Use it only as research
and as a source of reproductions.
