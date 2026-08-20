<!-- Generated from paper-xlsx 0.2.0 with griffe2md 1.5.0. Do not edit by hand. -->

> Generated from the `openpyxl.oracle` docstrings in paper-xlsx 0.2.0.

# `openpyxl.oracle`

Bounded recalculation and certification via headless LibreOffice.

This library never calculates — a partial engine is a silent-wrongness
machine. Instead it routes to a real implementation of
Excel's semantics and reports MEASUREMENTS, never judgments:

- `recalc` recomputes all cells on a TEMP COPY and scans for Excel
  error tokens. When given an output path, it splices eligible calculated
  caches into a separate copy of the original package; LibreOffice's
  rewritten archive is never delivered.
- `certify` checks whether LibreOffice reproduces the file's own
  cached values (Excel's answer key for its current inputs) within the
  pinned tolerance, excluding cells downstream of nondeterministic volatile
  functions.

Driver rules, all measured:
the caller's file is NEVER handed to LibreOffice (temp copies only — tested
invariant); every invocation gets its own `-env:UserInstallation` profile
(shared profiles fail nondeterministically); success is `returncode == 0 AND the output file exists` (soffice exits 0 on unloadable input); stderr
is never parsed (successful runs emit noise); timeouts kill the whole
process group. Custody never depends on this module: everything
preservation-related works with no LibreOffice installed.


## `openpyxl.oracle.CertificationResult`

```python
CertificationResult(status, checked, divergences, volatile_excluded, unverifiable, external_excluded = None, unsupported_excluded = None, input_excluded = None, artifact_sha256 = None)
```

### `to_dict`

```python
to_dict()
```


## `openpyxl.oracle.DATE_SERIAL_ABS_FLOOR`

```python
DATE_SERIAL_ABS_FLOOR = 1e-11
```


## `openpyxl.oracle.ERROR_TOKENS`

```python
ERROR_TOKENS = frozenset(EXCEL_ERROR_CODES)
```


## `openpyxl.oracle.Evaluation`

```python
Evaluation(inputs, outputs, errors, certification, artifact_sha256 = None)
```

One what-if run: inputs applied to a TEMP COPY through the spine,
LibreOffice recalculated, outputs harvested. Pinned surface.

### `status`

```python
status
```

### `error_cells`

```python
error_cells
```

### `to_dict`

```python
to_dict()
```


## `openpyxl.oracle.NUMERIC_ULPS`

```python
NUMERIC_ULPS = 4
```


## `openpyxl.oracle.ORACLE_UNSUPPORTED_FUNCS`

```python
ORACLE_UNSUPPORTED_FUNCS = frozenset(['LAMBDA', 'LET', 'MAP', 'REDUCE', 'SCAN', 'BYROW', 'BYCOL', 'MAKEARRAY', 'ISOMITTED', 'STOCKHISTORY', 'RTD', 'WEBSERVICE', 'FILTERXML', 'IMAGE', 'PY', 'CUBEVALUE', 'CUBEMEMBER', 'CUBESET', 'CUBESETCOUNT', 'CUBERANKEDMEMBER', 'CUBEMEMBERPROPERTY', 'CUBEKPIMEMBER'])
```


## `openpyxl.oracle.RecalcResult`

```python
RecalcResult(cells_scanned, formula_cells, errors, *, output_kind = None, written = None, verified_unchanged = None, excluded = None, package_diff = None, artifact_sha256 = None, calculation_artifact_sha256 = None)
```

### `status`

```python
status
```

### `cells_written`

```python
cells_written
```

Return the number of formula caches written to the candidate.

**Returns:**

Type | Description
---- | -----------
<code>int</code> | Number of entries in `written`.

### `to_dict`

```python
to_dict()
```


## `openpyxl.oracle.available`

```python
available()
```

True when a LibreOffice installation can be found.


## `openpyxl.oracle.certify`

```python
certify(source, *, timeout = 120.0)
```

The divergence check: does LibreOffice reproduce the file's own
cached values? Pre-flights on an untouched temp copy; the caller's file
is never modified. Returns measurements, never judgments.


## `openpyxl.oracle.evaluate`

```python
evaluate(source, set, read, *, timeout = 120.0)
```

Scenario run against `source` (path/bytes/file-like): apply
`set` inputs to a temp copy through the preserve spine, recalculate
with LibreOffice, harvest `read` outputs. The source and every
caller file stay untouched. One LibreOffice run serves both the
outputs and the certification (original caches vs computed, with
inputs' downstream cells excluded as `input_excluded`).


## `openpyxl.oracle.evaluate_many`

```python
evaluate_many(source, cases, read, *, pool_size = 2, timeout = 120.0)
```

`evaluate` for a list of input dicts, sharing warm LibreOffice
profiles across cases (the pool is an implementation
detail — `pool_size` per-thread-isolated profiles, created lazily,
crash-replaced once, destroyed before return).


## `openpyxl.oracle.find_soffice`

```python
find_soffice()
```

Locate the LibreOffice binary, or None.


## `openpyxl.oracle.recalc`

```python
recalc(source, *, output_path = None, timeout = 120.0)
```

Recalculate a temporary copy with LibreOffice.

With no `output_path`, return error-scan evidence and write nothing.
With a separate `output_path`, build a Paper-preserved candidate by
splicing eligible calculated caches into the original package structure.
LibreOffice's rewritten package is never delivered, and the result makes
no claim of Excel equivalence or financial correctness.
