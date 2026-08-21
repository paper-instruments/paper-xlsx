<!-- Generated from paper-xlsx 0.2.0 with griffe2md 1.5.0. Do not edit by hand. -->

> Generated from the `openpyxl.package` docstrings in paper-xlsx 0.2.0.

# `openpyxl.package`

Package-level perception: semantic XML comparison and part-by-part
package diffing.

Under preserve mode, patch-writing *is* the save path; this module exists so
tests and agents can verify what a save did. The byte-identity invariant is
defined on part payloads, never whole-archive bytes: zip entry metadata
(timestamps, permissions) is out of scope.


## `openpyxl.package.CellsDiff`

```python
CellsDiff(changes, sheets_added, sheets_removed)
```

### `clean`

```python
clean
```

### `to_dict`

```python
to_dict()
```


## `openpyxl.package.PackageDiff`

```python
PackageDiff(added, removed, changed, identical, equivalent)
```

Part-by-part diff of two OOXML packages.

XML parts are compared semantically; binary parts by size and SHA-256.
`identical` counts parts whose payloads are byte-identical (a stricter
condition than semantic equivalence; byte-identical XML parts are never
parsed at all).

### `clean`

```python
clean
```

No parts added, removed, or semantically changed.

### `to_dict`

```python
to_dict()
```


## `openpyxl.package.PartChange`

```python
PartChange(part, kind, detail)
```

One changed part in a package diff.

### `to_dict`

```python
to_dict()
```


## `openpyxl.package.diff_cells`

```python
diff_cells(a, b)
```

Cell-level semantic diff of two packages (paths, bytes, or binary
file-likes). Deterministic order: sheet, then row, then column.


## `openpyxl.package.diff_package`

```python
diff_package(a, b, max_detail = 25)
```

Diff two packages (paths, bytes, or binary file-likes) part by part.


## `openpyxl.package.xml_equivalent`

```python
xml_equivalent(a, b)
```

True when two XML payloads are semantically equivalent.

Never normalizes cell text content.


## `openpyxl.package.xml_semantic_diff`

```python
xml_semantic_diff(a, b, max_diffs = 25)
```

Semantic differences between two XML payloads (paths/bytes/file-likes).

Compared: element structure, Clark-qualified tags (namespace *prefixes*
are insignificant), attributes (order-insensitive), and text content —
which is never normalized. Returns a list of human-readable differences,
empty when equivalent.
