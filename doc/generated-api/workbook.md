<!-- Generated from paper-xlsx 0.2.0 with griffe2md 1.5.0. Do not edit by hand. -->

> Generated from the `openpyxl.workbook.Workbook` docstrings in paper-xlsx 0.2.0.

# `openpyxl.workbook.Workbook`

```python
Workbook(write_only = False, iso_dates = False)
```

Workbook is the container for all other parts of the document.

## `save`

```python
save(filename, *, allow_formula_loss = False, receipt = False)
```

Save the current workbook under the given `filename`.
Use this function instead of using an `ExcelWriter`.

**Parameters:**

Name | Type | Description | Default
---- | ---- | ----------- | -------
`allow_formula_loss` | | a workbook loaded with `data_only=True` holds cached values instead of formulas, so saving destroys formulas. Under preserve mode such a save refuses unless this flag is set (and even then only cells you actually edited lose their formulas — untouched cells keep them in the original bytes). On the stock path the flag silences the loud warning. | <code>False</code>
`receipt` | | preserve mode only — return an `openpyxl.preserve.receipts.EditReceipt` comparing the saved file against the AS-LOADED source bytes. NOTE: after several saves from one session the receipt is cumulative — it describes the session, not the last call. **Warning:** When creating your workbook using `write_only` set to True, you will only be able to call this function once. Subsequent attempts to modify or save the file will raise an `openpyxl.shared.exc.WorkbookAlreadySaved` exception. | <code>False</code>

## `validate`

```python
validate()
```

Run the preserve save planner without assembling an archive.

## `search`

```python
search(text_or_regex, *, regex = False, values = True, formulas = True)
```

Find text across the workbook (paper-xlsx).
Returns `[{"address", "match", "kind"}, ...]` where kind is
"value" or "formula".

## `set_pivot_refresh_on_load`

```python
set_pivot_refresh_on_load(pivots = None, *, all = False)
```

Request refresh-on-load for named pivots or explicitly all pivots.

Pivot names may be sheet-qualified (`"Sheet!PivotName"`). Multiple
pivots that share a cache necessarily share this cache-level setting.
