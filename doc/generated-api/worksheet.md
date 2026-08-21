<!-- Generated from paper-xlsx 0.2.0 with griffe2md 1.5.0. Do not edit by hand. -->

> Generated from the `openpyxl.worksheet.Worksheet` docstrings in paper-xlsx 0.2.0.

# `openpyxl.worksheet.Worksheet`

```python
Worksheet(parent, title = None)
```

Represents a worksheet.

Do not create worksheets yourself,
use `openpyxl.workbook.Workbook.create_sheet` instead

## `allowed_values`

```python
allowed_values(cell)
```

The data-validation vocabulary for `cell` (address string or
Cell), or None when no list-type validation covers it
(paper-xlsx).

## `append_table_row`

```python
append_table_row(table_name, values)
```

Append one row to a supported named table atomically.

**Parameters:**

Name | Type | Description | Default
---- | ---- | ----------- | -------
`table_name` | <code>str</code> | Name of the table to expand. | *required*
`values` | <code>iterable | mapping</code> | Row values as a sequence or column-name mapping. | *required*

**Returns:**

Type | Description
---- | -----------
<code>None</code> | `None`.

## `replace_image`

```python
replace_image(target, replacement, *, name = None)
```

Replace one loaded image while preserving its drawing anchor.

**Parameters:**

Name | Type | Description | Default
---- | ---- | ----------- | -------
`target` | <code>Image | str</code> | Loaded image or its anchor coordinate. | *required*
`replacement` | <code>Image | path-like</code> | Replacement image or image source. | *required*
`name` | <code>str | None</code> | Optional image name used to resolve an ambiguous anchor. | <code>None</code>

**Returns:**

Type | Description
---- | -----------
<code>openpyxl.drawing.image.Image</code> | The loaded image selected for replacement.
