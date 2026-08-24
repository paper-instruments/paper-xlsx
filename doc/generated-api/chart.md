<!-- Generated from paper-xlsx 0.2.1 with griffe2md 1.5.0. Do not edit by hand. -->

> Generated from the `openpyxl.chart.ChartBase` docstrings in paper-xlsx 0.2.1.

# `openpyxl.chart.ChartBase`

```python
ChartBase(axId = (), **kw)
```

Base class for all charts

## `repoint`

```python
repoint(series_index, new_range)
```

Point a series' VALUES at `new_range` — "the chart now covers
Q1-Q4" (paper-xlsx). `new_range` must be a
sheet-qualified single-area range like "'Data'!$B$2:$B$13"; it is
validated here, and under preserve mode the save expresses the
change as a byte patch of the chart's `<c:f>` text.
