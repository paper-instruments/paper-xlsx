Pivot Tables
============

paper-xlsx adds preserve-mode ``Worksheet.pivots`` for inspection and for
Paper-managed classic worksheet PivotTables. See :ref:`paper-pivottables`.
Direct mutation of inherited ``Worksheet._pivots`` is not the safe Paper API.

Upstream openpyxl still documents the low-level serializer objects below.
Those objects remain available, but they are not the Paper mutation
contract and must not be used to bypass typed refusals.


Example
-------

.. code::

    from openpyxl import load_workbook
    wb = load_workbook("campaign.xlsx")
    ws = wb["Results"]
    pivot = ws._pivots[0] # any will do as they share the same cache
    pivot.cache.refreshOnLoad = True


For further information see :class:`openpyxl.pivot.cache.CacheDefinition`
