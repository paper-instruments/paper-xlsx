paper-xlsx: agent-first editing of Excel files
===============================================


:Based on: openpyxl by Eric Gazoni and Charlie Clark
:Source code: https://github.com/paper-instruments/paper-xlsx
:Issues: https://github.com/paper-instruments/paper-xlsx/issues
:Generated: |today|
:License: MIT/Expat
:Version: |release|


``paper-xlsx`` is an agent-first Python library for safely inspecting, editing,
and verifying existing Excel (``.xlsx``) workbooks. It is a strict-superset hard
fork of ``openpyxl`` 3.1.5 and a drop-in replacement: the distribution is
renamed, and the import name stays ``openpyxl``, so existing imports do not
change.

Under the default preserve mode the original package bytes are the source of
truth. Edits are spliced into the original parts, untouched content survives
byte-identical, and operations that cannot be performed safely raise a typed
refusal instead of guessing. Formula calculation is delegated to a LibreOffice
oracle; the library itself never computes.

Start with the preserve-mode guide at :doc:`paper` for loading and saving,
perception, editing, the oracle, delivery, and the refusal taxonomy. The
project overview and quick start live in the repository ``README.md`` at
https://github.com/paper-instruments/paper-xlsx. Everything inherited from
openpyxl works as documented in the remaining upstream documentation.


Source and support
------------------

``paper-xlsx`` is an MIT-licensed, strict-superset hard fork of openpyxl. The
source distribution contains the fork source and the preserve-mode guide. For
the upstream project ``paper-xlsx`` builds on, see
`openpyxl <https://foss.heptapod.net/openpyxl/openpyxl>`_.


.. toctree::
    :maxdepth: 1
    :caption: Introduction
    :hidden:

    tutorial
    usage


.. toctree::
    :maxdepth: 1
    :caption: Preserve mode (paper-xlsx)
    :hidden:

    paper


.. toctree::
    :caption: Styling
    :maxdepth: 1
    :hidden:

    styles
    rich_text
    formatting

.. toctree::
    :maxdepth: 1
    :caption: Worksheets
    :hidden:

    editing_worksheets
    worksheet_properties
    validation
    worksheet_tables
    filters
    print_settings
    pivot
    comments
    datetime
    simple_formulae

.. toctree::
    :maxdepth: 1
    :caption: Workbooks
    :hidden:

    defined_names
    workbook_custom_doc_props
    protection

.. toctree::
    :maxdepth: 1
    :caption: Charts
    :hidden:

    charts/introduction

.. toctree::
    :maxdepth: 1
    :caption: Images
    :hidden:

    images

.. toctree::
    :caption: Pandas
    :maxdepth: 1
    :hidden:

    pandas

.. toctree::
    :caption: Performance
    :maxdepth: 1
    :hidden:

    optimized
    performance
    
    
.. toctree::
    :caption: Developers
    :maxdepth: 1
    :hidden:

    development
    api/openpyxl
    formula
       
.. toctree::
    :maxdepth: 1
    :caption: Release Notes
    :hidden:

    changes


API Documentation
------------------

Key Classes
+++++++++++

* :class:`openpyxl.workbook.workbook.Workbook`
* :class:`openpyxl.worksheet.worksheet.Worksheet`
* :class:`openpyxl.cell.cell.Cell`


Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`
