paper-xlsx
==========

``paper-xlsx`` is Paper Instruments' hard fork of ``openpyxl``, the standard
Python library for reading and writing Excel ``.xlsx``, ``.xlsm``, ``.xltx``,
and ``.xltm`` files.

This fork starts from upstream ``openpyxl`` tag ``3.1.5``. The package is
renamed for distribution, but the Python import name remains ``openpyxl``.
That mismatch is intentional and must not be "fixed": this package is meant to
be a drop-in replacement for existing code that says ``import openpyxl``.

Name map
--------

* GitHub repository: ``paper-xlsx``
* PyPI distribution: ``paper-xlsx``
* Built wheel/sdist names: ``paper_xlsx-*``
* Python import: ``openpyxl``
* Fork sentinel: ``openpyxl.__paper_version__``

Installation
------------

From PyPI::

    pip install paper-xlsx

From the repository::

    pip install "paper-xlsx @ git+https://github.com/The-LLM-Data-Company/paper-xlsx.git@main"

Verification
------------

::

    python -c "import openpyxl; print(openpyxl.__paper_version__)"

Expected output::

    0.1.0

Fork ledger
-----------

``PAPER.md`` records the upstream base tag, conversion notes, baseline test
results, sanctioned deviations, release safety notes, and future upstream merge
policy.
