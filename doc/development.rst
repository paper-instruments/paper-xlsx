Development
===========

The canonical contribution guide is
`CONTRIBUTING.md <https://github.com/paper-instruments/paper-xlsx/blob/main/CONTRIBUTING.md>`_.
It describes the fork's safety rules, fixture policy, and pull request process.


Set up a development environment
--------------------------------

Clone the GitHub repository and create a virtual environment::

    git clone https://github.com/paper-instruments/paper-xlsx.git
    cd paper-xlsx
    python -m venv .venv
    source .venv/bin/activate

On Windows, activate the environment with
``.venv\Scripts\activate`` instead. Then install the package and test
dependencies::

    python -m pip install -e .
    python -m pip install pytest "lxml>=6.1,<7" pillow pandas


Run tests
---------

Run the full test suite with::

    python -m pytest -q

Tests that need LibreOffice skip when ``soffice`` is not available. Install
LibreOffice Calc to run those tests locally. GitHub Actions runs the supported
Python versions and the additional dependency, XML backend, Windows,
LibreOffice, documentation, and distribution checks.


Build the documentation
-----------------------

Use the same commands as the documentation CI job::

    python -m pip install -e . -r doc/requirements.txt "lxml>=6.1,<7"
    sphinx-build -W --keep-going -b html doc doc/.build/html

Update the documentation when a public API or behavior changes. Add a matching
entry to ``doc/changes.rst`` under "Unreleased" for user-visible changes.
