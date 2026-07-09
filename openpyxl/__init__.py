# Copyright (c) 2010-2024 openpyxl

"""paper-xlsx — Paper Instruments' hard fork of openpyxl for LOSSLESS,
SAFE editing of existing Excel files.

The import name stays ``openpyxl`` so existing code is unchanged. The
fork adds a **preserve mode**: ``load_workbook(path, preserve=True)``
keeps the original package bytes as the source of truth, so saving
splices your edits back in without destroying the charts, pivots, VBA,
or formatting a normal openpyxl round-trip drops — and any edit it
cannot make safely refuses loudly (a typed :mod:`openpyxl.errors`
exception) instead of corrupting the file.

Start here: :attr:`openpyxl.__paper_version__` (fork sentinel), and the
:mod:`openpyxl.preserve`, :mod:`openpyxl.oracle` and
:mod:`openpyxl.errors` modules. The project README and ``doc/paper.rst``
give the full tour.
"""

DEBUG = False

from openpyxl.compat.numbers import NUMPY
from openpyxl.xml import DEFUSEDXML, LXML
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook as open
from openpyxl.reader.excel import load_workbook
import openpyxl._constants as constants

# Expose constants especially the version number

__author__ = constants.__author__
__author_email__ = constants.__author_email__
__license__ = constants.__license__
__maintainer_email__ = constants.__maintainer_email__
__url__ = constants.__url__
__version__ = constants.__version__
__paper_version__ = "0.1.0"
