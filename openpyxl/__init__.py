# Copyright (c) 2010-2024 openpyxl

"""paper-xlsx is Paper Instruments' import-compatible fork of openpyxl for
creating and safely editing Excel workbooks.

The import name stays ``openpyxl`` so existing imports are unchanged. The
standard object model creates new workbooks. For existing files, preserve mode
keeps the original package bytes as the source of truth and splices supported
edits into them. Unsafe edits raise a typed `openpyxl.errors` exception instead
of writing a damaged file.

Start here: `openpyxl.__paper_version__` (fork sentinel), and the
`openpyxl.preserve`, `openpyxl.oracle` and
`openpyxl.errors` modules. The project README and ``doc/paper.rst``
give the full tour.
"""

DEBUG = False

from openpyxl._distribution import assert_single_openpyxl_distribution
from openpyxl._paper_version import __paper_version__

assert_single_openpyxl_distribution()
del assert_single_openpyxl_distribution

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
