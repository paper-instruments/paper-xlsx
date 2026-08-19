#!/usr/bin/env python

"""Setup script for packaging paper-xlsx."""

import os
import sys
from importlib.util import module_from_spec, spec_from_file_location

from setuptools import setup, find_packages

here = os.path.abspath(os.path.dirname(__file__))
try:
    with open(os.path.join(here, 'README.md')) as f:
        README = f.read()
except IOError:
    README = ''

spec = spec_from_file_location(
    "constants", os.path.join(here, "openpyxl", "_constants.py"))
constants = module_from_spec(spec)
spec.loader.exec_module(constants)

paper_spec = spec_from_file_location(
    "paper_version", os.path.join(here, "openpyxl", "_paper_version.py"))
paper_version = module_from_spec(paper_spec)
paper_spec.loader.exec_module(paper_version)

__author__ = constants.__author__
__author_email__ = constants.__author_email__
__license__ = constants.__license__
__maintainer_email__ = constants.__maintainer_email__
__url__ = constants.__url__
__version__ = constants.__version__
__python__ = constants.__python__

PAPER_VERSION = paper_version.__paper_version__

def cythonize_modules():
    from Cython.Build import cythonize
    return cythonize([
        "openpyxl/worksheet/_reader.py",
        "openpyxl/worksheet/_writer.py",
        "openpyxl/utils/cell.py",
        ],
        nthreads=3,
        language_level=3,
    )


try:
    sys.argv.remove("--with-cython")
except ValueError:
    ext_modules = None
else:
    ext_modules = cythonize_modules()


setup(
    name='paper-xlsx',
    packages=find_packages(".",
        exclude=["*.tests", "tests", "tests.*", "scratchpad*", "*.c",]
        ),
    ext_modules=ext_modules,
    package_dir={},
    # metadata
    version=PAPER_VERSION,
    description=("A drop-in openpyxl fork for safe inspection and editing "
                 "of existing Excel files"),
    long_description=README,
    long_description_content_type="text/markdown",
    author=f"{__author__}; Paper Instruments, Inc.",
    author_email=__author_email__,
    url="https://github.com/paper-instruments/paper-xlsx",
    license=__license__,
    maintainer="Paper Instruments, Inc.",
    python_requires=f">={__python__}",
    install_requires=[
        'et_xmlfile',
        ],
    entry_points={
        'console_scripts': [
            'paper-xlsx-doctor=paper_xlsx_doctor:main',
        ],
    },
    project_urls={
        'Source': 'https://github.com/paper-instruments/paper-xlsx',
        'Tracker': 'https://github.com/paper-instruments/paper-xlsx/issues',
        'Changelog': 'https://github.com/paper-instruments/paper-xlsx/releases',
        'Upstream': 'https://foss.heptapod.net/openpyxl/openpyxl',
    },
    classifiers=[
                 'Development Status :: 3 - Alpha',
                 'Operating System :: MacOS :: MacOS X',
                 'Operating System :: Microsoft :: Windows',
                 'Operating System :: POSIX',
                 'License :: OSI Approved :: MIT License',
                 'Programming Language :: Python',
                 'Programming Language :: Python :: 3.9',
                 'Programming Language :: Python :: 3.10',
                 'Programming Language :: Python :: 3.11',
                 'Programming Language :: Python :: 3.12',
                 'Programming Language :: Python :: 3.13',
                 ],
    )
