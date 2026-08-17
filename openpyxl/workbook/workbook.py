# Copyright (c) 2010-2024 openpyxl

"""Workbook is the top-level container for all document information."""
from copy import copy

from openpyxl.compat import deprecated
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet._write_only import WriteOnlyWorksheet
from openpyxl.worksheet.copier import WorksheetCopy

from openpyxl.utils import quote_sheetname
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.utils.datetime  import WINDOWS_EPOCH, MAC_EPOCH
from openpyxl.utils.exceptions import ReadOnlyWorkbookException

from openpyxl.writer.excel import save_workbook

from openpyxl.styles.cell_style import StyleArray
from openpyxl.styles.named_styles import NamedStyle
from openpyxl.styles.differential import DifferentialStyleList
from openpyxl.styles.alignment import Alignment
from openpyxl.styles.borders import DEFAULT_BORDER
from openpyxl.styles.fills import DEFAULT_EMPTY_FILL, DEFAULT_GRAY_FILL
from openpyxl.styles.fonts import DEFAULT_FONT
from openpyxl.styles.protection import Protection
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.styles.named_styles import NamedStyleList
from openpyxl.styles.table import TableStyleList

from openpyxl.chartsheet import Chartsheet
from openpyxl.preserve import ledger as _ledger
from .defined_name import DefinedName, DefinedNameDict
from openpyxl.packaging.core import DocumentProperties
from openpyxl.packaging.custom import CustomPropertyList
from openpyxl.packaging.relationship import RelationshipList
from .child import _WorkbookChild
from .protection import DocumentSecurity
from .properties import CalcProperties
from .views import BookView


from openpyxl.xml.constants import (
    XLSM,
    XLSX,
    XLTM,
    XLTX
)

INTEGER_TYPES = (int,)


def _require_materialized_cells(wb, api):
    """The perception verbs read ws._cells; read-only and write-only
    workbooks never materialize it (raw AttributeError /
    silently empty results)."""
    if getattr(wb, "_read_only", False) or wb.write_only:
        raise ValueError(
            "{0} needs materialized cells; read-only and write-only "
            "workbooks do not hold them. Load normally (or with "
            "preserve=True) instead.".format(api))


def _guard_data_only_cell_mutation(workbook, target, operation="cell write"):
    """Prove a data-only mutation target was not a source formula."""
    _guard_data_only_range_mutation(
        workbook, target.parent,
        (target.column, target.row, target.column, target.row), operation,
        anchor="{0}!{1}".format(target.parent.title, target.coordinate))


def _guard_data_only_range_mutation(workbook, worksheet, bounds, operation,
                                    anchor=None):
    """Prove a data-only mutation range does not intersect a formula."""
    if not workbook.data_only:
        return
    from openpyxl.errors import UnsupportedStructureError

    ledger = getattr(workbook, "_paper_ledger", None)
    anchor = anchor or "{0}!{1}".format(worksheet.title, bounds)
    if ledger is None or not ledger.armed:
        return
    if worksheet in ledger.added_sheets:
        return

    import io
    import zipfile

    from openpyxl.preserve.saver import _package_info
    from openpyxl.xml.functions import fromstring

    original_title = ledger.renames.get(worksheet, worksheet.title)
    source = getattr(workbook, "_paper_source", None)
    try:
        with zipfile.ZipFile(io.BytesIO(source)) as archive:
            _workbook_part, sheet_parts = _package_info(archive)
            sheet_part = sheet_parts.get(original_title)
            payload = archive.read(sheet_part) if sheet_part else None
    except (KeyError, TypeError, ValueError, zipfile.BadZipFile):
        payload = None
    if payload is None:
        raise UnsupportedStructureError(
            "{0} cannot prove whether {1} was a formula because its "
            "retained worksheet structure is unavailable. Nothing was "
            "changed.".format(operation, anchor),
            kind="data-only-input-model-unavailable",
            anchor=anchor,
        )
    root = fromstring(payload)
    from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries

    min_col, min_row, max_col, max_row = bounds
    formula_hit = False
    for source_cell in root.iter():
        if source_cell.tag.rsplit("}", 1)[-1] != "c":
            continue
        formulas = [child for child in source_cell
                    if child.tag.rsplit("}", 1)[-1] == "f"]
        if not formulas:
            continue
        coordinate = source_cell.get("r")
        if coordinate:
            row, column = coordinate_to_tuple(coordinate)
            if min_row <= row <= max_row and min_col <= column <= max_col:
                formula_hit = True
                break
        for formula in formulas:
            ref = formula.get("ref")
            if not ref:
                continue
            try:
                f_min_col, f_min_row, f_max_col, f_max_row = \
                    range_boundaries(ref)
            except (TypeError, ValueError):
                formula_hit = True
                break
            if not (f_max_col < min_col or f_min_col > max_col
                    or f_max_row < min_row or f_min_row > max_row):
                formula_hit = True
                break
        if formula_hit:
            break
    if formula_hit:
        raise UnsupportedStructureError(
            "{0} cannot replace {1}: the retained source range contains "
            "a formula. Nothing was changed.".format(operation, anchor),
            kind="input-is-calculation",
            anchor=anchor,
        )


class Workbook:
    """Workbook is the container for all other parts of the document."""

    _read_only = False
    _data_only = False
    # paper-xlsx preserve mode: set by the reader, never directly
    _preserve = False
    _paper_source = None            # retained source-package bytes
    _paper_source_identity = None   # content-bound source path identity
    _paper_content_type = None      # source workbook type under preserve
    _paper_ledger = None            # the dirty ledger; armed after load
    # distinguishes stock-loaded workbooks from new in-memory workbooks
    _paper_loaded_from_package = False
    # protection awareness: True turns writes to locked
    # cells on protected sheets into typed refusals (default: warn once
    # per sheet). Protection is reported, never enforced or bypassed.
    # preserve-mode workbooks only (the check rides the armed
    # ledger); on stock loads the flag is inert.
    strict_protection = False
    template = False
    path = "/xl/workbook.xml"

    def __init__(self,
                 write_only=False,
                 iso_dates=False,
                 ):
        self._sheets = []
        self._pivots = []
        self._active_sheet_index = 0
        self.defined_names = DefinedNameDict()
        self._external_links = []
        self.properties = DocumentProperties()
        self.custom_doc_props = CustomPropertyList()
        self.security = DocumentSecurity()
        self.__write_only = write_only
        self.shared_strings = IndexedList()

        self._setup_styles()

        self.loaded_theme = None
        self.vba_archive = None
        self.is_template = False
        self.code_name = None
        self.epoch = WINDOWS_EPOCH
        self.encoding = "utf-8"
        self.iso_dates = iso_dates

        if not self.write_only:
            self._sheets.append(Worksheet(self))

        self.rels = RelationshipList()
        self.calculation = CalcProperties()
        self.views = [BookView()]


    def _setup_styles(self):
        """Bootstrap styles"""

        self._fonts = IndexedList()
        self._fonts.add(DEFAULT_FONT)

        self._alignments = IndexedList([Alignment()])

        self._borders = IndexedList()
        self._borders.add(DEFAULT_BORDER)

        self._fills = IndexedList()
        self._fills.add(DEFAULT_EMPTY_FILL)
        self._fills.add(DEFAULT_GRAY_FILL)

        self._number_formats = IndexedList()
        self._date_formats = {}
        self._timedelta_formats = {}

        self._protections = IndexedList([Protection()])

        self._colors = COLOR_INDEX
        self._cell_styles = IndexedList([StyleArray()])
        self._named_styles = NamedStyleList()
        self.add_named_style(NamedStyle(font=copy(DEFAULT_FONT), border=copy(DEFAULT_BORDER), builtinId=0))
        self._table_styles = TableStyleList()
        self._differential_styles = DifferentialStyleList()


    @property
    def epoch(self):
        if self._epoch == WINDOWS_EPOCH:
            return WINDOWS_EPOCH
        return MAC_EPOCH


    @epoch.setter
    def epoch(self, value):
        if value not in (WINDOWS_EPOCH, MAC_EPOCH):
            raise ValueError("The epoch must be either 1900 or 1904")
        self._epoch = value


    @property
    def read_only(self):
        return self._read_only

    @property
    def preserve(self):
        """True when this workbook was loaded with ``preserve=True``: the
        original package bytes are the source of truth and save is a
        lossless splice of recorded edits into them."""
        return self._preserve

    def set_pivot_refresh_on_load(self, pivots=None, *, all=False):
        """Request refresh-on-load for named pivots or explicitly all pivots.

        Pivot names may be sheet-qualified (``"Sheet!PivotName"``). Multiple
        pivots that share a cache necessarily share this cache-level setting.
        """
        if self._paper_ledger is None or not self._paper_ledger.armed:
            raise ValueError(
                "pivot refresh is only available under preserve mode")
        from openpyxl.preserve.pivots import resolve_requests

        parts = resolve_requests(self, pivots, all=all)
        self._paper_ledger.pivot_refresh_requests.update(parts)
        return parts

    def search(self, text_or_regex, *, regex=False, values=True,
               formulas=True):
        """Find text across the workbook (paper-xlsx).
        Returns ``[{"address", "match", "kind"}, ...]`` where kind is
        "value" or "formula"."""
        import re as _re

        _require_materialized_cells(self, "search()")
        if regex:
            try:
                pattern = _re.compile(text_or_regex)
            except _re.error as exc:
                raise ValueError(
                    "search(regex=True) got an invalid pattern "
                    "{0!r}: {1}".format(text_or_regex, exc))
        else:
            pattern = None
        results = []
        for ws in self.worksheets:
            for (row, col), cell in sorted(ws._cells.items()):
                value = cell._value
                if value is None:
                    continue
                is_formula = cell.data_type == "f"
                if is_formula and not formulas:
                    continue
                if not is_formula and not values:
                    continue
                if is_formula and not isinstance(value, str):
                    # ArrayFormula/DataTableFormula objects: search their
                    # TEXT, never the Python repr (repr
                    # fabricated matches and hid real ones)
                    text = getattr(value, "text", None)
                    if not isinstance(text, str):
                        continue
                else:
                    text = value if isinstance(value, str) else str(value)
                if pattern is not None:
                    m = pattern.search(text)
                    if m is None:
                        continue
                    match = m.group(0)
                else:
                    if str(text_or_regex) not in text:
                        continue
                    match = str(text_or_regex)
                results.append({
                    "address": "{0}!{1}".format(ws.title,
                                                cell.coordinate),
                    "match": match,
                    "kind": "formula" if is_formula else "value",
                })
        return results

    def validate(self):
        """Run the preserve save planner without assembling an archive."""
        if not self._preserve or self._paper_ledger is None:
            raise ValueError(
                "validate() runs the preserve save planner and is "
                "only available on workbooks loaded with preserve=True.")
        from openpyxl.preserve.saver import validate_preserved

        validate_preserved(self)
        return None

    @property
    def data_only(self):
        return self._data_only

    @property
    def write_only(self):
        return self.__write_only


    @property
    def excel_base_date(self):
        return self.epoch

    @property
    def active(self):
        """Get the currently active sheet or None

        :type: `openpyxl.worksheet.worksheet.Worksheet`
        """
        try:
            return self._sheets[self._active_sheet_index]
        except IndexError:
            pass

    @active.setter
    def active(self, value):
        """Set the active sheet"""
        if not isinstance(value, (_WorkbookChild, INTEGER_TYPES)):
            raise TypeError("Value must be either a worksheet, chartsheet or numerical index")
        if isinstance(value, INTEGER_TYPES):
            self._active_sheet_index = value
            return
            #if self._sheets and 0 <= value < len(self._sheets):
                #value = self._sheets[value]
            #else:
                #raise ValueError("Sheet index is outside the range of possible values", value)
        if value not in self._sheets:
            raise ValueError("Worksheet is not in the workbook")
        if value.sheet_state != "visible":
            raise ValueError("Only visible sheets can be made active")

        idx = self._sheets.index(value)
        self._active_sheet_index = idx


    def create_sheet(self, title=None, index=None):
        """Create a worksheet (at an optional index).

        :param title: optional title of the sheet
        :type title: str
        :param index: optional position at which the sheet will be inserted
        :type index: int

        """
        if self.read_only:
            raise ReadOnlyWorkbookException('Cannot create new sheet in a read-only workbook')

        ledger = getattr(self, "_paper_ledger", None)
        snapshot = None
        sheets = self._sheets
        sheet_values = list(sheets)
        if ledger is not None and ledger.armed:
            from openpyxl.preserve.structural import _capture_structural_state
            snapshot = _capture_structural_state(self)
        try:
            if self.write_only:
                new_ws = WriteOnlyWorksheet(parent=self, title=title)
            else:
                new_ws = Worksheet(parent=self, title=title)

            self._add_sheet(sheet=new_ws, index=index)
            _ledger.mark_sheet_added(self, new_ws)
            return new_ws
        except BaseException:
            if snapshot is not None:
                from openpyxl.preserve.structural import \
                    _restore_structural_state
                _restore_structural_state(self, snapshot)
                self._sheets = sheets
                sheets[:] = sheet_values
            raise


    def _add_sheet(self, sheet, index=None):
        """Add an worksheet (at an optional index)."""

        if not isinstance(sheet, (Worksheet, WriteOnlyWorksheet, Chartsheet)):
            raise TypeError("Cannot be added to a workbook")

        if sheet.parent != self:
            raise ValueError("You cannot add worksheets from another workbook.")

        if index is None:
            self._sheets.append(sheet)
        else:
            self._sheets.insert(index, sheet)


    def move_sheet(self, sheet, offset=0):
        """
        Move a sheet or sheetname
        """
        if not isinstance(sheet, Worksheet):
            sheet = self[sheet]
        # reorder is expressed at save by reordering the
        # ORIGINAL <sheet> entry bytes; definedNames/bookViews re-render
        # (localSheetId and activeTab are position-derived by the writer)
        idx = self._sheets.index(sheet)
        del self._sheets[idx]
        new_pos = idx + offset
        self._sheets.insert(new_pos, sheet)


    def remove(self, worksheet):
        """Remove `worksheet` from this workbook.

        Under preserve mode a loaded sheet's removal runs the reference
        audit first. The package-part cascade happens at save."""
        if getattr(worksheet, "parent", None) is not self \
                or worksheet not in self._sheets:
            raise ValueError("Worksheet is not part of this workbook.")

        ledger = getattr(self, "_paper_ledger", None)
        snapshot = None
        sheets = self._sheets
        sheet_values = list(sheets)
        if ledger is not None and ledger.armed:
            from openpyxl.preserve.structural import _capture_structural_state
            snapshot = _capture_structural_state(self)
        try:
            if not _ledger.allow_sheet_removal(self, worksheet):
                _ledger.audit_sheet_removal(self, worksheet)
                _ledger.record_sheet_removal(self, worksheet)
            self._sheets.remove(worksheet)
        except BaseException:
            if snapshot is not None:
                from openpyxl.preserve.structural import \
                    _restore_structural_state
                _restore_structural_state(self, snapshot)
                self._sheets = sheets
                sheets[:] = sheet_values
            raise
        return None


    @deprecated("Use wb.remove(worksheet) or del wb[sheetname]")
    def remove_sheet(self, worksheet):
        """Remove `worksheet` from this workbook."""
        self.remove(worksheet)


    def create_chartsheet(self, title=None, index=None):
        if self.read_only:
            raise ReadOnlyWorkbookException("Cannot create new sheet in a read-only workbook")
        _ledger.refuse_sheet_lifecycle(
            self, "create_chartsheet",
            "generating chartsheet and drawing parts alongside the "
            "preserved package is not supported in v0; the chartsheet "
            "would otherwise be silently absent from the saved file.")
        cs = Chartsheet(parent=self, title=title)

        self._add_sheet(cs, index)
        return cs


    @deprecated("Use wb[sheetname]")
    def get_sheet_by_name(self, name):
        """Returns a worksheet by its name.

        :param name: the name of the worksheet to look for
        :type name: string

        """
        return self[name]

    def __contains__(self, key):
        return key in self.sheetnames


    def index(self, worksheet):
        """Return the index of a worksheet."""
        return self.worksheets.index(worksheet)


    @deprecated("Use wb.index(worksheet)")
    def get_index(self, worksheet):
        """Return the index of the worksheet."""
        return self.index(worksheet)

    def __getitem__(self, key):
        """Returns a worksheet by its name.

        :param name: the name of the worksheet to look for
        :type name: string

        """
        for sheet in self.worksheets + self.chartsheets:
            if sheet.title == key:
                return sheet
        raise KeyError("Worksheet {0} does not exist.".format(key))

    def __delitem__(self, key):
        sheet = self[key]
        self.remove(sheet)

    def __iter__(self):
        return iter(self.worksheets)


    @deprecated("Use wb.sheetnames")
    def get_sheet_names(self):
        return self.sheetnames

    @property
    def worksheets(self):
        """A list of sheets in this workbook

        :type: list of `openpyxl.worksheet.worksheet.Worksheet`
        """
        return [s for s in self._sheets if isinstance(s, (Worksheet, ReadOnlyWorksheet, WriteOnlyWorksheet))]

    @property
    def chartsheets(self):
        """A list of Chartsheets in this workbook

        :type: list of `openpyxl.chartsheet.chartsheet.Chartsheet`
        """
        return [s for s in self._sheets if isinstance(s, Chartsheet)]

    @property
    def sheetnames(self):
        """Returns the list of the names of worksheets in this workbook.

        Names are returned in the worksheets order.

        :type: list of strings

        """
        return [s.title for s in self._sheets]


    @deprecated("Assign scoped named ranges directly to worksheets or global ones to the workbook. Deprecated in 3.1")
    def create_named_range(self, name, worksheet=None, value=None, scope=None):
        """Create a new named_range on a worksheet

        """
        defn = DefinedName(name=name)
        if worksheet is not None:
            defn.value = "{0}!{1}".format(quote_sheetname(worksheet.title), value)
        else:
            defn.value = value

        self.defined_names[name] = defn


    def add_named_style(self, style):
        """
        Add a named style
        """
        self._named_styles.append(style)
        style.bind(self)


    @property
    def named_styles(self):
        """
        List available named styles
        """
        return self._named_styles.names


    @property
    def mime_type(self):
        """
        The mime type is determined by whether a workbook is a template or
        not and whether it contains macros or not. Excel requires the file
        extension to match but openpyxl does not enforce this.

        """
        preserved = getattr(self, "_paper_content_type", None)
        if self._preserve and preserved in (XLSX, XLSM, XLTX, XLTM):
            return preserved
        ct = self.template and XLTX or XLSX
        if self.vba_archive:
            ct = self.template and XLTM or XLSM
        return ct


    def save(self, filename, *, allow_formula_loss=False, receipt=False):
        """Save the current workbook under the given `filename`.
        Use this function instead of using an `ExcelWriter`.

        :param allow_formula_loss: a workbook loaded with ``data_only=True``
            holds cached values instead of formulas, so saving destroys
            formulas. Under preserve mode such a save refuses unless this
            flag is set (and even then only cells you actually edited lose
            their formulas — untouched cells keep them in the original
            bytes). On the stock path the flag silences the loud warning.
        :param receipt: preserve mode only — return an
            `openpyxl.preserve.receipts.EditReceipt` comparing the
            saved file against the AS-LOADED source bytes. NOTE: after several saves from one session the receipt
            is cumulative — it describes the session, not the last call.

        .. warning::
            When creating your workbook using `write_only` set to True,
            you will only be able to call this function once. Subsequent attempts to
            modify or save the file will raise an `openpyxl.shared.exc.WorkbookAlreadySaved` exception.
        """
        if self.read_only:
            raise TypeError("""Workbook is read-only""")
        if receipt and (not self._preserve or self._paper_source is None):
            raise ValueError(
                "save(receipt=True) compares against the preserved source "
                "bytes and is only available under preserve mode.")
        if self.write_only and not self.worksheets:
            self.create_sheet()
        if receipt:
            from io import BytesIO

            from openpyxl.preserve import zipio
            from openpyxl.preserve.receipts import receipt as _receipt
            from openpyxl.preserve.saver import _expected_delivery_identity

            zipio.validate_target(filename)
            expected_identity = _expected_delivery_identity(self, filename)
            staged = BytesIO()
            save_workbook(
                self, staged, allow_formula_loss=allow_formula_loss)
            data = staged.getvalue()
            result = _receipt(
                self._paper_source, data, _ledger=self._paper_ledger,
                _workbook=self)

            def validate_source():
                if self._paper_source_identity is not None:
                    zipio._assert_path_identity(
                        self._paper_source_identity)

            committed_identity = zipio.deliver(
                data, filename, expected_identity=expected_identity,
                precommit=validate_source,
                postcommit=(None if self._paper_source_identity is not None
                            and expected_identity is not None
                            and (expected_identity.requested ==
                                 self._paper_source_identity.requested
                                 or zipio._same_occupant(
                                     expected_identity,
                                     self._paper_source_identity))
                            else validate_source))
            if self._paper_source_identity is not None \
                    and expected_identity is not None \
                    and (expected_identity.requested ==
                         self._paper_source_identity.requested
                         or zipio._same_occupant(
                             expected_identity,
                             self._paper_source_identity)):
                self._paper_source_identity = committed_identity
            return result
        save_workbook(self, filename, allow_formula_loss=allow_formula_loss)
        return None


    @property
    def style_names(self):
        """
        List of named styles
        """
        return [s.name for s in self._named_styles]


    def copy_worksheet(self, from_worksheet):
        """Copy an existing worksheet in the current workbook

        .. warning::
            This function cannot copy worksheets between workbooks.
            worksheets can only be copied within the workbook that they belong

        :param from_worksheet: the worksheet to be copied from
        :return: copy of the initial worksheet
        """
        if self.__write_only or self._read_only:
            raise ValueError("Cannot copy worksheets in read-only or write-only mode")
        ledger = getattr(self, "_paper_ledger", None)
        if ledger is not None and ledger.armed \
                and self.data_only \
                and getattr(from_worksheet, "parent", None) is self \
                and from_worksheet in self.worksheets \
                and from_worksheet not in ledger.added_sheets:
            from openpyxl.errors import UnsupportedStructureError

            raise UnsupportedStructureError(
                "copy_worksheet() cannot faithfully copy loaded sheet {0!r} "
                "after loading with data_only=True. The live model contains "
                "cached values, not the source formulas the copy would need. "
                "Nothing was changed.".format(from_worksheet.title),
                kind="data-only-reference-model-unavailable",
                anchor=from_worksheet.title,
            )
        # the copy registers as an ADDED sheet (create_sheet
        # below is ledger-hooked) and is generated whole at save; charts/
        # images do not copy (upstream's copier skips them), comments and
        # hyperlinks ride the added-sheet generators
        new_title = u"{0} Copy".format(from_worksheet.title)
        to_worksheet = self.create_sheet(title=new_title)
        cp = WorksheetCopy(source_worksheet=from_worksheet, target_worksheet=to_worksheet)
        cp.copy_worksheet()
        return to_worksheet


    def close(self):
        """
        Close workbook file if open. Only affects read-only and write-only modes.
        """
        if hasattr(self, '_archive'):
            self._archive.close()


    def _duplicate_name(self, name):
        """
        Check for duplicate name in defined name list and table list of each worksheet.
        Names are not case sensitive.
        """
        name = name.lower()
        for sheet in self.worksheets:
            for t in sheet.tables:
                if name == t.lower():
                    return True

        if name in self.defined_names:
            return True
