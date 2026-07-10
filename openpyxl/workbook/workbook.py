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

class Workbook:
    """Workbook is the container for all other parts of the document."""

    _read_only = False
    _data_only = False
    # paper-xlsx preserve mode: set by the reader, never directly
    _preserve = False
    _paper_source = None            # retained source-package bytes
    _paper_loss_inventory = None    # content the stock save cannot preserve
    _paper_ledger = None            # the dirty ledger; armed after load
    # protection awareness: True turns writes to locked
    # cells on protected sheets into typed refusals (default: warn once
    # per sheet). Protection is reported, never enforced or bypassed.
    # preserve-mode workbooks only (the check rides the armed
    # ledger); on stock loads the flag is inert.
    strict_protection = False
    # formula pre-flight lint mode at the value-bind chokepoint
    # (preserve-mode workbooks only): "off"|"warn"|"refuse"
    formula_lint = "warn"
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

    def mark_dirty(self, target):
        """Escape hatch for anyone reaching below the public API in
        preserve mode: declare that ``target`` was
        mutated so the splice save re-emits it from the model.

        ``target`` is either a sheet-qualified A1 range (``"Model!B7"``,
        ``"'My Sheet'!B2:D10"``) or an exact package part name
        (``"xl/media/image1.png"``). Raises
        :class:`openpyxl.errors.TargetNotFoundError` for unknown targets and
        ``ValueError`` outside preserve mode.
        """
        _ledger.mark_dirty_target(self, target)

    def replace_part(self, name, payload):
        """Raw byte swap of one unmanaged package part under preserve
        mode — media swaps are the intended use
        (``wb.replace_part("xl/media/image1.png", new_png_bytes)``).

        The part must exist (:class:`~openpyxl.errors.TargetNotFoundError`
        otherwise); parts the model actively manages (sheets, workbook,
        styles, sharedStrings, content types) refuse with
        :class:`~openpyxl.errors.RelationshipPolicyError` — replacing them
        raw would desync the model. Guards run NOW; bytes land at save.
        """
        if self._paper_ledger is None or not self._paper_ledger.armed:
            raise ValueError(
                "replace_part is only meaningful in preserve mode "
                "(load_workbook(..., preserve=True)).")
        from openpyxl.preserve.lifecycle import check_replace_part

        check_replace_part(self, name)
        if not isinstance(payload, bytes):
            raise TypeError("payload must be bytes")
        self._paper_ledger.replaced_parts[name] = payload

    def set_input(self, name_or_label, value):
        """Set a model INPUT by defined name or text label (paper-xlsx): resolution order is defined names, then
        ``locate`` over every sheet (a label found on several sheets is
        ambiguous). Refuses to overwrite a formula cell — set_input never
        destroys a calculation. Returns the Cell written."""
        from openpyxl.errors import (
            AmbiguousTargetError,
            TargetNotFoundError,
            UnsupportedStructureError,
        )

        target = None
        dn = self.defined_names.get(name_or_label)
        if dn is not None:
            destinations = list(dn.destinations)
            if len(destinations) != 1:
                raise AmbiguousTargetError(
                    "defined name {0!r} resolves to {1} areas; single "
                    "cells only.".format(name_or_label, len(destinations)),
                    kind="ambiguous-name", options=[
                        "{0}!{1}".format(t, r) for t, r in destinations])
            title, coord = destinations[0]
            if ":" in coord:
                raise AmbiguousTargetError(
                    "defined name {0!r} resolves to a RANGE ({1}); "
                    "set_input takes single cells only.".format(
                        name_or_label, coord),
                    kind="ambiguous-name",
                    options=["{0}!{1}".format(title, coord)])
            target = self[title][coord.replace("$", "")]
        else:
            hits = []
            for ws in self.worksheets:
                try:
                    hits.append(ws.locate(name_or_label))
                except TargetNotFoundError:
                    continue
            if not hits:
                raise TargetNotFoundError(
                    "{0!r} is neither a defined name nor a label on any "
                    "sheet.".format(name_or_label),
                    kind="input-not-found")
            if len(hits) > 1:
                options = ["{0}!{1}".format(c.parent.title, c.coordinate)
                           for c in hits]
                raise AmbiguousTargetError(
                    "label {0!r} resolves on {1} sheets: {2}. Qualify the "
                    "request.".format(name_or_label, len(hits),
                                      ", ".join(options)),
                    kind="ambiguous-input", options=options)
            target = hits[0]
        from openpyxl.cell.cell import MergedCell

        if isinstance(target, MergedCell):
            raise UnsupportedStructureError(
                "{0}!{1} is inside a merged range; write the input to the "
                "merge's anchor cell instead. Nothing was changed.".format(
                    target.parent.title, target.coordinate),
                kind="input-is-merged-interior",
                anchor="{0}!{1}".format(target.parent.title,
                                        target.coordinate))
        if target.data_type == "f":
            raise UnsupportedStructureError(
                "{0}!{1} holds a formula; set_input never overwrites "
                "calculations. Nothing was changed.".format(
                    target.parent.title, target.coordinate),
                kind="input-is-calculation",
                anchor="{0}!{1}".format(target.parent.title,
                                        target.coordinate))
        target.value = value
        return target

    def protect_for_delivery(self, password=None):
        """Lock every populated cell EXCEPT the model map's classified
        inputs, and enable sheet protection (paper-xlsx). Every non-input cell is ACTIVELY locked (a workbook authored
        with locked=False cells — templates, LibreOffice output — would
        otherwise ship editable under a "protected" sheet). Each cell's other protection flags (hidden) are preserved.
        Protection is advisory in the file format and REPORTED here —
        returns {"locked_sheets", "unlocked_inputs", "locked_cells"}."""
        from copy import copy as _copy

        from openpyxl.styles import Protection

        mm = self.model_map()
        unlocked = []
        locked_count = 0
        locked_sheets = []

        def _set_locked(cell, locked):
            prot = _copy(cell.protection) if cell.protection is not None \
                else Protection()
            prot.locked = locked
            cell.protection = prot

        for ws in self.worksheets:
            inputs = set(mm.sheets.get(ws.title, {}).get("inputs", []))
            for (row, col), cell in sorted(ws._cells.items()):
                if cell._value is None:
                    continue
                address = cell.coordinate
                if address in inputs:
                    _set_locked(cell, False)
                    unlocked.append("{0}!{1}".format(ws.title, address))
                else:
                    _set_locked(cell, True)
                    locked_count += 1
            ws.protection.sheet = True
            if password:
                ws.protection.password = password
            locked_sheets.append(ws.title)
        return {"locked_sheets": locked_sheets,
                "unlocked_inputs": unlocked,
                "locked_cells": locked_count}

    def scrub(self, remove=("comments", "metadata", "personal",
                            "hidden-sheets")):
        """Strip delivery-inappropriate content (paper-xlsx). Returns a scrub REPORT — everything removed is listed,
        everything that could NOT be removed is reported with its reason
        (hidden sheets whose removal would strand references refuse the
        removal and land in "skipped"; never silent).

        remove: any of "comments" (in-session comment objects; sheets
        with PRESERVED comment machinery are reported, not silently
        stripped), "metadata" (core document properties reset),
        "personal" (creator/lastModifiedBy cleared), "hidden-sheets"
        (removed via the audited removal path)."""
        from openpyxl.errors import PaperRefusal

        report = {"removed": [], "skipped": []}
        options = set(remove)
        unknown = options - {"comments", "metadata", "personal",
                             "hidden-sheets"}
        if unknown:
            raise ValueError("unknown scrub targets: {0}".format(
                sorted(unknown)))
        if "comments" in options:
            # a sheet whose comments come from PRESERVED machinery cannot
            # have them removed (the saver refuses editing preserved
            # comment parts) — nulling them would both LIE in the report
            # and brick the save. Detect
            # per sheet FIRST, then only null in-session comments on
            # comment-free sheets (compute-then-mutate keeps scrub atomic
            # and its report honest).
            preserved_sheets = set()
            if self._paper_source is not None:
                import io as _io
                import zipfile as _zipfile

                from openpyxl.preserve.comments import (
                    sheet_has_comment_machinery,
                )
                from openpyxl.preserve.saver import _package_info

                with _zipfile.ZipFile(
                        _io.BytesIO(self._paper_source)) as zin:
                    names = set(zin.namelist())
                    _wb, mapping = _package_info(zin)
                    led = self._paper_ledger
                    for ws in self.worksheets:
                        original = (led.renames.get(ws, ws.title)
                                    if led is not None else ws.title)
                        part = mapping.get(original)
                        if part is not None and sheet_has_comment_machinery(
                                zin, part, names):
                            preserved_sheets.add(ws.title)
            for ws in self.worksheets:
                if ws.title in preserved_sheets:
                    commented = sorted(
                        cell.coordinate for cell in ws._cells.values()
                        if cell._comment is not None)
                    if commented:
                        report["skipped"].append(
                            "sheet {0!r} carries preserved comment "
                            "machinery; its comments ({1}) cannot be "
                            "removed without rewriting preserved parts "
                            "(unsupported). Nothing on this sheet was "
                            "changed.".format(ws.title,
                                              ", ".join(commented)))
                    continue
                for (row, col), cell in sorted(ws._cells.items()):
                    if cell._comment is not None:
                        cell.comment = None
                        report["removed"].append(
                            "comment at {0}!{1}".format(ws.title,
                                                        cell.coordinate))
        if "metadata" in options:
            props = self.properties
            for attr in ("title", "subject", "description", "keywords",
                         "category", "contentStatus", "identifier"):
                if getattr(props, attr, None):
                    setattr(props, attr, None)
                    report["removed"].append(
                        "core property {0}".format(attr))
        if "personal" in options:
            props = self.properties
            for attr in ("creator", "lastModifiedBy"):
                if getattr(props, attr, None):
                    setattr(props, attr, None)
                    report["removed"].append(
                        "core property {0}".format(attr))
        if "hidden-sheets" in options:
            for ws in list(self.worksheets):
                if ws.sheet_state == "visible":
                    continue
                try:
                    self.remove(ws)
                    report["removed"].append(
                        "hidden sheet {0!r}".format(ws.title))
                except PaperRefusal as exc:
                    report["skipped"].append(
                        "hidden sheet {0!r}: {1}".format(ws.title,
                                                         exc))
        return report

    def set_pivot_refresh_on_load(self):
        """Byte-patch ``refreshOnLoad="1"`` onto every pivotCacheDefinition
        in the preserved package (paper-xlsx): pivots
        are preserved verbatim, so refresh-on-load is how their data stays
        honest after cell edits. Preserve mode only. Returns the list of
        parts patched."""
        if self._paper_ledger is None or not self._paper_ledger.armed:
            raise ValueError(
                "set_pivot_refresh_on_load() patches preserved pivot "
                "parts and is only available under preserve mode.")
        import io as _io
        import zipfile as _zipfile

        from openpyxl.preserve import crosspart

        patched = []
        with _zipfile.ZipFile(_io.BytesIO(self._paper_source)) as zin:
            for name in sorted(zin.namelist()):
                if not (name.startswith("xl/pivotCache/pivotCacheDefinition")
                        and name.endswith(".xml")):
                    continue
                payload = zin.read(name)
                root = crosspart.scan_small(payload,
                                            "pivotCacheDefinition",
                                            max_depth=1)
                if root.attrs.get("refreshOnLoad") == "1":
                    continue
                start, end, head = crosspart._patch_attr(
                    payload, root, "refreshOnLoad", "1")
                self._paper_ledger.replaced_parts[name] = (
                    payload[:start] + head + payload[end:])
                patched.append(name)
        return patched

    def model_map(self):
        """Role classification of every populated cell on formula-bearing
        sheets — inputs / calculations / outputs / constants (paper-xlsx). Returns
        :class:`openpyxl.preserve.modelmap.ModelMap`."""
        from openpyxl.preserve.modelmap import build_model_map

        _require_materialized_cells(self, "model_map()")
        return build_model_map(self)

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
        """Run the preserve saver's FULL validation pass without
        delivering a file (paper-xlsx): every refusal
        a save would raise is raised now; on success returns None and
        nothing is written anywhere."""
        if not self._preserve or self._paper_ledger is None:
            raise ValueError(
                "validate() replays the preserve save machinery and is "
                "only available on workbooks loaded with preserve=True.")
        import io as _io

        self.save(_io.BytesIO())
        return None

    def evaluate(self, set, read, *, timeout=120.0):
        """What-if scenario against THIS workbook's preserved source
        bytes: inputs applied to a temp copy through the
        spine, LibreOffice recalculates, outputs harvested. Neither the
        original file nor this live workbook is touched.

        NOTE: the run starts from the preserved AS-LOADED bytes — unsaved
        in-session edits are not part of the scenario (save first if they
        should be).

        ``set``: {address: value} single-cell inputs; ``read``: list of
        addresses to harvest. Addresses are sheet-qualified A1
        ("Model!B2") or defined names. Returns
        :class:`openpyxl.oracle.Evaluation`.
        """
        if not self._preserve or self._paper_source is None:
            raise ValueError(
                "evaluate() runs against the preserved source bytes and "
                "is only available on workbooks loaded with "
                "preserve=True.")
        from openpyxl import oracle

        return oracle.evaluate(self._paper_source, set, read,
                               timeout=timeout)

    def manifest(self):
        """A structured description of this workbook: sheets, formulas,
        defined names, volatile functions, a confession block enumerating
        content the package carries (charts, pivots, VBA, extensions), and
        what survives a save under the active mode.

        Returns a :class:`openpyxl.preserve.perception.WorkbookManifest`;
        call ``.to_dict()`` for the stable JSON form (schema
        ``workbook_manifest`` v1).
        """
        from openpyxl.preserve.perception import build_manifest

        return build_manifest(self)

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

        :type: :class:`openpyxl.worksheet.worksheet.Worksheet`
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

        if self.write_only :
            new_ws = WriteOnlyWorksheet(parent=self, title=title)
        else:
            new_ws = Worksheet(parent=self, title=title)

        self._add_sheet(sheet=new_ws, index=index)
        _ledger.mark_sheet_added(self, new_ws)
        return new_ws


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

        Under preserve mode a LOADED sheet's removal runs the reference
        audit first (anything on another sheet pointing at the victim
        refuses with the enumeration) and returns a
        :class:`~openpyxl.preserve.ledger.RemovalReport`; the part
        cascade happens at save."""
        report = None
        if not _ledger.allow_sheet_removal(self, worksheet):
            _ledger.audit_sheet_removal(self, worksheet)
            _ledger.record_sheet_removal(self, worksheet)
            report = True
        idx = self._sheets.index(worksheet)
        self._sheets.remove(worksheet)
        if report:
            from openpyxl.preserve.ledger import RemovalReport

            # parts enumerate at save; the report carries what is known now
            return RemovalReport([], remapped_names=len([
                n for ws in self._sheets
                for n in getattr(ws, "defined_names", {})]))
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

        :type: list of :class:`openpyxl.worksheet.worksheet.Worksheet`
        """
        return [s for s in self._sheets if isinstance(s, (Worksheet, ReadOnlyWorksheet, WriteOnlyWorksheet))]

    @property
    def chartsheets(self):
        """A list of Chartsheets in this workbook

        :type: list of :class:`openpyxl.chartsheet.chartsheet.Chartsheet`
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
            :class:`openpyxl.preserve.receipts.EditReceipt` comparing the
            saved file against the AS-LOADED source bytes. NOTE: after several saves from one session the receipt
            is cumulative — it describes the session, not the last call.

        .. warning::
            When creating your workbook using `write_only` set to True,
            you will only be able to call this function once. Subsequent attempts to
            modify or save the file will raise an :class:`openpyxl.shared.exc.WorkbookAlreadySaved` exception.
        """
        if self.read_only:
            raise TypeError("""Workbook is read-only""")
        if receipt and (not self._preserve or self._paper_source is None):
            raise ValueError(
                "save(receipt=True) compares against the preserved source "
                "bytes and is only available under preserve mode.")
        if self.write_only and not self.worksheets:
            self.create_sheet()
        save_workbook(self, filename, allow_formula_loss=allow_formula_loss)
        if receipt:
            from openpyxl.preserve.receipts import receipt as _receipt

            return _receipt(self._paper_source, filename)
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

