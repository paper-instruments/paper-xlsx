# paper-xlsx: model-to-file style index translation (PR-0 D2)

"""Translate the model's style indices into the ORIGINAL file's xf indices.

The model's style numbering drifts from the file's at load time:
``_normalise_numbers`` (styles/stylesheet.py:95,165-190) rewrites custom
numFmtIds in place, and ``apply_stylesheet`` may append arrays the file never
had (e.g. the Normal named-style bootstrap). Emitting model indices into a
spliced sheet therefore corrupts styles on any non-openpyxl producer —
measured on the LibreOffice-authored fixtures.

The translator parses the ORIGINAL styles.xml through upstream's own
``Stylesheet`` machinery, giving arrays that are positionally aligned with
the file's cellXfs but expressed in model semantics — directly comparable to
the model's arrays. Cells whose style exists in the file reuse its index;
anything else becomes an appended xf (never a rewrite of an existing entry),
with custom number formats re-translated into the FILE's numFmt numbering.
"""

from openpyxl.styles.numbers import (
    BUILTIN_FORMATS_MAX_SIZE,
    NumberFormat,
)
from openpyxl.xml.functions import fromstring, tostring


class StyleTranslator:

    def __init__(self, wb, original_styles_bytes):
        from openpyxl.styles.stylesheet import Stylesheet

        self._wb = wb
        sheet = Stylesheet.from_tree(fromstring(original_styles_bytes))
        file_arrays = list(sheet.cell_styles)
        self.file_xf_count = len(file_arrays)
        self._map = {}
        for idx, arr in enumerate(file_arrays):
            self._map.setdefault(tuple(arr), idx)   # first wins on duplicates
        # the file's custom number formats, in FILE numbering
        self._file_code_to_id = {}
        highest = BUILTIN_FORMATS_MAX_SIZE - 1
        for fmt_id, code in sheet.custom_formats.items():
            self._file_code_to_id.setdefault(code, fmt_id)
            highest = max(highest, fmt_id)
        self._next_numfmt_id = highest + 1
        self._new_xfs = []        # model StyleArrays, in allocation order
        self._new_numfmts = []    # (file_id, code)

    # -- resolution -----------------------------------------------------

    def resolve(self, style_array):
        """The FILE xf index for a model StyleArray (allocating an appended
        xf when the file has no equivalent)."""
        if style_array is None:
            return None
        key = tuple(style_array)
        idx = self._map.get(key)
        if idx is not None:
            return idx
        idx = self.file_xf_count + len(self._new_xfs)
        self._map[key] = idx
        self._new_xfs.append(style_array)
        return idx

    def resolver(self):
        """A per-cell callable for the splice writer."""
        def _resolve(cell):
            if cell._style is None:
                return None
            return self.resolve(cell._style)
        return _resolve

    # -- rendering for the styles planner ---------------------------------

    def _file_numfmt_id(self, model_numfmt_id):
        """Translate a model numFmtId into the FILE's numbering."""
        if model_numfmt_id < BUILTIN_FORMATS_MAX_SIZE:
            return model_numfmt_id                  # builtin: universal
        code = self._wb._number_formats[
            model_numfmt_id - BUILTIN_FORMATS_MAX_SIZE]
        file_id = self._file_code_to_id.get(code)
        if file_id is None:
            file_id = self._next_numfmt_id
            self._next_numfmt_id += 1
            self._file_code_to_id[code] = file_id
            self._new_numfmts.append((file_id, code))
        return file_id

    def render_new_xfs(self):
        """Serialized <xf> elements for the appended entries (mirrors
        styles/stylesheet.py write_stylesheet), numFmtIds in FILE numbering.
        Call after every resolve() — i.e. after all sheets are planned."""
        from openpyxl.styles.cell_style import CellStyle

        rendered = []
        for style in self._new_xfs:
            xf = CellStyle.from_array(style)
            xf.numFmtId = self._file_numfmt_id(style.numFmtId)
            if style.alignmentId:
                xf.alignment = self._wb._alignments[style.alignmentId]
            if style.protectionId:
                xf.protection = self._wb._protections[style.protectionId]
            rendered.append(tostring(xf.to_tree()))
        return rendered

    def render_new_numfmts(self):
        """Serialized <numFmt> elements allocated during xf rendering."""
        return [
            tostring(NumberFormat(fmt_id, code).to_tree(tagname="numFmt"))
            for fmt_id, code in self._new_numfmts
        ]

    def model_to_file_table(self):
        """{model index: file index} for every current model array — used to
        rewrite the s attributes of freshly generated (added-sheet) parts."""
        table = {}
        for model_idx, arr in enumerate(self._wb._cell_styles):
            table[model_idx] = self.resolve(arr)
        return table
