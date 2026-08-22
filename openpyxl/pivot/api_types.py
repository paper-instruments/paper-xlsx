# paper-xlsx: immutable public PivotTable value types

"""Caller-facing PivotTable vocabulary.

These types encode semantic intent. They do not expose OOXML field indexes,
relationship IDs, or serializer objects. Instances are immutable and have
deterministic ``to_dict()`` output.
"""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl.utils.cell import range_boundaries


SUPPORTED_AGGREGATES = (
    "sum",
    "count",
    "count_numbers",
    "average",
    "min",
    "max",
)
SUPPORTED_LAYOUTS = ("compact", "outline", "tabular")
SUPPORTED_VALUES_AXES = ("columns", "rows")
_AGGREGATE_SET = frozenset(SUPPORTED_AGGREGATES)
_LAYOUT_SET = frozenset(SUPPORTED_LAYOUTS)
_VALUES_AXIS_SET = frozenset(SUPPORTED_VALUES_AXES)


def _require_nonempty_string(value, label):
    if not isinstance(value, str) or not value:
        raise ValueError("%s must be a nonempty string" % label)
    return value


def _freeze_sequence(value, label):
    if value is None:
        return None
    try:
        items = tuple(value)
    except TypeError:
        raise TypeError("%s must be a sequence" % label)
    return items


def _axis_item_identity(value):
    """Return the equality key used by the typed pivot source model."""
    if value is None or value == "":
        return ("blank",)
    if isinstance(value, bool):
        return ("boolean", value)
    return ("value", value)


@dataclass(frozen=True)
class PivotSource:
    """A same-workbook table or rectangular range source."""

    kind: str
    name: str | None = None
    sheet: str | None = None
    ref: str | None = None

    def __post_init__(self):
        if self.kind not in ("table", "range", "defined-name"):
            raise ValueError(
                "PivotSource.kind must be 'table', 'range', or "
                "'defined-name'")
        if self.kind == "table":
            _require_nonempty_string(self.name, "PivotSource.table name")
        elif self.kind == "range":
            _require_nonempty_string(self.sheet, "PivotSource.range sheet")
            _require_nonempty_string(self.ref, "PivotSource.range ref")
            try:
                range_boundaries(self.ref)
            except (TypeError, ValueError) as exc:
                raise ValueError(
                    "PivotSource.range ref must be a valid A1 range"
                ) from exc
        else:
            _require_nonempty_string(self.name, "PivotSource defined-name")

    @classmethod
    def table(cls, name):
        return cls(kind="table", name=name)

    @classmethod
    def range(cls, sheet, ref):
        return cls(kind="range", sheet=sheet, ref=ref)

    @classmethod
    def parse(cls, value):
        """Parse unambiguous string shorthand or return an existing source."""
        if isinstance(value, cls):
            return value
        if not isinstance(value, str):
            raise TypeError("PivotSource must be a PivotSource or string")
        text = value.strip()
        if not text:
            raise ValueError("PivotSource string shorthand is empty")
        if "!" in text:
            return _parse_sheet_range(text)
        if any(char in text for char in " :$"):
            raise ValueError(
                "range shorthand must be sheet-qualified, for example "
                "\"'Raw Data'!A1:H5000\"")
        return cls.table(text)

    def to_dict(self):
        if self.kind == "table":
            return {"kind": "table", "name": self.name}
        if self.kind == "range":
            return {"kind": "range", "sheet": self.sheet, "ref": self.ref}
        payload = {"kind": "defined-name", "name": self.name}
        if self.sheet is not None:
            payload["sheet"] = self.sheet
        return payload


def _parse_sheet_range(text):
    from openpyxl.utils.cell import SHEETRANGE_RE

    match = SHEETRANGE_RE.match(text)
    if match is None:
        raise ValueError(
            "range shorthand must be an explicit sheet-qualified A1 range")
    sheet = match.group("quoted")
    if sheet is None:
        sheet = match.group("notquoted")
    else:
        sheet = sheet.replace("''", "'")
    cells = match.group("cells")
    if not sheet or not cells:
        raise ValueError(
            "range shorthand must be an explicit sheet-qualified A1 range")
    return PivotSource.range(sheet, cells)


@dataclass(frozen=True)
class PivotAxisField:
    field: str
    items: tuple | None = None

    def __post_init__(self):
        _require_nonempty_string(self.field, "PivotAxisField.field")
        object.__setattr__(self, "items", _freeze_sequence(self.items, "items"))
        if self.items is not None:
            seen = set()
            for item in self.items:
                identity = _axis_item_identity(item)
                if identity in seen:
                    raise ValueError(
                        "PivotAxisField.items must not contain duplicates")
                seen.add(identity)

    def to_dict(self):
        payload = {"field": self.field}
        if self.items is not None:
            payload["items"] = list(self.items)
        return payload


@dataclass(frozen=True)
class PivotItemFilter:
    field: str
    include: tuple | None = None
    exclude: tuple | None = None

    def __post_init__(self):
        _require_nonempty_string(self.field, "PivotItemFilter.field")
        include = _freeze_sequence(self.include, "include")
        exclude = _freeze_sequence(self.exclude, "exclude")
        if include is not None and exclude is not None:
            raise ValueError(
                "PivotItemFilter accepts include or exclude, not both")
        if include is not None and not include:
            raise ValueError("PivotItemFilter.include must be nonempty")
        object.__setattr__(self, "include", include)
        object.__setattr__(self, "exclude", exclude)

    def to_dict(self):
        payload = {"field": self.field}
        if self.include is not None:
            payload["include"] = list(self.include)
        if self.exclude is not None:
            payload["exclude"] = list(self.exclude)
        return payload


@dataclass(frozen=True)
class PivotMeasure:
    field: str
    aggregate: str = "sum"
    caption: str | None = None
    number_format: str | None = None

    def __post_init__(self):
        _require_nonempty_string(self.field, "PivotMeasure.field")
        if self.aggregate not in _AGGREGATE_SET:
            raise ValueError(
                "PivotMeasure.aggregate must be one of %s"
                % ", ".join(SUPPORTED_AGGREGATES))
        if self.caption is not None:
            _require_nonempty_string(self.caption, "PivotMeasure.caption")
        if self.number_format is not None:
            _require_nonempty_string(
                self.number_format, "PivotMeasure.number_format")

    def to_dict(self):
        payload = {"field": self.field, "aggregate": self.aggregate}
        if self.caption is not None:
            payload["caption"] = self.caption
        if self.number_format is not None:
            payload["number_format"] = self.number_format
        return payload


@dataclass(frozen=True)
class PivotCapabilities:
    can_refresh_on_open: bool = False
    can_headless_refresh: bool = False
    can_rebuild_cache: bool = False
    can_edit_layout: bool = False
    can_repoint_source: bool = False
    can_move: bool = False
    can_rename: bool = False
    can_delete: bool = False

    def to_dict(self):
        return {
            "can_refresh_on_open": self.can_refresh_on_open,
            "can_headless_refresh": self.can_headless_refresh,
            "can_rebuild_cache": self.can_rebuild_cache,
            "can_edit_layout": self.can_edit_layout,
            "can_repoint_source": self.can_repoint_source,
            "can_move": self.can_move,
            "can_rename": self.can_rename,
            "can_delete": self.can_delete,
        }


@dataclass(frozen=True)
class PivotSpec:
    """Complete resolved representation of a supported pivot."""

    name: str
    source: PivotSource
    destination: str
    rows: tuple = ()
    columns: tuple = ()
    filters: tuple = ()
    values: tuple = ()
    layout: str = "tabular"
    values_axis: str = "columns"
    row_grand_totals: bool = True
    column_grand_totals: bool = True
    subtotals: bool = False
    style: str | None = None

    def __post_init__(self):
        _require_nonempty_string(self.name, "PivotSpec.name")
        if not isinstance(self.source, PivotSource):
            raise TypeError("PivotSpec.source must be a PivotSource")
        _require_nonempty_string(self.destination, "PivotSpec.destination")
        object.__setattr__(self, "rows", tuple(self.rows))
        object.__setattr__(self, "columns", tuple(self.columns))
        object.__setattr__(self, "filters", tuple(self.filters))
        object.__setattr__(self, "values", tuple(self.values))
        if self.layout not in _LAYOUT_SET:
            raise ValueError(
                "PivotSpec.layout must be one of %s"
                % ", ".join(SUPPORTED_LAYOUTS))
        if self.values_axis not in _VALUES_AXIS_SET:
            raise ValueError(
                "PivotSpec.values_axis must be 'columns' or 'rows'")
        if self.style is not None:
            _require_nonempty_string(self.style, "PivotSpec.style")
        for item in self.rows + self.columns:
            if not isinstance(item, PivotAxisField):
                raise TypeError("axis fields must be PivotAxisField values")
        for item in self.filters:
            if not isinstance(item, PivotItemFilter):
                raise TypeError("filters must be PivotItemFilter values")
        if not self.values:
            raise ValueError("PivotSpec requires at least one measure")
        for item in self.values:
            if not isinstance(item, PivotMeasure):
                raise TypeError("values must be PivotMeasure values")

    def to_dict(self):
        payload = {
            "name": self.name,
            "source": self.source.to_dict(),
            "destination": self.destination,
            "rows": [item.field for item in self.rows],
            "columns": [item.field for item in self.columns],
            "filters": [item.to_dict() for item in self.filters],
            "values": [item.to_dict() for item in self.values],
            "layout": self.layout,
            "values_axis": self.values_axis,
            "row_grand_totals": self.row_grand_totals,
            "column_grand_totals": self.column_grand_totals,
            "subtotals": self.subtotals,
        }
        if self.style is not None:
            payload["style"] = self.style
        return payload


ADOPTION_TO_DICT_SCHEMA = "pivot_adoption_qualification"
ADOPTION_TO_DICT_VERSION = 1
ADOPTION_STRATEGIES = ("dedicated-replacement", "shared-isolation")


@dataclass(frozen=True)
class PivotAdoptionQualification:
    """Read-only eligibility of one foreign pivot for explicit adoption.

    This type does not grant mutation. ``eligible=True`` means no known
    structural blocker remains; it is not a promise that operation-time
    calculation or later ``adopt()`` will succeed. The schema is independent
    of ``PivotTable.to_dict()``.
    """

    eligible: bool
    strategy: str | None = None
    requires_calculation: bool = False
    calculation_engine: str | None = None
    operation_constraints: tuple = ()
    reasons: tuple = ()

    def __post_init__(self):
        if self.strategy is not None and self.strategy not in ADOPTION_STRATEGIES:
            raise ValueError(
                "PivotAdoptionQualification.strategy must be "
                "'dedicated-replacement', 'shared-isolation', or None")
        if self.calculation_engine is not None \
                and self.calculation_engine != "libreoffice":
            raise ValueError(
                "PivotAdoptionQualification.calculation_engine must be "
                "'libreoffice' or None")
        object.__setattr__(
            self, "operation_constraints", tuple(self.operation_constraints))
        object.__setattr__(self, "reasons", tuple(self.reasons))

    def to_dict(self):
        return {
            "schema": ADOPTION_TO_DICT_SCHEMA,
            "version": ADOPTION_TO_DICT_VERSION,
            "eligible": self.eligible,
            "strategy": self.strategy,
            "requires_calculation": self.requires_calculation,
            "calculation_engine": self.calculation_engine,
            "operation_constraints": [
                item.to_dict() for item in self.operation_constraints
            ],
            "reasons": [item.to_dict() for item in self.reasons],
        }
