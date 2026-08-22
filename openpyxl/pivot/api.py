# paper-xlsx: targeted PivotTable collection and identity-bound handles

"""Preserve-mode inspection of loaded PivotTables.

``Worksheet.pivots`` is a semantic overlay over the relationship-resolved
package graph. It does not mutate ``Worksheet._pivots``, deserialize foreign
parts through inherited serializers, or enumerate source records. Public
mutation methods are withheld until later stacked PRs.
"""

from __future__ import annotations

from openpyxl.errors import (
    AmbiguousTargetError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from openpyxl.pivot.graph import load_workbook_pivot_graph
from openpyxl.pivot.inspect import project_pivot
from openpyxl.pivot.qualify import qualify_pivot


TO_DICT_SCHEMA = "pivot_table"
TO_DICT_VERSION = 1
_SESSION_ATTR = "_paper_pivot_session"


class _PivotSession:
    __slots__ = ("generation", "graph", "states", "closed")

    def __init__(self, generation, graph, states, closed=False):
        self.generation = generation
        self.graph = graph
        self.states = states
        self.closed = closed


class _PivotState:
    __slots__ = (
        "identity", "sheet_title", "name", "projection", "qualification",
    )

    def __init__(self, identity, sheet_title, name, projection, qualification):
        self.identity = identity
        self.sheet_title = sheet_title
        self.name = name
        self.projection = projection
        self.qualification = qualification


def require_pivot_inspection(worksheet, api="Worksheet.pivots"):
    """Refuse unless ``worksheet`` belongs to a materialized preserve load."""
    workbook = getattr(worksheet, "parent", None)
    if workbook is None:
        raise UnsupportedStructureError(
            "%s requires a materialized preserve-mode workbook. "
            "Nothing was changed." % api,
            kind="invalid-pivot-graph",
        )
    if getattr(workbook, "write_only", False) \
            or getattr(workbook, "read_only", False):
        raise UnsupportedStructureError(
            "%s needs a materialized preserve-mode workbook; read-only "
            "and write-only loads do not retain the source package graph. "
            "Nothing was changed." % api,
            kind="invalid-pivot-graph",
        )
    if not getattr(workbook, "_preserve", False) \
            or getattr(workbook, "_paper_source", None) is None \
            or getattr(workbook, "_paper_ledger", None) is None:
        raise UnsupportedStructureError(
            "%s requires a workbook loaded with preserve=True so the "
            "source package graph remains authoritative. Nothing was "
            "changed." % api,
            kind="invalid-pivot-graph",
        )
    if getattr(workbook, "_paper_pivot_closed", False):
        raise UnsupportedStructureError(
            "%s cannot inspect a closed workbook. Nothing was changed."
            % api,
            kind="invalid-pivot-graph",
        )


def pivot_collection_for(worksheet):
    require_pivot_inspection(worksheet)
    session = _session_for(worksheet.parent)
    return PivotTableCollection(worksheet, session)


def _session_for(workbook):
    session = getattr(workbook, _SESSION_ATTR, None)
    if session is not None and not session.closed:
        return session
    graph = load_workbook_pivot_graph(workbook)
    states = {}
    source = workbook._paper_source
    for node in graph.pivots:
        cache = None
        if node.cache_definition_part:
            cache = graph.caches_by_part.get(node.cache_definition_part)
        if cache is None and node.cache_id:
            cache = graph.caches_by_id.get(node.cache_id)
        projection = project_pivot(
            node, cache, source=source, workbook=workbook)
        qualification = qualify_pivot(
            node, cache, projection, graph, workbook=workbook)
        states[node.identity] = _PivotState(
            identity=node.identity,
            sheet_title=node.sheet_title,
            name=node.identity.name,
            projection=projection,
            qualification=qualification,
        )
    session = _PivotSession(1, graph, states)
    setattr(workbook, _SESSION_ATTR, session)
    return session


def invalidate_pivot_overlay(workbook):
    """Mark cached inspection state stale. Used after overlay changes."""
    session = getattr(workbook, _SESSION_ATTR, None)
    if session is None:
        return
    session.generation += 1
    session.closed = True


def close_pivot_overlay(workbook):
    workbook._paper_pivot_closed = True
    invalidate_pivot_overlay(workbook)


class PivotTableCollection:
    """Worksheet-scoped, read-only collection of identity-bound handles."""

    def __init__(self, worksheet, session):
        self._worksheet = worksheet
        self._session = session
        self._handles = tuple(
            PivotTable(worksheet, state.identity, session.generation)
            for state in session.states.values()
            if state.sheet_title == worksheet.title
        )

    def __iter__(self):
        return iter(self._handles)

    def __len__(self):
        return len(self._handles)

    def __bool__(self):
        return bool(self._handles)

    def __getitem__(self, name):
        if not isinstance(name, str):
            raise TypeError("pivot names must be strings")
        if not name:
            raise TargetNotFoundError(
                "pivot name is empty",
                kind="pivot-not-found",
                options=[handle.name for handle in self._handles if handle.name],
            )
        exact = [handle for handle in self._handles if handle.name == name]
        if len(exact) == 1:
            return exact[0]
        if len(exact) > 1:
            raise AmbiguousTargetError(
                "pivot name %r matches more than one pivot on %r"
                % (name, self._worksheet.title),
                kind="ambiguous-pivot",
                options=[handle.name for handle in exact],
            )
        folded = name.casefold()
        casefold_matches = [
            handle for handle in self._handles
            if handle.name.casefold() == folded
        ]
        if len(casefold_matches) == 1:
            return casefold_matches[0]
        if len(casefold_matches) > 1:
            raise AmbiguousTargetError(
                "pivot name %r is ambiguous on %r"
                % (name, self._worksheet.title),
                kind="ambiguous-pivot",
                options=[handle.name for handle in casefold_matches],
            )
        raise TargetNotFoundError(
            "no pivot named %r on %r" % (name, self._worksheet.title),
            kind="pivot-not-found",
            options=[handle.name for handle in self._handles if handle.name],
        )


class PivotTable:
    """Identity-bound read-only handle for one loaded pivot."""

    def __init__(self, worksheet, identity, generation):
        self._worksheet = worksheet
        self._identity = identity
        self._generation = generation

    def _state(self):
        worksheet = self._worksheet
        workbook = getattr(worksheet, "parent", None)
        if workbook is None or getattr(workbook, "_paper_pivot_closed", False):
            raise TargetNotFoundError(
                "pivot handle is stale after workbook close",
                kind="stale-pivot-handle",
                anchor=self._identity.pivot_part,
            )
        session = getattr(workbook, _SESSION_ATTR, None)
        if session is None or session.closed \
                or session.generation != self._generation:
            raise TargetNotFoundError(
                "pivot handle is stale after an overlay change",
                kind="stale-pivot-handle",
                anchor=self._identity.pivot_part,
            )
        try:
            return session.states[self._identity]
        except KeyError:
            raise TargetNotFoundError(
                "pivot handle is stale after an overlay change",
                kind="stale-pivot-handle",
                anchor=self._identity.pivot_part,
            )

    @property
    def name(self):
        return self._state().name

    @property
    def worksheet(self):
        self._state()
        return self._worksheet

    @property
    def source(self):
        return self._state().projection.source

    @property
    def destination(self):
        return self._state().projection.destination

    @property
    def output_range(self):
        return self._state().projection.output_range

    @property
    def spec(self):
        return self._state().projection.spec

    @property
    def origin(self):
        return self._state().qualification.origin

    @property
    def valid(self):
        return self._state().qualification.valid

    @property
    def capabilities(self):
        return self._state().qualification.capabilities

    @property
    def qualification_reasons(self):
        return self._state().qualification.reasons

    @property
    def refresh_on_open_scope(self):
        return self._state().qualification.refresh_on_open_scope

    def to_dict(self):
        state = self._state()
        payload = {
            "schema": TO_DICT_SCHEMA,
            "version": TO_DICT_VERSION,
            "name": state.name,
            "sheet": state.sheet_title,
        }
        payload.update(state.projection.to_dict_fields())
        payload["origin"] = state.qualification.origin
        payload["valid"] = state.qualification.valid
        payload["capabilities"] = state.qualification.capabilities.to_dict()
        payload["refresh_on_open_scope"] = list(
            state.qualification.refresh_on_open_scope)
        payload["qualification_reasons"] = [
            item.to_dict() for item in state.qualification.reasons
        ]
        return payload
