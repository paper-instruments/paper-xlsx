# paper-xlsx: the part-lifecycle engine (PLAN-v0.1 Batch 2; PR-1 §1.1)

"""One primitive, both directions: every part a save creates or deletes
routes through a :class:`PartPlan` — part payload + content-type override
+ relationship, planned together, applied by the saver's build loop. No
bespoke cascades.

The engine is PLANNING-level: nothing here touches bytes until the saver
applies the plan, so refusals raised during planning stay atomic.
"""

from openpyxl.errors import (
    RelationshipPolicyError,
    TargetNotFoundError,
)

from . import crosspart

# parts the model actively manages: replacing them raw would desync the
# model from the file (PR-1 §1.4)
_MODEL_MANAGED = (
    "[Content_Types].xml",
    "xl/workbook.xml",
    "xl/styles.xml",
    "xl/sharedStrings.xml",
)
_MANAGED_PREFIXES = ("xl/worksheets/", "xl/tables/")


class PartPlan:
    """Adds/removals of whole parts for one save, with their content-type
    and relationship consequences planned in lockstep."""

    def __init__(self, existing_names):
        self.existing = set(existing_names)
        self.added = {}          # name -> payload
        self.dropped = set()     # names omitted from the copy loop
        self.ct_overrides = []   # (part_name, content_type)
        self.ct_defaults = []    # (extension, content_type)
        self.ct_removals = []    # part_names
        self.rel_appends = {}    # rels_part -> [(rid, type, target, mode)]
        self.rel_removals = {}   # rels_part -> [target suffix]
        self._rid_base = {}      # rels_part -> first reserved number
        self._rid_reserved = {}  # rels_part -> count reserved

    def reserve_rid(self, rels_part, existing_payload):
        """Sequential rId allocation shared by every planner touching one
        rels part (two independent next_rid computations collide)."""
        base = self._rid_base.get(rels_part)
        if base is None:
            base = crosspart.rels_next_rid(existing_payload) \
                if existing_payload else 1
            self._rid_base[rels_part] = base
        n = self._rid_reserved.get(rels_part, 0)
        self._rid_reserved[rels_part] = n + 1
        return "rId{0}".format(base + n)

    def add_default(self, extension, content_type):
        self.ct_defaults.append((extension, content_type))

    # -- the two verbs ---------------------------------------------------

    def add_part(self, name, payload, content_type=None,
                 relate_from=None, rel_type=None, rel_id=None):
        """Plan a new part; returns the allocated relationship id (or
        ``rel_id`` when given). ``relate_from`` names the part whose rels
        must point at the new part (its rels part is created if absent —
        the saver resolves ids at apply time via ``resolve_rel_ids``)."""
        if name in self.existing or name in self.added:
            raise RelationshipPolicyError(
                "part {0!r} already exists in the package; the lifecycle "
                "engine never overwrites parts (use wb.replace_part for "
                "byte swaps). Nothing was written.".format(name))
        self.added[name] = payload
        if content_type is not None:
            # crosspart.ct_append_overrides prefixes the "/" itself
            self.ct_overrides.append((name, content_type))
        if relate_from is not None:
            rels_part = _rels_path(relate_from)
            entry = (rel_id, rel_type, _relative_target(relate_from, name),
                     None)
            self.rel_appends.setdefault(rels_part, []).append(entry)
        return rel_id

    def remove_part(self, name, referencing_rels=()):
        """Plan a part's removal: dropped from the copy loop, its
        content-type override removed, and the named relationships cut.
        ``referencing_rels`` is [(rels_part, target_suffix)]."""
        if name not in self.existing:
            raise TargetNotFoundError(
                "part {0!r} does not exist in the package; nothing to "
                "remove.".format(name))
        self.dropped.add(name)
        # the part's own rels part (if any) goes with it
        own_rels = _rels_path(name)
        if own_rels in self.existing:
            self.dropped.add(own_rels)
        self.ct_removals.append(name)
        for rels_part, suffix in referencing_rels:
            self.rel_removals.setdefault(rels_part, []).append(suffix)

    def __bool__(self):
        return bool(self.added or self.dropped or self.ct_overrides
                    or self.ct_removals or self.rel_appends
                    or self.rel_removals)

    # -- application (called by the saver) --------------------------------

    def apply_content_types(self, payload):
        for name in self.ct_removals:
            payload = crosspart.ct_remove_override(payload, name)
        if self.ct_overrides:
            payload = crosspart.ct_append_overrides(payload,
                                                    self.ct_overrides)
        for ext, ctype in self.ct_defaults:
            existing = _default_content_type(payload, ext)
            if existing is None:
                payload = crosspart.ct_append_defaults(
                    payload, [(ext, ctype)])
            elif existing != ctype:
                raise RelationshipPolicyError(
                    "the package already types extension {0!r} as {1!r}; "
                    "the part being created needs {2!r} and re-typing a "
                    "preserved Default is not supported. Nothing was "
                    "written.".format(ext, existing, ctype))
        return payload

    def apply_rels(self, rels_part, payload):
        """The updated payload for one rels part (``payload`` is None when
        the part does not exist yet — a fresh rels document is built)."""
        removals = self.rel_removals.get(rels_part, ())
        if removals and payload is None:
            raise TargetNotFoundError(
                "internal: relationship removals planned against a rels "
                "part that does not exist ({0!r}).".format(rels_part))
        if removals:
            payload = _rels_remove_exact(rels_part, payload, removals)
        appends = self.rel_appends.get(rels_part, ())
        if appends:
            if payload is None:
                payload = crosspart.render_rels_document([])
            base = self._rid_base.get(rels_part)
            reserved = self._rid_reserved.get(rels_part, 0)
            next_rid = crosspart.rels_next_rid(payload)
            if base is not None:
                next_rid = max(next_rid, base + reserved)
            auto = 0
            resolved = []
            for (rid, rel_type, target, mode) in appends:
                if rid is None:
                    rid = "rId{0}".format(next_rid + auto)
                    auto += 1
                resolved.append((rid, rel_type, target, mode))
            payload = crosspart.rels_append(payload, resolved)
        return payload

    def touched_rels_parts(self):
        return set(self.rel_appends) | set(self.rel_removals)


def _default_content_type(ct_payload, extension):
    """The ContentType an existing <Default> gives ``extension`` (matched
    case-insensitively, either quote style — the first-cut substring check
    missed both, Batch-2 gate), else None."""
    root = crosspart.scan_small(ct_payload, "Types", max_depth=1)
    for child in root.children:
        if child.local() != "Default":
            continue
        if child.attrs.get("Extension", "").lower() == extension.lower():
            return child.attrs.get("ContentType")
    return None


def _rels_remove_exact(rels_part, payload, part_names):
    """Remove relationships whose RESOLVED target equals one of
    ``part_names`` — never suffix matching (a sibling named mytable1.xml
    must survive table1.xml's removal; Batch-2 gate)."""
    owner = _owner_of_rels(rels_part)
    targets = set(part_names)
    root = crosspart.scan_small(payload, "Relationships", max_depth=1)
    edits = []
    for child in root.children:
        if child.local() != "Relationship":
            continue
        resolved = _resolve_target(owner, child.attrs.get("Target", ""))
        if resolved in targets:
            edits.append((child.start, child.end, b""))
    if not edits:
        return payload
    return crosspart.apply_edits(payload, edits)


def _owner_of_rels(rels_part):
    """xl/worksheets/_rels/sheet1.xml.rels -> xl/worksheets/sheet1.xml;
    the package rels _rels/.rels -> "" (targets resolve from the root)."""
    if rels_part.startswith("_rels/"):
        base = rels_part[len("_rels/"):]
        return base[:-5] if base.endswith(".rels") else base
    folder, _, base = rels_part.rpartition("/_rels/")
    name = base[:-5] if base.endswith(".rels") else base
    return "{0}/{1}".format(folder, name) if folder else name


def _resolve_target(from_part, target):
    if target.startswith("/"):
        return target[1:]
    base = from_part.rpartition("/")[0].split("/") if "/" in from_part         else []
    for piece in target.split("/"):
        if piece == "..":
            base = base[:-1]
        elif piece != ".":
            base.append(piece)
    return "/".join(base)


def check_replace_part(wb, name):
    """Guards for Workbook.replace_part (PR-1 §1.4): the part must exist,
    and model-managed or sheet parts refuse — replacing them raw would
    desync the model."""
    import io
    import zipfile

    source = wb._paper_source
    with zipfile.ZipFile(io.BytesIO(source)) as z:
        names = set(z.namelist())
    if name not in names:
        raise TargetNotFoundError(
            "part {0!r} does not exist in the package. Nothing was "
            "changed.".format(name))
    if name in _MODEL_MANAGED or name.startswith(_MANAGED_PREFIXES) \
            or "_rels" in name.split("/"):
        raise RelationshipPolicyError(
            "part {0!r} is actively managed by the model; replacing its "
            "bytes raw would desync the model from the file. Edit it "
            "through the API instead. Nothing was changed.".format(name))


def _rels_path(part_name):
    folder, _, base = part_name.rpartition("/")
    return "{0}/_rels/{1}.rels".format(folder, base) if folder \
        else "_rels/{0}.rels".format(base)


def _relative_target(from_part, to_part):
    """OPC relationship targets are relative to the source part's folder."""
    from_dir = from_part.rpartition("/")[0]
    if from_dir and to_part.startswith(from_dir + "/"):
        return to_part[len(from_dir) + 1:]
    # walk up: ../ per level, then the full path from the common root
    from_parts = from_dir.split("/") if from_dir else []
    to_parts = to_part.split("/")
    common = 0
    while (common < len(from_parts) and common < len(to_parts) - 1
            and from_parts[common] == to_parts[common]):
        common += 1
    ups = "../" * (len(from_parts) - common)
    return ups + "/".join(to_parts[common:])
