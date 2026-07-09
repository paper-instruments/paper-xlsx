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


class PartPlan:
    """Adds/removals of whole parts for one save, with their content-type
    and relationship consequences planned in lockstep."""

    def __init__(self, existing_names):
        self.existing = set(existing_names)
        self.added = {}          # name -> payload
        self.dropped = set()     # names omitted from the copy loop
        self.ct_overrides = []   # (part_name, content_type)
        self.ct_removals = []    # part_names
        self.rel_appends = {}    # rels_part -> [(rid, type, target, mode)]
        self.rel_removals = {}   # rels_part -> [target suffix]

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
        return payload

    def apply_rels(self, rels_part, payload):
        """The updated payload for one rels part (``payload`` is None when
        the part does not exist yet — a fresh rels document is built)."""
        removals = self.rel_removals.get(rels_part, ())
        if removals and payload is None:
            raise TargetNotFoundError(
                "internal: relationship removals planned against a rels "
                "part that does not exist ({0!r}).".format(rels_part))
        for suffix in removals:
            payload = crosspart.rels_remove_by_target_suffix(payload, suffix)
        appends = self.rel_appends.get(rels_part, ())
        if appends:
            if payload is None:
                payload = crosspart.render_rels_document([])
            next_rid = crosspart.rels_next_rid(payload)
            resolved = []
            for i, (rid, rel_type, target, mode) in enumerate(appends):
                resolved.append((rid or "rId{0}".format(next_rid + i),
                                 rel_type, target, mode))
            payload = crosspart.rels_append(payload, resolved)
        return payload

    def touched_rels_parts(self):
        return set(self.rel_appends) | set(self.rel_removals)


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
    if name in _MODEL_MANAGED or name.startswith("xl/worksheets/"):
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
