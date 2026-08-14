# paper-xlsx: isolated image replacement

"""Replace one drawing relationship without rewriting the drawing."""

from openpyxl.errors import (
    AmbiguousTargetError,
    TargetNotFoundError,
    UnsupportedStructureError,
)

from . import crosspart
from .drawings import _IMAGE_MIME, _image_payload, _next_number
from .tables import _rels_path, _resolve_target


def _anchor_coordinate(image):
    anchor = getattr(image, "anchor", None)
    marker = getattr(anchor, "_from", None)
    if marker is None:
        return None
    from openpyxl.utils import get_column_letter

    return "{0}{1}".format(get_column_letter(marker.col + 1), marker.row + 1)


def request_replacement(ws, target, replacement, *, name=None):
    """Validate and record one relationship-owned image replacement."""
    ledger = getattr(ws.parent, "_paper_ledger", None)
    if ledger is None or not ledger.armed:
        raise ValueError("replace_image() is only available in preserve mode")
    images = list(getattr(ws, "_images", ()) or ())
    if target in images:
        matches = [target]
    elif isinstance(target, str):
        normalized = target.replace("$", "").upper()
        matches = [image for image in images
                   if _anchor_coordinate(image) == normalized]
    else:
        raise TypeError("target must be an anchor coordinate or loaded image")
    if name is not None:
        matches = [image for image in matches
                   if getattr(image, "_paper_name", None) == name]
    if not matches:
        raise TargetNotFoundError(
            "no loaded image matches {0!r}".format(target),
            kind="image-not-found", anchor=str(target))
    if len(matches) > 1:
        options = [getattr(image, "_paper_name", None) or
                   "image[{0}]".format(images.index(image))
                   for image in matches]
        raise AmbiguousTargetError(
            "multiple loaded images match anchor {0!r}; pass name=".format(
                target), kind="ambiguous-image", anchor=str(target),
            options=options)
    image = matches[0]
    index = images.index(image)
    if index not in ledger.object_snapshots.get(ws, {}).get("image", {}):
        raise UnsupportedStructureError(
            "replace_image() requires an image loaded from the source "
            "package", kind="image-not-loaded", anchor=str(target))
    old_part = getattr(image, "_paper_part", None)
    drawing_part = getattr(image, "_paper_drawing_part", None)
    rel_id = getattr(image, "_paper_rel_id", None)
    if not old_part or not drawing_part or not rel_id:
        raise UnsupportedStructureError(
            "the selected image relationship cannot be resolved safely",
            kind="unresolved-image-relationship", anchor=str(target))
    from openpyxl.drawing.image import Image

    candidate = replacement if isinstance(replacement, Image) \
        else Image(replacement)
    data, fmt = _image_payload(candidate)
    key = (ws, index)
    ledger.image_replacements[key] = {
        "old_part": old_part,
        "drawing_part": drawing_part,
        "rels_part": _rels_path(drawing_part),
        "rel_id": rel_id,
        "data": data,
        "format": fmt,
    }
    return image


def _patch_relationship(payload, request, new_part):
    root = crosspart.scan_small(payload, "Relationships", max_depth=1)
    matches = [child for child in root.children
               if child.local() == "Relationship"
               and child.attrs.get("Id") == request["rel_id"]]
    if len(matches) != 1:
        raise UnsupportedStructureError(
            "image relationship {0!r} is missing or ambiguous".format(
                request["rel_id"]), kind="unresolved-image-relationship")
    relationship = matches[0]
    if not relationship.attrs.get("Type", "").endswith("/image"):
        raise UnsupportedStructureError(
            "relationship {0!r} is not an image".format(request["rel_id"]),
            kind="unresolved-image-relationship")
    owner = request["drawing_part"]
    if _resolve_target(owner, relationship.attrs.get("Target", "")) != \
            request["old_part"]:
        raise UnsupportedStructureError(
            "the image relationship target changed since load",
            kind="image-relationship-drift")
    target = _relative_target(owner, new_part)
    start, end, head = crosspart._patch_attr(
        payload, relationship, "Target", target)
    return payload[:start] + head + payload[end:]


def _relative_target(owner, target):
    owner_dir = owner.rpartition("/")[0].split("/")
    target_parts = target.split("/")
    common = 0
    for left, right in zip(owner_dir, target_parts):
        if left != right:
            break
        common += 1
    return "/".join([".."] * (len(owner_dir) - common)
                    + target_parts[common:])


def plan_replacements(zin, requests, part_plan, names, plan):
    """Add fresh media parts and compose exact relationship retargets."""
    taken = set(names) | set(part_plan.added)
    next_media = _next_number(taken, r"xl/media/image(\d+)\.\w+$")
    effects = []
    for offset, (_key, request) in enumerate(sorted(
            requests.items(), key=lambda item: (item[0][0].title,
                                                item[0][1]))):
        fmt = request["format"]
        part = "xl/media/image{0}.{1}".format(next_media + offset, fmt)
        part_plan.add_part(part, request["data"])
        part_plan.add_default(fmt, _IMAGE_MIME[fmt])
        rels_part = request["rels_part"]
        base = plan.get(rels_part)
        if base is None:
            if rels_part not in names:
                raise UnsupportedStructureError(
                    "the selected image has no drawing relationships part",
                    kind="unresolved-image-relationship")
            base = zin.read(rels_part)
        plan[rels_part] = _patch_relationship(base, request, part)
        effects.append({
            "kind": "relationship_retargeted",
            "part": rels_part,
            "cause": "image_replaced",
            "target_part": part,
        })
    return effects
