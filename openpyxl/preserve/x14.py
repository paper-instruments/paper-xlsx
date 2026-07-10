# paper-xlsx: x14 twin-sync

"""Conditional-formatting edits on sheets whose rules carry x14 twins.

The model DROPS the <x14:id> twin pointer when re-rendering classic rules
(measured), so twin-bearing CF is never re-rendered here: the classic run
is COMPOSED from original bytes — surviving blocks verbatim, deleted
blocks omitted, sqref-only changes patched in place — and only genuinely
NEW blocks are model-rendered (new rules have no twins). The extLst twin
entries are patched in lockstep: xm:sqref text follows the classic sqref,
and deleted rules' twin entries are removed. Anything this composition
cannot express refuses, exactly as v0 did.
"""

import re

from openpyxl.errors import UnsupportedStructureError

_X14_ID_RE = re.compile(br"<x14:id>(\{[^}]+\})</x14:id>")
_SQREF_ATTR_RE = re.compile(br'(<conditionalFormatting[^>]*\ssqref=")([^"]*)(")')


def _refuse(msg):
    raise UnsupportedStructureError(msg + " Nothing was written.")


def sheet_has_cf_twins(scan, original):
    """True when classic rules carry twin pointers or the extLst carries
    x14 conditional formattings."""
    for span in scan.regions.get("conditionalFormatting", []):
        if b"extLst" in original[span.start:span.end]:
            return True
    for span in scan.regions.get("extLst", []):
        if b"conditionalFormattings" in original[span.start:span.end]:
            return True
    return False


def _rules_signature(rendered):
    """A block's identity minus its sqref (for sqref-only-change matching)."""
    return _SQREF_ATTR_RE.sub(rb"\1\3", rendered)


def _block_sqref(rendered):
    m = _SQREF_ATTR_RE.search(rendered)
    return m.group(2) if m else b""


def _render_block(wb, cf):
    """One classic CF block rendered for writing (mirrors
    regions.render_cf_for_write's dxf handling, per block)."""
    from openpyxl.styles.differential import DifferentialStyle
    from openpyxl.xml.functions import tostring

    empty = DifferentialStyle()
    for rule in cf.rules:
        if rule.dxf and rule.dxf != empty:
            rule.dxfId = wb._differential_styles.add(rule.dxf)
    return tostring(cf.to_tree())


def plan_cf_composed(wb, ws, scan, original, armed_blocks):
    """(classic_cf_replacement_bytes, extlst_replacement_or_None).

    ``armed_blocks``: the arm-time model renders (regions._render_cf
    tuple), positionally corresponding to the original document's classic
    CF elements."""
    spans = scan.regions.get("conditionalFormatting", [])
    if len(spans) != len(armed_blocks):
        _refuse("cannot sync conditional formatting on sheet {0!r}: the "
                "original document has {1} classic blocks but the model "
                "loaded {2} — the positional twin mapping is not "
                "trustworthy.".format(ws.title, len(spans),
                                      len(armed_blocks)))
    original_bytes = [original[s.start:s.end] for s in spans]
    armed_sigs = [_rules_signature(b) for b in armed_blocks]

    current = [_render_block(wb, cf) for cf in ws.conditional_formatting]
    current_sigs = [_rules_signature(b) for b in current]

    consumed = [False] * len(armed_blocks)
    pieces = [None] * len(armed_blocks)   # survivors, original order
    new_blocks = []
    sqref_patches = {}                    # guid -> new sqref bytes

    def _match(cur, cur_sig, position):
        # ORDER-PRESERVING: prefer the same position, then the earliest
        # unconsumed match (byte-identical blocks must never cross-steal
        # each other's twins)
        if position < len(armed_blocks) and not consumed[position] \
                and armed_blocks[position] == cur:
            return ("same", position)
        for j, armed in enumerate(armed_blocks):
            if not consumed[j] and armed == cur:
                return ("same", j)
        if position < len(armed_blocks) and not consumed[position] \
                and armed_sigs[position] == cur_sig:
            return ("sqref", position)
        for j, sig in enumerate(armed_sigs):
            if not consumed[j] and sig == cur_sig:
                return ("sqref", j)
        return None

    for pos, (cur, cur_sig) in enumerate(zip(current, current_sigs)):
        matched = _match(cur, cur_sig, pos)
        if matched is None:
            new_blocks.append(cur)
            continue
        kind, j = matched
        consumed[j] = True
        if kind == "same":
            pieces[j] = original_bytes[j]
        else:
            new_sqref = _block_sqref(cur)
            patched, n = _SQREF_ATTR_RE.subn(
                rb"\g<1>" + new_sqref.replace(b"\\", rb"\\") + rb"\g<3>",
                original_bytes[j], count=1)
            if n != 1:
                _refuse("cannot sync conditional formatting on sheet "
                        "{0!r}: the original block's sqref could not be "
                        "patched.".format(ws.title))
            pieces[j] = patched
            for guid in _X14_ID_RE.findall(original_bytes[j]):
                sqref_patches[guid] = new_sqref

    # an unconsumed twin-bearing block whose range a NEW block claims is a
    # MODIFICATION, not a deletion: re-rendering would lose the <x14:id>
    # pointer, so it refuses (modification silently deleted
    # the twin)
    new_sqrefs = {_block_sqref(nb) for nb in new_blocks}
    deleted_guids = set()
    for j, used in enumerate(consumed):
        if not used:
            guids = _X14_ID_RE.findall(original_bytes[j])
            if guids and _block_sqref(armed_blocks[j]) in new_sqrefs:
                _refuse("cannot sync conditional formatting on sheet "
                        "{0!r}: a rule carrying an x14 twin (range {1}) "
                        "was MODIFIED; re-rendering it would orphan the "
                        "twin. Delete and recreate the rule, or edit "
                        "without preserve=True.".format(
                            ws.title,
                            _block_sqref(armed_blocks[j]).decode("ascii",
                                                                 "replace")))
            deleted_guids.update(guids)
            # a deleted block with NO twin needs no extLst work; blocks
            # with twins get their entries removed below

    classic = b"".join(p for p in pieces if p is not None) \
        + b"".join(new_blocks)

    extlst_replacement = None
    if sqref_patches or deleted_guids:
        ext_spans = scan.regions.get("extLst", [])
        if len(ext_spans) != 1:
            _refuse("cannot sync x14 twins on sheet {0!r}: expected one "
                    "extLst element, found {1}.".format(
                        ws.title, len(ext_spans)))
        ext_original = original[ext_spans[0].start:ext_spans[0].end]
        extlst_replacement = _patch_twins(
            ws, ext_original, sqref_patches, deleted_guids)

    return classic, extlst_replacement


def _patch_twins(ws, ext_bytes, sqref_patches, deleted_guids):
    """Patch xm:sqref texts and delete whole twin entries by GUID inside
    the ORIGINAL extLst bytes."""
    entry_re = re.compile(
        br"<x14:conditionalFormatting\b.*?</x14:conditionalFormatting>",
        re.S)
    if (b"conditionalFormattings" in ext_bytes
            and b"<x14:conditionalFormatting" not in ext_bytes):
        _refuse("cannot sync x14 twins on sheet {0!r}: the extension uses "
                "an unexpected namespace prefix.".format(ws.title))

    out = []
    pos = 0
    for m in entry_re.finditer(ext_bytes):
        out.append(ext_bytes[pos:m.start()])
        entry = m.group(0)
        guids = set(_X14_ID_RE.findall(entry)) or set(
            g1 or g2 for g1, g2 in re.findall(
                br'\sid=(?:"(\{[^}]+\})"|\'(\{[^}]+\})\')', entry))
        if guids and guids <= deleted_guids:
            pos = m.end()
            continue                          # twin removed with its rule
        if guids & deleted_guids:
            _refuse("cannot sync x14 twins on sheet {0!r}: a twin entry "
                    "carries rules for BOTH deleted and surviving classic "
                    "blocks; per-rule twin surgery is not "
                    "supported.".format(ws.title))
        patch_guids = guids & set(sqref_patches)
        if patch_guids:
            new_sqref = sqref_patches[next(iter(patch_guids))]
            entry, n = re.subn(
                br"(<xm:sqref>)[^<]*(</xm:sqref>)",
                rb"\g<1>" + new_sqref.replace(b"\\", rb"\\") + rb"\g<2>",
                entry, count=1)
            if n != 1:
                _refuse("cannot sync x14 twins on sheet {0!r}: a twin "
                        "entry has no patchable xm:sqref.".format(ws.title))
        out.append(entry)
        pos = m.end()
    out.append(ext_bytes[pos:])
    return b"".join(out)


# ---------------------------------------------------------------------
# data validations: x14 DVs are independent validations (no id linkage);
# classic edits are safe while the two sqref sets stay disjoint

_XM_SQREF_RE = re.compile(br"<xm:sqref>([^<]*)</xm:sqref>")


def check_dv_coexistence(ws, scan, original):
    """Refuse only when a classic DV range intersects an x14 DV range —
    otherwise classic edits proceed and the x14 block stays verbatim."""
    from openpyxl.utils.cell import range_boundaries

    x14_refs = []
    for span in scan.regions.get("extLst", []):
        blob = original[span.start:span.end]
        dv_m = re.search(
            br"<(\w+):dataValidations\b.*?</\1:dataValidations>", blob,
            re.S)
        if dv_m is None:
            continue
        blob = dv_m.group(0)                 # only the DV ext's sqrefs
        for m in _XM_SQREF_RE.finditer(blob):
            for ref in m.group(1).decode("utf-8", "replace").split():
                try:
                    x14_refs.append(range_boundaries(ref.replace("$", "")))
                except Exception:
                    _refuse("cannot change data validations on sheet "
                            "{0!r}: an x14 validation range ({1!r}) could "
                            "not be parsed.".format(ws.title, ref))
    if not x14_refs:
        return
    for dv in ws.data_validations.dataValidation:
        for rng in dv.sqref.ranges:
            b = (rng.min_col, rng.min_row, rng.max_col, rng.max_row)
            for (xc1, xr1, xc2, xr2) in x14_refs:
                if not (b[2] < xc1 or b[0] > xc2
                        or b[3] < xr1 or b[1] > xr2):
                    _refuse("cannot change data validations on sheet "
                            "{0!r}: the classic range {1} overlaps an x14 "
                            "validation; editing would double-validate "
                            "those cells.".format(ws.title, rng))
