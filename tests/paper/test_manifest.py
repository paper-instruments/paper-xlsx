"""The fixture corpus is frozen: every fixture's SHA-256 must match
MANIFEST.sha256, every fixture must have a sidecar with the pinned schema,
and no unmanifested fixture may appear."""
from __future__ import annotations

import hashlib
import json
import os

from .conftest import FIXTURES_DIR

SIDECAR_REQUIRED_KEYS = {"fixture", "provenance", "features", "ground_truth",
                         "verified_by", "date"}


def _manifest_entries():
    manifest = os.path.join(FIXTURES_DIR, "MANIFEST.sha256")
    entries = {}
    with open(manifest) as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            digest, rel = line.split("  ", 1)
            entries[rel] = digest
    return entries


def _fixture_files():
    found = {}
    for root, dirs, files in os.walk(FIXTURES_DIR):
        if os.path.basename(root) == "generators":
            dirs[:] = []
            continue
        for name in files:
            if name.endswith((".json", ".md", ".sha256")):
                continue
            path = os.path.join(root, name)
            rel = os.path.relpath(path, FIXTURES_DIR).replace(os.sep, "/")
            found[rel] = path
    return found


def test_every_fixture_matches_manifest():
    entries = _manifest_entries()
    files = _fixture_files()
    assert set(files) == set(entries), (
        "fixture set drifted from MANIFEST.sha256: unmanifested={0}, missing={1}".format(
            sorted(set(files) - set(entries)), sorted(set(entries) - set(files)))
    )
    mismatched = []
    for rel, path in sorted(files.items()):
        with open(path, "rb") as f:
            digest = hashlib.sha256(f.read()).hexdigest()
        if digest != entries[rel]:
            mismatched.append(rel)
    assert not mismatched, (
        "fixture bytes changed (the corpus is FROZEN; regeneration is an explicit "
        "human decision): {0}".format(mismatched)
    )


def test_every_fixture_has_a_conforming_sidecar():
    problems = []
    for rel, path in sorted(_fixture_files().items()):
        sidecar = path + ".json"
        if not os.path.exists(sidecar):
            problems.append("{0}: sidecar missing".format(rel))
            continue
        with open(sidecar) as f:
            doc = json.load(f)
        missing = SIDECAR_REQUIRED_KEYS - set(doc)
        if missing:
            problems.append("{0}: sidecar missing keys {1}".format(rel, sorted(missing)))
        if doc.get("fixture") != os.path.basename(rel):
            problems.append("{0}: sidecar fixture name mismatch".format(rel))
        prov = doc.get("provenance", {})
        if not isinstance(prov, dict) or "app" not in prov or "notes" not in prov:
            problems.append("{0}: sidecar provenance malformed".format(rel))
        # provenance honesty: nothing in this corpus is Excel-authored
        if "excel" in str(prov.get("app", "")).lower():
            problems.append(
                "{0}: provenance claims Excel — this corpus cannot contain "
                "Excel-authored fixtures".format(rel))
    assert not problems, "\n".join(problems)
