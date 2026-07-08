# paper-xlsx Fork Ledger

Based on upstream tag `3.1.5`, forked 2026-07-07, marker tag `paper-base`.

Upstream source is the official Mercurial repository at
`https://foss.heptapod.net/openpyxl/openpyxl`. The GitHub repository for this
fork was bootstrapped by cloning that Mercurial repository and converting it to
Git with `hg-fast-export`, then checking out release tag `3.1.5` as `main`.
The upstream tag check showed newer branch commits after `3.1.5`, but no newer
stable release tag.

## Baseline Test Results

- Python 3.9.6 with upstream `requirements.txt` (`lxml==5.0.1`):
  `2592 passed, 6 skipped, 7 xfailed in 17.18s`.
- Python 3.13.3 with CI dependency constraint `lxml<6` (resolved to
  `lxml==5.4.0`): `2592 passed, 6 skipped, 7 xfailed in 3.94s`.
- Environment note: Python 3.13.3 with latest unconstrained `lxml==6.1.1`
  produced four pre-existing upstream failures in
  `openpyxl/xml/tests/test_functions.py::test_iterparse`; lxml now raises
  `TypeError` for the `BytesIO` input where the test expects `ValueError`.
  CI intentionally uses `lxml<6` until upstream handles that dependency change.

## Packaging Smoke Results

- Built with `python -m build`: `paper_xlsx-0.1.0.tar.gz` and
  `paper_xlsx-0.1.0-py2.py3-none-any.whl`.
- Wheel listing starts with `openpyxl/__init__.py` and
  `openpyxl/_constants.py`, confirming the import package was not renamed.
- Wheel smoke and sdist smoke both printed `0.1.0` from
  `openpyxl.__paper_version__`.

## Sanctioned Deviations From Upstream Behavior

None.

## Phase 0 — Orientation (2026-07-07)

- Baseline re-verified on the development machine (`.venv`, Python 3.13.3, lxml 5.4.0,
  pandas 3.0.3): `2592 passed, 6 skipped, 7 xfailed in 2.77s` — matches the fork-point baseline
  above. Raw log: `scratch/results/baseline_pytest.txt` (gitignored spike area).
- Provenance re-verified: full converted history (9,142 commits, 123 tags);
  `paper-base` == `3.1.5` == `c4986390b`; PyPI's latest openpyxl is still 3.1.5 as of
  2026-07-07, so the fork base is current upstream stable.
- Deliverables: `agent_docs/ARCHITECTURE-NOTES.md` (source tour),
  `agent_docs/OPEN-QUESTIONS.md` (ten open questions answered with evidence, cross-cutting
  gaps, and flags against pinned shapes for human decision), `FIXTURE-REQUESTS.md`
  (real-Excel fixtures a human must author).
- Performance seeds for the Phase-2 guardrail (large synthetic fixture, 3.39 MB / 600k cells):
  stock load 2.505 s, stock save 2.174 s, LibreOffice warm convert 2.09 s.
- Hygiene note: a leftover `soxhub` git remote points at `/tmp/soxhub-openpyxl` (the
  hg-conversion staging clone); candidate for removal, left untouched pending owner decision.

## Phase 1 — Test infrastructure (2026-07-08)

- Fixture corpus frozen: 18 fixtures under `tests/paper/fixtures/` with pinned-schema
  sidecars and `MANIFEST.sha256` (enforced by `tests/paper/test_manifest.py`). All
  provenance is openpyxl-authored / zip surgery / LibreOffice conversion — honestly
  labeled; the real-Excel bucket is requested in `FIXTURE-REQUESTS.md`.
- Five-job battery in `tests/paper/test_battery.py`: `TestStockCarnageBaseline`
  (passes today; regression-guards the damage model with the Phase-0-corrected claims)
  and `TestBatterySafety` (the forever criterion as strict xfails, each naming the
  phase that must flip it).
- Contract-harness helpers in `tests/paper/support/` (part-payload diff, semantic XML
  diff that never normalizes cell text, refusal-atomicity assertion, LibreOffice test
  driver with per-invocation profile isolation and temp-copy discipline).
- `pytest.ini`: registered the `lo_smoke` marker. CI: added a `test-libreoffice` job
  (ubuntu, Python 3.13, `libreoffice-calc`, `PAPER_REQUIRE_LO=1` promotes skips to
  failures).
- `setup.py`: `find_packages` exclude extended with `"tests", "tests.*"` so the new
  top-level test package cannot ship in the wheel.
- Full suite after Phase 1: 2617 passed, 6 skipped, 12 xfailed (2592 upstream tests
  unchanged and green).


## Phase 1.5 — PR-0 API proposal (2026-07-08)

- `agent_docs/PR0-API-PROPOSAL.md`: the v0 design contract. Freezes the delegated
  decisions (inline strings everywhere; per-operation-class collateral sets; no
  core.xml auto-stamp under preserve; performance budget 1.5x stock save, evidence
  0.16x composed prototype; frozen three-tier chokepoint inventory; splice guard
  set; shared-formula dissolve-on-touch; sheet delete/rename/reorder refuse in v0;
  mixed-chart semantics; rels append-only policy; calcChain deletion cascade).
- Sanctioned deviations register grows by one (recorded in PR-0 §10 and below):
  preserve-mode save does not auto-stamp `properties.modified` (stock path
  unchanged) — required by the pinned no-op payload-identity invariant.
- G6/G9 evidence spikes: `scratch/probes/pr0_composed_save.py` (preserve save
  prototype 0.381s vs stock 2.325s on 600k cells; untouched payloads verified
  byte-identical), `scratch/probes/pr0_g9_chokepoints.py` (chart mutation -> chart
  part; ws._rels.append discarded by stock save; code_name -> workbook.xml;
  template -> [Content_Types]).

## Release Safety

The repository is private. The release workflow targets the `pypi` environment
and the publish step is additionally guarded by `vars.PUBLISH_ENABLED == 'true'`.
Create required reviewers on the `pypi` environment in GitHub before any
release. Publishing is intentionally disabled by default.

Do not push upstream release tags to origin. Only the `paper-base` marker tag is
pushed during bootstrap. Future `v*` release tags are pushed deliberately only
when publishing is intended.

## Upstream Merge Policy

Quarterly, clone or pull the official Mercurial upstream, convert the updated
history to Git in a fresh staging repository, identify the newest release tag,
merge that release into this repository, and run the full baseline suite.
Resolve conflicts using this ledger as the map. Merge, never rebase, after the
fork is published.
