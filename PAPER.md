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
