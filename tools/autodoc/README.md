# Plain-Markdown API reference

`generate_api_reference.py` renders the Paper-added public API from the source
docstrings into `doc/generated-api/`. It deliberately omits the inherited
openpyxl surface and produces plain Markdown suitable for repository browsing
or lazy inclusion in an agent skill.

Install the pinned generator and regenerate:

```bash
python -m pip install -r tools/autodoc/requirements.txt
python tools/autodoc/generate_api_reference.py
```

The Documentation job runs the non-writing check on every pull request and on
every push to `main`:

```bash
python tools/autodoc/generate_api_reference.py --check
```

CI does not rewrite or commit generated files. When the check reports drift,
run the writing command locally and commit the updated Markdown with the source
change.
