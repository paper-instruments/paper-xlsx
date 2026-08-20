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

CI uses the non-writing check mode:

```bash
python tools/autodoc/generate_api_reference.py --check
```
