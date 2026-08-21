"""Generate the Paper-added public API reference as plain Markdown."""

from __future__ import annotations

import argparse
import ast
import difflib
import re
import sys
from importlib.metadata import version as distribution_version
from pathlib import Path
from typing import Dict, List, Optional

import griffe
from griffe2md import render_object_docs
from griffe_typingdoc import TypingDocExtension


GENERATOR_VERSION = "1.5.0"
LOCAL_CROSSREF = re.compile(r"\[([^\]]+)\]\(#[^)]+\)")
REST_WARNING = re.compile(r"(?:<br>|\s)\.\. warning::\s*")
TYPE_CODE = re.compile(r"<code>([^<]*)</code>")
OUTPUT_NAMES = (
    "INDEX.md",
    "chart.md",
    "errors.md",
    "oracle.md",
    "package.md",
    "preserve.md",
    "workbook.md",
    "worksheet.md",
)


def paper_version(source: Path) -> str:
    """Read ``__paper_version__`` from a source checkout without importing it."""
    version_file = source / "openpyxl" / "_paper_version.py"
    for node in ast.walk(ast.parse(version_file.read_text(encoding="utf-8"))):
        if not isinstance(node, ast.Assign):
            continue
        for target in node.targets:
            if isinstance(target, ast.Name) and target.id == "__paper_version__":
                return ast.literal_eval(node.value)
    raise RuntimeError(f"No __paper_version__ found in {version_file}")


def config(*, heading_level: int, members: Optional[List[str]] = None) -> dict:
    """Return the pinned griffe2md rendering configuration."""
    return {
        "heading_level": heading_level,
        "members": members,
        "members_order": "source",
        "inherited_members": False,
        "show_submodules": False,
        "show_root_heading": True,
        "show_root_full_path": True,
        "show_root_members_full_path": False,
        "show_object_full_path": False,
        "show_bases": False,
        "show_signature": True,
        "show_signature_annotations": True,
        "separate_signature": True,
        "docstring_section_style": "table",
        "summary": False,
    }


def banner(version: str, subject: str) -> str:
    return (
        f"<!-- Generated from paper-xlsx {version} with griffe2md "
        f"{GENERATOR_VERSION}. Do not edit by hand. -->\n\n"
        f"> Generated from the `{subject}` docstrings in paper-xlsx {version}.\n\n"
    )


def exported(module: griffe.Module) -> List[str]:
    """Return the module's declared public exports or fail loudly."""
    names = sorted(str(name) for name in (module.exports or []))
    if not names:
        raise RuntimeError(f"{module.path} has no __all__")
    return names


def class_members(obj: griffe.Class) -> List[str]:
    """Return public methods and properties, excluding implementation assignments."""
    names = []
    for name, member in obj.members.items():
        if name.startswith("_"):
            continue
        if isinstance(member, griffe.Function):
            names.append(name)
        elif isinstance(member, griffe.Attribute) and "property" in member.labels:
            names.append(name)
    return names


def find_reexport(module: griffe.Module, name: str):
    """Resolve one export, including lazy package-level re-exports."""
    if name in module.members:
        return module.members[name]
    matches = [
        child.members[name]
        for child in module.modules.values()
        if name in child.members
    ]
    if len(matches) != 1:
        raise RuntimeError(
            f"Expected one definition for {module.path}.{name}, found {len(matches)}"
        )
    return matches[0]


def replace_root_heading(markdown: str, public_path: str, level: int) -> str:
    """Replace griffe2md's root heading with the public import path."""
    _, separator, rest = markdown.partition("\n")
    return f"{'#' * level} `{public_path}`{separator}{rest}"


def normalize_markdown(markdown: str) -> str:
    """Remove dead links and normalize renderer-dependent Markdown details."""
    markdown = LOCAL_CROSSREF.sub(r"\1", markdown)
    markdown = REST_WARNING.sub(" **Warning:** ", markdown)
    return TYPE_CODE.sub(
        lambda match: f"<code>{match.group(1).replace(' - ', '-')}</code>",
        markdown,
    )


def render_public_member(module: griffe.Module, name: str) -> str:
    obj = find_reexport(module, name)
    target = obj.final_target if isinstance(obj, griffe.Alias) else obj
    members = class_members(target) if isinstance(target, griffe.Class) else None
    rendered = render_object_docs(
        obj,
        config(heading_level=2, members=members),
        format_md=True,
    )
    return normalize_markdown(
        replace_root_heading(rendered, f"{module.path}.{name}", 2)
    )


def render_module(module: griffe.Module, version: str) -> str:
    intro = render_object_docs(
        module,
        config(heading_level=1, members=[]),
        format_md=True,
    )
    sections = [render_public_member(module, name) for name in exported(module)]
    return normalize_markdown(
        banner(version, module.path) + intro + "\n\n" + "\n\n".join(sections)
    )


def render_class_surface(
    obj: griffe.Class,
    public_path: str,
    methods: List[str],
    version: str,
) -> str:
    """Render only the Paper-added methods on an inherited class."""
    missing = sorted(set(methods) - set(obj.functions))
    if missing:
        raise RuntimeError(f"{obj.path} is missing {missing}")
    rendered = render_object_docs(
        obj,
        config(heading_level=1, members=methods),
        format_md=True,
    )
    rendered = rendered.replace(obj.path, public_path)
    return normalize_markdown(
        banner(version, public_path)
        + replace_root_heading(rendered, public_path, 1)
    )


def render_outputs(source: Path) -> Dict[str, str]:
    """Render every generated file in memory so failures cannot leave partial output."""
    installed = distribution_version("griffe2md")
    if installed != GENERATOR_VERSION:
        raise RuntimeError(
            f"griffe2md {GENERATOR_VERSION} is required; found {installed}"
        )

    version = paper_version(source)
    root = griffe.load(
        "openpyxl",
        search_paths=[str(source)],
        docstring_parser="auto",
        store_source=False,
        extensions=griffe.load_extensions(TypingDocExtension),
    )

    files = {
        "errors.md": render_module(root.modules["errors"], version),
        "oracle.md": render_module(root.modules["oracle"], version),
        "package.md": render_module(root.modules["package"], version),
        "preserve.md": render_module(root.modules["preserve"], version),
        "workbook.md": render_class_surface(
            root.modules["workbook"].modules["workbook"].classes["Workbook"],
            "openpyxl.workbook.Workbook",
            ["save", "validate", "search", "set_pivot_refresh_on_load"],
            version,
        ),
        "worksheet.md": render_class_surface(
            root.modules["worksheet"].modules["worksheet"].classes["Worksheet"],
            "openpyxl.worksheet.Worksheet",
            ["allowed_values", "append_table_row", "replace_image"],
            version,
        ),
        "chart.md": render_class_surface(
            root.modules["chart"].modules["_chart"].classes["ChartBase"],
            "openpyxl.chart.ChartBase",
            ["repoint"],
            version,
        ),
    }
    files["INDEX.md"] = f"""# paper-xlsx {version} generated API reference

<!-- Generated from paper-xlsx {version} with griffe2md {GENERATOR_VERSION}. Do not edit by hand. -->

This is the Paper-added public surface, generated from the package docstrings as plain Markdown.
Search by exact symbol, then open only the relevant file.

- [Typed refusals and warnings](errors.md)
- [Calculation, evaluation, and certification](oracle.md)
- [Semantic package and cell diffs](package.md)
- [Preservation helpers and receipts](preserve.md)
- [Workbook methods](workbook.md)
- [Worksheet methods](worksheet.md)
- [Chart methods](chart.md)

High-value symbols include `set_pivot_refresh_on_load`, `allowed_values`,
`append_table_row`, `replace_image`, and `repoint`.
"""
    if tuple(sorted(files)) != OUTPUT_NAMES:
        raise RuntimeError("Generated file inventory does not match OUTPUT_NAMES")
    return files


def check_outputs(output: Path, files: Dict[str, str]) -> int:
    """Return nonzero when the checked-in reference is missing, stale, or changed."""
    actual = {path.name for path in output.glob("*.md")}
    expected = set(files)
    problems = []
    diffs = []
    for name in sorted(expected - actual):
        problems.append(f"missing: {output / name}")
    for name in sorted(actual - expected):
        problems.append(f"stale: {output / name}")
    for name in sorted(actual & expected):
        current = (output / name).read_text(encoding="utf-8")
        if current != files[name]:
            problems.append(f"changed: {output / name}")
            diffs.extend(
                difflib.unified_diff(
                    current.splitlines(),
                    files[name].splitlines(),
                    fromfile=f"checked-in/{name}",
                    tofile=f"generated/{name}",
                    lineterm="",
                )
            )
    if problems:
        print("Generated API reference is out of date:", file=sys.stderr)
        print("\n".join(f"  {problem}" for problem in problems), file=sys.stderr)
        print(
            "Run: python tools/autodoc/generate_api_reference.py",
            file=sys.stderr,
        )
        if diffs:
            print("\n".join(diffs), file=sys.stderr)
        return 1
    print(f"Generated API reference is current: {len(files)} files in {output}")
    return 0


def write_outputs(output: Path, files: Dict[str, str]) -> None:
    """Replace the generated Markdown inventory after rendering succeeds."""
    output.mkdir(parents=True, exist_ok=True)
    for path in output.glob("*.md"):
        if path.name not in files:
            path.unlink()
    for name, content in files.items():
        (output / name).write_text(content, encoding="utf-8")
    print(f"Generated {len(files)} files in {output}")


def main() -> int:
    repo_root = Path(__file__).resolve().parents[2]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=repo_root)
    parser.add_argument("--out", type=Path, default=repo_root / "doc/generated-api")
    parser.add_argument(
        "--check",
        action="store_true",
        help="fail when the checked-in output differs instead of writing it",
    )
    args = parser.parse_args()

    source = args.source.resolve()
    output = args.out.resolve()
    files = render_outputs(source)
    if args.check:
        return check_outputs(output, files)
    write_outputs(output, files)
    return 0


if __name__ == "__main__":
    sys.exit(main())
