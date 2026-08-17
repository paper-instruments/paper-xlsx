"""The paper-added modules must declare their public surface, accurately.

`__all__` has two consumers that fail quietly when it is wrong. `griffe check` reads it to
decide whether a change breaks the public API, so a name missing from it can be removed without
the gate objecting. The docs generator reads it to decide what to publish, so a name missing
from it silently vanishes from the reference rather than failing a build.

Both directions matter: a name in `__all__` that does not exist is a lie, and a public name
absent from `__all__` is an omission.

`openpyxl.preserve` deserves a note. It re-exports public names through a module-level
`__getattr__`, which exists to keep the package importable from anywhere without import cycles.
Those names resolve fine at runtime - `hasattr` finds them - but a static reader cannot follow
`__getattr__`, so the docs generator has to resolve them from the submodules that define them.
The last test here pins that down, because a name that resolves at runtime and vanishes
statically is precisely the kind of gap nobody notices.
"""
from __future__ import annotations

import ast
import importlib
import pathlib

import pytest

# -- modules paper added wholly; upstream openpyxl modules are not ours to declare --
PAPER_MODULES = ("errors", "oracle", "package", "preserve")
PACKAGE_ROOT = pathlib.Path(__file__).resolve().parents[2] / "openpyxl"


def _module_file(name):
    """Return the file defining `openpyxl.<name>`, whether module or package."""
    direct = PACKAGE_ROOT / f"{name}.py"
    return direct if direct.is_file() else PACKAGE_ROOT / name / "__init__.py"


def _defined_public_names(name):
    """Public names `openpyxl.<name>` defines itself, read statically."""
    tree = ast.parse(_module_file(name).read_text(encoding="utf-8"))
    names = set()
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            if not node.name.startswith("_"):
                names.add(node.name)
        elif isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id.isupper():
                    if not target.id.startswith("_"):
                        names.add(target.id)
    return names


class TestThePublicSurface:

    @pytest.mark.parametrize("name", PAPER_MODULES)
    def test_it_declares_all(self, name):
        module = importlib.import_module(f"openpyxl.{name}")
        assert hasattr(module, "__all__"), (
            f"openpyxl.{name} declares no __all__. Both `griffe check` and the docs generator "
            f"read it; without one, neither knows what is public."
        )

    @pytest.mark.parametrize("name", PAPER_MODULES)
    def test_every_exported_name_resolves(self, name):
        module = importlib.import_module(f"openpyxl.{name}")
        missing = sorted(n for n in module.__all__ if not hasattr(module, n))
        assert not missing, (
            f"openpyxl.{name}.__all__ exports {missing}, which do not resolve. Remove them or "
            f"restore the names."
        )

    @pytest.mark.parametrize("name", PAPER_MODULES)
    def test_no_public_name_is_left_unexported(self, name):
        module = importlib.import_module(f"openpyxl.{name}")
        unexported = sorted(_defined_public_names(name) - set(module.__all__))
        assert not unexported, (
            f"openpyxl.{name} defines public {unexported} but does not export them. An "
            f"unexported public name is invisible to `griffe check` and absent from the "
            f"generated reference. Add them to __all__, or underscore-prefix them if they were "
            f"never meant to be public."
        )

    def test_every_export_is_reachable_without_running_the_module(self):
        """Each `preserve` export must be defined in exactly one submodule.

        The docs generator reads statically and cannot follow `preserve`'s `__getattr__`, so it
        resolves each re-export from the submodule that defines it. That only works while every
        name has exactly one definer: none means the name silently disappears from the published
        reference, and more than one means there is no single right answer.
        """
        preserve = importlib.import_module("openpyxl.preserve")
        directory = PACKAGE_ROOT / "preserve"
        defined_in = {}

        for submodule in sorted(directory.glob("*.py")):
            if submodule.name == "__init__.py":
                continue
            tree = ast.parse(submodule.read_text(encoding="utf-8"))
            for node in tree.body:
                if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                    defined_in.setdefault(node.name, []).append(submodule.stem)

        here = _defined_public_names("preserve")
        problems = {}
        for name in preserve.__all__:
            if name in here:
                continue
            definers = defined_in.get(name, [])
            if len(definers) != 1:
                problems[name] = definers or "defined nowhere"

        assert not problems, (
            f"these openpyxl.preserve exports cannot be resolved to a single defining "
            f"submodule, so the generated reference would drop or misattribute them: {problems}"
        )
