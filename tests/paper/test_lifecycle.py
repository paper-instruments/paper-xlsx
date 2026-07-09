"""The part-lifecycle engine and its dividends (PLAN-v0.1 Batch 2;
PR-1 §1). Every part a save creates or deletes routes through
PartPlan.add_part/remove_part — no bespoke cascades."""
from __future__ import annotations

import io
import zipfile

import pytest

from openpyxl import load_workbook
from openpyxl.errors import (
    RelationshipPolicyError,
    TargetNotFoundError,
)

from .support.partdiff import part_payloads


class TestCustomPropsLifecycle:

    def test_first_custom_prop_creates_the_part(self, fixture_copy,
                                                 tmp_path):
        # v0 refused; the engine creates part + CT override + package rel
        src = fixture_copy("minimal/minimal_clean.xlsx")
        assert "docProps/custom.xml" not in part_payloads(src)
        wb = load_workbook(src, preserve=True)
        from openpyxl.packaging.custom import StringProperty

        wb.custom_doc_props.append(StringProperty(name="Reviewed",
                                                  value="yes"))
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        parts = part_payloads(out)
        assert b"Reviewed" in parts["docProps/custom.xml"]
        assert b"custom.xml" in parts["[Content_Types].xml"]
        assert b"custom-properties" in parts["_rels/.rels"]
        wb2 = load_workbook(out)
        assert wb2.custom_doc_props["Reviewed"].value == "yes"

    def test_removing_last_custom_prop_drops_the_part(self, fixture_copy,
                                                      tmp_path):
        # build a file WITH one custom prop first, then remove it
        src = fixture_copy("minimal/minimal_clean.xlsx")
        wb = load_workbook(src, preserve=True)
        from openpyxl.packaging.custom import StringProperty

        wb.custom_doc_props.append(StringProperty(name="Tmp", value="x"))
        staged = str(tmp_path / "staged.xlsx")
        wb.save(staged)

        wb2 = load_workbook(staged, preserve=True)
        del wb2.custom_doc_props["Tmp"]
        out = str(tmp_path / "o.xlsx")
        wb2.save(out)
        parts = part_payloads(out)
        assert "docProps/custom.xml" not in parts
        assert b"custom.xml" not in parts["[Content_Types].xml"]
        assert b"custom-properties" not in parts["_rels/.rels"]
        load_workbook(out)                        # reloadable


class TestStylesPartCreation:

    def _styleless(self, fixture_copy, tmp_path):
        import re

        src = fixture_copy("minimal/minimal_clean.xlsx")
        out = str(tmp_path / "styleless.xlsx")
        with zipfile.ZipFile(src) as zin, zipfile.ZipFile(out, "w") as zout:
            for name in zin.namelist():
                if name == "xl/styles.xml":
                    continue
                payload = zin.read(name)
                if name == "[Content_Types].xml":
                    payload = payload.replace(
                        b'<Override PartName="/xl/styles.xml" ContentType='
                        b'"application/vnd.openxmlformats-officedocument.'
                        b'spreadsheetml.styles+xml"/>', b"")
                if name == "xl/_rels/workbook.xml.rels":
                    payload = re.sub(
                        br'<Relationship [^>]*Target="styles.xml"[^>]*/>',
                        b"", payload)
                if name.startswith("xl/worksheets/sheet"):
                    payload = re.sub(br' s="\d+"', b"", payload)
                zout.writestr(name, payload)
        return out

    def test_styled_write_creates_styles_part(self, fixture_copy, tmp_path):
        src = self._styleless(fixture_copy, tmp_path)
        wb = load_workbook(src, preserve=True)
        from openpyxl.styles import Font

        wb["Sheet1"]["B2"] = 42
        wb["Sheet1"]["B2"].font = Font(bold=True)
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        parts = part_payloads(out)
        assert "xl/styles.xml" in parts
        assert b"/xl/styles.xml" in parts["[Content_Types].xml"]
        wb2 = load_workbook(out)
        assert wb2["Sheet1"]["B2"].font.bold is True
        assert wb2["Sheet1"]["B2"].value == 42


class TestReplacePart:

    def test_media_swap_lands_and_guards_hold(self, fixture_copy, tmp_path):
        src = fixture_copy("features/chart_image.xlsx")
        wb = load_workbook(src, preserve=True)
        new_png = part_payloads(src)["xl/media/image1.png"] + b"\x00"
        wb.replace_part("xl/media/image1.png", new_png)
        out = str(tmp_path / "o.xlsx")
        wb.save(out)
        assert part_payloads(out)["xl/media/image1.png"] == new_png

        wb2 = load_workbook(src, preserve=True)
        with pytest.raises(TargetNotFoundError):
            wb2.replace_part("xl/media/nope.png", b"x")
        with pytest.raises(RelationshipPolicyError, match="managed"):
            wb2.replace_part("xl/workbook.xml", b"x")
        with pytest.raises(RelationshipPolicyError, match="managed"):
            wb2.replace_part("xl/worksheets/sheet1.xml", b"x")

    def test_replace_part_meaningless_on_stock(self, fixture_copy):
        wb = load_workbook(fixture_copy("features/chart_image.xlsx"))
        with pytest.raises(ValueError, match="preserve"):
            wb.replace_part("xl/media/image1.png", b"x")


class TestEnginePrimitive:

    def test_add_part_refuses_existing_names(self):
        from openpyxl.preserve.lifecycle import PartPlan

        plan = PartPlan({"xl/workbook.xml"})
        with pytest.raises(RelationshipPolicyError, match="never overwrites"):
            plan.add_part("xl/workbook.xml", b"x")

    def test_remove_part_requires_existence(self):
        from openpyxl.preserve.lifecycle import PartPlan

        plan = PartPlan({"xl/workbook.xml"})
        with pytest.raises(TargetNotFoundError):
            plan.remove_part("xl/ghost.xml")

    def test_relative_targets(self):
        from openpyxl.preserve.lifecycle import _relative_target

        assert _relative_target("xl/workbook.xml",
                                "xl/styles.xml") == "styles.xml"
        assert _relative_target("", "docProps/custom.xml") == \
            "docProps/custom.xml"
        assert _relative_target("xl/worksheets/sheet1.xml",
                                "xl/comments1.xml") == "../comments1.xml"
