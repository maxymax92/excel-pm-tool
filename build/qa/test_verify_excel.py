"""Regression contracts for desktop-Excel semantic normalization."""

from __future__ import annotations

import io
import unittest
import zipfile

from .verify_excel import WorkbookSemanticError, _conditional_formats

MAIN_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
WORKSHEET_PART = "xl/worksheets/sheet1.xml"
STYLES_PART = "xl/styles.xml"
STYLES_XML = f"""\
<styleSheet xmlns="{MAIN_NAMESPACE}">
  <dxfs count="2">
    <dxf><font><b/></font></dxf>
    <dxf><font><i/></font></dxf>
  </dxfs>
</styleSheet>
"""


def _conditional_manifest(*conditional_formats: str) -> dict[str, tuple]:
    worksheet_xml = (
        f'<worksheet xmlns="{MAIN_NAMESPACE}">' + "".join(conditional_formats) + "</worksheet>"
    )
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as package:
        package.writestr(STYLES_PART, STYLES_XML)
        package.writestr(WORKSHEET_PART, worksheet_xml)
    buffer.seek(0)
    with zipfile.ZipFile(buffer) as package:
        return _conditional_formats(package)


def _rule(sqref: str, *, priority: int, dxf_id: int, formula: str) -> str:
    return (
        f'<conditionalFormatting sqref="{sqref}">'
        f'<cfRule type="expression" dxfId="{dxf_id}" priority="{priority}" stopIfTrue="1">'
        f"<formula>{formula}</formula>"
        "</cfRule>"
        "</conditionalFormatting>"
    )


class ConditionalFormatNormalizationTests(unittest.TestCase):
    """Protect equivalence and change detection across Excel rewrites."""

    def test_split_and_merged_relative_rules_are_equivalent(self) -> None:
        """Treat Excel's adjacent-rule coalescing as semantically neutral."""
        split = _conditional_manifest(
            _rule("B4", priority=1, dxf_id=0, formula="ISBLANK(B4)"),
            _rule("B5", priority=2, dxf_id=0, formula="ISBLANK(B5)"),
        )
        merged = _conditional_manifest(
            _rule("B4:B5", priority=1, dxf_id=0, formula="ISBLANK(B4)"),
        )

        self.assertEqual(split, merged)

    def test_formula_change_is_detected(self) -> None:
        """Retain exact formula behavior after origin normalization."""
        original = _conditional_manifest(
            _rule("B4:B5", priority=1, dxf_id=0, formula="ISBLANK(B4)"),
        )
        changed = _conditional_manifest(
            _rule("B4:B5", priority=1, dxf_id=0, formula="B4=0"),
        )

        self.assertNotEqual(original, changed)

    def test_overlapping_rule_order_change_is_detected(self) -> None:
        """Preserve precedence wherever conditional rules overlap."""
        original = _conditional_manifest(
            _rule("B4:B5", priority=1, dxf_id=0, formula="ISBLANK(B4)"),
            _rule("B5", priority=2, dxf_id=1, formula="B5=0"),
        )
        reordered = _conditional_manifest(
            _rule("B4:B5", priority=2, dxf_id=0, formula="ISBLANK(B4)"),
            _rule("B5", priority=1, dxf_id=1, formula="B5=0"),
        )

        self.assertNotEqual(original, reordered)

    def test_duplicate_priority_is_rejected(self) -> None:
        """Reject ambiguous global rule ordering instead of guessing."""
        with self.assertRaisesRegex(WorkbookSemanticError, "duplicate cfRule priority 1"):
            _conditional_manifest(
                _rule("B4", priority=1, dxf_id=0, formula="ISBLANK(B4)"),
                _rule("B5", priority=1, dxf_id=1, formula="B5=0"),
            )
