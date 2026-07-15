"""Regression contracts for desktop-Excel semantic normalization."""

from __future__ import annotations

import io
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import call, patch

from ..automation import workspace as excel_workspace
from .performance import benchmark_contract_failures
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


class ExcelWorkspaceFallbackTests(unittest.TestCase):
    """Keep desktop-Excel automation usable when macOS protects its container."""

    def test_permission_denial_uses_isolated_shared_temporary_directory(self) -> None:
        """Fall back loudly and clean the unique shared temporary directory."""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            preferred = root / "private-excel-root"
            fallback_root = root / "shared-temp"
            fallback_root.mkdir()
            fallback = fallback_root / "PMWorkbookAutomation.pm-build-test"
            fallback.mkdir(mode=0o700)
            denied = PermissionError(1, "Operation not permitted", preferred / "pm-build-test")

            with (
                patch.object(
                    excel_workspace,
                    "EXCEL_FALLBACK_ROOT",
                    fallback_root,
                    create=True,
                ),
                patch.object(
                    excel_workspace,
                    "excel_workspace_root",
                    return_value=preferred,
                ),
                patch.object(
                    excel_workspace.tempfile,
                    "mkdtemp",
                    side_effect=(denied, str(fallback)),
                ) as make_directory,
                self.assertLogs(excel_workspace.__name__, level="WARNING"),
                excel_workspace.excel_working_directory("pm-build-") as work_dir,
            ):
                self.assertEqual(work_dir, fallback)
                self.assertTrue(work_dir.is_dir())

            self.assertFalse(fallback.exists())
            self.assertEqual(
                make_directory.call_args_list,
                [
                    call(prefix="pm-build-", dir=preferred),
                    call(
                        prefix="PMWorkbookAutomation.pm-build-",
                        dir=fallback_root,
                    ),
                ],
            )

    def test_fallback_failure_reports_both_errors_and_preserves_private_cause(self) -> None:
        """Keep the private denial as the cause while exposing fallback failure."""
        preferred = Path("/private/excel")
        fallback_root = Path("/private/tmp")
        denied = PermissionError(1, "Operation not permitted", preferred / "pm-build-test")
        fallback_failure = OSError(28, "No space left on device", fallback_root)

        with (
            patch.object(excel_workspace, "EXCEL_FALLBACK_ROOT", fallback_root),
            patch.object(
                excel_workspace,
                "excel_workspace_root",
                return_value=preferred,
            ),
            patch.object(
                excel_workspace.tempfile,
                "mkdtemp",
                side_effect=(denied, fallback_failure),
            ),
            self.assertRaises(excel_workspace.ExcelWorkspaceError) as raised,
            excel_workspace.excel_working_directory("pm-build-"),
        ):
            self.fail("workspace creation should fail")

        self.assertIn("Operation not permitted", str(raised.exception))
        self.assertIn("No space left on device", str(raised.exception))
        self.assertIsInstance(raised.exception.__cause__, excel_workspace.ExcelWorkspaceError)
        self.assertIs(raised.exception.__cause__.__cause__, denied)

    def test_non_permission_error_does_not_use_shared_fallback(self) -> None:
        """Do not conceal a private workspace failure unrelated to macOS access."""
        preferred = Path("/private/excel")
        input_output_error = OSError(5, "Input/output error", preferred)

        with (
            patch.object(
                excel_workspace,
                "excel_workspace_root",
                return_value=preferred,
            ),
            patch.object(
                excel_workspace.tempfile,
                "mkdtemp",
                side_effect=input_output_error,
            ) as make_directory,
            self.assertRaises(excel_workspace.ExcelWorkspaceError) as raised,
            excel_workspace.excel_working_directory("pm-build-"),
        ):
            self.fail("workspace creation should fail")

        self.assertIn("Input/output error", str(raised.exception))
        make_directory.assert_called_once_with(prefix="pm-build-", dir=preferred)


class PerformanceBenchmarkContractTests(unittest.TestCase):
    """Keep performance probes aligned to the current Items schema."""

    def test_edit_probes_target_declared_operational_columns(self) -> None:
        """Measure Latest Status and Delivery Health rather than stale coordinates."""
        self.assertEqual(benchmark_contract_failures(), [])
