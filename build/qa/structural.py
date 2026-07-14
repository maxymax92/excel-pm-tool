"""Orchestrate package, workbook, formula, and cached-value QA.

``--recalc`` asks Excel to full-rebuild a disposable copy and requires exact
authored semantics to survive. Cached-value assertions inspect the supplied
artifact, so a release must already contain its published calculation cache.
"""

from __future__ import annotations

import argparse
import logging
import sys
from enum import Enum
from pathlib import Path
from typing import TYPE_CHECKING

import openpyxl

from ..paths import ROOT
from .structural_formula import formula_failures
from .structural_package import package_failures
from .structural_results import report
from .structural_values import cached_value_failures
from .structural_workbook import workbook_failures
from .verify_excel import recalculate_and_compare

if TYPE_CHECKING:
    from collections.abc import Sequence

LOGGER = logging.getLogger(__name__)
DEFAULT = ROOT / "dist" / "PM_Workbook.xlsx"
SUPPORTED_SUFFIXES = {".xlsx", ".xlsm"}


class _StructuralProblem(Enum):
    SUFFIX = "structural QA requires .xlsx or .xlsm: {}"


class StructuralUsageError(ValueError):
    """Report an unsupported structural-QA input path."""

    def __init__(self, problem: _StructuralProblem, *details: object) -> None:
        """Create a stable input diagnostic."""
        super().__init__(problem.value.format(*details))


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run structural and cached-value QA on one generated workbook."
    )
    parser.add_argument(
        "--recalc",
        action="store_true",
        help="full-rebuild a disposable copy and verify exact semantic preservation",
    )
    parser.add_argument(
        "path",
        nargs="?",
        type=Path,
        default=DEFAULT,
        help=".xlsx or .xlsm workbook; defaults to dist/PM_Workbook.xlsx",
    )
    return parser


def _resolve_path(path: Path) -> Path:
    resolved = path if path.is_absolute() else (ROOT / path).resolve()
    if not resolved.is_file():
        raise FileNotFoundError(resolved)
    if resolved.suffix.lower() not in SUPPORTED_SUFFIXES:
        raise StructuralUsageError(_StructuralProblem.SUFFIX, resolved)
    return resolved


def _recalculation_failures(path: Path) -> list[str]:
    issues = recalculate_and_compare(path)
    if not issues:
        LOGGER.info("excel disposable full rebuild preserved exact authored semantics")
    return [f"EXCEL SAVE {issue}" for issue in issues]


def _workbook_checks(path: Path, *, check_expected_values: bool) -> list[str]:
    cached_workbook = openpyxl.load_workbook(path, data_only=True)
    formula_workbook = openpyxl.load_workbook(path, data_only=False)
    try:
        return [
            *cached_value_failures(cached_workbook, check_expected=check_expected_values),
            *formula_failures(formula_workbook),
            *workbook_failures(formula_workbook),
        ]
    finally:
        cached_workbook.close()
        formula_workbook.close()


def main(argv: Sequence[str] | None = None) -> int:
    """Run structural QA and return a shell-compatible status.

    Returns:
        Zero for a clean workbook and one for observed violations.

    """
    arguments = _parser().parse_args(sys.argv[1:] if argv is None else argv)
    path = _resolve_path(arguments.path)
    failures = _recalculation_failures(path) if arguments.recalc else []
    package_issues = package_failures(path)
    failures.extend(package_issues)
    if not package_issues:
        failures.extend(_workbook_checks(path, check_expected_values=arguments.recalc))
    return report(path, failures)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    raise SystemExit(main())
