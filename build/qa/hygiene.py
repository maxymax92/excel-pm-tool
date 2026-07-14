"""Validate repository source hygiene, current paths and documentation links."""

from __future__ import annotations

import logging
import re
from pathlib import Path

from ..paths import ROOT
from .vba_source import source_failures

LOGGER = logging.getLogger(__name__)
TEXT_SUFFIXES = {".py", ".bas", ".txt", ".applescript", ".md", ".toml"}
EXCLUDED_PARTS = {
    ".git",
    ".remember",
    ".ruff_cache",
    ".venv",
    "__pycache__",
    "dist",
}
RULES = {
    "VBA error suppression": re.compile(r"On\s+Error\s+Resume\s+Next", re.IGNORECASE),
    "VBA append-only file output": re.compile(r"Open\s+[^\n]+\s+For\s+Append\b", re.IGNORECASE),
    "VBA debug output": re.compile(r"\bDebug\.Print\b", re.IGNORECASE),
    "lossy text decoding": re.compile(
        r"(?:decode\([^\n]*(?:ignore|replace)|errors\s*=\s*['\"](?:ignore|replace))",
        re.IGNORECASE,
    ),
    "warning suppression": re.compile(
        r"warnings\.(?:filterwarnings|simplefilter)|catch_warnings", re.IGNORECASE
    ),
    "optional file deletion": re.compile(r"missing_ok\s*=\s*True"),
    "empty exception handler": re.compile(
        r"except\s+(?:Exception|BaseException)(?:\s+as\s+\w+)?\s*:\s*pass\b"
    ),
    "lint or coverage suppression": re.compile(
        r"#\s*(?:noqa\b|type:\s*ignore\b|pragma:\s*no\s*cover\b)", re.IGNORECASE
    ),
}
JOURNAL_PATTERN = re.compile(
    r"\b(?:legacy|superseded|recovered|removed|retired|formerly)\b|"
    r"\bno longer\b|\bused to\b|\bmigration notes?\b",
    re.IGNORECASE,
)
OBSOLETE_COMMAND_PATHS = (
    "build/build.py",
    "build/qa.py",
    "build/qa_abuse.py",
    "build/qa_design.py",
    "build/qa_empty.py",
    "build/qa_overview.py",
    "build/qa_scenario.py",
    "build/verify_excel.py",
    "build/verify_vba.py",
)


def source_files() -> list[Path]:
    """Collect current repository text sources covered by hygiene checks.

    Returns:
        Sorted source paths, excluding generated and environment directories.

    """
    current = Path(__file__).resolve()
    return sorted(
        path
        for path in ROOT.rglob("*")
        if path.is_file()
        and path.suffix.lower() in TEXT_SUFFIXES
        and not EXCLUDED_PARTS.intersection(path.parts)
        and path.resolve() != current
    )


def documentation_link_failures(path: Path, content: str) -> list[str]:
    """Return missing relative Markdown links with source line numbers.

    Returns:
        Missing-target diagnostics with source line numbers.

    """
    failures: list[str] = []
    for match in re.finditer(r"(?<!!)\[[^\]]+\]\(([^)]+)\)", content):
        target = match.group(1).strip().strip("<>")
        if not target or target.startswith(("#", "http://", "https://", "mailto:")):
            continue
        target = target.split("#", 1)[0].split("?", 1)[0]
        resolved = (path.parent / target).resolve()
        if not resolved.exists():
            line = content.count("\n", 0, match.start()) + 1
            failures.append(f"{path.relative_to(ROOT)}:{line}: missing Markdown target {target}")
    return failures


def vba_registry_failures() -> list[str]:
    """Return full-registry VBA source-contract failures.

    Returns:
        Registry diagnostics labelled for the hygiene report.

    """
    return [f"VBA source: {failure}" for failure in source_failures()]


def audit() -> list[str]:
    """Run every repository hygiene and VBA source-contract check.

    Returns:
        All observed hygiene violations in deterministic traversal order.

    """
    failures: list[str] = []
    for path in source_files():
        try:
            source_text = path.read_text(encoding="utf-8")
        except UnicodeError as exc:
            failures.append(f"{path.relative_to(ROOT)}: invalid UTF-8: {exc}")
            continue
        for label, pattern in RULES.items():
            for match in pattern.finditer(source_text):
                line = source_text.count("\n", 0, match.start()) + 1
                failures.append(f"{path.relative_to(ROOT)}:{line}: {label}")
        for line_number, line_text in enumerate(source_text.splitlines(), start=1):
            stripped = line_text.lstrip()
            is_comment = (
                (path.suffix.lower() in {".py", ".toml"} and stripped.startswith("#"))
                or (path.suffix.lower() in {".bas", ".txt"} and stripped.startswith("'"))
                or (path.suffix.lower() == ".applescript" and stripped.startswith("--"))
            )
            if (path.suffix.lower() == ".md" or is_comment) and JOURNAL_PATTERN.search(line_text):
                failures.append(
                    f"{path.relative_to(ROOT)}:{line_number}: repository journal language"
                )
        for command_path in OBSOLETE_COMMAND_PATHS:
            index = source_text.find(command_path)
            if index >= 0:
                line = source_text.count("\n", 0, index) + 1
                failures.append(
                    f"{path.relative_to(ROOT)}:{line}: obsolete command path {command_path}"
                )
        if path.suffix.lower() == ".md":
            failures.extend(documentation_link_failures(path, source_text))
    failures.extend(vba_registry_failures())
    return failures


def main() -> int:
    """Report repository hygiene results.

    Returns:
        A process exit status: zero for success and one for failure.

    """
    failures = audit()
    if failures:
        LOGGER.error("HYGIENE QA FAIL (%s issue(s))", len(failures))
        for failure in failures:
            LOGGER.error("  - %s", failure)
        return 1
    LOGGER.info("HYGIENE QA PASS")
    return 0


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    raise SystemExit(main())
