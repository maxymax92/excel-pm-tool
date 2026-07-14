"""Run the fail-fast release matrix and bind live Excel evidence.

``prepare`` builds and verifies the release, writes a template bound to the
exact source, workbook, and installed Excel patch version, then exits with the
deliberate blocked status. ``final`` accepts only completed, fresh evidence for
that exact state and reruns every non-mutating gate sequentially.
"""

from __future__ import annotations

import argparse
import asyncio
import hashlib
import json
import logging
import os
import re
import shlex
import sys
import tempfile
from dataclasses import dataclass
from datetime import UTC, datetime, timedelta
from enum import Enum
from pathlib import Path
from typing import TYPE_CHECKING, cast

from ..paths import DIST, ROOT

if TYPE_CHECKING:
    from collections.abc import Iterable, Mapping, Sequence, Set

LOGGER = logging.getLogger(__name__)

RELEASE_WORKBOOK = ROOT / "PM_Workbook.xlsm"
DIST_WORKBOOK = DIST / "PM_Workbook.xlsm"
QA_WORKBOOK = DIST / "PM_Workbook.xlsx"
EVIDENCE_SCHEMA = 3
EVIDENCE_MAX_AGE = timedelta(hours=24)
SUPPORTED_EXCEL_TRAIN = "16.110"
PREPARED_BUT_BLOCKED = 2
SOURCE_SUFFIXES = {".applescript", ".bas", ".bin", ".md", ".py", ".toml", ".txt"}
SOURCE_ROOT_FILES = {"README.md", "pyproject.toml", "uv.lock"}
SHA256_HEX_LENGTH = 64
MIN_MARKDOWN_TABLES = 6
EXCEL_VERSION_TIMEOUT_SECONDS = 30
STRUCTURAL_MODULE = "build.qa.structural"
STRUCTURAL_GATE_COUNT = 2
MIN_MODULE_ARGUMENTS = 2
LIVE_CHECKS = {
    "all_sheets_visual_inspection",
    "calc_hidden_and_protected",
    "config_in_use_role_edits_rollback",
    "editable_tables_unprotected",
    "export_button_runs",
    "export_cancel_preserves_files",
    "export_markdown_content_complete",
    "export_markdown_no_sidecars",
    "export_markdown_utf8",
    "export_write_failure_restores_destination",
    "item_bulk_edit_rejected_without_partial_stamps",
    "item_created_and_updated_stamped",
    "item_delivery_health_blocked_stamped",
    "item_id_created",
    "item_invalid_edits_rejected_without_partial_stamps",
    "item_latest_status_stamped",
    "item_status_lifecycle_stamped",
    "organise_blank_rows_ignored",
    "organise_button_runs",
    "organise_partial_and_invalid_rows_rejected",
    "organise_rows_sorted_and_outlined",
    "protected_views_remain_protected",
    "raid_bulk_edit_rejected_without_partial_stamps",
    "raid_id_created",
    "raid_invalid_edits_rejected_without_partial_stamps",
    "raid_lifecycle_stamped",
    "vba_project_compiles",
    "workbook_open_maximised",
    "workbook_open_no_repair_notice",
    "workbook_open_overview_first",
}
MARKDOWN_HEADINGS = (
    "# PM workbook status",
    "## Overview",
    "### Executive Status Summary",
    "### Top RAID",
    "### Coming up",
    "### Recent progress",
    "## Source data",
    "### Items",
    "### RAID",
)


class _ReleaseProblem(Enum):
    GATE_EXIT = "release gate {!r} failed with exit {}"
    EVIDENCE_EXISTS = "evidence destination already exists: {}"
    EVIDENCE_IN_REPOSITORY = "macro evidence must be stored outside the repository"
    EVIDENCE_WRITE_CLEANUP = (
        "evidence-template write failed and temporary-file cleanup also failed: {}: {}"
    )
    ROOT_NOT_OBJECT = "macro evidence root must be a JSON object"
    INVALID_JSON = "macro evidence is not valid UTF-8 JSON: {}"
    KEYS_DIFFER = "{} keys differ: missing={}, extra={}"
    TESTED_AT_TYPE = "tested_at must be a timezone-aware ISO-8601 string"
    TESTED_AT_VALUE = "tested_at is not valid ISO-8601"
    TESTED_AT_OFFSET = "tested_at must include a UTC offset"
    TESTED_AT_FUTURE = "tested_at is in the future"
    TESTED_AT_STALE = "live macro evidence is more than 24 hours old"
    EXPORT_PATH_TYPE = "export.path must be an absolute path"
    EXPORT_PATH_VALUE = "export.path must be an absolute .md path"
    EXPORT_DIGEST = "export.sha256 must be a lowercase SHA-256 digest"
    EXPORT_DIGEST_MISMATCH = "exported Markdown SHA-256 does not match the evidence"
    EXPORT_ENTRIES_TYPE = "{} must be a list of names"
    EXPORT_ENTRIES_DUPLICATE = "export directory evidence contains duplicate names"
    EXPORT_DELTA = "export directory evidence does not prove one new Markdown file"
    EXPORT_CHANGED = "export directory contents changed after live evidence was recorded"
    EXPORT_UTF8 = "exported Markdown is not valid UTF-8"
    EXPORT_HEADING = "exported Markdown is missing heading {!r}"
    EXPORT_CONTROL = "exported Markdown contains control characters"
    EXPORT_TABLES = "exported Markdown does not contain every required table"
    SCHEMA = "macro evidence schema must be {}"
    SOURCE_DIGEST = "macro evidence was recorded for different source files"
    WORKBOOK_DIGEST = "macro evidence was recorded for a different release workbook"
    WORKBOOK_COPIES = "root and dist macro workbooks are not byte-identical"
    EXCEL_VERSION_TYPE = "excel_version must be a dotted numeric version"
    EXCEL_VERSION_QUERY = "could not read the installed Excel version: {}"
    EXCEL_TRAIN = "installed Excel {} is outside the supported {} train"
    EXCEL_CHANGED = "installed Excel changed after prepare: evidence {}, current {}"
    TESTER = "tester must identify the person who ran the live checks"
    CHECKS_TYPE = "checks must be a JSON object"
    CHECKS_FAILED = "live macro checks are not PASS: {}"
    EXPORT_TYPE = "export must be a JSON object"
    STRUCTURAL_COUNT = "artifact gate matrix must contain exactly two structural invocations"
    STRUCTURAL_UNIQUE = "structural gate invocations must be unique"
    STRUCTURAL_FORMATS = "structural gates must cover one .xlsx and one .xlsm path"
    STRUCTURAL_PATHS = "structural gates target {}, expected {}"


class ReleaseError(RuntimeError):
    """Report an automated release-gate failure."""

    def __init__(self, problem: _ReleaseProblem, *details: object) -> None:
        """Create a stable release diagnostic."""
        super().__init__(problem.value.format(*details))


class EvidenceValidationError(ValueError):
    """Report invalid or stale live-macro evidence."""

    def __init__(self, problem: _ReleaseProblem, *details: object) -> None:
        """Create a stable evidence diagnostic."""
        super().__init__(problem.value.format(*details))


class EvidenceTypeError(TypeError):
    """Report a JSON evidence value with the wrong type."""

    def __init__(self, problem: _ReleaseProblem, *details: object) -> None:
        """Create a stable evidence type diagnostic."""
        super().__init__(problem.value.format(*details))


@dataclass(frozen=True, slots=True)
class Gate:
    """One labelled Python-module invocation in the release matrix."""

    label: str
    arguments: tuple[str, ...]

    @property
    def command(self) -> tuple[str, ...]:
        """Return the complete command for the active Python interpreter.

        Returns:
            Executable and arguments suitable for direct process execution.

        """
        return (sys.executable, *self.arguments)


@dataclass(frozen=True, slots=True)
class _CommandResult:
    returncode: int
    stdout: str
    stderr: str


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run the complete release gate and require bound live-macro evidence."
    )
    subparsers = parser.add_subparsers(dest="phase", required=True)
    prepare = subparsers.add_parser(
        "prepare",
        help="build and verify, then write a live-macro evidence template",
    )
    prepare.add_argument(
        "--evidence-template",
        required=True,
        type=Path,
        help="new JSON file to create outside the repository",
    )
    final = subparsers.add_parser(
        "final",
        help="verify fresh live-macro evidence for the exact built artifact",
    )
    final.add_argument(
        "--macro-evidence",
        required=True,
        type=Path,
        help="completed JSON evidence created from the prepare template",
    )
    return parser


async def _execute_gate(command: tuple[str, ...]) -> int:
    process = await asyncio.create_subprocess_exec(*command, cwd=ROOT)
    return await process.wait()


async def _capture_command(command: tuple[str, ...]) -> _CommandResult:
    process = await asyncio.create_subprocess_exec(
        *command,
        cwd=ROOT,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    try:
        async with asyncio.timeout(EXCEL_VERSION_TIMEOUT_SECONDS):
            stdout_bytes, stderr_bytes = await process.communicate()
    except TimeoutError as timeout_error:
        process.kill()
        await process.wait()
        detail = f"version query exceeded {EXCEL_VERSION_TIMEOUT_SECONDS} seconds"
        raise ReleaseError(_ReleaseProblem.EXCEL_VERSION_QUERY, detail) from timeout_error
    return _CommandResult(
        process.returncode or 0,
        stdout_bytes.decode("utf-8").strip(),
        stderr_bytes.decode("utf-8").strip(),
    )


def _run_gate(gate: Gate) -> None:
    LOGGER.info("\n=== %s ===", gate.label)
    LOGGER.info("%s", shlex.join(gate.command))
    returncode = asyncio.run(_execute_gate(gate.command))
    if returncode != 0:
        raise ReleaseError(_ReleaseProblem.GATE_EXIT, gate.label, returncode)


def _run_gates(gates: Iterable[Gate]) -> None:
    for gate in gates:
        _run_gate(gate)


def _python_gate(label: str, *arguments: str) -> Gate:
    return Gate(label, arguments)


def _source_gate_matrix() -> tuple[Gate, ...]:
    return (
        _python_gate("Ruff format", "-m", "ruff", "format", "--check", "build"),
        _python_gate("Ruff lint", "-m", "ruff", "check", "build"),
        _python_gate("Python compilation", "-m", "compileall", "-q", "build"),
        _python_gate("source hygiene", "-m", "build.qa.hygiene"),
        _python_gate(
            "Excel semantic comparison contracts",
            "-m",
            "unittest",
            "-q",
            "build.qa.test_verify_excel",
        ),
        _python_gate("VBA source contract", "-m", "build.qa.vba_source"),
        _python_gate("compiled VBA source match", "-m", "build.qa.verify_vba"),
    )


def _artifact_gate_matrix() -> tuple[Gate, ...]:
    return (
        _python_gate(
            "structural xlsx",
            "-m",
            STRUCTURAL_MODULE,
            "--recalc",
            str(QA_WORKBOOK),
        ),
        _python_gate(
            "structural xlsm",
            "-m",
            STRUCTURAL_MODULE,
            "--recalc",
            str(DIST_WORKBOOK),
        ),
        _python_gate(
            "design",
            "-m",
            "build.qa.design",
            str(QA_WORKBOOK),
            str(DIST_WORKBOOK),
        ),
        _python_gate("empty state", "-m", "build.qa.empty_state"),
        _python_gate("formula scenarios", "-m", "build.qa.formula_scenarios"),
        _python_gate("abuse scenarios", "-m", "build.qa.abuse"),
        _python_gate("Overview scenarios", "-m", "build.qa.overview"),
        _python_gate(
            "Excel save preservation",
            "-m",
            "build.qa.verify_excel",
            str(QA_WORKBOOK),
            str(DIST_WORKBOOK),
        ),
        _python_gate(
            "desktop Excel performance",
            "-m",
            "build.qa.performance",
            str(DIST_WORKBOOK),
        ),
        _python_gate("populated ship demo", "-m", "build.scenarios.ship_demo"),
    )


def _structural_paths(gates: Sequence[Gate]) -> tuple[Path, ...]:
    return tuple(
        Path(gate.arguments[-1]).expanduser().resolve()
        for gate in gates
        if len(gate.arguments) >= MIN_MODULE_ARGUMENTS
        and gate.arguments[0] == "-m"
        and gate.arguments[1] == STRUCTURAL_MODULE
    )


def _validate_artifact_gate_matrix(gates: Sequence[Gate]) -> None:
    paths = _structural_paths(gates)
    if len(paths) != STRUCTURAL_GATE_COUNT:
        raise ReleaseError(_ReleaseProblem.STRUCTURAL_COUNT)
    if len(set(paths)) != len(paths):
        raise ReleaseError(_ReleaseProblem.STRUCTURAL_UNIQUE)
    if {path.suffix.lower() for path in paths} != {".xlsx", ".xlsm"}:
        raise ReleaseError(_ReleaseProblem.STRUCTURAL_FORMATS)
    actual = set(paths)
    expected = {QA_WORKBOOK.resolve(), DIST_WORKBOOK.resolve()}
    if actual != expected:
        raise ReleaseError(
            _ReleaseProblem.STRUCTURAL_PATHS,
            sorted(map(str, actual)),
            sorted(map(str, expected)),
        )


def _source_files() -> list[Path]:
    files: list[Path] = []
    for name in sorted(SOURCE_ROOT_FILES):
        path = ROOT / name
        if not path.is_file():
            raise FileNotFoundError(path)
        files.append(path)
    for directory in (ROOT / "build", ROOT / "docs"):
        files.extend(
            path
            for path in sorted(directory.rglob("*"))
            if "__pycache__" not in path.parts
            and path.is_file()
            and path.suffix.lower() in SOURCE_SUFFIXES
        )
    return files


def _source_digest() -> str:
    digest = hashlib.sha256()
    for path in _source_files():
        relative = path.relative_to(ROOT).as_posix().encode("utf-8")
        data = path.read_bytes()
        digest.update(len(relative).to_bytes(8, "big"))
        digest.update(relative)
        digest.update(len(data).to_bytes(8, "big"))
        digest.update(data)
    return digest.hexdigest()


def _file_digest(path: Path) -> str:
    if not path.is_file():
        raise FileNotFoundError(path)
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _installed_excel_version() -> str:
    command = (
        "osascript",
        "-e",
        'tell application "Microsoft Excel" to get version',
    )
    completed = asyncio.run(_capture_command(command))
    if completed.returncode != 0 or completed.stderr:
        detail = completed.stderr or completed.stdout or f"exit {completed.returncode}"
        raise ReleaseError(_ReleaseProblem.EXCEL_VERSION_QUERY, detail)
    if re.fullmatch(r"\d+(?:\.\d+)+", completed.stdout) is None:
        raise EvidenceValidationError(_ReleaseProblem.EXCEL_VERSION_TYPE)
    return completed.stdout


def _require_supported_excel_version() -> str:
    version = _installed_excel_version()
    train = ".".join(version.split(".")[:2])
    if train != SUPPORTED_EXCEL_TRAIN:
        raise ReleaseError(_ReleaseProblem.EXCEL_TRAIN, version, SUPPORTED_EXCEL_TRAIN)
    return version


def _write_new_json(path: Path, value: Mapping[str, object]) -> None:
    destination = path.expanduser().resolve()
    if destination.exists():
        raise EvidenceValidationError(_ReleaseProblem.EVIDENCE_EXISTS, destination)
    if destination.parent == ROOT or ROOT in destination.parents:
        raise EvidenceValidationError(_ReleaseProblem.EVIDENCE_IN_REPOSITORY)
    if not destination.parent.is_dir():
        raise FileNotFoundError(destination.parent)

    payload = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    descriptor, name = tempfile.mkstemp(
        prefix=f".{destination.stem}.",
        suffix=destination.suffix or ".json",
        dir=destination.parent,
    )
    temporary = Path(name)
    try:
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        temporary.replace(destination)
    except OSError as write_error:
        cleanup_error: OSError | None = None
        if temporary.exists():
            try:
                temporary.unlink()
            except OSError as error:
                cleanup_error = error
        if cleanup_error is not None:
            raise ReleaseError(
                _ReleaseProblem.EVIDENCE_WRITE_CLEANUP,
                type(cleanup_error).__name__,
                cleanup_error,
            ) from write_error
        raise


def _evidence_template(*, excel_version: str) -> dict[str, object]:
    release_digest = _file_digest(RELEASE_WORKBOOK)
    if _file_digest(DIST_WORKBOOK) != release_digest:
        raise EvidenceValidationError(_ReleaseProblem.WORKBOOK_COPIES)
    return {
        "schema": EVIDENCE_SCHEMA,
        "source_sha256": _source_digest(),
        "release_workbook_sha256": release_digest,
        "excel_version": excel_version,
        "tested_at": "REPLACE_WITH_TIMEZONE_AWARE_ISO_8601",
        "tester": "REPLACE_WITH_TESTER_NAME",
        "checks": dict.fromkeys(sorted(LIVE_CHECKS), "TODO"),
        "export": {
            "path": "REPLACE_WITH_ABSOLUTE_MARKDOWN_PATH",
            "sha256": "REPLACE_WITH_SHA256",
            "directory_entries_before": [],
            "directory_entries_after": [],
        },
    }


def _source_gates() -> None:
    _run_gates(_source_gate_matrix())


def _artifact_gates() -> None:
    gates = _artifact_gate_matrix()
    _validate_artifact_gate_matrix(gates)
    _run_gates(gates)


def _build_release() -> None:
    _run_gate(_python_gate("release build", "-m", "build"))


def _load_evidence(path: Path) -> dict[str, object]:
    evidence_path = path.expanduser().resolve()
    if not evidence_path.is_file():
        raise FileNotFoundError(evidence_path)
    try:
        value = json.loads(evidence_path.read_text(encoding="utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise EvidenceValidationError(_ReleaseProblem.INVALID_JSON, evidence_path) from error
    if not isinstance(value, dict):
        raise EvidenceTypeError(_ReleaseProblem.ROOT_NOT_OBJECT)
    return cast("dict[str, object]", value)


def _require_exact_keys(
    value: Mapping[str, object],
    expected: Set[str],
    label: str,
) -> None:
    actual = set(value)
    missing = expected - actual
    extra = actual - expected
    if missing or extra:
        raise EvidenceValidationError(
            _ReleaseProblem.KEYS_DIFFER,
            label,
            sorted(missing),
            sorted(extra),
        )


def _parse_test_time(value: object) -> datetime:
    if not isinstance(value, str):
        raise EvidenceTypeError(_ReleaseProblem.TESTED_AT_TYPE)
    try:
        tested_at = datetime.fromisoformat(value)
    except ValueError as error:
        raise EvidenceValidationError(_ReleaseProblem.TESTED_AT_VALUE) from error
    if tested_at.tzinfo is None or tested_at.utcoffset() is None:
        raise EvidenceValidationError(_ReleaseProblem.TESTED_AT_OFFSET)
    now = datetime.now(UTC)
    tested_utc = tested_at.astimezone(UTC)
    if tested_utc > now + timedelta(minutes=5):
        raise EvidenceValidationError(_ReleaseProblem.TESTED_AT_FUTURE)
    if now - tested_utc > EVIDENCE_MAX_AGE:
        raise EvidenceValidationError(_ReleaseProblem.TESTED_AT_STALE)
    return tested_at


def _require_string_list(value: object, label: str) -> list[str]:
    if not isinstance(value, list) or not all(isinstance(name, str) for name in value):
        raise EvidenceTypeError(_ReleaseProblem.EXPORT_ENTRIES_TYPE, label)
    return cast("list[str]", value)


def _validate_export_path(value: object) -> Path:
    if not isinstance(value, str):
        raise EvidenceTypeError(_ReleaseProblem.EXPORT_PATH_TYPE)
    export_path = Path(value).expanduser()
    if not export_path.is_absolute() or export_path.suffix.lower() != ".md":
        raise EvidenceValidationError(_ReleaseProblem.EXPORT_PATH_VALUE)
    resolved = export_path.resolve()
    if not resolved.is_file():
        raise FileNotFoundError(resolved)
    return resolved


def _validate_markdown(export: Mapping[str, object]) -> None:
    _require_exact_keys(
        export,
        {"path", "sha256", "directory_entries_before", "directory_entries_after"},
        "export evidence",
    )
    export_path = _validate_export_path(export["path"])
    expected_digest = export["sha256"]
    if (
        not isinstance(expected_digest, str)
        or len(expected_digest) != SHA256_HEX_LENGTH
        or re.fullmatch(r"[0-9a-f]+", expected_digest) is None
    ):
        raise EvidenceValidationError(_ReleaseProblem.EXPORT_DIGEST)
    if _file_digest(export_path) != expected_digest:
        raise EvidenceValidationError(_ReleaseProblem.EXPORT_DIGEST_MISMATCH)

    before = _require_string_list(export["directory_entries_before"], "before entries")
    after = _require_string_list(export["directory_entries_after"], "after entries")
    if len(before) != len(set(before)) or len(after) != len(set(after)):
        raise EvidenceValidationError(_ReleaseProblem.EXPORT_ENTRIES_DUPLICATE)
    if set(after) - set(before) != {export_path.name} or set(before) - set(after):
        raise EvidenceValidationError(_ReleaseProblem.EXPORT_DELTA)
    actual_entries = sorted(path.name for path in export_path.parent.iterdir())
    if sorted(after) != actual_entries:
        raise EvidenceValidationError(_ReleaseProblem.EXPORT_CHANGED)

    _validate_markdown_content(export_path)


def _validate_markdown_content(export_path: Path) -> None:
    try:
        markdown = export_path.read_text(encoding="utf-8", errors="strict")
    except UnicodeDecodeError as error:
        raise EvidenceValidationError(_ReleaseProblem.EXPORT_UTF8) from error
    markdown = markdown.removeprefix("\ufeff")
    for heading in MARKDOWN_HEADINGS:
        if heading not in markdown:
            raise EvidenceValidationError(_ReleaseProblem.EXPORT_HEADING, heading)
    if re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", markdown):
        raise EvidenceValidationError(_ReleaseProblem.EXPORT_CONTROL)
    if markdown.count("| ---") < MIN_MARKDOWN_TABLES:
        raise EvidenceValidationError(_ReleaseProblem.EXPORT_TABLES)


def _validate_excel_version(value: object) -> None:
    if not isinstance(value, str) or re.fullmatch(r"\d+(?:\.\d+)+", value) is None:
        raise EvidenceTypeError(_ReleaseProblem.EXCEL_VERSION_TYPE)
    current = _require_supported_excel_version()
    if current != value:
        raise EvidenceValidationError(_ReleaseProblem.EXCEL_CHANGED, value, current)


def _validate_checks(value: object) -> None:
    if not isinstance(value, dict):
        raise EvidenceTypeError(_ReleaseProblem.CHECKS_TYPE)
    checks = cast("dict[str, object]", value)
    _require_exact_keys(checks, LIVE_CHECKS, "live macro checks")
    failed = sorted(name for name, result in checks.items() if result != "PASS")
    if failed:
        raise EvidenceValidationError(_ReleaseProblem.CHECKS_FAILED, ", ".join(failed))


def _validate_evidence(evidence: Mapping[str, object]) -> None:
    _require_exact_keys(
        evidence,
        {
            "schema",
            "source_sha256",
            "release_workbook_sha256",
            "excel_version",
            "tested_at",
            "tester",
            "checks",
            "export",
        },
        "macro evidence",
    )
    if evidence["schema"] != EVIDENCE_SCHEMA:
        raise EvidenceValidationError(_ReleaseProblem.SCHEMA, EVIDENCE_SCHEMA)
    if evidence["source_sha256"] != _source_digest():
        raise EvidenceValidationError(_ReleaseProblem.SOURCE_DIGEST)
    release_digest = _file_digest(RELEASE_WORKBOOK)
    if evidence["release_workbook_sha256"] != release_digest:
        raise EvidenceValidationError(_ReleaseProblem.WORKBOOK_DIGEST)
    if _file_digest(DIST_WORKBOOK) != release_digest:
        raise EvidenceValidationError(_ReleaseProblem.WORKBOOK_COPIES)
    if not QA_WORKBOOK.is_file():
        raise FileNotFoundError(QA_WORKBOOK)

    _validate_excel_version(evidence["excel_version"])
    _parse_test_time(evidence["tested_at"])
    tester = evidence["tester"]
    if not isinstance(tester, str) or not tester.strip() or tester.startswith("REPLACE_"):
        raise EvidenceValidationError(_ReleaseProblem.TESTER)
    _validate_checks(evidence["checks"])

    export = evidence["export"]
    if not isinstance(export, dict):
        raise EvidenceTypeError(_ReleaseProblem.EXPORT_TYPE)
    _validate_markdown(cast("dict[str, object]", export))


def _prepare(template_path: Path) -> int:
    excel_version = _require_supported_excel_version()
    _source_gates()
    _build_release()
    _artifact_gates()
    destination = template_path.expanduser().resolve()
    _write_new_json(destination, _evidence_template(excel_version=excel_version))
    LOGGER.info(
        "\nRELEASE BLOCKED - automated QA passed, but fresh live macro evidence is required."
    )
    LOGGER.info("Evidence template: %s", destination)
    LOGGER.info(
        "Complete every check in desktop Excel, record the export evidence, then run:\n  %s",
        shlex.join((
            sys.executable,
            "-m",
            "build.qa.release",
            "final",
            "--macro-evidence",
            str(destination),
        )),
    )
    return PREPARED_BUT_BLOCKED


def _final(evidence_path: Path) -> int:
    evidence = _load_evidence(evidence_path)
    _validate_evidence(evidence)
    _source_gates()
    _artifact_gates()
    _validate_evidence(evidence)
    LOGGER.info("\nRELEASE QA PASS - automated gates and bound live macro evidence are current")
    return 0


def main(argv: list[str] | None = None) -> int:
    """Run the requested release phase and return its process status.

    Returns:
        Status two for a successful prepare or zero for a successful final.

    """
    arguments = _parser().parse_args(sys.argv[1:] if argv is None else argv)
    if arguments.phase == "prepare":
        return _prepare(arguments.evidence_template)
    return _final(arguments.macro_evidence)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    raise SystemExit(main())
