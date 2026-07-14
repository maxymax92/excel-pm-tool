"""Strict real-Excel execution helpers shared by the QA suites."""

import asyncio
from dataclasses import dataclass
from enum import Enum
from pathlib import Path

from ..core.design import MACRO_ACTIONS
from ..core.package_style import patch_workbook_package
from ..paths import AUTOMATION

RECALC_SUCCESS = "RECALCULATED sentinel=All"


class _RecalculationProblem(Enum):
    INCOMPLETE_EDIT = "sheet, cell, and value must be supplied together"
    TIMEOUT = "Excel recalculation exceeded {} seconds"
    TIMEOUT_CLEANUP = "Excel recalculation timed out and process cleanup failed: {}: {}"
    NONZERO_EXIT = "Excel recalculation failed for {} (exit {}): {}"
    ERROR_STREAM = "Excel recalculation wrote an unexpected error stream for {}: {}"
    BAD_SENTINEL = "Excel recalculation returned {!r}; expected {!r}"
    ACTION_COUNT = "{} contains {} macro actions after recalculation; expected {}"


class _IncompleteEditError(ValueError):
    def __init__(self) -> None:
        super().__init__(_RecalculationProblem.INCOMPLETE_EDIT.value)


class _ExcelRecalculationError(RuntimeError):
    def __init__(self, problem: _RecalculationProblem, *details: object) -> None:
        super().__init__(problem.value.format(*details))


@dataclass(frozen=True, slots=True)
class _AutomationResult:
    returncode: int
    stdout: str
    stderr: str


async def _run_automation(command: tuple[str, ...], timeout_seconds: float) -> _AutomationResult:
    process = await asyncio.create_subprocess_exec(
        *command,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    try:
        async with asyncio.timeout(timeout_seconds):
            stdout_bytes, stderr_bytes = await process.communicate()
    except TimeoutError as timeout_error:
        try:
            process.kill()
            await process.wait()
        except (ChildProcessError, OSError, ProcessLookupError) as cleanup_error:
            raise _ExcelRecalculationError(
                _RecalculationProblem.TIMEOUT_CLEANUP,
                type(cleanup_error).__name__,
                cleanup_error,
            ) from timeout_error
        raise _ExcelRecalculationError(
            _RecalculationProblem.TIMEOUT,
            timeout_seconds,
        ) from timeout_error

    return _AutomationResult(
        returncode=process.returncode or 0,
        stdout=stdout_bytes.decode("utf-8").strip(),
        stderr=stderr_bytes.decode("utf-8").strip(),
    )


def recalculate(
    path: str | Path,
    *,
    sheet: str | None = None,
    cell: str | None = None,
    value: str | None = None,
    timeout: float = 630,
) -> str:
    """Recalculate and save a workbook in Excel, with an optional literal edit.

    Returns:
        The validated automation success sentinel.

    Raises:
        FileNotFoundError: If the workbook path does not identify a file.
        _IncompleteEditError: If only part of an optional cell edit is supplied.
        _ExcelRecalculationError: If automation, sentinel or macro-action checks fail.

    """
    workbook = Path(path).expanduser().resolve()
    if not workbook.is_file():
        raise FileNotFoundError(workbook)

    edit = (sheet, cell, value)
    supplied = tuple(part is not None for part in edit)
    if any(supplied) and not all(supplied):
        raise _IncompleteEditError

    command = (
        "osascript",
        str(AUTOMATION / "excel_recalc.applescript"),
        str(workbook),
    )
    if all(supplied):
        command += (str(sheet), str(cell), str(value))

    completed = asyncio.run(_run_automation(command, timeout))
    if completed.returncode != 0:
        diagnostic = completed.stderr or completed.stdout or "no diagnostic"
        raise _ExcelRecalculationError(
            _RecalculationProblem.NONZERO_EXIT,
            workbook.name,
            completed.returncode,
            diagnostic,
        )
    if completed.stderr:
        raise _ExcelRecalculationError(
            _RecalculationProblem.ERROR_STREAM,
            workbook.name,
            completed.stderr,
        )
    if completed.stdout != RECALC_SUCCESS:
        raise _ExcelRecalculationError(
            _RecalculationProblem.BAD_SENTINEL,
            completed.stdout,
            RECALC_SUCCESS,
        )
    styled = patch_workbook_package(workbook, calculation_complete=True)
    expected_actions = len(MACRO_ACTIONS) if workbook.suffix.lower() == ".xlsm" else 0
    if styled.button_count != expected_actions:
        raise _ExcelRecalculationError(
            _RecalculationProblem.ACTION_COUNT,
            workbook.name,
            styled.button_count,
            expected_actions,
        )
    return completed.stdout
