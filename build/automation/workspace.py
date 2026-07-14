"""Provide a prompt-free disposable workspace inside Excel's macOS sandbox."""

from __future__ import annotations

import shutil
import sys
import tempfile
from contextlib import contextmanager
from enum import Enum
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from collections.abc import Iterator

EXCEL_DOCUMENTS = (
    Path.home() / "Library" / "Containers" / "com.microsoft.Excel" / "Data" / "Documents"
)
EXCEL_AUTOMATION_WORKSPACE = EXCEL_DOCUMENTS / "PMWorkbookAutomation"


class _WorkspaceProblem(Enum):
    DOCUMENTS = "Excel sandbox Documents directory is unavailable: {}"
    CREATE = "cannot create Excel automation workspace {}: {}: {}"
    TEMPORARY = "cannot create a disposable Excel workspace in {}: {}: {}"
    COPY = "cannot copy {} into the Excel automation workspace: {}: {}"
    CLEANUP = "cannot clean Excel automation workspace {}: {}: {}"


class ExcelWorkspaceError(RuntimeError):
    """Report an unavailable or uncleanable Excel automation workspace."""

    def __init__(self, problem: _WorkspaceProblem, *details: object) -> None:
        """Create one stable workspace diagnostic."""
        super().__init__(problem.value.format(*details))


def excel_workspace_root() -> Path:
    """Return Excel's private automation root, creating its leaf directory.

    Returns:
        The stable workspace beneath Excel's sandboxed Documents directory.

    Raises:
        ExcelWorkspaceError: If the supported Mac Excel container is unavailable.

    """
    if not EXCEL_DOCUMENTS.is_dir():
        raise ExcelWorkspaceError(_WorkspaceProblem.DOCUMENTS, EXCEL_DOCUMENTS)
    try:
        EXCEL_AUTOMATION_WORKSPACE.mkdir(mode=0o700, exist_ok=True)
    except OSError as error:
        raise ExcelWorkspaceError(
            _WorkspaceProblem.CREATE,
            EXCEL_AUTOMATION_WORKSPACE,
            type(error).__name__,
            error,
        ) from error
    return EXCEL_AUTOMATION_WORKSPACE


@contextmanager
def excel_working_directory(prefix: str) -> Iterator[Path]:
    """Yield one isolated directory that Excel can access without a grant prompt.

    Yields:
        A unique directory beneath :data:`EXCEL_AUTOMATION_WORKSPACE`.

    Raises:
        ExcelWorkspaceError: If creation or cleanup fails.

    """
    root = excel_workspace_root()
    try:
        directory = Path(tempfile.mkdtemp(prefix=prefix, dir=root))
    except OSError as error:
        raise ExcelWorkspaceError(
            _WorkspaceProblem.TEMPORARY,
            root,
            type(error).__name__,
            error,
        ) from error
    try:
        yield directory
    finally:
        active_error = sys.exception()
        try:
            shutil.rmtree(directory)
        except OSError as cleanup_error:
            if active_error is not None:
                active_error.add_note(
                    "Excel automation workspace cleanup also failed: "
                    f"{type(cleanup_error).__name__}: {cleanup_error}"
                )
            else:
                raise ExcelWorkspaceError(
                    _WorkspaceProblem.CLEANUP,
                    directory,
                    type(cleanup_error).__name__,
                    cleanup_error,
                ) from cleanup_error


@contextmanager
def excel_workbook_copy(source: Path, *, prefix: str) -> Iterator[Path]:
    """Yield an exact disposable workbook copy inside Excel's sandbox.

    Yields:
        The copied workbook path.

    Raises:
        FileNotFoundError: If the source workbook does not exist.
        ExcelWorkspaceError: If copying or cleanup fails.

    """
    workbook = source.expanduser().resolve()
    if not workbook.is_file():
        raise FileNotFoundError(workbook)
    with excel_working_directory(prefix) as directory:
        destination = directory / workbook.name
        try:
            shutil.copy2(workbook, destination)
        except OSError as error:
            raise ExcelWorkspaceError(
                _WorkspaceProblem.COPY,
                workbook,
                type(error).__name__,
                error,
            ) from error
        yield destination


__all__ = [
    "EXCEL_AUTOMATION_WORKSPACE",
    "ExcelWorkspaceError",
    "excel_workbook_copy",
    "excel_working_directory",
    "excel_workspace_root",
]
