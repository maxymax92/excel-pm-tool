"""Shared state, artifact and workbook-error helpers for scenario QA."""

from __future__ import annotations

import re
from contextlib import contextmanager
from copy import deepcopy
from enum import Enum
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from collections.abc import Iterator
    from pathlib import Path
    from types import ModuleType

    from openpyxl import Workbook

from ..automation.workspace import excel_working_directory
from ..spec import examples

ERROR_VALUE_RE = re.compile(
    r"^#(REF!|NAME\?|VALUE!|DIV/0!|N/A|SPILL!|CALC!|NULL!|NUM!|FIELD!|"
    r"BLOCKED!|GETTING_DATA|CIRCULAR!|UNKNOWN!)"
)
EXAMPLE_ATTRIBUTES = ("ITEMS_EXAMPLES", "PEOPLE_EXAMPLES", "RAID_EXAMPLES")


class _CleanupResource(Enum):
    EXAMPLES = "example-fixture restoration"


class _ScenarioCleanupError(RuntimeError):
    def __init__(self, resource: _CleanupResource, cleanup_error: BaseException) -> None:
        super().__init__(
            f"scenario failed and {resource.value} cleanup also failed: "
            f"{type(cleanup_error).__name__}: {cleanup_error}"
        )


def workbook_error_cells(workbook: Workbook, *, limit: int | None = None) -> list[str]:
    """Return coordinates and values for calculated Excel errors.

    Returns:
        Error coordinates and cached values, capped at ``limit`` when supplied.

    """
    errors: list[str] = []
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and ERROR_VALUE_RE.match(cell.value):
                    errors.append(f"{worksheet.title}!{cell.coordinate}={cell.value}")
                    if limit is not None and len(errors) >= limit:
                        return errors
    return errors


@contextmanager
def temporary_examples() -> Iterator[ModuleType]:
    """Restore the in-process example fixtures after a scenario build.

    Yields:
        The mutable example-fixture module.

    Raises:
        _ScenarioCleanupError: If a scenario fails and fixture restoration also fails.

    """
    original = {name: deepcopy(getattr(examples, name)) for name in EXAMPLE_ATTRIBUTES}
    try:
        yield examples
    except BaseException as operation_error:
        try:
            for name, value in original.items():
                setattr(examples, name, value)
        except (AttributeError, TypeError) as cleanup_error:
            raise _ScenarioCleanupError(
                _CleanupResource.EXAMPLES,
                cleanup_error,
            ) from operation_error
        raise
    else:
        for name, value in original.items():
            setattr(examples, name, value)


@contextmanager
def temporary_workbook(prefix: str) -> Iterator[Path]:
    """Provide a unique QA workbook path and remove it after use.

    Yields:
        A unique path beneath Excel's private automation workspace.

    """
    with excel_working_directory(f"{prefix}.") as directory:
        yield directory / f"{prefix}.xlsx"
