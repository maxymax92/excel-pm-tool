"""Logging and exit-status reporting for structural workbook QA."""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pathlib import Path

LOGGER = logging.getLogger(__name__)


def report(path: Path, failures: list[str]) -> int:
    """Log structural-QA failures or success and return the process status.

    Returns:
        One when failures exist, otherwise zero.

    """
    if failures:
        LOGGER.error("QA: %s FAILURES", len(failures))
        for failure in failures:
            LOGGER.error("  FAIL %s", failure)
        return 1
    LOGGER.info("QA: ALL PASS (%s)", path.name)
    return 0
