"""Export, rebuild-with-injection, verify and publish one live workbook.

Migration reuses the release pipeline untouched: the injected build goes
through the same authoring path, desktop-Excel recalculation, semantic
comparison and rollback-capable publication as ``python -m build``. The
snapshot is persisted before anything else happens and the replaced workbook
is kept in a bounded backup ring, so no step can endanger authored data.
"""

from __future__ import annotations

import logging
from datetime import UTC, datetime
from enum import Enum
from pathlib import Path
from typing import TYPE_CHECKING

from ..automation.workspace import excel_working_directory
from ..paths import DIST, ROOT
from ..pipeline import (
    build_one,
    publish_transaction,
    recalculate_stage,
    require_current_vba,
    require_semantic_preservation,
)
from .export import export_workbook
from .inject import injected_source, validate_snapshot
from .schema import DATA_TABLES
from .snapshot import RING_LIMIT, write_snapshot

if TYPE_CHECKING:
    from .export import ExportResult
    from .inject import Reconciliation
    from .snapshot import Snapshot

LOGGER = logging.getLogger(__name__)

DEFAULT_WORKBOOK = ROOT / "PM_Workbook.xlsm"
BACKUP_DIR = DIST / "backups"


class _MigrateProblem(Enum):
    MISSING_WORKBOOK = "workbook {} does not exist"
    NOT_MACRO = "migration targets the macro-enabled artifact; {} is not an .xlsm workbook"
    LOCKED = (
        "{} appears open in Excel (lock file {}); close every workbook in Excel "
        "or remove the stale lock, then re-run"
    )
    TABLE_MISMATCH = (
        "migrated {} does not hold the snapshot's rows and values "
        "({} rows built, {} expected); no artifact was published"
    )
    SETTING_MISMATCH = "migrated Config {} is {!r}; expected {!r}"


class MigrationError(RuntimeError):
    """Report an unusable migration target or a failed verification."""

    def __init__(self, problem: _MigrateProblem, *details: object) -> None:
        """Create an error from a stable diagnostic template."""
        super().__init__(problem.value.format(*details))


def _require_workbook(workbook: Path) -> Path:
    source = workbook.expanduser().resolve()
    if not source.is_file():
        raise MigrationError(_MigrateProblem.MISSING_WORKBOOK, source)
    if source.suffix.lower() != ".xlsm":
        raise MigrationError(_MigrateProblem.NOT_MACRO, source.name)
    lock = _workbook_lock(source)
    if lock.exists():
        raise MigrationError(_MigrateProblem.LOCKED, source.name, lock)
    return source


def _workbook_lock(source: Path) -> Path:
    return source.parent / f"~${source.name}"


def _verify_populated(package: Path, snapshot: Snapshot, reconciliation: Reconciliation) -> None:
    built = export_workbook(package)
    for table_schema in DATA_TABLES:
        built_rows = list(built.snapshot.tables.get(table_schema.table, ()))
        expected_rows = list(snapshot.tables.get(table_schema.table, ()))
        if built_rows != expected_rows:
            raise MigrationError(
                _MigrateProblem.TABLE_MISMATCH,
                table_schema.table,
                len(built_rows),
                len(expected_rows),
            )
    for name, expected in reconciliation.settings.items():
        if built.snapshot.settings.get(name) != expected:
            raise MigrationError(
                _MigrateProblem.SETTING_MISMATCH,
                name,
                built.snapshot.settings.get(name),
                expected,
            )


def _backup_plan(workbook: Path) -> tuple[Path, tuple[Path, ...]]:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(UTC).strftime("%Y%m%dT%H%M%SZ")
    target = BACKUP_DIR / f"{workbook.stem}-{stamp}{workbook.suffix}"
    aged = sorted({*BACKUP_DIR.glob(f"{workbook.stem}-*{workbook.suffix}"), target})
    stale = tuple(aged[: max(0, len(aged) - RING_LIMIT)])
    return target, stale


def _log_export(result: ExportResult) -> None:
    for table, count in sorted(result.skipped_examples.items()):
        LOGGER.info("%s: skipped %s marked example row(s)", table, count)
    for table, columns in sorted(result.added_columns.items()):
        LOGGER.info("%s: new column(s) %s start blank", table, ", ".join(columns))
    for table, columns in sorted(result.unknown_columns.items()):
        LOGGER.info("%s: exported out-of-schema column(s) %s", table, ", ".join(columns))
    for note in result.notes:
        LOGGER.info("%s", note)


def export_command(workbook: str | Path = DEFAULT_WORKBOOK) -> Path:
    """Export one workbook's authored data into the snapshot ring.

    Returns:
        The written snapshot path.

    """
    source = _require_workbook(Path(workbook))
    result = export_workbook(source)
    snapshot_path = write_snapshot(result.snapshot)
    _log_export(result)
    LOGGER.info("exported %s", source)
    LOGGER.info("snapshot: %s", snapshot_path)
    return snapshot_path


def rebuild_and_publish(source: Path, snapshot: Snapshot, reconciliation: Reconciliation) -> Path:
    """Inject, rebuild, recalculate, verify and transactionally publish one workbook.

    Returns:
        The backup path holding the replaced workbook.

    """
    with excel_working_directory("pm-migrate-") as stage:
        raw_dir = stage / "raw"
        calculated_dir = stage / "calculated"
        raw_dir.mkdir()
        calculated_dir.mkdir()
        raw_workbook = raw_dir / source.name
        calculated_workbook = calculated_dir / source.name

        with injected_source(snapshot, reconciliation):
            build_one(raw_workbook, with_vba=True)
        recalculate_stage(raw_workbook, calculated_workbook)
        require_semantic_preservation(((raw_workbook, calculated_workbook),))
        _verify_populated(
            calculated_workbook,
            reconciliation.normalized_snapshot,
            reconciliation,
        )

        backup, stale_backups = _backup_plan(source)
        publish_transaction(
            {backup: source, source: calculated_workbook},
            removals=stale_backups,
            expected_digests={source: snapshot.workbook_digest},
            required_absent=(_workbook_lock(source),),
        )
    return backup


def migrate_command(workbook: str | Path = DEFAULT_WORKBOOK) -> None:
    """Rebuild one workbook from the current source with its own data."""
    source = _require_workbook(Path(workbook))
    require_current_vba()

    result = export_workbook(source)
    snapshot_path = write_snapshot(result.snapshot)
    reconciliation = validate_snapshot(result.snapshot)

    _log_export(result)
    for line in reconciliation.lines():
        LOGGER.info("%s", line)
    LOGGER.info("snapshot: %s", snapshot_path)

    backup = rebuild_and_publish(source, result.snapshot, reconciliation)
    LOGGER.info("backup: %s", backup)
    LOGGER.info("migrated %s from the current source", source)
