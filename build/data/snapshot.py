"""Persisted authored-data snapshots with a bounded retention ring.

A snapshot is the durable JSON interchange for everything a user authored in
one workbook: table rows keyed by column name plus Config settings values. It
is written before any migration touches an artifact, so a failed rebuild can
never endanger data, and it doubles as the import format for external sources.
"""

from __future__ import annotations

import json
import os
import tempfile
from dataclasses import dataclass
from datetime import date
from enum import Enum
from pathlib import Path

from ..paths import DIST
from .schema import SETTINGS_TYPES, TABLES_BY_NAME

SNAPSHOT_FORMAT = 1
SNAPSHOT_DIR = DIST / "snapshots"
RING_LIMIT = 20


class _SnapshotProblem(Enum):
    FORMAT = "snapshot {} declares unsupported format {!r}; expected {}"
    MALFORMED = "snapshot {} is malformed: {}"
    VALUE = "snapshot {} has an invalid value at {}: expected {}, got {!r}"


class SnapshotError(RuntimeError):
    """Report an unreadable or badly typed snapshot document."""

    def __init__(self, problem: _SnapshotProblem, *details: object) -> None:
        """Create an error from a stable diagnostic template."""
        super().__init__(problem.value.format(*details))


@dataclass(frozen=True, kw_only=True, slots=True)
class Snapshot:
    """One complete authored-data capture from a workbook."""

    schema_fingerprint: str
    workbook_digest: str
    exported_at: str
    settings: dict[str, object]
    tables: dict[str, tuple[dict[str, object], ...]]


def _encode_value(value: object) -> object:
    if isinstance(value, date):
        return value.isoformat()
    return value


def _document(snapshot: Snapshot) -> dict[str, object]:
    return {
        "format": SNAPSHOT_FORMAT,
        "schema_fingerprint": snapshot.schema_fingerprint,
        "workbook_digest": snapshot.workbook_digest,
        "exported_at": snapshot.exported_at,
        "settings": {name: _encode_value(value) for name, value in snapshot.settings.items()},
        "tables": {
            table: [{column: _encode_value(value) for column, value in row.items()} for row in rows]
            for table, rows in snapshot.tables.items()
        },
    }


def _decode_typed(source: Path, location: str, value: object, value_type: str) -> object:
    if value_type == "date":
        if isinstance(value, str):
            try:
                return date.fromisoformat(value)
            except ValueError as error:
                raise SnapshotError(
                    _SnapshotProblem.VALUE,
                    source.name,
                    location,
                    "an ISO date",
                    value,
                ) from error
        raise SnapshotError(_SnapshotProblem.VALUE, source.name, location, "an ISO date", value)
    if value_type == "bool":
        if isinstance(value, bool):
            return value
        raise SnapshotError(_SnapshotProblem.VALUE, source.name, location, "a boolean", value)
    if value_type == "int":
        if isinstance(value, int) and not isinstance(value, bool):
            return value
        raise SnapshotError(_SnapshotProblem.VALUE, source.name, location, "an integer", value)
    if isinstance(value, str):
        return value
    raise SnapshotError(_SnapshotProblem.VALUE, source.name, location, "text", value)


def _decode_row(
    source: Path,
    table: str,
    index: int,
    row: object,
) -> dict[str, object]:
    if not isinstance(row, dict):
        raise SnapshotError(
            _SnapshotProblem.MALFORMED,
            source.name,
            f"{table} row {index} is not an object",
        )
    table_schema = TABLES_BY_NAME.get(table)
    column_types = (
        {column.name: column.value_type for column in table_schema.columns}
        if table_schema is not None
        else {}
    )
    decoded: dict[str, object] = {}
    for column, value in row.items():
        location = f"{table} row {index} column {column!r}"
        value_type = column_types.get(column)
        # Columns unknown to the current schema stay raw; injection reports
        # them as orphans before any value is used.
        decoded[column] = (
            value if value_type is None else _decode_typed(source, location, value, value_type)
        )
    return decoded


def _decode_settings(source: Path, settings: object) -> dict[str, object]:
    if not isinstance(settings, dict):
        raise SnapshotError(_SnapshotProblem.MALFORMED, source.name, "settings is not an object")
    decoded: dict[str, object] = {}
    for name, value in settings.items():
        value_type = SETTINGS_TYPES.get(name)
        location = f"setting {name!r}"
        decoded[name] = (
            value if value_type is None else _decode_typed(source, location, value, value_type)
        )
    return decoded


def read_snapshot(path: str | Path) -> Snapshot:
    """Read and type-check one persisted snapshot document.

    Returns:
        The decoded snapshot.

    Raises:
        SnapshotError: If the document is malformed or badly typed.

    """
    source = Path(path)
    try:
        document = json.loads(source.read_text(encoding="utf-8"))
    except json.JSONDecodeError as error:
        raise SnapshotError(_SnapshotProblem.MALFORMED, source.name, error) from error
    if not isinstance(document, dict):
        raise SnapshotError(_SnapshotProblem.MALFORMED, source.name, "document is not an object")
    if document.get("format") != SNAPSHOT_FORMAT:
        raise SnapshotError(
            _SnapshotProblem.FORMAT,
            source.name,
            document.get("format"),
            SNAPSHOT_FORMAT,
        )

    tables_field = document.get("tables")
    if not isinstance(tables_field, dict):
        raise SnapshotError(_SnapshotProblem.MALFORMED, source.name, "tables is not an object")
    tables: dict[str, tuple[dict[str, object], ...]] = {}
    for table, rows in tables_field.items():
        if not isinstance(rows, list):
            raise SnapshotError(
                _SnapshotProblem.MALFORMED,
                source.name,
                f"{table} rows are not an array",
            )
        tables[table] = tuple(
            _decode_row(source, table, index, row) for index, row in enumerate(rows, start=1)
        )

    for field in ("schema_fingerprint", "workbook_digest", "exported_at"):
        if not isinstance(document.get(field), str):
            raise SnapshotError(
                _SnapshotProblem.MALFORMED,
                source.name,
                f"{field} is not a string",
            )

    return Snapshot(
        schema_fingerprint=str(document["schema_fingerprint"]),
        workbook_digest=str(document["workbook_digest"]),
        exported_at=str(document["exported_at"]),
        settings=_decode_settings(source, document.get("settings")),
        tables=tables,
    )


def prune_ring(directory: Path, pattern: str, *, limit: int = RING_LIMIT) -> None:
    """Keep only the newest ``limit`` files matching one ring pattern."""
    aged = sorted(directory.glob(pattern))
    for stale in aged[: max(0, len(aged) - limit)]:
        stale.unlink()


def atomic_write_json(target: Path, document: dict[str, object]) -> None:
    """Write one JSON document through a same-directory temporary file."""
    payload = json.dumps(document, ensure_ascii=False, indent=2, sort_keys=True)
    descriptor, name = tempfile.mkstemp(
        prefix=f".{target.stem}.",
        suffix=".json",
        dir=target.parent,
    )
    temporary = Path(name)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
            stream.write(payload + "\n")
            stream.flush()
            os.fsync(stream.fileno())
        temporary.replace(target)
    finally:
        if temporary.exists():
            temporary.unlink()


def write_snapshot(snapshot: Snapshot, *, directory: str | Path = SNAPSHOT_DIR) -> Path:
    """Atomically persist one snapshot and prune the retention ring.

    Returns:
        The written snapshot path.

    """
    target_dir = Path(directory)
    target_dir.mkdir(parents=True, exist_ok=True)
    stamp = snapshot.exported_at.replace("-", "").replace(":", "")
    target = target_dir / f"pm-data-{stamp}-{snapshot.workbook_digest[:12]}.json"
    atomic_write_json(target, _document(snapshot))
    prune_ring(target_dir, "pm-data-*.json")
    return target
