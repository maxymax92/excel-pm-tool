"""Validate a snapshot and inject it in-process into a fresh build.

Injection never edits a finished workbook: it swaps the example and Config
module state exactly the way the ship-demo scenario does, so the standard
pipeline writes the authored rows while composing a fresh package. Hard gates
halt before any build when data would be lost or the build would fail; value
findings the workbook itself flags in red are reported, never silently fixed.
"""

from __future__ import annotations

from contextlib import contextmanager
from dataclasses import dataclass
from enum import Enum
from itertools import starmap
from typing import TYPE_CHECKING

from ..core.design import HIERARCHY
from ..spec import config, examples
from .schema import (
    ID_COUNTERS,
    SETTINGS_DEFAULTS,
    SETTINGS_DESCRIPTIONS,
    TABLES_BY_NAME,
    key_problems,
    schema_fingerprint,
)
from .snapshot import Snapshot

if TYPE_CHECKING:
    from collections.abc import Callable, Iterator


class _InjectProblem(Enum):
    ORPHAN_TABLES = (
        "snapshot holds tables outside the current schema: {}; "
        "resolve them before injection so no data is dropped"
    )
    ORPHAN_COLUMNS = (
        "snapshot {} rows hold columns outside the current schema: {}; "
        "resolve them before injection so no data is dropped"
    )
    ORPHAN_SETTINGS = (
        "snapshot holds settings outside the current schema: {}; "
        "resolve them before injection so no data is dropped"
    )
    CAPACITY = "{} holds {} rows; the workbook capacity is {}"
    BLANK_KEY = "{} row {} has no {}"
    DUPLICATE_KEY = "{} holds duplicate {} values: {}"
    UNKNOWN_TYPE = "Items rows use types missing from tblTypes: {}; the build would fail"
    INVALID_LEVEL = "tblTypes levels for used types must be 1-{}: {}"
    MISSING_TYPE = "Items rows have no Type: {}; assign one in the workbook, then re-run"
    MISSING_TITLE = "Items rows have no Title: {}; add one in the workbook, then re-run"
    DELETED_AMBIGUOUS = "{} has multiple deletion-role candidates: {}"
    DELETED_CAPACITY = "{} cannot append Deleted because Config capacity is {} rows"
    RECONCILIATION_SOURCE = "reconciliation does not belong to workbook digest {}"
    RESTORE = "injected build failed and source restoration also failed: {}: {}"


class InjectError(RuntimeError):
    """Report a snapshot the current structure cannot absorb."""

    def __init__(self, problem: _InjectProblem, *details: object) -> None:
        """Create an error from a stable diagnostic template."""
        super().__init__(problem.value.format(*details))


@dataclass(frozen=True, kw_only=True, slots=True)
class Reconciliation:
    """Validated injection inputs and everything worth telling the user."""

    fingerprint_matches: bool
    row_counts: dict[str, int]
    settings: dict[str, object]
    defaulted_settings: tuple[str, ...]
    counter_bumps: dict[str, tuple[int, int]]
    empty_tables: tuple[str, ...]
    warnings: tuple[str, ...]
    adjustments: tuple[str, ...]
    normalized_snapshot: Snapshot

    def lines(self) -> tuple[str, ...]:
        """Return the human reconciliation report.

        Returns:
            One formatted line per reconciliation fact.

        """
        lines = [f"{table}: {count} rows" for table, count in sorted(self.row_counts.items())]
        if not self.fingerprint_matches:
            lines.append(
                "snapshot was exported under a different schema; column mapping was re-validated"
            )
        lines.extend(
            f"{table}: absent from the snapshot; injected empty" for table in self.empty_tables
        )
        lines.extend(
            f"{name}: not in the snapshot; default {SETTINGS_DEFAULTS[name]!r} applied"
            for name in self.defaulted_settings
        )
        lines.extend(
            f"{name}: bumped {before} -> {after} to stay above every existing ID"
            for name, (before, after) in sorted(self.counter_bumps.items())
        )
        lines.extend(f"warning: {warning}" for warning in self.warnings)
        lines.extend(f"adjustment: {adjustment}" for adjustment in self.adjustments)
        return tuple(lines)


_DELETION_TABLES = {
    "tblStatuses": (
        "Status",
        {"IsActive": False, "IsDone": True, "IsCancelled": True, "IsDeleted": True},
    ),
    "tblRaidStatuses": (
        "RaidStatus",
        {"IsClosed": True, "IsDeleted": True},
    ),
}


def _normalize_deletion_table(
    table: str,
    rows: tuple[dict[str, object], ...],
) -> tuple[tuple[dict[str, object], ...], tuple[str, ...]]:
    label_column, required_roles = _DELETION_TABLES[table]
    normalized = [dict(row) for row in rows]
    explicit = {index for index, row in enumerate(normalized) if row.get("IsDeleted") is True}
    named = {
        index
        for index, row in enumerate(normalized)
        if str(row.get(label_column, "")).strip().casefold() == "deleted"
    }
    candidates = explicit | named
    if len(candidates) > 1:
        labels = ", ".join(str(normalized[index].get(label_column, "")) for index in candidates)
        raise InjectError(_InjectProblem.DELETED_AMBIGUOUS, table, labels)

    adjustments: list[str] = []
    missing_flags = sum("IsDeleted" not in row for row in normalized)
    if missing_flags:
        adjustments.append(f"{table}: added IsDeleted=False to {missing_flags} existing row(s)")

    appended = not candidates
    if appended:
        if len(normalized) >= TABLES_BY_NAME[table].capacity:
            raise InjectError(
                _InjectProblem.DELETED_CAPACITY,
                table,
                TABLES_BY_NAME[table].capacity,
            )
        target = len(normalized)
        normalized.append({label_column: "Deleted", **required_roles})
        adjustments.append(f"{table}: appended Deleted with the required deletion roles")
    else:
        target = next(iter(candidates))

    for index, row in enumerate(normalized):
        row["IsDeleted"] = index == target
    target_row = normalized[target]
    before = {name: target_row.get(name) for name in required_roles}
    target_row.update(required_roles)
    if before != required_roles and not appended:
        label = target_row.get(label_column, "Deleted")
        adjustments.append(f"{table} {label}: normalized the required deletion roles")
    return tuple(normalized), tuple(adjustments)


def _normalize_deletion_roles(snapshot: Snapshot) -> tuple[Snapshot, tuple[str, ...]]:
    tables = {table: tuple(dict(row) for row in rows) for table, rows in snapshot.tables.items()}
    adjustments: list[str] = []
    for table in _DELETION_TABLES:
        normalized, table_adjustments = _normalize_deletion_table(
            table,
            tables.get(table, ()),
        )
        tables[table] = normalized
        adjustments.extend(table_adjustments)
    return (
        Snapshot(
            schema_fingerprint=schema_fingerprint(),
            workbook_digest=snapshot.workbook_digest,
            exported_at=snapshot.exported_at,
            settings=dict(snapshot.settings),
            tables=tables,
        ),
        tuple(adjustments),
    )


def _require_schema_tables(snapshot: Snapshot) -> None:
    orphans = sorted(set(snapshot.tables) - set(TABLES_BY_NAME))
    if orphans:
        raise InjectError(_InjectProblem.ORPHAN_TABLES, ", ".join(orphans))


def _require_table_shape(table: str, rows: tuple[dict[str, object], ...]) -> None:
    table_schema = TABLES_BY_NAME[table]
    known = set(table_schema.column_names)
    orphans = sorted({column for row in rows for column in row} - known)
    if orphans:
        raise InjectError(_InjectProblem.ORPHAN_COLUMNS, table, ", ".join(orphans))
    if len(rows) > table_schema.capacity:
        raise InjectError(
            _InjectProblem.CAPACITY,
            table,
            len(rows),
            table_schema.capacity,
        )
    blank_row, duplicates = key_problems(table_schema.key, rows)
    if blank_row is not None:
        raise InjectError(_InjectProblem.BLANK_KEY, table, blank_row, table_schema.key)
    if duplicates:
        if table in _DELETION_TABLES:
            label_column, _required_roles = _DELETION_TABLES[table]
            candidates = [
                row
                for row in rows
                if row.get("IsDeleted") is True
                or str(row.get(label_column, "")).strip().casefold() == "deleted"
            ]
            if len(candidates) > 1:
                labels = ", ".join(str(row.get(label_column, "")) for row in candidates)
                raise InjectError(_InjectProblem.DELETED_AMBIGUOUS, table, labels)
        raise InjectError(
            _InjectProblem.DUPLICATE_KEY,
            table,
            table_schema.key,
            ", ".join(duplicates),
        )


def _require_buildable_items(snapshot: Snapshot) -> None:
    items = snapshot.tables.get("tblItems", ())
    untyped = sorted(
        str(row.get("ID", f"row {index}"))
        for index, row in enumerate(items, start=1)
        if row.get("Type") in {None, ""}
    )
    if untyped:
        raise InjectError(_InjectProblem.MISSING_TYPE, ", ".join(untyped))
    untitled = sorted(
        str(row.get("ID", f"row {index}"))
        for index, row in enumerate(items, start=1)
        if row.get("Title") in {None, ""}
    )
    if untitled:
        raise InjectError(_InjectProblem.MISSING_TITLE, ", ".join(untitled))


def _require_usable_types(snapshot: Snapshot) -> None:
    items = snapshot.tables.get("tblItems", ())
    used = {
        str(row["Type"]).casefold(): str(row["Type"])
        for row in items
        if row.get("Type") not in {None, ""}
    }
    levels = {
        str(row.get("Type", "")).casefold(): row.get("Level")
        for row in snapshot.tables.get("tblTypes", ())
    }
    unknown = sorted(label for folded, label in used.items() if folded not in levels)
    if unknown:
        raise InjectError(_InjectProblem.UNKNOWN_TYPE, ", ".join(unknown))
    invalid = sorted(
        f"{label} (Level {levels[folded]!r})"
        for folded, label in used.items()
        if levels[folded] not in HIERARCHY
    )
    if invalid:
        raise InjectError(_InjectProblem.INVALID_LEVEL, max(HIERARCHY), ", ".join(invalid))


def _merge_settings(snapshot: Snapshot) -> tuple[dict[str, object], tuple[str, ...]]:
    orphans = sorted(set(snapshot.settings) - set(SETTINGS_DEFAULTS))
    if orphans:
        raise InjectError(_InjectProblem.ORPHAN_SETTINGS, ", ".join(orphans))
    merged = dict(SETTINGS_DEFAULTS)
    merged.update(snapshot.settings)
    defaulted = tuple(name for name in SETTINGS_DEFAULTS if name not in snapshot.settings)
    return merged, defaulted


def _bump_counters(
    snapshot: Snapshot,
    settings: dict[str, object],
    warnings: list[str],
) -> dict[str, tuple[int, int]]:
    bumps: dict[str, tuple[int, int]] = {}
    for table, (prefix_setting, counter_setting) in ID_COUNTERS.items():
        table_schema = TABLES_BY_NAME[table]
        prefix = str(settings[prefix_setting])
        highest = 0
        for row in snapshot.tables.get(table, ()):
            key_value = str(row.get(table_schema.key, ""))
            suffix = key_value.removeprefix(prefix)
            if suffix == key_value or not suffix.isdecimal():
                warnings.append(
                    f"{table} {key_value}: does not follow the {prefix_setting} "
                    f"prefix {prefix!r}; the workbook flags it and the counter ignores it"
                )
                continue
            highest = max(highest, int(suffix))
        counter = int(str(settings[counter_setting]))
        if counter <= highest:
            settings[counter_setting] = highest + 1
            bumps[counter_setting] = (counter, highest + 1)
    return bumps


def _domain(snapshot: Snapshot, table: str, column: str) -> set[str]:
    return {
        str(row[column]).casefold()
        for row in snapshot.tables.get(table, ())
        if row.get(column) is not None
    }


def _value_warnings(snapshot: Snapshot) -> list[str]:
    checks = (
        ("tblItems", "Status", _domain(snapshot, "tblStatuses", "Status")),
        ("tblItems", "Priority", _domain(snapshot, "tblPriorities", "Priority")),
        ("tblItems", "Owner", _domain(snapshot, "tblPeople", "Person")),
        ("tblItems", "Delivery Health", _domain(snapshot, "tblDeliveryHealth", "Delivery Health")),
        ("tblItems", "Parent", _domain(snapshot, "tblItems", "ID")),
        ("tblRAID", "Type", _domain(snapshot, "tblRaidTypes", "RaidType")),
        ("tblRAID", "Status", _domain(snapshot, "tblRaidStatuses", "RaidStatus")),
        ("tblRAID", "Owner", _domain(snapshot, "tblPeople", "Person")),
        ("tblRAID", "RelatedID", _domain(snapshot, "tblItems", "ID")),
        ("tblPeople", "Team", _domain(snapshot, "tblTeams", "Team")),
    )
    warnings: list[str] = []
    for table, column, domain in checks:
        table_schema = TABLES_BY_NAME[table]
        for row in snapshot.tables.get(table, ()):
            value = row.get(column)
            if value in {None, ""} or str(value).casefold() in domain:
                continue
            label = row.get(table_schema.key or "", "?")
            warnings.append(
                f"{table} {label}: {column} {value!r} is outside the configured list; "
                "the workbook flags it in red"
            )
    return warnings


def validate_snapshot(snapshot: Snapshot) -> Reconciliation:
    """Run every injection gate and assemble the reconciliation report.

    Returns:
        The validated injection inputs and report material.

    """
    original_fingerprint = snapshot.schema_fingerprint
    original_tables = set(snapshot.tables)
    empty_critical = tuple(
        table
        for table in ("tblStatuses", "tblTypes", "tblDeliveryHealth", "tblRaidStatuses")
        if not snapshot.tables.get(table, ())
    )
    _require_schema_tables(snapshot)
    for table in TABLES_BY_NAME:
        _require_table_shape(table, snapshot.tables.get(table, ()))
    snapshot, adjustments = _normalize_deletion_roles(snapshot)
    _require_buildable_items(snapshot)
    _require_usable_types(snapshot)

    settings, defaulted = _merge_settings(snapshot)
    warnings = _value_warnings(snapshot)
    counter_bumps = _bump_counters(snapshot, settings, warnings)
    empty_tables = tuple(sorted(set(TABLES_BY_NAME) - original_tables))
    warnings.extend(
        f"{table} is empty; the workbook loses its workflow behaviour" for table in empty_critical
    )

    normalized_snapshot = Snapshot(
        schema_fingerprint=schema_fingerprint(),
        workbook_digest=snapshot.workbook_digest,
        exported_at=snapshot.exported_at,
        settings=dict(settings),
        tables=snapshot.tables,
    )
    return Reconciliation(
        fingerprint_matches=original_fingerprint == schema_fingerprint(),
        row_counts={table: len(snapshot.tables.get(table, ())) for table in TABLES_BY_NAME},
        settings=settings,
        defaulted_settings=defaulted,
        counter_bumps=counter_bumps,
        empty_tables=empty_tables,
        warnings=tuple(warnings),
        adjustments=adjustments,
        normalized_snapshot=normalized_snapshot,
    )


@contextmanager
def swapped_module_state(
    replacements: dict[tuple[object, str], object],
    on_restore_failure: Callable[[BaseException], Exception],
) -> Iterator[None]:
    """Swap module attributes for the context body and always restore them.

    When a failing body's state restoration also fails, the exception built
    by ``on_restore_failure`` is raised from the original error.

    Yields:
        Nothing; the swapped state lives for the context body.

    """
    original = {
        (module, attribute): getattr(module, attribute) for module, attribute in replacements
    }
    for (module, attribute), value in replacements.items():
        setattr(module, attribute, value)

    def _restore() -> None:
        for (module, attribute), value in original.items():
            setattr(module, attribute, value)

    try:
        yield
    except BaseException as operation_error:
        try:
            _restore()
        except (AttributeError, TypeError) as cleanup_error:
            raise on_restore_failure(cleanup_error) from operation_error
        raise
    else:
        _restore()


def _config_rows(
    snapshot: Snapshot,
    table: str,
    columns: tuple[tuple[str, object], ...],
) -> list[tuple[object, ...]]:
    return [tuple(starmap(row.get, columns)) for row in snapshot.tables.get(table, ())]


def _config_values(snapshot: Snapshot, table: str, column: str) -> list[object]:
    return [row.get(column, "") for row in snapshot.tables.get(table, ())]


@contextmanager
def injected_source(snapshot: Snapshot, reconciliation: Reconciliation) -> Iterator[None]:
    """Swap the authored source state for the duration of one build.

    Yields:
        Nothing; the swapped state lives for the context body.

    Raises:
        InjectError: If the reconciliation belongs to another workbook snapshot.

    """
    effective = reconciliation.normalized_snapshot
    if effective.workbook_digest != snapshot.workbook_digest:
        raise InjectError(_InjectProblem.RECONCILIATION_SOURCE, snapshot.workbook_digest)
    replacements: dict[tuple[object, str], object] = {
        (examples, "ITEMS_EXAMPLES"): [dict(row) for row in effective.tables.get("tblItems", ())],
        (examples, "PEOPLE_EXAMPLES"): [dict(row) for row in effective.tables.get("tblPeople", ())],
        (examples, "RAID_EXAMPLES"): [dict(row) for row in effective.tables.get("tblRAID", ())],
        (config, "SETTINGS"): [
            (name, reconciliation.settings[name], SETTINGS_DESCRIPTIONS[name])
            for name in SETTINGS_DEFAULTS
        ],
        (config, "STATUSES"): _config_rows(
            effective,
            "tblStatuses",
            (
                ("Status", ""),
                ("IsActive", False),
                ("IsDone", False),
                ("IsCancelled", False),
                ("IsDeleted", False),
            ),
        ),
        (config, "TYPES"): _config_rows(effective, "tblTypes", (("Type", ""), ("Level", 0))),
        (config, "PRIORITIES"): _config_values(effective, "tblPriorities", "Priority"),
        (config, "TEAMS"): _config_values(effective, "tblTeams", "Team"),
        (config, "RAID_TYPES"): _config_rows(
            effective,
            "tblRaidTypes",
            (("RaidType", ""), ("IsAlert", False), ("IsDecision", False)),
        ),
        (config, "RAID_STATUSES"): _config_rows(
            effective,
            "tblRaidStatuses",
            (("RaidStatus", ""), ("IsClosed", False), ("IsDeleted", False)),
        ),
        (config, "SEVERITY"): _config_rows(
            effective,
            "tblSeverity",
            (("Severity", ""), ("MinScore", 0)),
        ),
        (config, "DELIVERY_HEALTH"): _config_values(
            effective,
            "tblDeliveryHealth",
            "Delivery Health",
        ),
    }

    def _restore_error(cleanup_error: BaseException) -> Exception:
        return InjectError(
            _InjectProblem.RESTORE,
            type(cleanup_error).__name__,
            cleanup_error,
        )

    with swapped_module_state(replacements, _restore_error):
        yield
