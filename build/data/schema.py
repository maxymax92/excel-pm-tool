"""Authored-data schema shared by export, snapshot, injection and migration.

The registry names every durable workbook value: kind ``I`` (input), kind
``V`` (VBA-stamped) and kind ``S`` (system-managed source identity) table
columns plus the Config settings band. Kind ``F`` formula columns are never
exported or injected; the rebuilt structure recomputes them. The schema
fingerprint records exactly which shape a snapshot was exported under, so
cross-version column mapping stays explicit.
"""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass

from ..spec import config
from ..spec.capacity import CONFIG_ROWS, DATA_ROWS
from ..spec.items import ITEMS_COLUMNS, RAID_COLUMNS

SCHEMA_FORMAT = 1
AUTHORED_KINDS = frozenset({"I", "S", "V"})
VALUE_TYPES = frozenset({"text", "int", "date", "bool"})

# The exact copy identifying a shipped delete-me example row in any cell.
EXAMPLE_MARKER = "EXAMPLE — delete this row"


@dataclass(frozen=True, kw_only=True, slots=True)
class ColumnSchema:
    """One authored column of an exported workbook table."""

    name: str
    value_type: str
    kind: str


@dataclass(frozen=True, kw_only=True, slots=True)
class TableSchema:
    """Location, authored columns and capacity for one exported table."""

    sheet: str
    table: str
    key: str | None
    columns: tuple[ColumnSchema, ...]
    workbook_columns: tuple[str, ...]
    capacity: int

    @property
    def column_names(self) -> tuple[str, ...]:
        """Return the authored column names in workbook order.

        Returns:
            The authored column names.

        """
        return tuple(column.name for column in self.columns)


def _spec_value_type(fmt: object) -> str:
    if fmt == "date":
        return "date"
    if fmt == "int":
        return "int"
    return "text"


def _data_table(
    *,
    sheet: str,
    table: str,
    key: str,
    specs: list[dict[str, object]],
) -> TableSchema:
    columns = tuple(
        ColumnSchema(
            name=str(spec["name"]),
            value_type=_spec_value_type(spec["fmt"]),
            kind=str(spec["kind"]),
        )
        for spec in specs
        if spec["kind"] in AUTHORED_KINDS
    )
    return TableSchema(
        sheet=sheet,
        table=table,
        key=key,
        columns=columns,
        workbook_columns=tuple(str(spec["name"]) for spec in specs),
        capacity=DATA_ROWS,
    )


def _config_table(table: str, key: str, columns: tuple[tuple[str, str], ...]) -> TableSchema:
    schema_columns = tuple(
        ColumnSchema(name=name, value_type=value_type, kind="I") for name, value_type in columns
    )
    return TableSchema(
        sheet="Config",
        table=table,
        key=key,
        columns=schema_columns,
        workbook_columns=tuple(name for name, _value_type in columns),
        capacity=CONFIG_ROWS,
    )


DATA_TABLES = (
    _data_table(sheet="Items", table="tblItems", key="ID", specs=ITEMS_COLUMNS),
    _data_table(sheet="RAID", table="tblRAID", key="RaidID", specs=RAID_COLUMNS),
    _config_table(
        "tblStatuses",
        "Status",
        (
            ("Status", "text"),
            ("IsActive", "bool"),
            ("IsDone", "bool"),
            ("IsCancelled", "bool"),
            ("IsDeleted", "bool"),
        ),
    ),
    _config_table("tblTypes", "Type", (("Type", "text"), ("Level", "int"))),
    _config_table("tblPriorities", "Priority", (("Priority", "text"),)),
    _config_table("tblTeams", "Team", (("Team", "text"),)),
    _config_table(
        "tblRaidTypes",
        "RaidType",
        (("RaidType", "text"), ("IsAlert", "bool"), ("IsDecision", "bool")),
    ),
    _config_table(
        "tblRaidStatuses",
        "RaidStatus",
        (("RaidStatus", "text"), ("IsClosed", "bool"), ("IsDeleted", "bool")),
    ),
    _config_table("tblSeverity", "Severity", (("Severity", "text"), ("MinScore", "int"))),
    _config_table("tblDeliveryHealth", "Delivery Health", (("Delivery Health", "text"),)),
    _config_table(
        "tblPeople",
        "Person",
        (("Person", "text"), ("Role", "text"), ("Team", "text")),
    ),
)

TABLES_BY_NAME = {table_schema.table: table_schema for table_schema in DATA_TABLES}

# Settings shape captured from the pristine spec at import time, before any
# in-process injection swaps the config module state.
SETTINGS_DEFAULTS = {name: value for name, value, _description in config.SETTINGS}
SETTINGS_DESCRIPTIONS = {name: description for name, _value, description in config.SETTINGS}
SETTINGS_TYPES = {
    name: ("text" if isinstance(value, str) else "int") for name, value in SETTINGS_DEFAULTS.items()
}

# Data table -> the Config identifier settings the VBA uses to assign its IDs.
ID_COUNTERS = {
    "tblItems": ("cfgItemIDPrefix", "cfgNextItemID"),
    "tblRAID": ("cfgRaidIDPrefix", "cfgNextRaidID"),
}


def key_problems(
    key: str | None,
    rows: tuple[dict[str, object], ...],
) -> tuple[int | None, tuple[str, ...]]:
    """Find the first blank key and every duplicate key value in one table.

    Returns:
        The first blank-key row number (1-based) or None, and the sorted
        duplicate key values.

    """
    if key is None:
        return None, ()
    blank_row: int | None = None
    seen: dict[object, int] = {}
    labels: dict[object, object] = {}
    for index, row in enumerate(rows, start=1):
        key_value = row.get(key)
        if key_value in {None, ""}:
            if blank_row is None:
                blank_row = index
            continue
        normalized = key_value.casefold() if isinstance(key_value, str) else key_value
        labels.setdefault(normalized, key_value)
        seen[normalized] = seen.get(normalized, 0) + 1
    duplicates = tuple(sorted(str(labels[value]) for value, count in seen.items() if count > 1))
    return blank_row, duplicates


def schema_fingerprint() -> str:
    """Return the stable digest of the authored-data schema.

    Returns:
        The SHA-256 hex digest of the canonical schema description.

    """
    description = {
        "format": SCHEMA_FORMAT,
        "tables": [
            {
                "sheet": table_schema.sheet,
                "table": table_schema.table,
                "key": table_schema.key,
                "capacity": table_schema.capacity,
                "columns": [
                    [column.name, column.value_type, column.kind] for column in table_schema.columns
                ],
                "workbook_columns": list(table_schema.workbook_columns),
            }
            for table_schema in DATA_TABLES
        ],
        "settings": sorted(SETTINGS_TYPES.items()),
    }
    canonical = json.dumps(description, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()
