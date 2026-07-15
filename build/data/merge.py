"""Deterministic identity, relationship and field merge for agent change sets."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from datetime import date
from typing import Literal

from .contract import ITEM_WRITABLE_FIELDS, RAID_WRITABLE_FIELDS
from .diagnostics import Diagnostic
from .schema import ID_COUNTERS, SETTINGS_DEFAULTS, TABLES_BY_NAME
from .snapshot import Snapshot
from .validation import compare_snapshot_findings

type Entity = Literal["item", "raid"]
type MergeAction = Literal["create", "update", "mark_deleted", "noop"]

_ENTITY_TABLE = {"item": "tblItems", "raid": "tblRAID"}
_ENTITY_KEY = {"item": "ID", "raid": "RaidID"}
_WRITABLE = {
    "item": frozenset(ITEM_WRITABLE_FIELDS),
    "raid": frozenset(RAID_WRITABLE_FIELDS),
}
_REFERENCE_FIELDS = {
    "item": frozenset({"Parent", "BlockedBy"}),
    "raid": frozenset({"RelatedID"}),
}
_DATE_FIELDS = {
    "item": frozenset({"Start", "Due"}),
    "raid": frozenset({"NextReview"}),
}


@dataclass(frozen=True, kw_only=True, slots=True)
class FieldDiff:
    """One exact authored-field transition."""

    field: str
    before: object | None
    after: object | None

    def as_dict(self) -> dict[str, object | None]:
        """Return one JSON-compatible field diff.

        Returns:
            The field name and before/after values.

        """
        return {
            "field": self.field,
            "before": _json_value(self.before),
            "after": _json_value(self.after),
        }


@dataclass(frozen=True, kw_only=True, slots=True)
class OperationDiff:
    """One resolved operation and its deterministic field transitions."""

    operation_id: str
    entity: Entity
    action: MergeAction
    workbook_id: str
    diffs: tuple[FieldDiff, ...]

    def as_dict(self) -> dict[str, object]:
        """Return the public JSON representation.

        Returns:
            The resolved operation result.

        """
        return {
            "operation_id": self.operation_id,
            "entity": self.entity,
            "action": self.action,
            "workbook_id": self.workbook_id,
            "diffs": [diff.as_dict() for diff in self.diffs],
        }


@dataclass(frozen=True, kw_only=True, slots=True)
class MergeResult:
    """The complete atomic merge outcome."""

    snapshot: Snapshot
    operations: tuple[OperationDiff, ...]
    diagnostics: tuple[Diagnostic, ...]

    @property
    def valid(self) -> bool:
        """Return whether no merge error was found."""
        return not any(diagnostic.severity == "error" for diagnostic in self.diagnostics)

    @property
    def changed(self) -> bool:
        """Return whether a valid merge changes authored state."""
        return self.valid and any(operation.action != "noop" for operation in self.operations)


@dataclass(slots=True)
class _Pending:
    operation_index: int
    operation_id: str
    entity: Entity
    row_index: int
    workbook_id: str
    before: dict[str, object]
    created: bool
    operation: Mapping[str, object]


@dataclass(slots=True)
class _Indexes:
    ids: dict[Entity, dict[str, list[int]]]
    sources: dict[Entity, dict[tuple[str, str], list[int]]]


@dataclass(frozen=True, slots=True)
class _OperationRequest:
    operation: Mapping[str, object]
    entity: Entity
    index: int
    operation_id: str


@dataclass(frozen=True, slots=True)
class _IdentityLookup:
    request: _OperationRequest
    requested_id: str | None
    source: tuple[str, str] | None
    id_matches: tuple[int, ...]
    active_source_matches: tuple[int, ...]
    deleted_source_matches: tuple[int, ...]
    rows: Sequence[dict[str, object]]
    deletion_label: str


@dataclass(frozen=True, slots=True)
class _Resolution:
    row_index: int | None
    source: tuple[str, str] | None
    diagnostics: tuple[Diagnostic, ...] = ()


@dataclass(frozen=True, slots=True)
class _ReferenceRequest:
    reference: object
    index: int
    operation_id: str
    field: str


@dataclass(frozen=True, slots=True)
class _ReferenceContext:
    indexes: _Indexes
    rows: Mapping[Entity, list[dict[str, object]]]
    client_refs: Mapping[str, _Pending]
    deletion_roles: _DeletionRoles


@dataclass(frozen=True, slots=True)
class _DeletionRoles:
    item: str
    raid: str

    def for_entity(self, entity: Entity) -> str:
        """Return the configured deletion label for an entity.

        Returns:
            The exact current Config label.

        """
        return self.item if entity == "item" else self.raid


@dataclass(frozen=True, slots=True)
class _ItemStatusRole:
    active: bool
    done: bool
    cancelled: bool
    deleted: bool


@dataclass(frozen=True, slots=True)
class _LifecycleConfig:
    item_statuses: Mapping[str, _ItemStatusRole]
    raid_statuses: Mapping[str, bool]
    blocked_health: str | None


def _same_config_text(value: object, configured: str) -> bool:
    """Match Config-backed text with VBA's ``vbTextCompare`` semantics.

    Returns:
        True when the value is text equal to the configured label ignoring case.

    """
    return isinstance(value, str) and value.casefold() == configured.casefold()


def _json_value(value: object | None) -> object | None:
    return value.isoformat() if isinstance(value, date) else value


def _diagnostic(
    code: str,
    message: str,
    hint: str,
    *,
    pointer: str,
    operation_id: str | None,
) -> Diagnostic:
    return Diagnostic(
        code=code,
        severity="error",
        phase="plan",
        pointer=pointer,
        operation_id=operation_id,
        message=message,
        hint=hint,
    )


def _operation_pointer(index: int, suffix: str = "") -> str:
    return f"/operations/{index}{suffix}"


def _entity(value: object) -> Entity | None:
    if value == "item":
        return "item"
    if value == "raid":
        return "raid"
    return None


def _source_pair(value: object) -> tuple[str, str] | None:
    if not isinstance(value, Mapping):
        return None
    namespace = value.get("namespace")
    record_id = value.get("record_id")
    if not isinstance(namespace, str) or not isinstance(record_id, str):
        return None
    if not namespace or not record_id:
        return None
    return namespace, record_id


def _row_source(row: Mapping[str, object]) -> tuple[str, str] | None:
    namespace = row.get("Source")
    record_id = row.get("Source ID")
    if isinstance(namespace, str) and namespace and isinstance(record_id, str) and record_id:
        return namespace, record_id
    return None


def _build_indexes(rows: Mapping[Entity, list[dict[str, object]]]) -> _Indexes:
    ids: dict[Entity, dict[str, list[int]]] = {"item": {}, "raid": {}}
    sources: dict[Entity, dict[tuple[str, str], list[int]]] = {"item": {}, "raid": {}}
    for entity, entity_rows in rows.items():
        key = _ENTITY_KEY[entity]
        for index, row in enumerate(entity_rows):
            workbook_id = row.get(key)
            if isinstance(workbook_id, str) and workbook_id:
                ids[entity].setdefault(workbook_id.casefold(), []).append(index)
            source = _row_source(row)
            if source is not None:
                sources[entity].setdefault(source, []).append(index)
    return _Indexes(ids=ids, sources=sources)


def _deletion_label(
    tables: Mapping[str, list[dict[str, object]]],
    entity: Entity,
) -> tuple[str, Diagnostic | None]:
    table = "tblStatuses" if entity == "item" else "tblRaidStatuses"
    label_column = "Status" if entity == "item" else "RaidStatus"
    matches = [row for row in tables.get(table, ()) if row.get("IsDeleted") is True]
    if len(matches) == 1 and isinstance(matches[0].get(label_column), str):
        return str(matches[0][label_column]), None
    return "", _diagnostic(
        "config.deleted_role",
        f"{table} must contain exactly one nonblank IsDeleted role.",
        "Reconcile legacy Config deletion roles before planning changes.",
        pointer=f"/config/{table}",
        operation_id=None,
    )


def _deletion_roles(
    tables: Mapping[str, list[dict[str, object]]],
) -> tuple[_DeletionRoles, tuple[Diagnostic, ...]]:
    item, item_error = _deletion_label(tables, "item")
    raid, raid_error = _deletion_label(tables, "raid")
    diagnostics = tuple(error for error in (item_error, raid_error) if error is not None)
    return _DeletionRoles(item=item, raid=raid), diagnostics


def _is_deleted(
    row: Mapping[str, object],
    entity: Entity,
    roles: _DeletionRoles,
) -> bool:
    return _same_config_text(row.get("Status"), roles.for_entity(entity))


def _partition_source_matches(
    matches: Sequence[int],
    rows: Sequence[dict[str, object]],
    entity: Entity,
    roles: _DeletionRoles,
) -> tuple[tuple[int, ...], tuple[int, ...]]:
    active = tuple(index for index in matches if not _is_deleted(rows[index], entity, roles))
    deleted = tuple(index for index in matches if _is_deleted(rows[index], entity, roles))
    return active, deleted


def _active_source_diagnostics(
    indexes: _Indexes,
    rows: Mapping[Entity, list[dict[str, object]]],
    roles: _DeletionRoles,
) -> tuple[Diagnostic, ...]:
    diagnostics: list[Diagnostic] = []
    for entity in ("item", "raid"):
        for source, matches in sorted(indexes.sources[entity].items()):
            active, _deleted = _partition_source_matches(
                matches,
                rows[entity],
                entity,
                roles,
            )
            if len(active) <= 1:
                continue
            diagnostics.append(
                _diagnostic(
                    "identity.duplicate_active_source",
                    f"{entity} source identity {source!r} appears on multiple active rows.",
                    "Keep at most one non-Deleted row for each Source and Source ID pair.",
                    pointer=f"/records/{entity}",
                    operation_id=None,
                )
            )
    return tuple(diagnostics)


def _lifecycle_config(
    tables: Mapping[str, list[dict[str, object]]],
) -> _LifecycleConfig:
    item_statuses = {
        str(row["Status"]).casefold(): _ItemStatusRole(
            active=row.get("IsActive") is True,
            done=row.get("IsDone") is True,
            cancelled=row.get("IsCancelled") is True,
            deleted=row.get("IsDeleted") is True,
        )
        for row in tables.get("tblStatuses", ())
        if isinstance(row.get("Status"), str) and row.get("Status")
    }
    raid_statuses = {
        str(row["RaidStatus"]).casefold(): row.get("IsClosed") is True
        for row in tables.get("tblRaidStatuses", ())
        if isinstance(row.get("RaidStatus"), str) and row.get("RaidStatus")
    }
    health_values = [
        str(row["Delivery Health"])
        for row in tables.get("tblDeliveryHealth", ())
        if isinstance(row.get("Delivery Health"), str) and row.get("Delivery Health")
    ]
    return _LifecycleConfig(
        item_statuses=item_statuses,
        raid_statuses=raid_statuses,
        blocked_health=health_values[-1] if health_values else None,
    )


def _is_blank(value: object | None) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _field_changed(pending: _Pending, row: Mapping[str, object], field: str) -> bool:
    return pending.before.get(field) != row.get(field)


def _operational_change(pending: _Pending, row: Mapping[str, object]) -> bool:
    return pending.created or any(
        _field_changed(pending, row, field) for field in _WRITABLE[pending.entity]
    )


def _apply_item_status_lifecycle(
    pending: _Pending,
    row: dict[str, object],
    effective_date: date,
    config: _LifecycleConfig,
) -> None:
    status_touched = (
        pending.created
        or _field_changed(pending, row, "Type")
        or _field_changed(pending, row, "Status")
    )
    if not status_touched:
        return
    status = row.get("Status")
    role = config.item_statuses.get(status.casefold()) if isinstance(status, str) else None
    if role is None:
        return
    if role.active and _is_blank(row.get("InProgressSince")):
        row["InProgressSince"] = effective_date
    if role.done and not role.cancelled and not role.deleted:
        if _is_blank(row.get("DoneDate")):
            row["DoneDate"] = effective_date
    else:
        row.pop("DoneDate", None)


def _apply_item_health_lifecycle(
    pending: _Pending,
    row: dict[str, object],
    effective_date: date,
    config: _LifecycleConfig,
) -> None:
    health_touched = (
        pending.created
        or _field_changed(pending, row, "Type")
        or _field_changed(pending, row, "Delivery Health")
    )
    if not health_touched or config.blocked_health is None:
        return
    if _same_config_text(row.get("Delivery Health"), config.blocked_health):
        if _is_blank(row.get("BlockedSince")):
            row["BlockedSince"] = effective_date
    else:
        row.pop("BlockedSince", None)


def _apply_item_lifecycle(
    pending: _Pending,
    row: dict[str, object],
    effective_date: date,
    config: _LifecycleConfig,
) -> None:
    if not _operational_change(pending, row):
        return
    if _is_blank(row.get("Created")):
        row["Created"] = effective_date
    _apply_item_status_lifecycle(pending, row, effective_date, config)
    _apply_item_health_lifecycle(pending, row, effective_date, config)
    if pending.created or _field_changed(pending, row, "Latest Status"):
        if _is_blank(row.get("Latest Status")):
            row.pop("LatestUpdateOn", None)
        else:
            row["LatestUpdateOn"] = effective_date
    row["Updated"] = effective_date


def _apply_raid_lifecycle(
    pending: _Pending,
    row: dict[str, object],
    effective_date: date,
    config: _LifecycleConfig,
) -> None:
    if not _operational_change(pending, row):
        return
    if _is_blank(row.get("Raised")):
        row["Raised"] = effective_date
    status_touched = (
        pending.created
        or _field_changed(pending, row, "Type")
        or _field_changed(pending, row, "Status")
    )
    status = row.get("Status")
    closed = config.raid_statuses.get(status.casefold()) if isinstance(status, str) else None
    if status_touched and closed is not None:
        if closed:
            if _is_blank(row.get("Closed")):
                row["Closed"] = effective_date
        else:
            row.pop("Closed", None)
    row["Updated"] = effective_date


def _apply_lifecycle(
    pending: _Pending,
    row: dict[str, object],
    effective_date: date,
    config: _LifecycleConfig,
) -> None:
    if pending.entity == "item":
        _apply_item_lifecycle(pending, row, effective_date, config)
    else:
        _apply_raid_lifecycle(pending, row, effective_date, config)


def _copy_tables(snapshot: Snapshot) -> dict[str, list[dict[str, object]]]:
    return {table: [dict(row) for row in rows] for table, rows in snapshot.tables.items()}


def _rows_by_entity(
    tables: dict[str, list[dict[str, object]]],
) -> dict[Entity, list[dict[str, object]]]:
    return {
        "item": tables.setdefault("tblItems", []),
        "raid": tables.setdefault("tblRAID", []),
    }


def _duplicate_source_diagnostics(
    operation: Mapping[str, object],
    entity: Entity,
    index: int,
    operation_id: str | None,
    seen: dict[tuple[Entity, str, str], int],
) -> tuple[Diagnostic, ...]:
    """Register one explicit source identity and report a repeated target.

    Returns:
        One duplicate-source diagnostic, or an empty tuple after registration.

    """
    identity = operation.get("identity")
    identity_map = identity if isinstance(identity, Mapping) else {}
    source = _source_pair(identity_map.get("source"))
    if source is None:
        return ()
    source_key = (entity, *source)
    if source_key not in seen:
        seen[source_key] = index
        return ()
    return (
        _diagnostic(
            "operation.duplicate_source_identity",
            f"Source identity {source!r} is targeted more than once for {entity}.",
            "Combine all work for one source record into a single operation.",
            pointer=_operation_pointer(index, "/identity/source"),
            operation_id=operation_id,
        ),
    )


def _preflight_operations(operations: Sequence[object]) -> list[Diagnostic]:
    diagnostics: list[Diagnostic] = []
    operation_ids: dict[str, int] = {}
    client_refs: dict[str, int] = {}
    source_identities: dict[tuple[Entity, str, str], int] = {}
    for index, raw in enumerate(operations):
        pointer = _operation_pointer(index)
        if not isinstance(raw, Mapping):
            diagnostics.append(
                _diagnostic(
                    "operation.invalid",
                    "Operation must be an object.",
                    "Use one upsert or mark_deleted object from the embedded schema.",
                    pointer=pointer,
                    operation_id=None,
                )
            )
            continue
        operation_id = raw.get("operation_id")
        operation_name = operation_id if isinstance(operation_id, str) else None
        if operation_name is not None and operation_name in operation_ids:
            diagnostics.append(
                _diagnostic(
                    "operation.duplicate_id",
                    f"Operation ID {operation_name!r} is repeated.",
                    "Give every operation a unique operation_id.",
                    pointer=f"{pointer}/operation_id",
                    operation_id=operation_name,
                )
            )
        elif operation_name is not None:
            operation_ids[operation_name] = index

        client_ref = raw.get("client_ref")
        if isinstance(client_ref, str) and client_ref in client_refs:
            diagnostics.append(
                _diagnostic(
                    "operation.duplicate_client_ref",
                    f"Client reference {client_ref!r} is repeated.",
                    "Give every batch-local client_ref a unique value.",
                    pointer=f"{pointer}/client_ref",
                    operation_id=operation_name,
                )
            )
        elif isinstance(client_ref, str):
            client_refs[client_ref] = index

        entity = _entity(raw.get("entity"))
        if entity is None:
            diagnostics.append(
                _diagnostic(
                    "operation.entity",
                    "Operation entity must be item or raid.",
                    "Use an entity allowed by the contract schema.",
                    pointer=f"{pointer}/entity",
                    operation_id=operation_name,
                )
            )
            continue
        diagnostics.extend(
            _duplicate_source_diagnostics(
                raw,
                entity,
                index,
                operation_name,
                source_identities,
            )
        )
        operation_kind = raw.get("op")
        if operation_kind not in {"upsert", "mark_deleted"}:
            diagnostics.append(
                _diagnostic(
                    "operation.unsupported",
                    "Operation must be upsert or mark_deleted.",
                    "Use one operation kind from the embedded contract schema.",
                    pointer=f"{pointer}/op",
                    operation_id=operation_name,
                )
            )
            continue
        if operation_kind == "upsert":
            diagnostics.extend(_field_diagnostics(raw, entity, index, operation_name))
    return diagnostics


def _field_diagnostics(
    operation: Mapping[str, object],
    entity: Entity,
    index: int,
    operation_id: str | None,
) -> list[Diagnostic]:
    diagnostics: list[Diagnostic] = []
    set_fields = operation.get("set", {})
    clear_fields = operation.get("clear", [])
    set_names = set(set_fields) if isinstance(set_fields, Mapping) else set()
    clear_names = set(clear_fields) if isinstance(clear_fields, list) else set()
    diagnostics.extend(
        _diagnostic(
            "field.not_writable",
            f"{field!r} is not an agent-writable {entity} field.",
            "Write only fields listed by describe; identity metadata is attached separately.",
            pointer=_operation_pointer(index, f"/set/{field}"),
            operation_id=operation_id,
        )
        for field in sorted((set_names | clear_names) - _WRITABLE[entity])
    )
    diagnostics.extend(
        _diagnostic(
            "field.conflict",
            f"{field!r} is present in both set and clear.",
            "Remove the field from either set or clear.",
            pointer=_operation_pointer(index, f"/set/{field}"),
            operation_id=operation_id,
        )
        for field in sorted(set_names & clear_names)
    )
    return diagnostics


def _matches(
    mapping: Mapping[object, list[int]],
    key: object | None,
) -> list[int]:
    return [] if key is None else list(mapping.get(key, ()))


def _identity_diagnostic(
    code: str,
    message: str,
    hint: str,
    pending_index: int,
    operation_id: str,
) -> Diagnostic:
    return _diagnostic(
        code,
        message,
        hint,
        pointer=_operation_pointer(pending_index, "/identity"),
        operation_id=operation_id,
    )


def _identity_lookup(
    request: _OperationRequest,
    indexes: _Indexes,
    rows: Mapping[Entity, list[dict[str, object]]],
    roles: _DeletionRoles,
) -> _IdentityLookup:
    identity = request.operation.get("identity")
    identity_map = identity if isinstance(identity, Mapping) else {}
    workbook_id = identity_map.get("workbook_id")
    requested_id = workbook_id if isinstance(workbook_id, str) and workbook_id else None
    source = _source_pair(identity_map.get("source"))
    id_matches = _matches(
        indexes.ids[request.entity],
        requested_id.casefold() if requested_id is not None else None,
    )
    source_matches = _matches(indexes.sources[request.entity], source)
    active, deleted = _partition_source_matches(
        source_matches,
        rows[request.entity],
        request.entity,
        roles,
    )
    return _IdentityLookup(
        request=request,
        requested_id=requested_id,
        source=source,
        id_matches=tuple(id_matches),
        active_source_matches=active,
        deleted_source_matches=deleted,
        rows=rows[request.entity],
        deletion_label=roles.for_entity(request.entity),
    )


def _resolution_error(lookup: _IdentityLookup, diagnostic: Diagnostic) -> _Resolution:
    return _Resolution(
        row_index=None,
        source=lookup.source,
        diagnostics=(diagnostic,),
    )


def _resolve_requested_id(lookup: _IdentityLookup) -> _Resolution:
    request = lookup.request
    if not lookup.id_matches:
        return _resolution_error(
            lookup,
            _identity_diagnostic(
                "identity.workbook_id_missing",
                f"Workbook ID {lookup.requested_id!r} does not exist.",
                "Use a current workbook ID from describe, or create by source identity only.",
                request.index,
                request.operation_id,
            ),
        )
    row_index = lookup.id_matches[0]
    if request.operation.get("op") == "upsert" and _same_config_text(
        lookup.rows[row_index].get("Status"), lookup.deletion_label
    ):
        return _resolution_error(
            lookup,
            _identity_diagnostic(
                "identity.deleted_history",
                "A Deleted historical row cannot be changed by workbook ID.",
                "Upsert by source identity to create a fresh row, or leave history unchanged.",
                request.index,
                request.operation_id,
            ),
        )
    active_conflicts = [index for index in lookup.active_source_matches if index != row_index]
    if active_conflicts:
        return _resolution_error(
            lookup,
            _identity_diagnostic(
                "identity.conflict",
                "Workbook ID and source identity resolve to different rows.",
                "Use the matching identity pair or correct the source record mapping.",
                request.index,
                request.operation_id,
            ),
        )
    existing_source = _row_source(lookup.rows[row_index])
    if lookup.source is not None and existing_source not in {None, lookup.source}:
        return _resolution_error(
            lookup,
            _identity_diagnostic(
                "identity.source_immutable",
                "The target row already carries a different source identity.",
                "Do not reassign an existing Source and Source ID pair.",
                request.index,
                request.operation_id,
            ),
        )
    return _Resolution(row_index=row_index, source=lookup.source)


def _resolve_source_only(lookup: _IdentityLookup) -> _Resolution:
    request = lookup.request
    if lookup.source is None:
        return _resolution_error(
            lookup,
            _identity_diagnostic(
                "identity.source_required",
                "A new row requires a source namespace and source record ID.",
                "Supply identity.source with both namespace and record_id.",
                request.index,
                request.operation_id,
            ),
        )
    if lookup.active_source_matches:
        return _Resolution(
            row_index=lookup.active_source_matches[0],
            source=lookup.source,
        )
    if request.operation.get("op") == "upsert":
        return _Resolution(row_index=None, source=lookup.source)
    if len(lookup.deleted_source_matches) == 1:
        return _Resolution(
            row_index=lookup.deleted_source_matches[0],
            source=lookup.source,
        )
    if len(lookup.deleted_source_matches) > 1:
        return _resolution_error(
            lookup,
            _identity_diagnostic(
                "identity.ambiguous_deleted_history",
                f"Source identity {lookup.source!r} has multiple Deleted history rows.",
                "Target one historical row by its exact workbook ID.",
                request.index,
                request.operation_id,
            ),
        )
    return _resolution_error(
        lookup,
        _identity_diagnostic(
            "identity.source_missing",
            f"Source identity {lookup.source!r} does not exist.",
            "Use a current identity from describe.",
            request.index,
            request.operation_id,
        ),
    )


def _resolve_existing(lookup: _IdentityLookup) -> _Resolution:
    request = lookup.request
    if len(lookup.id_matches) > 1:
        return _resolution_error(
            lookup,
            _identity_diagnostic(
                "identity.ambiguous_workbook_id",
                f"Workbook ID {lookup.requested_id!r} matches multiple rows.",
                "Correct duplicate workbook identifiers before planning changes.",
                request.index,
                request.operation_id,
            ),
        )
    if len(lookup.active_source_matches) > 1:
        return _resolution_error(
            lookup,
            _identity_diagnostic(
                "identity.duplicate_active_source",
                f"Source identity {lookup.source!r} matches multiple active rows.",
                "Correct duplicate active source identities before planning changes.",
                request.index,
                request.operation_id,
            ),
        )
    if lookup.requested_id is not None:
        return _resolve_requested_id(lookup)
    return _resolve_source_only(lookup)


def _allocate_id(
    entity: Entity,
    indexes: _Indexes,
    settings: dict[str, object],
) -> str:
    table = _ENTITY_TABLE[entity]
    prefix_name, counter_name = ID_COUNTERS[table]
    prefix = str(settings.get(prefix_name, SETTINGS_DEFAULTS[prefix_name]))
    number = int(settings.get(counter_name, SETTINGS_DEFAULTS[counter_name]))
    used = indexes.ids[entity]
    while f"{prefix}{number}".casefold() in used:
        number += 1
    workbook_id = f"{prefix}{number}"
    settings[counter_name] = number + 1
    return workbook_id


def _attach_source(row: dict[str, object], source: tuple[str, str] | None) -> None:
    if source is None or _row_source(row) == source:
        return
    row["Source"], row["Source ID"] = source


def _coerce_field(
    entity: Entity,
    field: str,
    value: object,
    index: int,
    operation_id: str,
) -> tuple[object | None, Diagnostic | None]:
    if field not in _DATE_FIELDS[entity]:
        return value, None
    if isinstance(value, date):
        return value, None
    if isinstance(value, str):
        try:
            return date.fromisoformat(value), None
        except ValueError:
            pass
    return None, _diagnostic(
        "field.invalid_date",
        f"{field!r} must be an RFC 3339 full date.",
        "Use YYYY-MM-DD.",
        pointer=_operation_pointer(index, f"/set/{field}"),
        operation_id=operation_id,
    )


def _apply_scalar_fields(
    pending: _Pending,
    row: dict[str, object],
) -> list[Diagnostic]:
    operation = pending.operation
    set_fields = operation.get("set", {})
    clear_fields = operation.get("clear", [])
    diagnostics: list[Diagnostic] = []
    if isinstance(clear_fields, list):
        for field in clear_fields:
            if isinstance(field, str):
                row.pop(field, None)
    if not isinstance(set_fields, Mapping):
        return diagnostics
    for field, raw_value in set_fields.items():
        if field in _REFERENCE_FIELDS[pending.entity] or field not in _WRITABLE[pending.entity]:
            continue
        value, error = _coerce_field(
            pending.entity,
            field,
            raw_value,
            pending.operation_index,
            pending.operation_id,
        )
        if error is not None:
            diagnostics.append(error)
        else:
            row[field] = value
    return diagnostics


def _resolve_reference(
    request: _ReferenceRequest,
    context: _ReferenceContext,
) -> tuple[str | None, Diagnostic | None]:
    reference_map = request.reference if isinstance(request.reference, Mapping) else {}
    workbook_id = reference_map.get("workbook_id")
    if isinstance(workbook_id, str):
        matches = _matches(context.indexes.ids["item"], workbook_id.casefold())
    else:
        source = _source_pair(reference_map.get("source"))
        if source is not None:
            matches = _matches(context.indexes.sources["item"], source)
            active, deleted = _partition_source_matches(
                matches,
                context.rows["item"],
                "item",
                context.deletion_roles,
            )
            matches = list(active or deleted)
        else:
            client_ref = reference_map.get("client_ref")
            pending = context.client_refs.get(client_ref) if isinstance(client_ref, str) else None
            if pending is None:
                matches = []
            elif pending.entity != "item":
                return None, _diagnostic(
                    "reference.entity",
                    f"{request.field} can reference Items only.",
                    "Point the client_ref to an Item upsert.",
                    pointer=_operation_pointer(request.index, f"/set/{request.field}"),
                    operation_id=request.operation_id,
                )
            else:
                return pending.workbook_id, None
    if len(matches) == 1:
        return str(context.rows["item"][matches[0]]["ID"]), None
    code = "reference.ambiguous" if len(matches) > 1 else "reference.missing"
    message = (
        f"{request.field} reference matches multiple Items."
        if matches
        else f"{request.field} reference does not match an Item."
    )
    return None, _diagnostic(
        code,
        message,
        "Use one current Item workbook ID, source identity or batch-local client_ref.",
        pointer=_operation_pointer(request.index, f"/set/{request.field}"),
        operation_id=request.operation_id,
    )


def _apply_reference_fields(
    pending: _Pending,
    row: dict[str, object],
    context: _ReferenceContext,
) -> list[Diagnostic]:
    set_fields = pending.operation.get("set", {})
    if not isinstance(set_fields, Mapping):
        return []
    diagnostics: list[Diagnostic] = []
    for field in _REFERENCE_FIELDS[pending.entity]:
        if field not in set_fields:
            continue
        raw_value = set_fields[field]
        references = (
            raw_value if field == "BlockedBy" and isinstance(raw_value, list) else [raw_value]
        )
        resolved: list[str] = []
        for reference in references:
            workbook_id, error = _resolve_reference(
                _ReferenceRequest(
                    reference=reference,
                    index=pending.operation_index,
                    operation_id=pending.operation_id,
                    field=field,
                ),
                context,
            )
            if error is not None:
                diagnostics.append(error)
            elif workbook_id is not None:
                resolved.append(workbook_id)
        if diagnostics:
            continue
        folded = [value.casefold() for value in resolved]
        if len(folded) != len(set(folded)):
            diagnostics.append(
                _diagnostic(
                    "reference.duplicate",
                    f"{field} resolves the same Item more than once.",
                    "Remove duplicate references after identity resolution.",
                    pointer=_operation_pointer(pending.operation_index, f"/set/{field}"),
                    operation_id=pending.operation_id,
                )
            )
            continue
        row[field] = ", ".join(resolved) if field == "BlockedBy" else resolved[0]
    return diagnostics


def _required_diagnostics(pending: _Pending, row: Mapping[str, object]) -> list[Diagnostic]:
    if not pending.created:
        return []
    missing = [field for field in ("Type", "Title", "Status") if row.get(field) in {None, ""}]
    if not missing:
        return []
    return [
        _diagnostic(
            "record.required",
            f"New {pending.entity} row is missing required fields: {', '.join(missing)}.",
            "Set valid Type, Title and Status values for every new row.",
            pointer=_operation_pointer(pending.operation_index, "/set"),
            operation_id=pending.operation_id,
        )
    ]


def _field_diffs(
    entity: Entity,
    before: Mapping[str, object],
    after: Mapping[str, object],
) -> tuple[FieldDiff, ...]:
    table_schema = TABLES_BY_NAME[_ENTITY_TABLE[entity]]
    return tuple(
        FieldDiff(field=name, before=before.get(name), after=after.get(name))
        for name in table_schema.column_names
        if before.get(name) != after.get(name)
    )


def _snapshot_from(
    baseline: Snapshot,
    tables: Mapping[str, list[dict[str, object]]],
    settings: Mapping[str, object],
) -> Snapshot:
    return Snapshot(
        schema_fingerprint=baseline.schema_fingerprint,
        workbook_digest=baseline.workbook_digest,
        exported_at=baseline.exported_at,
        settings=dict(settings),
        tables={table: tuple(dict(row) for row in rows) for table, rows in tables.items()},
    )


def _invalid(
    snapshot: Snapshot,
    diagnostics: Sequence[Diagnostic],
    operations: Sequence[OperationDiff] = (),
) -> MergeResult:
    return MergeResult(
        snapshot=snapshot,
        operations=tuple(operations),
        diagnostics=tuple(diagnostics),
    )


class _Merger:
    """Hold mutable planning state while preserving an atomic public result."""

    def __init__(
        self,
        snapshot: Snapshot,
        operations: Sequence[object],
        effective_date: date,
    ) -> None:
        self.snapshot = snapshot
        self.operations = operations
        self.effective_date = effective_date
        self.tables = _copy_tables(snapshot)
        self.rows = _rows_by_entity(self.tables)
        self.settings = dict(snapshot.settings)
        self.indexes = _build_indexes(self.rows)
        self.deletion_roles, role_errors = _deletion_roles(self.tables)
        self.lifecycle_config = _lifecycle_config(self.tables)
        self.pending: list[_Pending] = []
        self.client_refs: dict[str, _Pending] = {}
        self.claimed_targets: dict[tuple[Entity, int], str] = {}
        self.diagnostics = list(role_errors)
        if not role_errors:
            self.diagnostics.extend(
                _active_source_diagnostics(
                    self.indexes,
                    self.rows,
                    self.deletion_roles,
                )
            )

    @staticmethod
    def _request(index: int, raw: object) -> _OperationRequest | None:
        operation = raw if isinstance(raw, Mapping) else {}
        entity = _entity(operation.get("entity"))
        if entity is None:
            return None
        return _OperationRequest(
            operation=operation,
            entity=entity,
            index=index,
            operation_id=str(operation.get("operation_id", f"operation-{index}")),
        )

    def _materialize(self, request: _OperationRequest, resolution: _Resolution) -> _Pending:
        row_index = resolution.row_index
        created = row_index is None
        if created:
            workbook_id = _allocate_id(request.entity, self.indexes, self.settings)
            row = {_ENTITY_KEY[request.entity]: workbook_id}
            _attach_source(row, resolution.source)
            row_index = len(self.rows[request.entity])
            self.rows[request.entity].append(row)
            self.indexes = _build_indexes(self.rows)
            before: dict[str, object] = {}
        else:
            row = self.rows[request.entity][row_index]
            workbook_id = str(row[_ENTITY_KEY[request.entity]])
            before = dict(row)
            _attach_source(row, resolution.source)
        return _Pending(
            operation_index=request.index,
            operation_id=request.operation_id,
            entity=request.entity,
            row_index=row_index,
            workbook_id=workbook_id,
            before=before,
            created=created,
            operation=request.operation,
        )

    def _claim(self, pending: _Pending) -> bool:
        target = (pending.entity, pending.row_index)
        previous = self.claimed_targets.get(target)
        if previous is None:
            self.claimed_targets[target] = pending.operation_id
            return True
        self.diagnostics.append(
            _diagnostic(
                "operation.duplicate_target",
                f"Operation targets the same {pending.entity} row as {previous!r}.",
                "Combine all requested writes for one logical row into one operation.",
                pointer=_operation_pointer(pending.operation_index, "/identity"),
                operation_id=pending.operation_id,
            )
        )
        return False

    def _register(self, pending: _Pending) -> None:
        self.pending.append(pending)
        client_ref = pending.operation.get("client_ref")
        if isinstance(client_ref, str):
            self.client_refs[client_ref] = pending
        row = self.rows[pending.entity][pending.row_index]
        if pending.operation.get("op") == "mark_deleted":
            if not _is_deleted(row, pending.entity, self.deletion_roles):
                row["Status"] = self.deletion_roles.for_entity(pending.entity)
        else:
            self.diagnostics.extend(_apply_scalar_fields(pending, row))

    def _deleted_route_error(self, request: _OperationRequest) -> Diagnostic | None:
        if request.operation.get("op") != "upsert":
            return None
        set_fields = request.operation.get("set")
        status = set_fields.get("Status") if isinstance(set_fields, Mapping) else None
        if not _same_config_text(status, self.deletion_roles.for_entity(request.entity)):
            return None
        return _diagnostic(
            "operation.deleted_status_requires_mark",
            "An upsert cannot set the configured Deleted status directly.",
            "Use mark_deleted for an explicit history-preserving transition.",
            pointer=_operation_pointer(request.index, "/set/Status"),
            operation_id=request.operation_id,
        )

    def resolve_operations(self) -> None:
        """Resolve identities and apply non-reference fields in source order."""
        if any(diagnostic.code == "config.deleted_role" for diagnostic in self.diagnostics):
            return
        for index, raw in enumerate(self.operations):
            request = self._request(index, raw)
            if request is None:
                continue
            route_error = self._deleted_route_error(request)
            if route_error is not None:
                self.diagnostics.append(route_error)
                continue
            lookup = _identity_lookup(
                request,
                self.indexes,
                self.rows,
                self.deletion_roles,
            )
            resolution = _resolve_existing(lookup)
            self.diagnostics.extend(resolution.diagnostics)
            if resolution.diagnostics:
                continue
            pending = self._materialize(request, resolution)
            if self._claim(pending):
                self._register(pending)

    def resolve_relationships(self) -> None:
        """Resolve structured references after every batch ID is known."""
        self.indexes = _build_indexes(self.rows)
        context = _ReferenceContext(
            indexes=self.indexes,
            rows=self.rows,
            client_refs=self.client_refs,
            deletion_roles=self.deletion_roles,
        )
        for pending in self.pending:
            row = self.rows[pending.entity][pending.row_index]
            self.diagnostics.extend(_apply_reference_fields(pending, row, context))
            self.diagnostics.extend(_required_diagnostics(pending, row))
            _apply_lifecycle(
                pending,
                row,
                self.effective_date,
                self.lifecycle_config,
            )

    def _operation_diffs(self) -> tuple[OperationDiff, ...]:
        results: list[OperationDiff] = []
        for pending in self.pending:
            row = self.rows[pending.entity][pending.row_index]
            diffs = _field_diffs(pending.entity, pending.before, row)
            action: MergeAction = (
                "create"
                if pending.created
                else "mark_deleted"
                if diffs and pending.operation.get("op") == "mark_deleted"
                else "update"
                if diffs
                else "noop"
            )
            results.append(
                OperationDiff(
                    operation_id=pending.operation_id,
                    entity=pending.entity,
                    action=action,
                    workbook_id=pending.workbook_id,
                    diffs=diffs,
                )
            )
        return tuple(results)

    def result(self) -> MergeResult:
        """Return the unchanged baseline on error or the complete merged state.

        Returns:
            The atomic merge outcome.

        """
        if any(diagnostic.severity == "error" for diagnostic in self.diagnostics):
            return _invalid(self.snapshot, self.diagnostics)
        merged = _snapshot_from(self.snapshot, self.tables, self.settings)
        operation_diffs = self._operation_diffs()
        operation_ids = {
            (pending.entity, pending.workbook_id.casefold()): pending.operation_id
            for pending in self.pending
        }
        operation_fields = {
            (operation.entity, operation.workbook_id.casefold()): frozenset(
                diff.field for diff in operation.diffs
            )
            for operation in operation_diffs
        }
        diagnostics = (
            *self.diagnostics,
            *compare_snapshot_findings(
                self.snapshot,
                merged,
                operation_ids,
                operation_fields,
            ),
        )
        if any(diagnostic.severity == "error" for diagnostic in diagnostics):
            return _invalid(self.snapshot, diagnostics, operation_diffs)
        return MergeResult(
            snapshot=merged,
            operations=operation_diffs,
            diagnostics=diagnostics,
        )


def merge_changes(
    snapshot: Snapshot,
    change_set: Mapping[str, object],
    *,
    effective_date: date,
) -> MergeResult:
    """Merge one parsed change set into an authored snapshot without I/O.

    Args:
        snapshot: Current workbook-authored values.
        change_set: A strict contract document, normally from ``parse_changeset``.
        effective_date: Local lifecycle date reserved for the transition stage.

    Returns:
        An atomic result containing the unchanged input snapshot on any error.

    """
    raw_operations = change_set.get("operations", [])
    operations = raw_operations if isinstance(raw_operations, list) else []
    diagnostics = _preflight_operations(operations)
    if diagnostics:
        return _invalid(snapshot, diagnostics)
    merger = _Merger(snapshot, operations, effective_date)
    merger.resolve_operations()
    merger.resolve_relationships()
    return merger.result()


__all__ = [
    "FieldDiff",
    "MergeResult",
    "OperationDiff",
    "merge_changes",
]
