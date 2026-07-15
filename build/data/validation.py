"""Stable baseline-versus-merged validation findings for the agent bridge."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime
from typing import TYPE_CHECKING

from .diagnostics import Diagnostic
from .schema import DATA_TABLES

if TYPE_CHECKING:
    from collections.abc import Iterable, Mapping, Sequence

    from .snapshot import Snapshot

MINIMUM_DATE = date(2020, 1, 1)
EXCEL_CELL_LIMIT = 32_767
MAX_HIERARCHY_LEVEL = 6
MAX_ANCESTORS = MAX_HIERARCHY_LEVEL - 1
MAX_SEVERITY_SCORE = 25
MAX_RAID_RATING = 5


@dataclass(frozen=True, slots=True)
class Finding:
    """One stable workbook validation finding."""

    code: str
    entity: str
    record: str
    field: str
    message: str
    hint: str

    @property
    def identity(self) -> tuple[str, str, str, str]:
        """Return the comparison identity for baseline and merged states."""
        return (
            self.code,
            self.entity,
            self.record.casefold(),
            self.field.casefold(),
        )


@dataclass(frozen=True, slots=True)
class _ItemRole:
    active: bool
    done: bool
    cancelled: bool
    deleted: bool


@dataclass(frozen=True, slots=True)
class _Domains:
    type_levels: Mapping[str, object]
    item_statuses: Mapping[str, _ItemRole]
    priorities: frozenset[str]
    health: tuple[str, ...]
    people: frozenset[str]
    raid_types: Mapping[str, tuple[bool, bool]]
    raid_statuses: Mapping[str, bool]
    teams: frozenset[str]

    @property
    def blocked_health(self) -> str | None:
        """Return the final configured Delivery Health value."""
        return self.health[-1] if self.health else None


def _finding(
    code: str,
    entity: str,
    record: object,
    field: str,
    *text: str,
) -> Finding:
    message, hint = text
    return Finding(
        code=code,
        entity=entity,
        record=str(record),
        field=field,
        message=message,
        hint=hint,
    )


def _folded(values: Iterable[object]) -> frozenset[str]:
    return frozenset(str(value).casefold() for value in values if value not in {None, ""})


def _domains(snapshot: Snapshot) -> _Domains:
    tables = snapshot.tables
    return _Domains(
        type_levels={
            str(row["Type"]).casefold(): row.get("Level")
            for row in tables.get("tblTypes", ())
            if row.get("Type") not in {None, ""}
        },
        item_statuses={
            str(row["Status"]).casefold(): _ItemRole(
                active=row.get("IsActive") is True,
                done=row.get("IsDone") is True,
                cancelled=row.get("IsCancelled") is True,
                deleted=row.get("IsDeleted") is True,
            )
            for row in tables.get("tblStatuses", ())
            if row.get("Status") not in {None, ""}
        },
        priorities=_folded(row.get("Priority") for row in tables.get("tblPriorities", ())),
        health=tuple(
            str(row["Delivery Health"])
            for row in tables.get("tblDeliveryHealth", ())
            if row.get("Delivery Health") not in {None, ""}
        ),
        people=_folded(row.get("Person") for row in tables.get("tblPeople", ())),
        raid_types={
            str(row["RaidType"]).casefold(): (
                row.get("IsAlert") is True,
                row.get("IsDecision") is True,
            )
            for row in tables.get("tblRaidTypes", ())
            if row.get("RaidType") not in {None, ""}
        },
        raid_statuses={
            str(row["RaidStatus"]).casefold(): row.get("IsClosed") is True
            for row in tables.get("tblRaidStatuses", ())
            if row.get("RaidStatus") not in {None, ""}
        },
        teams=_folded(row.get("Team") for row in tables.get("tblTeams", ())),
    )


def _record_id(row: Mapping[str, object], key: str, index: int) -> str:
    value = row.get(key)
    return str(value) if value not in {None, ""} else f"row-{index}"


def _duplicate_findings(
    entity: str,
    rows: Sequence[Mapping[str, object]],
    key: str,
    code: str,
) -> list[Finding]:
    counts = Counter(
        str(row[key]).casefold() for row in rows if isinstance(row.get(key), str) and row.get(key)
    )
    return [
        _finding(
            code,
            entity,
            row[key],
            key,
            f"{key} {row[key]!r} is duplicated.",
            f"Keep one case-insensitively unique {key} value.",
        )
        for row in rows
        if isinstance(row.get(key), str) and row.get(key) and counts[str(row[key]).casefold()] > 1
    ]


def _config_shape_findings(snapshot: Snapshot) -> list[Finding]:
    findings: list[Finding] = []
    for table_schema in DATA_TABLES:
        rows = snapshot.tables.get(table_schema.table, ())
        if len(rows) > table_schema.capacity:
            findings.append(
                _finding(
                    "capacity.exceeded",
                    table_schema.table,
                    table_schema.table,
                    "rows",
                    f"{table_schema.table} has {len(rows)} rows; "
                    f"capacity is {table_schema.capacity}.",
                    "Reduce rows before applying the change set.",
                )
            )
        if table_schema.key is not None and table_schema.sheet == "Config":
            findings.extend(
                _duplicate_findings(
                    table_schema.table,
                    rows,
                    table_schema.key,
                    "config.duplicate_key",
                )
            )
    return findings


def _status_role_findings(snapshot: Snapshot) -> list[Finding]:
    findings: list[Finding] = []
    item_deleted = 0
    for row in snapshot.tables.get("tblStatuses", ()):
        label = row.get("Status", "?")
        active = row.get("IsActive") is True
        done = row.get("IsDone") is True
        cancelled = row.get("IsCancelled") is True
        deleted = row.get("IsDeleted") is True
        item_deleted += int(deleted)
        invalid = (
            (active and done)
            or (cancelled and not done)
            or (deleted and (active or not done or not cancelled))
        )
        if invalid:
            findings.append(
                _finding(
                    "config.item_status_roles",
                    "tblStatuses",
                    label,
                    "roles",
                    f"Item status {label!r} has contradictory role flags.",
                    "Use inactive/done/cancelled for the deletion role and coherent normal roles.",
                )
            )
    if item_deleted != 1:
        findings.append(
            _finding(
                "config.deleted_role",
                "tblStatuses",
                "IsDeleted",
                "IsDeleted",
                "Item statuses require exactly one deletion role.",
                "Reconcile Config before planning changes.",
            )
        )

    raid_deleted = 0
    for row in snapshot.tables.get("tblRaidStatuses", ()):
        label = row.get("RaidStatus", "?")
        closed = row.get("IsClosed") is True
        deleted = row.get("IsDeleted") is True
        raid_deleted += int(deleted)
        if deleted and not closed:
            findings.append(
                _finding(
                    "config.raid_status_roles",
                    "tblRaidStatuses",
                    label,
                    "roles",
                    f"RAID status {label!r} is deleted but not closed.",
                    "Mark every RAID deletion role closed.",
                )
            )
    if raid_deleted != 1:
        findings.append(
            _finding(
                "config.deleted_role",
                "tblRaidStatuses",
                "IsDeleted",
                "IsDeleted",
                "RAID statuses require exactly one deletion role.",
                "Reconcile Config before planning changes.",
            )
        )
    return findings


def _required_role_findings(snapshot: Snapshot) -> list[Finding]:
    tables = snapshot.tables
    requirements = (
        (
            "tblStatuses",
            "IsActive",
            any(row.get("IsActive") is True for row in tables.get("tblStatuses", ())),
        ),
        (
            "tblStatuses",
            "IsDone",
            any(row.get("IsDone") is True for row in tables.get("tblStatuses", ())),
        ),
        (
            "tblTypes",
            "Level 1",
            any(row.get("Level") == 1 for row in tables.get("tblTypes", ())),
        ),
        ("tblPriorities", "Priority", bool(tables.get("tblPriorities", ()))),
        (
            "tblRaidTypes",
            "IsAlert",
            any(row.get("IsAlert") is True for row in tables.get("tblRaidTypes", ())),
        ),
        (
            "tblRaidTypes",
            "IsDecision",
            any(row.get("IsDecision") is True for row in tables.get("tblRaidTypes", ())),
        ),
        (
            "tblRaidStatuses",
            "IsClosed",
            any(row.get("IsClosed") is True for row in tables.get("tblRaidStatuses", ())),
        ),
        (
            "tblDeliveryHealth",
            "Delivery Health",
            bool(tables.get("tblDeliveryHealth", ())),
        ),
    )
    return [
        _finding(
            "config.required_role",
            table,
            role,
            role,
            f"{table} has no configured {role} role.",
            "Restore the minimum Config role required by workbook formulas and automation.",
        )
        for table, role, present in requirements
        if not present
    ]


def _type_and_severity_findings(snapshot: Snapshot) -> list[Finding]:
    findings: list[Finding] = []
    for row in snapshot.tables.get("tblTypes", ()):
        label = row.get("Type", "?")
        level = row.get("Level")
        if (
            isinstance(level, bool)
            or not isinstance(level, int)
            or not 1 <= level <= MAX_HIERARCHY_LEVEL
        ):
            findings.append(
                _finding(
                    "config.type_level",
                    "tblTypes",
                    label,
                    "Level",
                    f"Type {label!r} has invalid level {level!r}.",
                    "Use one whole-number hierarchy level from 1 to 6.",
                )
            )
    previous = 0
    for index, row in enumerate(snapshot.tables.get("tblSeverity", ()), start=1):
        label = row.get("Severity", f"row-{index}")
        score = row.get("MinScore")
        invalid = (
            isinstance(score, bool)
            or not isinstance(score, int)
            or not 1 <= score <= MAX_SEVERITY_SCORE
            or (index == 1 and score != 1)
            or (isinstance(score, int) and score <= previous)
        )
        if invalid:
            findings.append(
                _finding(
                    "config.severity_order",
                    "tblSeverity",
                    label,
                    "MinScore",
                    f"Severity {label!r} has invalid threshold {score!r}.",
                    "Start at 1 and use strictly increasing whole-number thresholds through 25.",
                )
            )
        if isinstance(score, int) and not isinstance(score, bool):
            previous = score
    return findings


def _people_findings(snapshot: Snapshot, domains: _Domains) -> list[Finding]:
    findings: list[Finding] = []
    for index, row in enumerate(snapshot.tables.get("tblPeople", ()), start=1):
        person = _record_id(row, "Person", index)
        team = row.get("Team")
        if team not in {None, ""} and str(team).casefold() not in domains.teams:
            findings.append(
                _finding(
                    "choice.person_team",
                    "tblPeople",
                    person,
                    "Team",
                    f"Team {team!r} is not configured.",
                    "Choose a Team returned by describe.",
                )
            )
    return findings


def _config_findings(snapshot: Snapshot, domains: _Domains) -> list[Finding]:
    return [
        *_config_shape_findings(snapshot),
        *_status_role_findings(snapshot),
        *_required_role_findings(snapshot),
        *_type_and_severity_findings(snapshot),
        *_people_findings(snapshot, domains),
    ]


def _identifier_findings(
    snapshot: Snapshot,
    entity: str,
    table: str,
    key: str,
    prefix_setting: str,
) -> list[Finding]:
    rows = snapshot.tables.get(table, ())
    prefix = snapshot.settings.get(prefix_setting)
    prefix_text = str(prefix) if isinstance(prefix, str) else ""
    findings = _duplicate_findings(entity, rows, key, "id.duplicate")
    for index, row in enumerate(rows, start=1):
        record = _record_id(row, key, index)
        value = row.get(key)
        if value in {None, ""}:
            findings.append(
                _finding(
                    "id.missing",
                    entity,
                    record,
                    key,
                    f"{key} is blank on a populated row.",
                    "Assign a workbook identifier before applying changes.",
                )
            )
            continue
        suffix = str(value).removeprefix(prefix_text)
        malformed = (
            not isinstance(value, str)
            or not prefix_text
            or suffix == str(value)
            or not suffix.isdecimal()
            or int(suffix) < 1
        )
        if malformed:
            findings.append(
                _finding(
                    "id.malformed",
                    entity,
                    record,
                    key,
                    f"{key} {value!r} does not follow prefix {prefix_text!r} "
                    "plus a positive integer.",
                    "Correct the identifier or Config prefix.",
                )
            )
    return findings


def _source_findings(
    entity: str,
    rows: Sequence[Mapping[str, object]],
    key: str,
) -> list[Finding]:
    findings: list[Finding] = []
    for index, row in enumerate(rows, start=1):
        record = _record_id(row, key, index)
        source = row.get("Source")
        source_id = row.get("Source ID")
        source_present = source not in {None, ""}
        id_present = source_id not in {None, ""}
        if source_present != id_present:
            findings.append(
                _finding(
                    "source.incomplete_pair",
                    entity,
                    record,
                    "Source/Source ID",
                    "Source and Source ID must both be blank or both nonblank.",
                    "Complete or clear the system-managed identity pair through attachment.",
                )
            )
        for field, value in (("Source", source), ("Source ID", source_id)):
            if value not in {None, ""} and (
                not isinstance(value, str) or len(value) > EXCEL_CELL_LIMIT
            ):
                findings.append(
                    _finding(
                        "source.invalid_value",
                        entity,
                        record,
                        field,
                        f"{field} cannot be represented as one Excel text cell.",
                        "Use a text identifier no longer than 32,767 characters.",
                    )
                )
    return findings


def _required_and_choice_findings(
    entity: str,
    row: Mapping[str, object],
    record: str,
    choices: Sequence[tuple[str, frozenset[str], str]],
) -> list[Finding]:
    findings = [
        _finding(
            f"required.{field.casefold().replace(' ', '_')}",
            entity,
            record,
            field,
            f"{field} is required.",
            f"Set {field} to a valid nonblank value.",
        )
        for field in ("Type", "Title", "Status")
        if row.get(field) in {None, ""}
    ]
    for field, domain, code in choices:
        value = row.get(field)
        if value not in {None, ""} and str(value).casefold() not in domain:
            findings.append(
                _finding(
                    code,
                    entity,
                    record,
                    field,
                    f"{field} value {value!r} is not configured.",
                    f"Choose a {field} value returned by describe.",
                )
            )
    return findings


def _date_findings(
    entity: str,
    row: Mapping[str, object],
    record: str,
    fields: Sequence[str],
) -> list[Finding]:
    findings: list[Finding] = []
    for field in fields:
        value = row.get(field)
        if value in {None, ""}:
            continue
        if not isinstance(value, date) or isinstance(value, datetime):
            findings.append(
                _finding(
                    "date.invalid_type",
                    entity,
                    record,
                    field,
                    f"{field} is not a pure date.",
                    "Use an RFC 3339 full date with no time of day.",
                )
            )
        elif value < MINIMUM_DATE:
            findings.append(
                _finding(
                    "date.before_minimum",
                    entity,
                    record,
                    field,
                    f"{field} {value.isoformat()} is before 2020-01-01.",
                    "Use a date on or after 2020-01-01.",
                )
            )
    return findings


def _item_lifecycle_findings(
    row: Mapping[str, object],
    record: str,
    domains: _Domains,
) -> list[Finding]:
    findings = [
        _finding(
            "lifecycle.required_stamp",
            "item",
            record,
            field,
            f"{field} is required on an identified Item row.",
            "Let workbook automation stamp the lifecycle date.",
        )
        for field in ("Created", "Updated")
        if row.get(field) in {None, ""}
    ]
    status = row.get("Status")
    role = domains.item_statuses.get(str(status).casefold()) if status else None
    if role is not None:
        if role.active and row.get("InProgressSince") in {None, ""}:
            findings.append(
                _finding(
                    "lifecycle.active_stamp",
                    "item",
                    record,
                    "InProgressSince",
                    "Active status requires InProgressSince.",
                    "Let workbook automation stamp the active date.",
                )
            )
        delivered = role.done and not role.cancelled and not role.deleted
        done_present = row.get("DoneDate") not in {None, ""}
        if delivered != done_present:
            findings.append(
                _finding(
                    "lifecycle.done_stamp",
                    "item",
                    record,
                    "DoneDate",
                    "DoneDate contradicts the configured status roles.",
                    "Use the lifecycle stamp implied by the current Status.",
                )
            )
    health = row.get("Delivery Health")
    blocked = (
        domains.blocked_health is not None
        and isinstance(health, str)
        and health.casefold() == domains.blocked_health.casefold()
    )
    blocked_present = row.get("BlockedSince") not in {None, ""}
    if blocked != blocked_present:
        findings.append(
            _finding(
                "lifecycle.blocked_stamp",
                "item",
                record,
                "BlockedSince",
                "BlockedSince contradicts Delivery Health.",
                "Use the blocked stamp implied by the final Config health value.",
            )
        )
    narrative = row.get("Latest Status") not in {None, ""}
    narrative_stamp = row.get("LatestUpdateOn") not in {None, ""}
    if narrative != narrative_stamp:
        findings.append(
            _finding(
                "lifecycle.latest_status_stamp",
                "item",
                record,
                "LatestUpdateOn",
                "LatestUpdateOn contradicts Latest Status.",
                "Set both through a Latest Status edit.",
            )
        )
    return findings


def _date_order_finding(
    entity: str,
    record: str,
    names: tuple[str, str],
    values: tuple[object, object],
) -> Finding | None:
    start_name, end_name = names
    start, end = values
    if isinstance(start, date) and isinstance(end, date) and start > end:
        return _finding(
            "date.order",
            entity,
            record,
            f"{start_name}/{end_name}",
            f"{start_name} is after {end_name}.",
            f"Move {start_name} on or before {end_name}.",
        )
    return None


def _chronology_findings(
    entity: str,
    row: Mapping[str, object],
    record: str,
    pairs: Sequence[tuple[str, str]],
) -> list[Finding]:
    findings: list[Finding] = []
    for earlier_name, later_name in pairs:
        earlier = row.get(earlier_name)
        later = row.get(later_name)
        if not isinstance(earlier, date) or not isinstance(later, date) or earlier <= later:
            continue
        findings.append(
            _finding(
                "lifecycle.date_order",
                entity,
                record,
                f"{earlier_name}/{later_name}",
                f"{later_name} is before {earlier_name}.",
                "Correct the lifecycle stamp chronology.",
            )
        )
    return findings


def _item_findings(snapshot: Snapshot, domains: _Domains) -> list[Finding]:
    rows = snapshot.tables.get("tblItems", ())
    findings = [
        *_identifier_findings(snapshot, "item", "tblItems", "ID", "cfgItemIDPrefix"),
        *_source_findings("item", rows, "ID"),
    ]
    choices = (
        ("Type", frozenset(domains.type_levels), "choice.item_type"),
        ("Status", frozenset(domains.item_statuses), "choice.item_status"),
        ("Priority", domains.priorities, "choice.item_priority"),
        ("Delivery Health", _folded(domains.health), "choice.delivery_health"),
        ("Owner", domains.people, "choice.owner"),
    )
    date_fields = (
        "Start",
        "Due",
        "Created",
        "Updated",
        "InProgressSince",
        "DoneDate",
        "BlockedSince",
        "LatestUpdateOn",
    )
    for index, row in enumerate(rows, start=1):
        record = _record_id(row, "ID", index)
        findings.extend(_required_and_choice_findings("item", row, record, choices))
        findings.extend(_date_findings("item", row, record, date_fields))
        order = _date_order_finding(
            "item",
            record,
            ("Start", "Due"),
            (row.get("Start"), row.get("Due")),
        )
        if order is not None:
            findings.append(order)
        findings.extend(_item_lifecycle_findings(row, record, domains))
        findings.extend(
            _chronology_findings(
                "item",
                row,
                record,
                (
                    ("Created", "Updated"),
                    ("Created", "InProgressSince"),
                    ("Created", "DoneDate"),
                    ("Created", "BlockedSince"),
                    ("Created", "LatestUpdateOn"),
                ),
            )
        )
    return findings


def _item_indexes(
    rows: Sequence[Mapping[str, object]],
) -> dict[str, Mapping[str, object]]:
    return {
        str(row["ID"]).casefold(): row
        for row in rows
        if isinstance(row.get("ID"), str) and row.get("ID")
    }


def _item_level(row: Mapping[str, object], domains: _Domains) -> object:
    item_type = row.get("Type")
    return domains.type_levels.get(str(item_type).casefold()) if item_type else None


def _parent_findings(
    row: Mapping[str, object],
    record: str,
    item_index: Mapping[str, Mapping[str, object]],
    domains: _Domains,
) -> list[Finding]:
    findings: list[Finding] = []
    level = _item_level(row, domains)
    parent = row.get("Parent")
    if isinstance(level, int) and level > 1 and parent in {None, ""}:
        findings.append(
            _finding(
                "hierarchy.parent_required",
                "item",
                record,
                "Parent",
                f"Level-{level} Item requires a Parent.",
                "Reference one Item at a lower configured level.",
            )
        )
    if parent in {None, ""}:
        return findings
    parent_key = str(parent).casefold()
    if parent_key == record.casefold():
        findings.append(
            _finding(
                "hierarchy.self_parent",
                "item",
                record,
                f"Parent:{parent}",
                "An Item cannot be its own Parent.",
                "Choose another ancestor or clear Parent for Level 1.",
            )
        )
        return findings
    target = item_index.get(parent_key)
    if target is None:
        findings.append(
            _finding(
                "reference.parent_missing",
                "item",
                record,
                f"Parent:{parent}",
                f"Parent {parent!r} does not exist.",
                "Use a current Item ID.",
            )
        )
        return findings
    parent_level = _item_level(target, domains)
    if isinstance(level, int) and isinstance(parent_level, int) and parent_level >= level:
        findings.append(
            _finding(
                "hierarchy.parent_level",
                "item",
                record,
                f"Parent:{parent}",
                "Parent level must be lower than child level.",
                "Choose an ancestor Type with a lower configured level number.",
            )
        )
    target_status = target.get("Status")
    role = domains.item_statuses.get(str(target_status).casefold()) if target_status else None
    if role is not None and role.deleted:
        findings.append(
            _finding(
                "hierarchy.deleted_parent",
                "item",
                record,
                f"Parent:{parent}",
                "Parent points to a Deleted historical Item.",
                "Retain legacy history only; choose an active Parent for a new link.",
            )
        )
    return findings


def _dependency_findings(
    row: Mapping[str, object],
    record: str,
    item_index: Mapping[str, Mapping[str, object]],
) -> list[Finding]:
    value = row.get("BlockedBy")
    if value in {None, ""}:
        return []
    references = [part.strip() for part in str(value).split(",") if part.strip()]
    findings: list[Finding] = []
    for reference in references:
        if reference.casefold() == record.casefold():
            code = "reference.blocked_by_self"
            message = "BlockedBy cannot reference the Item itself."
        elif reference.casefold() not in item_index:
            code = "reference.blocked_by_missing"
            message = f"BlockedBy Item {reference!r} does not exist."
        else:
            continue
        findings.append(
            _finding(
                code,
                "item",
                record,
                f"BlockedBy:{reference}",
                message,
                "Use current Item IDs separated by commas.",
            )
        )
    return findings


def _ancestry_findings(
    rows: Sequence[Mapping[str, object]],
    item_index: Mapping[str, Mapping[str, object]],
) -> list[Finding]:
    findings: list[Finding] = []
    for index, row in enumerate(rows, start=1):
        record = _record_id(row, "ID", index)
        current = row
        seen = {record.casefold()}
        depth = 0
        while current.get("Parent") not in {None, ""}:
            parent = str(current["Parent"])
            folded = parent.casefold()
            depth += 1
            if folded in seen:
                findings.append(
                    _finding(
                        "hierarchy.cycle",
                        "item",
                        record,
                        f"Parent:{parent}",
                        "Item ancestry contains a cycle.",
                        "Break the Parent loop.",
                    )
                )
                break
            seen.add(folded)
            target = item_index.get(folded)
            if target is None:
                break
            current = target
        if depth > MAX_ANCESTORS:
            findings.append(
                _finding(
                    "hierarchy.depth",
                    "item",
                    record,
                    "Parent",
                    "Item ancestry exceeds the six-level hierarchy.",
                    "Shorten the Parent chain to six levels including the Item.",
                )
            )
    return findings


def _relationship_findings(snapshot: Snapshot, domains: _Domains) -> list[Finding]:
    rows = snapshot.tables.get("tblItems", ())
    item_index = _item_indexes(rows)
    findings: list[Finding] = []
    for index, row in enumerate(rows, start=1):
        record = _record_id(row, "ID", index)
        findings.extend(_parent_findings(row, record, item_index, domains))
        findings.extend(_dependency_findings(row, record, item_index))
    findings.extend(_ancestry_findings(rows, item_index))
    return findings


def _raid_lifecycle_findings(
    row: Mapping[str, object],
    record: str,
    domains: _Domains,
) -> list[Finding]:
    findings = [
        _finding(
            "lifecycle.required_stamp",
            "raid",
            record,
            field,
            f"{field} is required on an identified RAID row.",
            "Let workbook automation stamp the lifecycle date.",
        )
        for field in ("Raised", "Updated")
        if row.get(field) in {None, ""}
    ]
    status = row.get("Status")
    closed_role = domains.raid_statuses.get(str(status).casefold()) if status else None
    closed_present = row.get("Closed") not in {None, ""}
    if closed_role is not None and closed_role != closed_present:
        findings.append(
            _finding(
                "lifecycle.closed_stamp",
                "raid",
                record,
                "Closed",
                "Closed contradicts the configured RAID status role.",
                "Use the lifecycle stamp implied by Status.",
            )
        )
    order = _date_order_finding(
        "raid",
        record,
        ("Raised", "NextReview"),
        (row.get("Raised"), row.get("NextReview")),
    )
    if order is not None:
        findings.append(order)
    return findings


def _rating_findings(
    row: Mapping[str, object],
    record: str,
    domains: _Domains,
) -> list[Finding]:
    raid_type = row.get("Type")
    roles = domains.raid_types.get(str(raid_type).casefold()) if raid_type else None
    alert = roles[0] if roles is not None else False
    findings: list[Finding] = []
    for field in ("Prob", "Impact"):
        value = row.get(field)
        if roles is not None and not alert and value not in {None, ""}:
            findings.append(
                _finding(
                    "raid.rating_not_applicable",
                    "raid",
                    record,
                    field,
                    f"Non-alert RAID type does not use {field}.",
                    "Clear the rating or choose an alert RAID type.",
                )
            )
        elif alert and value in {None, ""}:
            findings.append(
                _finding(
                    "raid.rating_required",
                    "raid",
                    record,
                    field,
                    f"Alert RAID type requires {field}.",
                    "Set a whole-number rating from 1 to 5.",
                )
            )
        elif value not in {None, ""} and (
            isinstance(value, bool)
            or not isinstance(value, int)
            or not 1 <= value <= MAX_RAID_RATING
        ):
            findings.append(
                _finding(
                    "raid.rating_invalid",
                    "raid",
                    record,
                    field,
                    f"{field} {value!r} is outside 1 to 5.",
                    "Set a whole-number rating from 1 to 5.",
                )
            )
    return findings


def _raid_findings(snapshot: Snapshot, domains: _Domains) -> list[Finding]:
    rows = snapshot.tables.get("tblRAID", ())
    findings = [
        *_identifier_findings(snapshot, "raid", "tblRAID", "RaidID", "cfgRaidIDPrefix"),
        *_source_findings("raid", rows, "RaidID"),
    ]
    choices = (
        ("Type", frozenset(domains.raid_types), "choice.raid_type"),
        ("Status", frozenset(domains.raid_statuses), "choice.raid_status"),
        ("Owner", domains.people, "choice.owner"),
    )
    date_fields = ("NextReview", "Raised", "Closed", "Updated")
    item_index = _item_indexes(snapshot.tables.get("tblItems", ()))
    for index, row in enumerate(rows, start=1):
        record = _record_id(row, "RaidID", index)
        findings.extend(_required_and_choice_findings("raid", row, record, choices))
        findings.extend(_date_findings("raid", row, record, date_fields))
        findings.extend(_rating_findings(row, record, domains))
        findings.extend(_raid_lifecycle_findings(row, record, domains))
        findings.extend(
            _chronology_findings(
                "raid",
                row,
                record,
                (("Raised", "Updated"), ("Raised", "Closed")),
            )
        )
        related = row.get("RelatedID")
        if related not in {None, ""} and str(related).casefold() not in item_index:
            findings.append(
                _finding(
                    "reference.related_item_missing",
                    "raid",
                    record,
                    f"RelatedID:{related}",
                    f"RelatedID {related!r} does not exist.",
                    "Use a current Item ID.",
                )
            )
    return findings


def snapshot_findings(snapshot: Snapshot) -> tuple[Finding, ...]:
    """Return every stable Config, Item and RAID finding in sort order.

    Returns:
        Deterministically ordered validation findings.

    """
    domains = _domains(snapshot)
    findings = [
        *_config_findings(snapshot, domains),
        *_item_findings(snapshot, domains),
        *_relationship_findings(snapshot, domains),
        *_raid_findings(snapshot, domains),
    ]
    unique = {finding.identity: finding for finding in findings}
    return tuple(unique[identity] for identity in sorted(unique))


def _escape(value: object) -> str:
    return str(value).replace("~", "~0").replace("/", "~1")


def _finding_pointer(finding: Finding) -> str:
    return "/".join((
        "",
        "records" if finding.entity in {"item", "raid"} else "config",
        _escape(finding.entity),
        _escape(finding.record),
        _escape(finding.field),
    ))


def _item_by_id(snapshot: Snapshot, workbook_id: str) -> Mapping[str, object] | None:
    folded = workbook_id.casefold()
    return next(
        (
            row
            for row in snapshot.tables.get("tblItems", ())
            if isinstance(row.get("ID"), str) and str(row["ID"]).casefold() == folded
        ),
        None,
    )


def _retained_deleted_parent(
    finding: Finding,
    baseline: Snapshot,
    merged: Snapshot,
) -> bool:
    if finding.code != "hierarchy.deleted_parent":
        return False
    before = _item_by_id(baseline, finding.record)
    after = _item_by_id(merged, finding.record)
    return before is not None and after is not None and before.get("Parent") == after.get("Parent")


def _finding_fields(finding: Finding) -> frozenset[str]:
    """Return workbook columns named by a simple or composite finding field.

    Returns:
        The column names represented by the finding's field identifier.

    """
    return frozenset(part.split(":", maxsplit=1)[0] for part in finding.field.split("/"))


def compare_snapshot_findings(
    baseline: Snapshot,
    merged: Snapshot,
    operation_ids: Mapping[tuple[str, str], str],
    operation_fields: Mapping[tuple[str, str], frozenset[str]],
) -> tuple[Diagnostic, ...]:
    """Classify unchanged findings as warnings and new findings as errors.

    Returns:
        Stable diagnostics sorted with blocking errors before warnings.

    """
    baseline_identities = {finding.identity for finding in snapshot_findings(baseline)}
    diagnostics: list[Diagnostic] = []
    for finding in snapshot_findings(merged):
        record_key = (finding.entity, finding.record.casefold())
        changed_fields = operation_fields.get(record_key, frozenset())
        existing_untouched = (
            finding.identity in baseline_identities
            and not changed_fields.intersection(_finding_fields(finding))
        )
        unchanged = existing_untouched or _retained_deleted_parent(finding, baseline, merged)
        severity = "warning" if unchanged else "error"
        operation_id = operation_ids.get(record_key)
        diagnostics.append(
            Diagnostic(
                code=finding.code,
                severity=severity,
                phase="plan",
                pointer=_finding_pointer(finding),
                operation_id=operation_id,
                message=finding.message,
                hint=finding.hint,
            )
        )
    return tuple(
        sorted(
            diagnostics,
            key=lambda diagnostic: (
                diagnostic.severity == "warning",
                diagnostic.code,
                diagnostic.pointer,
            ),
        )
    )


__all__ = ["Finding", "compare_snapshot_findings", "snapshot_findings"]
