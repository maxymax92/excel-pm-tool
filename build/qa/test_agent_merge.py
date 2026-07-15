"""Regression contracts for provider-neutral identity and merge behaviour."""

from __future__ import annotations

import unittest
from copy import deepcopy
from dataclasses import replace
from datetime import date
from typing import TYPE_CHECKING

from ..data.merge import merge_changes
from ..data.validation import snapshot_findings
from .test_agent_contract import _change_set
from .test_data_layer import ITEM_ROWS, RAID_ROWS, _snapshot

if TYPE_CHECKING:
    from ..data.snapshot import Snapshot


def _source(namespace: str, record_id: str) -> dict[str, object]:
    return {"namespace": namespace, "record_id": record_id}


def _upsert(
    operation_id: str,
    *,
    entity: str = "item",
    **values: object,
) -> dict[str, object]:
    workbook_id = values.get("workbook_id")
    source = values.get("source")
    client_ref = values.get("client_ref")
    set_fields = values.get("set_fields")
    clear = values.get("clear")
    identity: dict[str, object] = {}
    if isinstance(workbook_id, str):
        identity["workbook_id"] = workbook_id
    if isinstance(source, dict):
        identity["source"] = source
    operation: dict[str, object] = {
        "operation_id": operation_id,
        "op": "upsert",
        "entity": entity,
        "identity": identity,
    }
    if isinstance(client_ref, str):
        operation["client_ref"] = client_ref
    if isinstance(set_fields, dict):
        operation["set"] = set_fields
    if isinstance(clear, list):
        operation["clear"] = clear
    return operation


def _changes(*operations: dict[str, object]) -> dict[str, object]:
    document = _change_set()
    document["operations"] = list(operations)
    return document


def _mark_deleted(
    operation_id: str,
    *,
    entity: str = "item",
    workbook_id: str | None = None,
    source: dict[str, object] | None = None,
) -> dict[str, object]:
    identity: dict[str, object] = {}
    if workbook_id is not None:
        identity["workbook_id"] = workbook_id
    if source is not None:
        identity["source"] = source
    return {
        "operation_id": operation_id,
        "op": "mark_deleted",
        "entity": entity,
        "identity": identity,
    }


def _item_registry_cases() -> tuple[tuple[str, Snapshot], ...]:
    item = dict(ITEM_ROWS[0])
    item.update({
        "Type": "Unknown type",
        "Title": "",
        "Status": "Unknown status",
        "Priority": "P9",
        "Delivery Health": "Unknown health",
        "Owner": "Unknown owner",
        "Start": "not-a-date",
        "Created": date(2026, 7, 10),
        "Updated": date(2026, 7, 1),
        "LatestUpdateOn": None,
    })
    item_snapshot = _snapshot(tables={"tblItems": (item,)})

    required = dict(ITEM_ROWS[0])
    required.update({"Type": "", "Title": "", "Status": ""})
    required_snapshot = _snapshot(tables={"tblItems": (required,)})

    unstamped = dict(ITEM_ROWS[0])
    unstamped.pop("Created", None)
    unstamped.pop("Updated", None)
    unstamped_snapshot = _snapshot(tables={"tblItems": (unstamped,)})

    lifecycle = dict(ITEM_ROWS[0])
    lifecycle.update({
        "Status": "Done",
        "Delivery Health": "Blocked",
        "Latest Status": "Narrative without stamp",
    })
    lifecycle.pop("DoneDate", None)
    lifecycle.pop("BlockedSince", None)
    lifecycle.pop("LatestUpdateOn", None)
    lifecycle_snapshot = _snapshot(tables={"tblItems": (lifecycle,)})

    return (
        ("required.type", required_snapshot),
        ("required.title", required_snapshot),
        ("required.status", required_snapshot),
        ("choice.item_type", item_snapshot),
        ("choice.item_status", item_snapshot),
        ("choice.item_priority", item_snapshot),
        ("choice.delivery_health", item_snapshot),
        ("choice.owner", item_snapshot),
        ("date.invalid_type", item_snapshot),
        ("lifecycle.date_order", item_snapshot),
        ("lifecycle.required_stamp", unstamped_snapshot),
        ("lifecycle.active_stamp", unstamped_snapshot),
        ("lifecycle.done_stamp", lifecycle_snapshot),
        ("lifecycle.blocked_stamp", lifecycle_snapshot),
        ("lifecycle.latest_status_stamp", lifecycle_snapshot),
    )


def _raid_registry_cases() -> tuple[tuple[str, Snapshot], ...]:
    domain = dict(RAID_ROWS[0])
    domain.update({
        "Type": "Unknown RAID type",
        "Status": "Unknown RAID status",
        "Owner": "Unknown owner",
        "Prob": 6,
        "Impact": 0,
    })
    domain_snapshot = _snapshot(tables={"tblRAID": (domain,)})

    closed = dict(RAID_ROWS[0])
    closed["Status"] = "Closed"
    closed.pop("Closed", None)
    closed_snapshot = _snapshot(tables={"tblRAID": (closed,)})
    return (
        ("choice.raid_type", domain_snapshot),
        ("choice.raid_status", domain_snapshot),
        ("raid.rating_invalid", domain_snapshot),
        ("lifecycle.closed_stamp", closed_snapshot),
    )


def _config_registry_cases() -> tuple[tuple[str, Snapshot], ...]:
    tables = deepcopy(_snapshot().tables)
    statuses = [dict(row) for row in tables["tblStatuses"]]
    next(row for row in statuses if row.get("IsDeleted") is True)["IsDone"] = False
    tables["tblStatuses"] = tuple(statuses)
    raid_statuses = [dict(row) for row in tables["tblRaidStatuses"]]
    next(row for row in raid_statuses if row.get("IsDeleted") is True)["IsClosed"] = False
    tables["tblRaidStatuses"] = tuple(raid_statuses)
    types = [dict(row) for row in tables["tblTypes"]]
    types[0]["Level"] = 7
    tables["tblTypes"] = tuple(types)
    severities = [dict(row) for row in tables["tblSeverity"]]
    severities[0]["MinScore"] = 2
    tables["tblSeverity"] = tuple(severities)
    config_snapshot = replace(_snapshot(), tables=tables)

    required_tables = deepcopy(_snapshot().tables)
    required_tables["tblStatuses"] = tuple(
        {**row, "IsActive": False} for row in required_tables["tblStatuses"]
    )
    required_snapshot = replace(_snapshot(), tables=required_tables)
    return (
        ("config.item_status_roles", config_snapshot),
        ("config.raid_status_roles", config_snapshot),
        ("config.required_role", required_snapshot),
        ("config.type_level", config_snapshot),
        ("config.severity_order", config_snapshot),
    )


class IdentityAndDiffTests(unittest.TestCase):
    """Resolve existing rows and preserve omission/set/clear semantics."""

    def test_case_insensitive_workbook_id_updates_with_exact_diffs(self) -> None:
        """Match IDs like Excel while preserving spelling and omitted fields."""
        changes = _changes(
            _upsert(
                "update-project",
                workbook_id="i-3001",
                set_fields={"Title": "Rollout programme v2", "Due": "2026-10-15"},
                clear=["Priority"],
            )
        )
        result = merge_changes(_snapshot(), changes, effective_date=date(2026, 7, 15))

        self.assertTrue(result.valid)
        self.assertTrue(result.changed)
        operation = result.operations[0]
        self.assertEqual(operation.action, "update")
        self.assertEqual(operation.workbook_id, "I-3001")
        self.assertEqual(
            {diff.field: (diff.before, diff.after) for diff in operation.diffs},
            {
                "Title": ("Rollout programme", "Rollout programme v2"),
                "Priority": ("P1", None),
                "Due": (date(2026, 9, 30), date(2026, 10, 15)),
                "Updated": (date(2026, 7, 10), date(2026, 7, 15)),
            },
        )
        row = result.snapshot.tables["tblItems"][0]
        self.assertEqual(row["ID"], "I-3001")
        self.assertEqual(row["Owner"], "Ana Ruiz")
        self.assertNotIn("Priority", row)

    def test_source_identity_updates_and_manual_attachment_are_exact(self) -> None:
        """Use exact source pairs and attach only the named blank manual row."""
        changes = _changes(
            _upsert(
                "source-update",
                source=_source("api:delivery", "project-3001"),
                set_fields={"Title": "Source-owned programme"},
            ),
            _upsert(
                "attach-manual",
                workbook_id="I-3002",
                source=_source("paste:weekly", "vendor-workstream"),
            ),
        )
        result = merge_changes(_snapshot(), changes, effective_date=date(2026, 7, 15))

        self.assertTrue(result.valid)
        first, second = result.operations
        self.assertEqual(first.workbook_id, "I-3001")
        self.assertEqual(second.workbook_id, "I-3002")
        self.assertEqual(
            {diff.field for diff in second.diffs},
            {"Source", "Source ID"},
        )
        attached = result.snapshot.tables["tblItems"][1]
        self.assertEqual(attached["Source"], "paste:weekly")
        self.assertEqual(attached["Source ID"], "vendor-workstream")

    def test_same_value_upsert_is_a_noop(self) -> None:
        """Avoid diffs and counter changes when requested values already match."""
        snapshot = _snapshot()
        changes = _changes(
            _upsert(
                "replay",
                workbook_id="I-3001",
                set_fields={"Title": "Rollout programme", "Due": "2026-09-30"},
            )
        )
        result = merge_changes(snapshot, changes, effective_date=date(2026, 7, 15))

        self.assertTrue(result.valid)
        self.assertFalse(result.changed)
        self.assertEqual(result.operations[0].action, "noop")
        self.assertEqual(result.operations[0].diffs, ())
        self.assertEqual(result.snapshot.settings, snapshot.settings)


class CreationAndRelationshipTests(unittest.TestCase):
    """Allocate IDs before resolving provider-neutral references."""

    def test_batch_local_item_and_raid_links_resolve_in_operation_order(self) -> None:
        """Create deterministic IDs, counters and workbook-native references."""
        changes = _changes(
            _upsert(
                "new-project",
                source=_source("api:portfolio", "project-42"),
                client_ref="project",
                set_fields={
                    "Type": "Project",
                    "Title": "New programme",
                    "Status": "Backlog",
                },
            ),
            _upsert(
                "new-release",
                source=_source("api:portfolio", "release-42"),
                client_ref="release",
                set_fields={
                    "Type": "Release",
                    "Title": "First release",
                    "Status": "Ready",
                    "Parent": {"client_ref": "project"},
                    "BlockedBy": [
                        {"workbook_id": "I-3003"},
                        {"client_ref": "project"},
                    ],
                },
            ),
            _upsert(
                "new-risk",
                entity="raid",
                source=_source("api:portfolio", "risk-42"),
                set_fields={
                    "Type": "Risk",
                    "Title": "Delivery risk",
                    "Status": "Open",
                    "RelatedID": {"client_ref": "release"},
                    "Prob": 3,
                    "Impact": 4,
                },
            ),
        )
        result = merge_changes(_snapshot(), changes, effective_date=date(2026, 7, 15))

        self.assertTrue(result.valid, result.diagnostics)
        self.assertEqual(
            [operation.workbook_id for operation in result.operations],
            ["I-3004", "I-3005", "R-102"],
        )
        project, release = result.snapshot.tables["tblItems"][-2:]
        risk = result.snapshot.tables["tblRAID"][-1]
        self.assertEqual(release["Parent"], project["ID"])
        self.assertEqual(release["BlockedBy"], "I-3003, I-3004")
        self.assertEqual(risk["RelatedID"], release["ID"])
        self.assertEqual(result.snapshot.settings["cfgNextItemID"], 3006)
        self.assertEqual(result.snapshot.settings["cfgNextRaidID"], 103)

    def test_exact_source_identity_does_not_fold_case(self) -> None:
        """Treat a differently cased namespace as a new portable identity."""
        changes = _changes(
            _upsert(
                "case-sensitive-source",
                source=_source("API:delivery", "project-3001"),
                set_fields={
                    "Type": "Project",
                    "Title": "Independent source row",
                    "Status": "Backlog",
                },
            )
        )
        result = merge_changes(_snapshot(), changes, effective_date=date(2026, 7, 15))

        self.assertTrue(result.valid)
        self.assertEqual(result.operations[0].action, "create")
        self.assertEqual(result.operations[0].workbook_id, "I-3004")

    def test_source_identity_relationships_resolve_to_workbook_ids(self) -> None:
        """Resolve Parent, BlockedBy and RelatedID through exact source pairs."""
        source_reference = {"source": _source("api:delivery", "project-3001")}
        changes = _changes(
            _upsert(
                "source-item-links",
                workbook_id="I-3003",
                set_fields={
                    "Parent": source_reference,
                    "BlockedBy": [source_reference],
                },
            ),
            _upsert(
                "source-raid-link",
                entity="raid",
                workbook_id="R-101",
                set_fields={"RelatedID": source_reference},
            ),
        )
        result = merge_changes(_snapshot(), changes, effective_date=date(2026, 7, 15))

        self.assertTrue(result.valid, result.diagnostics)
        item = result.snapshot.tables["tblItems"][2]
        raid = result.snapshot.tables["tblRAID"][0]
        self.assertEqual(item["Parent"], "I-3001")
        self.assertEqual(item["BlockedBy"], "I-3001")
        self.assertEqual(raid["RelatedID"], "I-3001")


class MergeRejectionTests(unittest.TestCase):
    """Reject ambiguous, unowned or unresolved batches atomically."""

    def test_conflicting_dual_identity_is_rejected(self) -> None:
        """Never attach a source pair that already identifies another row."""
        changes = _changes(
            _upsert(
                "conflict",
                workbook_id="I-3002",
                source=_source("api:delivery", "project-3001"),
            )
        )
        result = merge_changes(_snapshot(), changes, effective_date=date(2026, 7, 15))
        self.assertFalse(result.valid)
        self.assertEqual(result.diagnostics[0].code, "identity.conflict")

    def test_missing_relationship_is_rejected_without_partial_snapshot(self) -> None:
        """Fail the whole batch when a structured reference cannot resolve."""
        snapshot = _snapshot()
        changes = _changes(
            _upsert(
                "missing-parent",
                workbook_id="I-3002",
                set_fields={"Parent": {"workbook_id": "I-9999"}},
            )
        )
        result = merge_changes(snapshot, changes, effective_date=date(2026, 7, 15))
        self.assertFalse(result.valid)
        self.assertEqual(result.diagnostics[0].code, "reference.missing")
        self.assertEqual(result.snapshot, snapshot)

    def test_duplicate_operation_client_and_target_are_rejected(self) -> None:
        """Reject operation-order-dependent batches before mutating a snapshot."""
        cases = (
            (
                "operation.duplicate_id",
                _changes(
                    _upsert("same", workbook_id="I-3001"),
                    _upsert("same", workbook_id="I-3002"),
                ),
            ),
            (
                "operation.duplicate_client_ref",
                _changes(
                    _upsert("one", workbook_id="I-3001", client_ref="same"),
                    _upsert("two", workbook_id="I-3002", client_ref="same"),
                ),
            ),
            (
                "operation.duplicate_target",
                _changes(
                    _upsert("one", workbook_id="I-3001"),
                    _upsert(
                        "two",
                        source=_source("api:delivery", "project-3001"),
                    ),
                ),
            ),
            (
                "operation.duplicate_source_identity",
                _changes(
                    _upsert(
                        "source-one",
                        source=_source("api:delivery", "project-3001"),
                    ),
                    _upsert(
                        "source-two",
                        source=_source("api:delivery", "project-3001"),
                    ),
                ),
            ),
        )
        for code, changes in cases:
            with self.subTest(code=code):
                result = merge_changes(_snapshot(), changes, effective_date=date(2026, 7, 15))
                self.assertFalse(result.valid)
                self.assertIn(code, {diagnostic.code for diagnostic in result.diagnostics})

    def test_defensive_ownership_check_rejects_ids_stamps_and_formulas(self) -> None:
        """Enforce ownership even when a caller bypasses JSON Schema parsing."""
        for field in ("ID", "Updated", "WbsKey", "Source"):
            with self.subTest(field=field):
                changes = _changes(
                    _upsert(
                        "unowned",
                        workbook_id="I-3001",
                        set_fields={field: "forbidden"},
                    )
                )
                result = merge_changes(
                    _snapshot(),
                    changes,
                    effective_date=date(2026, 7, 15),
                )
                self.assertFalse(result.valid)
                self.assertEqual(result.diagnostics[0].code, "field.not_writable")

    def test_new_row_requires_source_and_core_fields(self) -> None:
        """Do not create rows without a durable source identity and core fields."""
        missing_source = _changes(
            _upsert(
                "missing-source",
                set_fields={"Type": "Project", "Title": "No source", "Status": "Backlog"},
            )
        )
        missing_fields = _changes(
            _upsert(
                "missing-fields",
                source=_source("paste:test", "empty"),
                set_fields={"Title": "Incomplete"},
            )
        )
        for code, changes in (
            ("identity.source_required", missing_source),
            ("record.required", missing_fields),
        ):
            with self.subTest(code=code):
                result = merge_changes(_snapshot(), changes, effective_date=date(2026, 7, 15))
                self.assertFalse(result.valid)
                self.assertIn(code, {diagnostic.code for diagnostic in result.diagnostics})

    def test_input_snapshot_and_change_set_are_not_mutated(self) -> None:
        """Keep merge planning referentially transparent for retries."""
        snapshot = _snapshot()
        changes = _changes(_upsert("update", workbook_id="I-3001", set_fields={"Title": "Changed"}))
        snapshot_before = deepcopy(snapshot)
        changes_before = deepcopy(changes)
        merge_changes(snapshot, changes, effective_date=date(2026, 7, 15))
        self.assertEqual(snapshot, snapshot_before)
        self.assertEqual(changes, changes_before)


class DeletedHistoryTests(unittest.TestCase):
    """Keep deletion explicit, visible, non-cascading and replay-safe."""

    @staticmethod
    def _deleted_item(
        *,
        workbook_id: str = "I-3001",
        source: tuple[str, str] = ("api:delivery", "project-3001"),
    ) -> dict[str, object]:
        row = dict(ITEM_ROWS[0])
        row.update({
            "ID": workbook_id,
            "Status": "Deleted",
            "Source": source[0],
            "Source ID": source[1],
        })
        return row

    def test_mark_deleted_changes_status_without_removal_or_cascade(self) -> None:
        """Retain the target, descendants and every unrelated authored value."""
        snapshot = _snapshot()
        changes = _changes(_mark_deleted("delete-project", workbook_id="I-3001"))
        result = merge_changes(snapshot, changes, effective_date=date(2026, 7, 15))

        self.assertTrue(result.valid)
        self.assertEqual(result.operations[0].action, "mark_deleted")
        self.assertEqual(len(result.snapshot.tables["tblItems"]), len(ITEM_ROWS))
        target = result.snapshot.tables["tblItems"][0]
        child = result.snapshot.tables["tblItems"][1]
        self.assertEqual(target["Status"], "Deleted")
        self.assertEqual(target["Source"], "api:delivery")
        self.assertEqual(target["Title"], ITEM_ROWS[0]["Title"])
        self.assertEqual(child["Parent"], "I-3001")

    def test_repeated_mark_deleted_is_a_noop_for_item_and_raid(self) -> None:
        """Do not restamp or rewrite rows already in their deletion role."""
        deleted_raid = dict(RAID_ROWS[0])
        deleted_raid["Status"] = "Deleted"
        snapshot = _snapshot(
            tables={
                "tblItems": (self._deleted_item(), *ITEM_ROWS[1:]),
                "tblRAID": (deleted_raid,),
            }
        )
        changes = _changes(
            _mark_deleted("delete-item-again", workbook_id="I-3001"),
            _mark_deleted("delete-raid-again", entity="raid", workbook_id="R-101"),
        )
        result = merge_changes(snapshot, changes, effective_date=date(2026, 7, 15))

        self.assertTrue(result.valid)
        self.assertFalse(result.changed)
        self.assertEqual([operation.action for operation in result.operations], ["noop", "noop"])

    def test_source_reappearance_creates_fresh_id_and_keeps_history(self) -> None:
        """Create a new active row when the source pair has only Deleted history."""
        snapshot = _snapshot(
            tables={
                "tblItems": (self._deleted_item(), *ITEM_ROWS[1:]),
            }
        )
        changes = _changes(
            _upsert(
                "source-reappeared",
                source=_source("api:delivery", "project-3001"),
                set_fields={
                    "Type": "Project",
                    "Title": "Reopened programme",
                    "Status": "Backlog",
                },
            )
        )
        result = merge_changes(snapshot, changes, effective_date=date(2026, 7, 15))

        self.assertTrue(result.valid, result.diagnostics)
        self.assertEqual(result.operations[0].action, "create")
        self.assertEqual(result.operations[0].workbook_id, "I-3004")
        rows = result.snapshot.tables["tblItems"]
        self.assertEqual(rows[0]["Status"], "Deleted")
        self.assertEqual(rows[-1]["Source"], rows[0]["Source"])
        self.assertEqual(rows[-1]["Source ID"], rows[0]["Source ID"])
        self.assertEqual(rows[-1]["Status"], "Backlog")

    def test_active_source_match_wins_over_deleted_history(self) -> None:
        """Resolve a source-only operation to its one active row."""
        history = self._deleted_item(workbook_id="I-2999")
        snapshot = _snapshot(tables={"tblItems": (history, *ITEM_ROWS)})
        changes = _changes(
            _upsert(
                "update-active",
                source=_source("api:delivery", "project-3001"),
                set_fields={"Title": "Current active title"},
            )
        )
        result = merge_changes(snapshot, changes, effective_date=date(2026, 7, 15))
        self.assertTrue(result.valid)
        self.assertEqual(result.operations[0].workbook_id, "I-3001")
        self.assertEqual(result.snapshot.tables["tblItems"][0]["Status"], "Deleted")

    def test_deleted_workbook_id_cannot_be_upserted(self) -> None:
        """Protect a historical row from mutation through its old workbook ID."""
        snapshot = _snapshot(
            tables={
                "tblItems": (self._deleted_item(), *ITEM_ROWS[1:]),
            }
        )
        changes = _changes(
            _upsert(
                "rewrite-history",
                workbook_id="I-3001",
                set_fields={"Title": "Should fail"},
            )
        )
        result = merge_changes(snapshot, changes, effective_date=date(2026, 7, 15))
        self.assertFalse(result.valid)
        self.assertEqual(result.diagnostics[0].code, "identity.deleted_history")

    def test_source_only_delete_with_multiple_history_rows_is_ambiguous(self) -> None:
        """Require an exact workbook ID when no active source row exists."""
        source = ("paste:weekly", "project-x")
        first = self._deleted_item(workbook_id="I-2998", source=source)
        second = self._deleted_item(workbook_id="I-2999", source=source)
        snapshot = _snapshot(tables={"tblItems": (first, second, *ITEM_ROWS[1:])})
        changes = _changes(
            _mark_deleted(
                "ambiguous-history",
                source=_source(*source),
            )
        )
        result = merge_changes(snapshot, changes, effective_date=date(2026, 7, 15))
        self.assertFalse(result.valid)
        self.assertEqual(result.diagnostics[0].code, "identity.ambiguous_deleted_history")

    def test_mark_deleted_missing_target_is_an_error(self) -> None:
        """Never treat an unresolved explicit deletion as a successful no-op."""
        cases = (
            (
                _mark_deleted("missing-workbook", workbook_id="I-9999"),
                "identity.workbook_id_missing",
            ),
            (
                _mark_deleted(
                    "missing-source",
                    source=_source("paste:missing", "item-1"),
                ),
                "identity.source_missing",
            ),
        )
        for operation, code in cases:
            with self.subTest(code=code):
                result = merge_changes(
                    _snapshot(),
                    _changes(operation),
                    effective_date=date(2026, 7, 15),
                )
                self.assertFalse(result.valid)
                self.assertIn(code, {diagnostic.code for diagnostic in result.diagnostics})

    def test_duplicate_active_source_identity_blocks_the_batch(self) -> None:
        """Permit history duplicates but never two non-Deleted source rows."""
        duplicate = dict(ITEM_ROWS[1])
        duplicate.update({
            "Source": "api:delivery",
            "Source ID": "project-3001",
        })
        snapshot = _snapshot(tables={"tblItems": (ITEM_ROWS[0], duplicate, ITEM_ROWS[2])})
        changes = _changes(
            _upsert("unrelated", workbook_id="I-3003", set_fields={"Title": "Still blocked"})
        )
        result = merge_changes(snapshot, changes, effective_date=date(2026, 7, 15))
        self.assertFalse(result.valid)
        self.assertIn(
            "identity.duplicate_active_source",
            {diagnostic.code for diagnostic in result.diagnostics},
        )

    def test_upsert_cannot_set_the_deleted_role_directly(self) -> None:
        """Route deletion through mark_deleted so history semantics stay explicit."""
        changes = _changes(
            _upsert(
                "wrong-delete-route",
                workbook_id="I-3001",
                set_fields={"Status": "Deleted"},
            )
        )
        result = merge_changes(_snapshot(), changes, effective_date=date(2026, 7, 15))
        self.assertFalse(result.valid)
        self.assertEqual(result.diagnostics[0].code, "operation.deleted_status_requires_mark")

    def test_case_variant_deleted_role_cannot_bypass_explicit_deletion(self) -> None:
        """Match Config deletion roles case-insensitively like desktop Excel."""
        changes = _changes(
            _upsert(
                "case-variant-delete",
                workbook_id="I-3001",
                set_fields={"Status": "deleted"},
            )
        )
        result = merge_changes(_snapshot(), changes, effective_date=date(2026, 7, 15))

        self.assertFalse(result.valid)
        self.assertEqual(result.diagnostics[0].code, "operation.deleted_status_requires_mark")


class LifecycleTransitionTests(unittest.TestCase):
    """Reproduce the source-controlled VBA lifecycle transition matrix."""

    effective_date = date(2026, 7, 15)

    def test_new_item_stamps_created_updated_and_role_dates(self) -> None:
        """Apply active, blocked and narrative roles on one created row."""
        changes = _changes(
            _upsert(
                "create-active-blocked",
                source=_source("paste:lifecycle", "item-1"),
                set_fields={
                    "Type": "Project",
                    "Title": "Lifecycle item",
                    "Status": "In Progress",
                    "Delivery Health": "Blocked",
                    "Latest Status": "Waiting for approval",
                },
            )
        )
        result = merge_changes(_snapshot(), changes, effective_date=self.effective_date)

        self.assertTrue(result.valid, result.diagnostics)
        row = result.snapshot.tables["tblItems"][-1]
        for field in (
            "Created",
            "Updated",
            "InProgressSince",
            "BlockedSince",
            "LatestUpdateOn",
        ):
            self.assertEqual(row[field], self.effective_date, field)
        self.assertNotIn("DoneDate", row)

    def test_every_item_status_role_matches_vba(self) -> None:
        """Set, retain or clear active/done stamps from Config roles."""
        cases = {
            "Backlog": {"done": False, "active": False},
            "Ready": {"done": False, "active": False},
            "In Progress": {"done": False, "active": True},
            "Review": {"done": False, "active": True},
            "Done": {"done": True, "active": False},
            "Cancelled": {"done": False, "active": False},
            "Deleted": {"done": False, "active": False},
        }
        for target, expected in cases.items():
            with self.subTest(status=target):
                row = dict(ITEM_ROWS[0])
                row["Status"] = "Ready" if target != "Ready" else "Backlog"
                row["DoneDate"] = date(2026, 7, 1)
                row.pop("InProgressSince", None)
                snapshot = _snapshot(tables={"tblItems": (row, *ITEM_ROWS[1:])})
                operation = (
                    _mark_deleted("status-change", workbook_id="I-3001")
                    if target == "Deleted"
                    else _upsert(
                        "status-change",
                        workbook_id="I-3001",
                        set_fields={"Status": target},
                    )
                )
                result = merge_changes(
                    snapshot,
                    _changes(operation),
                    effective_date=self.effective_date,
                )
                self.assertTrue(result.valid, result.diagnostics)
                changed = result.snapshot.tables["tblItems"][0]
                if expected["active"]:
                    self.assertEqual(changed["InProgressSince"], self.effective_date)
                else:
                    self.assertNotIn("InProgressSince", changed)
                if expected["done"]:
                    self.assertEqual(changed["DoneDate"], date(2026, 7, 1))
                else:
                    self.assertNotIn("DoneDate", changed)
                self.assertEqual(changed["Updated"], self.effective_date)

    def test_case_variant_roles_match_vba_text_comparison(self) -> None:
        """Apply Item, health and RAID roles with VBA's case-insensitive matching."""
        item = merge_changes(
            _snapshot(),
            _changes(
                _upsert(
                    "case-variant-item-roles",
                    workbook_id="I-3001",
                    set_fields={"Status": "done", "Delivery Health": "blocked"},
                )
            ),
            effective_date=self.effective_date,
        )
        self.assertTrue(item.valid, item.diagnostics)
        changed_item = item.snapshot.tables["tblItems"][0]
        self.assertEqual(changed_item["DoneDate"], self.effective_date)
        self.assertEqual(changed_item["BlockedSince"], self.effective_date)

        raid = merge_changes(
            _snapshot(),
            _changes(
                _upsert(
                    "case-variant-raid-role",
                    entity="raid",
                    workbook_id="R-101",
                    set_fields={"Status": "closed"},
                )
            ),
            effective_date=self.effective_date,
        )
        self.assertTrue(raid.valid, raid.diagnostics)
        self.assertEqual(raid.snapshot.tables["tblRAID"][0]["Closed"], self.effective_date)

    def test_type_change_reapplies_current_status_and_health_roles(self) -> None:
        """Mirror VBA's Type-touched role recalculation."""
        row = dict(ITEM_ROWS[0])
        row.update({
            "Status": "Done",
            "Delivery Health": "Blocked",
        })
        row.pop("DoneDate", None)
        row.pop("BlockedSince", None)
        snapshot = _snapshot(tables={"tblItems": (row, *ITEM_ROWS[1:])})
        changes = _changes(
            _upsert("change-type", workbook_id="I-3001", set_fields={"Type": "Product"})
        )
        result = merge_changes(snapshot, changes, effective_date=self.effective_date)

        self.assertTrue(result.valid)
        changed = result.snapshot.tables["tblItems"][0]
        self.assertEqual(changed["DoneDate"], self.effective_date)
        self.assertEqual(changed["BlockedSince"], self.effective_date)

    def test_delivery_health_enters_and_leaves_blocked_role(self) -> None:
        """Set a blank BlockedSince once and clear it when unblocked."""
        enter = merge_changes(
            _snapshot(),
            _changes(
                _upsert(
                    "block",
                    workbook_id="I-3001",
                    set_fields={"Delivery Health": "Blocked"},
                )
            ),
            effective_date=self.effective_date,
        )
        self.assertEqual(
            enter.snapshot.tables["tblItems"][0]["BlockedSince"],
            self.effective_date,
        )

        row = dict(ITEM_ROWS[0])
        row.update({
            "Delivery Health": "Blocked",
            "BlockedSince": date(2026, 7, 1),
        })
        leave = merge_changes(
            _snapshot(tables={"tblItems": (row, *ITEM_ROWS[1:])}),
            _changes(
                _upsert(
                    "unblock",
                    workbook_id="I-3001",
                    set_fields={"Delivery Health": "On track"},
                )
            ),
            effective_date=self.effective_date,
        )
        self.assertNotIn("BlockedSince", leave.snapshot.tables["tblItems"][0])

    def test_latest_status_edit_sets_or_clears_its_stamp(self) -> None:
        """Date a nonblank narrative and clear its date with the narrative."""
        changed = merge_changes(
            _snapshot(),
            _changes(
                _upsert(
                    "narrative",
                    workbook_id="I-3001",
                    set_fields={"Latest Status": "New narrative"},
                )
            ),
            effective_date=self.effective_date,
        )
        self.assertEqual(
            changed.snapshot.tables["tblItems"][0]["LatestUpdateOn"],
            self.effective_date,
        )
        cleared = merge_changes(
            changed.snapshot,
            _changes(
                _upsert(
                    "clear-narrative",
                    workbook_id="I-3001",
                    clear=["Latest Status"],
                )
            ),
            effective_date=self.effective_date,
        )
        row = cleared.snapshot.tables["tblItems"][0]
        self.assertNotIn("Latest Status", row)
        self.assertNotIn("LatestUpdateOn", row)

    def test_source_attachment_alone_does_not_touch_lifecycle_dates(self) -> None:
        """Keep system identity attachment separate from operational edits."""
        before = _snapshot()
        result = merge_changes(
            before,
            _changes(
                _upsert(
                    "attach-only",
                    workbook_id="I-3002",
                    source=_source("paste:lifecycle", "manual-2"),
                )
            ),
            effective_date=self.effective_date,
        )
        prior = before.tables["tblItems"][1]
        after = result.snapshot.tables["tblItems"][1]
        for field in (
            "Created",
            "Updated",
            "InProgressSince",
            "DoneDate",
            "BlockedSince",
            "LatestUpdateOn",
        ):
            self.assertEqual(after.get(field), prior.get(field), field)

    def test_raid_create_close_reopen_and_delete_follow_closed_role(self) -> None:
        """Stamp Raised/Updated and maintain Closed for every RAID transition."""
        created = merge_changes(
            _snapshot(),
            _changes(
                _upsert(
                    "create-closed-raid",
                    entity="raid",
                    source=_source("paste:lifecycle", "raid-1"),
                    set_fields={
                        "Type": "Risk",
                        "Title": "Closed at entry",
                        "Status": "Closed",
                        "Prob": 2,
                        "Impact": 2,
                    },
                )
            ),
            effective_date=self.effective_date,
        )
        row = created.snapshot.tables["tblRAID"][-1]
        self.assertEqual(row["Raised"], self.effective_date)
        self.assertEqual(row["Updated"], self.effective_date)
        self.assertEqual(row["Closed"], self.effective_date)

        reopened = merge_changes(
            created.snapshot,
            _changes(
                _upsert(
                    "reopen-raid",
                    entity="raid",
                    workbook_id=str(row["RaidID"]),
                    set_fields={"Status": "Open"},
                )
            ),
            effective_date=self.effective_date,
        )
        reopened_row = reopened.snapshot.tables["tblRAID"][-1]
        self.assertNotIn("Closed", reopened_row)
        self.assertEqual(reopened_row["Raised"], self.effective_date)

        deleted = merge_changes(
            reopened.snapshot,
            _changes(
                _mark_deleted(
                    "delete-raid",
                    entity="raid",
                    workbook_id=str(row["RaidID"]),
                )
            ),
            effective_date=date(2026, 7, 16),
        )
        deleted_row = deleted.snapshot.tables["tblRAID"][-1]
        self.assertEqual(deleted_row["Closed"], date(2026, 7, 16))
        self.assertEqual(deleted_row["Updated"], date(2026, 7, 16))


class ValidationRegistryTests(unittest.TestCase):
    """Compare stable baseline and merged red-state finding identities."""

    effective_date = date(2026, 7, 15)

    def _assert_finding(self, code: str, snapshot: object) -> None:
        """Assert that one deliberately invalid snapshot reaches the registry."""
        self.assertIn(code, {finding.code for finding in snapshot_findings(snapshot)})

    def test_unchanged_baseline_finding_is_a_warning(self) -> None:
        """Permit an unrelated edit without hiding a pre-existing red value."""
        invalid = dict(ITEM_ROWS[0])
        invalid["Status"] = "Legacy status"
        snapshot = _snapshot(tables={"tblItems": (invalid, *ITEM_ROWS[1:])})
        result = merge_changes(
            snapshot,
            _changes(
                _upsert(
                    "unrelated-edit",
                    workbook_id="I-3002",
                    set_fields={"Title": "Vendor workstream updated"},
                )
            ),
            effective_date=self.effective_date,
        )
        self.assertTrue(result.valid, result.diagnostics)
        warning_codes = {
            diagnostic.code for diagnostic in result.diagnostics if diagnostic.severity == "warning"
        }
        self.assertIn("choice.item_status", warning_codes)

    def test_touching_an_existing_red_field_must_not_keep_it_as_a_warning(self) -> None:
        """Block an operation that replaces one invalid value with another."""
        invalid = dict(ITEM_ROWS[0])
        invalid["Status"] = "Legacy status"
        snapshot = _snapshot(tables={"tblItems": (invalid, *ITEM_ROWS[1:])})

        result = merge_changes(
            snapshot,
            _changes(
                _upsert(
                    "rewrite-invalid-status",
                    workbook_id="I-3001",
                    set_fields={"Status": "Different invalid status"},
                )
            ),
            effective_date=self.effective_date,
        )

        self.assertFalse(result.valid)
        finding = next(
            diagnostic
            for diagnostic in result.diagnostics
            if diagnostic.code == "choice.item_status"
        )
        self.assertEqual(finding.severity, "error")
        self.assertEqual(finding.operation_id, "rewrite-invalid-status")

    def test_new_config_choice_violation_blocks_atomically(self) -> None:
        """Reject a newly introduced value outside the current Config domain."""
        snapshot = _snapshot()
        result = merge_changes(
            snapshot,
            _changes(
                _upsert(
                    "bad-status",
                    workbook_id="I-3001",
                    set_fields={"Status": "Not configured"},
                )
            ),
            effective_date=self.effective_date,
        )
        self.assertFalse(result.valid)
        self.assertEqual(result.snapshot, snapshot)
        self.assertIn(
            "choice.item_status",
            {diagnostic.code for diagnostic in result.diagnostics},
        )

    def test_hierarchy_requires_parent_and_rejects_cycles(self) -> None:
        """Validate level contracts after all batch references resolve."""
        no_parent = merge_changes(
            _snapshot(),
            _changes(
                _upsert(
                    "orphan-release",
                    source=_source("paste:validation", "release"),
                    set_fields={
                        "Type": "Release",
                        "Title": "Orphan release",
                        "Status": "Backlog",
                    },
                )
            ),
            effective_date=self.effective_date,
        )
        self.assertFalse(no_parent.valid)
        self.assertIn(
            "hierarchy.parent_required",
            {diagnostic.code for diagnostic in no_parent.diagnostics},
        )

        cycle = merge_changes(
            _snapshot(),
            _changes(
                _upsert(
                    "make-cycle",
                    workbook_id="I-3001",
                    set_fields={"Parent": {"workbook_id": "I-3003"}},
                )
            ),
            effective_date=self.effective_date,
        )
        self.assertFalse(cycle.valid)
        self.assertIn(
            "hierarchy.cycle",
            {diagnostic.code for diagnostic in cycle.diagnostics},
        )

    def test_new_parent_cannot_target_deleted_history(self) -> None:
        """Allow retained history links but reject a newly changed Parent link."""
        deleted = dict(ITEM_ROWS[0])
        deleted["Status"] = "Deleted"
        snapshot = _snapshot(tables={"tblItems": (deleted, *ITEM_ROWS[1:])})
        result = merge_changes(
            snapshot,
            _changes(
                _upsert(
                    "link-deleted-parent",
                    workbook_id="I-3003",
                    set_fields={"Parent": {"workbook_id": "I-3001"}},
                )
            ),
            effective_date=self.effective_date,
        )
        self.assertFalse(result.valid)
        self.assertIn(
            "hierarchy.deleted_parent",
            {diagnostic.code for diagnostic in result.diagnostics},
        )

    def test_date_order_and_minimum_date_are_validated(self) -> None:
        """Reject impossible schedules and dates Excel flags in red."""
        for code, set_fields in (
            ("date.before_minimum", {"Start": "2019-12-31"}),
            ("date.order", {"Start": "2026-10-01", "Due": "2026-09-01"}),
        ):
            with self.subTest(code=code):
                result = merge_changes(
                    _snapshot(),
                    _changes(
                        _upsert(
                            "bad-date",
                            workbook_id="I-3001",
                            set_fields=set_fields,
                        )
                    ),
                    effective_date=self.effective_date,
                )
                self.assertFalse(result.valid)
                self.assertIn(code, {diagnostic.code for diagnostic in result.diagnostics})

    def test_alert_raid_requires_valid_probability_and_impact(self) -> None:
        """Apply current alert-type rating rules to created RAID rows."""
        result = merge_changes(
            _snapshot(),
            _changes(
                _upsert(
                    "unrated-risk",
                    entity="raid",
                    source=_source("paste:validation", "risk"),
                    set_fields={
                        "Type": "Risk",
                        "Title": "Unrated risk",
                        "Status": "Open",
                    },
                )
            ),
            effective_date=self.effective_date,
        )
        self.assertFalse(result.valid)
        self.assertIn(
            "raid.rating_required",
            {diagnostic.code for diagnostic in result.diagnostics},
        )

    def test_non_alert_raid_rejects_probability_and_impact(self) -> None:
        """Reject ratings that the workbook marks invalid for a non-alert type."""
        result = merge_changes(
            _snapshot(),
            _changes(
                _upsert(
                    "decision-with-ratings",
                    entity="raid",
                    workbook_id="R-101",
                    set_fields={"Type": "Decision"},
                )
            ),
            effective_date=self.effective_date,
        )
        self.assertFalse(result.valid)
        self.assertIn(
            "raid.rating_not_applicable",
            {diagnostic.code for diagnostic in result.diagnostics},
        )

    def test_unchanged_non_alert_ratings_remain_baseline_warnings(self) -> None:
        """Preserve a legacy invalid rating while rejecting only new findings."""
        legacy = dict(RAID_ROWS[0])
        legacy["Type"] = "Decision"
        snapshot = _snapshot(tables={"tblRAID": (legacy,)})
        result = merge_changes(
            snapshot,
            _changes(
                _upsert(
                    "edit-legacy-decision",
                    entity="raid",
                    workbook_id="R-101",
                    set_fields={"Response": "Legacy scoring retained for review."},
                )
            ),
            effective_date=self.effective_date,
        )
        self.assertTrue(result.valid, result.diagnostics)
        matching = [
            diagnostic
            for diagnostic in result.diagnostics
            if diagnostic.code == "raid.rating_not_applicable"
        ]
        self.assertEqual(len(matching), 2)
        self.assertTrue(all(diagnostic.severity == "warning" for diagnostic in matching))

    def test_incomplete_source_pair_is_reported_without_mutation(self) -> None:
        """Keep both system identity cells paired even in existing rows."""
        partial = dict(ITEM_ROWS[1])
        partial["Source"] = "api:partial"
        snapshot = _snapshot(tables={"tblItems": (ITEM_ROWS[0], partial, ITEM_ROWS[2])})
        result = merge_changes(
            snapshot,
            _changes(),
            effective_date=self.effective_date,
        )
        self.assertTrue(result.valid)
        self.assertIn(
            "source.incomplete_pair",
            {diagnostic.code for diagnostic in result.diagnostics},
        )
        self.assertTrue(all(diagnostic.severity == "warning" for diagnostic in result.diagnostics))

    def test_capacity_breach_blocks_the_complete_batch(self) -> None:
        """Refuse a create once the shared Item capacity is full."""
        rows = tuple(
            {
                "ID": f"I-{10_000 + index}",
                "Type": "Project",
                "Title": f"Capacity row {index}",
                "Status": "Backlog",
                "Created": date(2026, 1, 1),
                "Updated": date(2026, 1, 1),
            }
            for index in range(2_000)
        )
        snapshot = _snapshot(tables={"tblItems": rows})
        result = merge_changes(
            snapshot,
            _changes(
                _upsert(
                    "over-capacity",
                    source=_source("paste:validation", "too-many"),
                    set_fields={
                        "Type": "Project",
                        "Title": "One too many",
                        "Status": "Backlog",
                    },
                )
            ),
            effective_date=self.effective_date,
        )
        self.assertFalse(result.valid)
        self.assertIn(
            "capacity.exceeded",
            {diagnostic.code for diagnostic in result.diagnostics},
        )

    def test_registry_covers_identifier_hierarchy_and_relationship_boundaries(self) -> None:
        """Pin every non-cycle hierarchy/reference finding used by plan comparison."""
        malformed = dict(ITEM_ROWS[0])
        malformed["ID"] = "BAD"

        self_parent = dict(ITEM_ROWS[1])
        self_parent["Parent"] = "I-3002"

        missing_parent = dict(ITEM_ROWS[1])
        missing_parent["Parent"] = "I-9999"

        wrong_level = dict(ITEM_ROWS[0])
        wrong_level["Parent"] = "I-3002"

        dependencies = dict(ITEM_ROWS[0])
        dependencies["BlockedBy"] = "I-3001, I-9999"

        missing_related = dict(RAID_ROWS[0])
        missing_related["RelatedID"] = "I-9999"

        chain = []
        for index in range(7):
            row = dict(ITEM_ROWS[0])
            row.update({
                "ID": f"I-{4001 + index}",
                "Title": f"Depth {index + 1}",
                "Parent": "" if index == 0 else f"I-{4000 + index}",
            })
            chain.append(row)

        cases = (
            (
                "id.malformed",
                _snapshot(tables={"tblItems": (malformed, *ITEM_ROWS[1:])}),
            ),
            (
                "hierarchy.self_parent",
                _snapshot(tables={"tblItems": (ITEM_ROWS[0], self_parent, ITEM_ROWS[2])}),
            ),
            (
                "reference.parent_missing",
                _snapshot(tables={"tblItems": (ITEM_ROWS[0], missing_parent, ITEM_ROWS[2])}),
            ),
            (
                "hierarchy.parent_level",
                _snapshot(tables={"tblItems": (wrong_level, *ITEM_ROWS[1:])}),
            ),
            (
                "reference.blocked_by_self",
                _snapshot(tables={"tblItems": (dependencies, *ITEM_ROWS[1:])}),
            ),
            (
                "reference.blocked_by_missing",
                _snapshot(tables={"tblItems": (dependencies, *ITEM_ROWS[1:])}),
            ),
            (
                "reference.related_item_missing",
                _snapshot(tables={"tblRAID": (missing_related,)}),
            ),
            ("hierarchy.depth", _snapshot(tables={"tblItems": tuple(chain)})),
        )
        for code, snapshot in cases:
            with self.subTest(code=code):
                self._assert_finding(code, snapshot)

    def test_registry_covers_domains_config_lifecycle_and_rating_boundaries(self) -> None:
        """Pin domain, role, chronology, lifecycle and RAID rating findings."""
        cases = (*_item_registry_cases(), *_raid_registry_cases(), *_config_registry_cases())
        for code, snapshot in cases:
            with self.subTest(code=code):
                self._assert_finding(code, snapshot)

    def test_registry_covers_duplicate_missing_people_and_source_boundaries(self) -> None:
        """Pin the remaining Config, identifier, People and source-pair branches."""
        duplicate_config_tables = deepcopy(_snapshot().tables)
        priorities = [dict(row) for row in duplicate_config_tables["tblPriorities"]]
        priorities.append(dict(priorities[0]))
        duplicate_config_tables["tblPriorities"] = tuple(priorities)

        invalid_team_tables = deepcopy(_snapshot().tables)
        people = [dict(row) for row in invalid_team_tables["tblPeople"]]
        people[0]["Team"] = "Unconfigured team"
        invalid_team_tables["tblPeople"] = tuple(people)

        duplicate_id = dict(ITEM_ROWS[1])
        duplicate_id["ID"] = ITEM_ROWS[0]["ID"]

        missing_id = dict(ITEM_ROWS[0])
        missing_id["ID"] = ""

        invalid_source = dict(ITEM_ROWS[0])
        invalid_source["Source"] = 42
        invalid_source["Source ID"] = "opaque-42"

        cases = (
            (
                "config.duplicate_key",
                replace(_snapshot(), tables=duplicate_config_tables),
            ),
            (
                "choice.person_team",
                replace(_snapshot(), tables=invalid_team_tables),
            ),
            (
                "id.duplicate",
                _snapshot(tables={"tblItems": (ITEM_ROWS[0], duplicate_id)}),
            ),
            ("id.missing", _snapshot(tables={"tblItems": (missing_id,)})),
            (
                "source.invalid_value",
                _snapshot(tables={"tblItems": (invalid_source,)}),
            ),
        )
        for code, snapshot in cases:
            with self.subTest(code=code):
                self._assert_finding(code, snapshot)


if __name__ == "__main__":
    unittest.main()
