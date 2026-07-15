"""Regression contracts for the monday.com board import."""

from __future__ import annotations

import json
import tempfile
import unittest
from datetime import date
from pathlib import Path
from typing import TYPE_CHECKING

from .. import pipeline
from ..data.export import export_workbook
from ..data.inject import injected_source, validate_snapshot
from ..data.monday import (
    MAX_ATTEMPTS,
    BoardData,
    MondayImportError,
    PostResult,
    fetch_board,
    map_board,
    read_map,
    write_map,
)
from ..data.snapshot import Snapshot
from .test_data_layer import ITEM_ROWS, _snapshot

if TYPE_CHECKING:
    from collections.abc import Callable

BOARD_ID = 4321098765
SUB_BOARD = "9990001"
CREDENTIAL = "unit-test-credential"

_COLUMNS = [
    {"id": "status", "title": "Status", "type": "status"},
    {"id": "people1", "title": "Owner", "type": "people"},
    {"id": "date4", "title": "Due date", "type": "date"},
    {"id": "timerange", "title": "Timeline", "type": "timeline"},
    {"id": "status_2", "title": "Priority", "type": "status"},
    {"id": "text7", "title": "Type", "type": "text"},
]
# The subitems board carries its own column identifiers for the same titles.
_SUB_COLUMNS = [
    {"id": "status_s", "title": "Status", "type": "status"},
    {"id": "date_s", "title": "Due date", "type": "date"},
    {"id": "people_s", "title": "Owner", "type": "people"},
]


def _value(column_id: str, text: str | None, value: object) -> dict[str, object]:
    return {"id": column_id, "text": text, "value": None if value is None else json.dumps(value)}


def _epic_node() -> dict[str, object]:
    return {
        "id": "9001",
        "name": "Supplier onboarding",
        "created_at": "2026-05-01T09:30:00Z",
        "updated_at": "2026-07-01T10:00:00Z",
        "board": {"id": str(BOARD_ID)},
        "parent_item": None,
        "column_values": [
            _value("status", "In Progress", {"index": 2}),
            _value("people1", "Ana Ruiz, Ben Cole", {"personsAndTeams": [{"id": 1}]}),
            _value(
                "timerange",
                "2026-06-01 - 2026-08-30",
                {"from": "2026-06-01", "to": "2026-08-30"},
            ),
            _value("text7", "Epic", "Epic"),
        ],
        "subitems": [
            {
                "id": "9101",
                "name": "Contract draft",
                "created_at": "2026-05-02T08:00:00Z",
                "updated_at": "2026-06-21T07:45:00Z",
                "board": {"id": SUB_BOARD},
                "parent_item": {"id": "9001"},
                "column_values": [
                    _value("status_s", "Done", {"index": 1}),
                    _value("date_s", "20 Jun 2026", {"date": "2026-06-20", "time": None}),
                    _value("people_s", "Ben Cole", {"personsAndTeams": [{"id": 3}]}),
                ],
            }
        ],
    }


def _task_node() -> dict[str, object]:
    # The date value is UTC 14 Jul late evening; a caller-timezone rendering
    # would display 15 Jul. The importer must use the raw UTC value.
    return {
        "id": "9002",
        "name": "Data migration rehearsal",
        "created_at": "2026-06-10T12:00:00Z",
        "updated_at": "2026-07-10T16:20:00Z",
        "board": {"id": str(BOARD_ID)},
        "parent_item": None,
        "column_values": [
            _value("status", "Working on it", {"index": 0}),
            _value("people1", "Cara Nowak", {"personsAndTeams": [{"id": 2}]}),
            _value("date4", "15 Jul 2026", {"date": "2026-07-14", "time": "23:30:00"}),
            _value("status_2", "High", {"index": 1}),
        ],
        "subitems": [],
    }


def _document(data: dict[str, object]) -> PostResult:
    return PostResult(status=200, retry_after=None, document={"data": data})


def _combined_first(items: list[dict[str, object]], cursor: str | None) -> PostResult:
    return _document({
        "boards": [
            {
                "id": str(BOARD_ID),
                "name": "Delivery board",
                "columns": _COLUMNS,
                "items_page": {"cursor": cursor, "items": items},
            }
        ]
    })


def _next_page(items: list[dict[str, object]], cursor: str | None) -> PostResult:
    return _document({"next_items_page": {"cursor": cursor, "items": items}})


def _catalog_response() -> PostResult:
    return _document({"boards": [{"id": SUB_BOARD, "columns": _SUB_COLUMNS}]})


def _paged_post(pages: list[PostResult]) -> Callable[[str, dict[str, object]], PostResult]:
    remaining = list(pages)

    def post(_token: str, payload: dict[str, object]) -> PostResult:
        query = str(payload["query"])
        if "columns { id title type }" in query and "items_page" not in query:
            return _catalog_response()
        return remaining.pop(0)

    return post


def _no_sleep(_seconds: float) -> None:
    return


def _fetch_board_with(pages: list[PostResult]) -> BoardData:
    return fetch_board(BOARD_ID, token=CREDENTIAL, post=_paged_post(pages), sleeper=_no_sleep)


def _fetch_default_board() -> BoardData:
    return _fetch_board_with([
        _combined_first([_epic_node()], "cursor-1"),
        _next_page([_task_node()], None),
    ])


def _without_rows(snapshot: Snapshot, row_ids: set[str]) -> Snapshot:
    tables = dict(snapshot.tables)
    tables["tblItems"] = tuple(
        row for row in snapshot.tables["tblItems"] if str(row.get("ID")) not in row_ids
    )
    return Snapshot(
        schema_fingerprint=snapshot.schema_fingerprint,
        workbook_digest=snapshot.workbook_digest,
        exported_at=snapshot.exported_at,
        settings=dict(snapshot.settings),
        tables=tables,
    )


class MondayFetchTests(unittest.TestCase):
    """Prove pagination, catalog fetching, retries and error surfacing."""

    def test_pagination_flattens_items_and_fetches_subitem_catalogs(self) -> None:
        """Walk the cursor chain and catalogue every board the items live on."""
        board = _fetch_default_board()
        self.assertEqual(board.name, "Delivery board")
        self.assertEqual(
            [(item.identifier, item.parent_id, item.board_id) for item in board.items],
            [
                ("9001", None, str(BOARD_ID)),
                ("9101", "9001", SUB_BOARD),
                ("9002", None, str(BOARD_ID)),
            ],
        )
        self.assertEqual(board.catalogs[str(BOARD_ID)]["date4"], ("Due date", "date"))
        self.assertEqual(board.catalogs[SUB_BOARD]["date_s"], ("Due date", "date"))

    def test_rate_limited_requests_honor_retry_after(self) -> None:
        """Wait the server-directed delay before retrying HTTP 429."""
        waits: list[float] = []
        board = fetch_board(
            BOARD_ID,
            token=CREDENTIAL,
            post=_paged_post([
                PostResult(status=429, retry_after=7, document=None),
                _combined_first([_task_node()], None),
            ]),
            sleeper=waits.append,
        )
        self.assertEqual(len(board.items), 1)
        self.assertEqual(waits, [7])

    def test_complexity_budget_exhaustion_waits_and_retries(self) -> None:
        """Honor retry_in_seconds carried inside a GraphQL error."""
        waits: list[float] = []
        throttled = PostResult(
            status=200,
            retry_after=None,
            document={
                "errors": [
                    {
                        "message": "Complexity budget exhausted",
                        "extensions": {"code": "ComplexityException", "retry_in_seconds": 12},
                    }
                ]
            },
        )
        board = fetch_board(
            BOARD_ID,
            token=CREDENTIAL,
            post=_paged_post([throttled, _combined_first([_task_node()], None)]),
            sleeper=waits.append,
        )
        self.assertEqual(len(board.items), 1)
        self.assertEqual(waits, [12])

    def test_transient_network_failures_retry(self) -> None:
        """Treat dropped connections as retryable, not as raw tracebacks."""
        waits: list[float] = []
        attempts = iter([
            ConnectionResetError("connection reset by peer"),
            TimeoutError("timed out"),
            _combined_first([_task_node()], None),
        ])

        def post(_token: str, _payload: dict[str, object]) -> PostResult:
            outcome = next(attempts)
            if isinstance(outcome, BaseException):
                raise outcome
            return outcome

        board = fetch_board(BOARD_ID, token=CREDENTIAL, post=post, sleeper=waits.append)
        self.assertEqual(len(board.items), 1)
        self.assertEqual(waits, [2, 4])

    def test_exhaustion_reports_the_last_failure(self) -> None:
        """Name the real final failure instead of blaming rate limits."""
        failing = PostResult(status=502, retry_after=None, document=None)
        waits: list[float] = []
        with self.assertRaisesRegex(MondayImportError, "last failure: HTTP 502"):
            fetch_board(
                BOARD_ID,
                token=CREDENTIAL,
                post=lambda _token, _payload: failing,
                sleeper=waits.append,
            )
        self.assertEqual(len(waits), MAX_ATTEMPTS - 1)

    def test_api_errors_surface_verbatim(self) -> None:
        """Report GraphQL errors instead of returning partial data."""
        failing = PostResult(
            status=200,
            retry_after=None,
            document={"errors": [{"message": "User unauthorized to perform action"}]},
        )
        with self.assertRaisesRegex(MondayImportError, "User unauthorized"):
            fetch_board(
                BOARD_ID,
                token=CREDENTIAL,
                post=lambda _token, _payload: failing,
                sleeper=_no_sleep,
            )

    def test_http_failures_surface_with_status(self) -> None:
        """Report non-retryable statuses with their code."""
        failing = PostResult(status=401, retry_after=None, document=None)
        with self.assertRaisesRegex(MondayImportError, "HTTP 401"):
            fetch_board(
                BOARD_ID,
                token=CREDENTIAL,
                post=lambda _token, _payload: failing,
                sleeper=_no_sleep,
            )

    def test_missing_board_is_reported(self) -> None:
        """Distinguish an absent board from an empty one."""
        empty = _document({"boards": []})
        with self.assertRaisesRegex(MondayImportError, str(BOARD_ID)):
            fetch_board(
                BOARD_ID,
                token=CREDENTIAL,
                post=lambda _token, _payload: empty,
                sleeper=_no_sleep,
            )

    def test_stalled_cursor_is_reported(self) -> None:
        """Refuse pagination that stops advancing instead of spinning."""
        stalled = [_combined_first([_task_node()], "same-cursor")] + [
            _next_page([], "same-cursor") for _ in range(3)
        ]
        with self.assertRaisesRegex(MondayImportError, "not advancing"):
            _fetch_board_with(stalled)

    def test_bad_timestamps_are_reported_as_timestamps(self) -> None:
        """Name unparseable stamps for what they are."""
        node = _task_node()
        node["created_at"] = "not-a-stamp"
        with self.assertRaisesRegex(MondayImportError, "unparseable created_at timestamp"):
            _fetch_board_with([_combined_first([node], None)])


class MondayMappingTests(unittest.TestCase):
    """Prove board rows map, merge and re-import faithfully."""

    def setUp(self) -> None:
        """Fetch the fixture board once per test."""
        self.board = _fetch_default_board()
        self.base = _snapshot()

    def test_fresh_import_builds_the_hierarchy(self) -> None:
        """Create a container, wire parents and honor the Type column."""
        outcome = map_board(self.board, self.base, {}, allow_duplicates=False)
        rows = {str(row["ID"]): row for row in outcome.snapshot.tables["tblItems"]}
        self.assertEqual(outcome.added, 4)
        self.assertEqual(outcome.updated, 0)

        container = rows["I-3004"]
        self.assertEqual(container["Type"], "Project")
        self.assertEqual(container["Title"], "monday.com — Delivery board")

        epic = rows["I-3005"]
        self.assertEqual(epic["Type"], "Epic")
        self.assertEqual(epic["Parent"], "I-3004")
        self.assertEqual(epic["Owner"], "Ana Ruiz")
        self.assertEqual(epic["Start"], date(2026, 6, 1))
        self.assertEqual(epic["Due"], date(2026, 8, 30))
        self.assertEqual(epic["Created"], date(2026, 5, 1))

    def test_subitem_values_map_through_their_own_board_catalog(self) -> None:
        """Carry subitem Status, Owner and dates despite distinct column ids."""
        outcome = map_board(self.board, self.base, {}, allow_duplicates=False)
        rows = {str(row["ID"]): row for row in outcome.snapshot.tables["tblItems"]}
        subitem = rows["I-3006"]
        self.assertEqual(subitem["Type"], "Sub Task")
        self.assertEqual(subitem["Parent"], "I-3005")
        self.assertEqual(subitem["Status"], "Done")
        self.assertEqual(subitem["Owner"], "Ben Cole")
        self.assertEqual(subitem["Due"], date(2026, 6, 20))

    def test_dates_come_from_utc_values_not_rendered_text(self) -> None:
        """Never let a caller-timezone rendering shift a date."""
        outcome = map_board(self.board, self.base, {}, allow_duplicates=False)
        rows = {str(row["ID"]): row for row in outcome.snapshot.tables["tblItems"]}
        self.assertEqual(rows["I-3007"]["Due"], date(2026, 7, 14))

    def test_dropped_co_owners_are_reported(self) -> None:
        """Keep one Owner but never drop co-owners silently."""
        outcome = map_board(self.board, self.base, {}, allow_duplicates=False)
        self.assertTrue(any("co-owner" in note and "Ben Cole" in note for note in outcome.notes))

    def test_counter_and_people_follow_the_import(self) -> None:
        """Advance the ID counter and register supplied owners."""
        outcome = map_board(self.board, self.base, {}, allow_duplicates=False)
        self.assertEqual(outcome.snapshot.settings["cfgNextItemID"], 3008)
        people = [row["Person"] for row in outcome.snapshot.tables["tblPeople"]]
        self.assertIn("Cara Nowak", people)
        self.assertNotIn("Ana Ruiz, Ben Cole", people)

    def test_merged_snapshot_passes_the_injection_gates(self) -> None:
        """Surface out-of-list monday values as warnings, not failures."""
        outcome = map_board(self.board, self.base, {}, allow_duplicates=False)
        reconciliation = validate_snapshot(outcome.snapshot)
        self.assertEqual(reconciliation.counter_bumps, {})
        self.assertTrue(any("Working on it" in warning for warning in reconciliation.warnings))
        self.assertTrue(any("High" in warning for warning in reconciliation.warnings))

    def test_reimport_updates_in_place_without_duplicates(self) -> None:
        """Key re-imports on the identifier map, not titles or position."""
        first = map_board(self.board, self.base, {}, allow_duplicates=False)
        changed_task = _task_node()
        changed_task["column_values"] = [_value("status", "Done", {"index": 1})]
        second_board = _fetch_board_with([_combined_first([_epic_node(), changed_task], None)])
        second = map_board(
            second_board,
            first.snapshot,
            first.map_document,
            allow_duplicates=False,
        )
        self.assertEqual(second.added, 0)
        self.assertEqual(second.updated, 1)
        self.assertEqual(second.unchanged, 2)
        rows = {str(row["ID"]): row for row in second.snapshot.tables["tblItems"]}
        self.assertEqual(rows["I-3007"]["Status"], "Done")
        self.assertEqual(second.snapshot.settings["cfgNextItemID"], 3008)
        self.assertEqual(
            len(second.snapshot.tables["tblItems"]),
            len(first.snapshot.tables["tblItems"]),
        )

    def test_cleared_monday_fields_clear_on_reimport_with_notes(self) -> None:
        """Propagate removals of Priority, Owner, Start and Due, and say so."""
        first = map_board(self.board, self.base, {}, allow_duplicates=False)
        stripped_task = _task_node()
        stripped_task["column_values"] = [_value("status", "Done", {"index": 1})]
        second_board = _fetch_board_with([_combined_first([_epic_node(), stripped_task], None)])
        second = map_board(
            second_board,
            first.snapshot,
            first.map_document,
            allow_duplicates=False,
        )
        rows = {str(row["ID"]): row for row in second.snapshot.tables["tblItems"]}
        task = rows["I-3007"]
        self.assertNotIn("Due", task)
        self.assertNotIn("Owner", task)
        self.assertNotIn("Priority", task)
        self.assertEqual(task["Status"], "Done")
        for field_name in ("Due", "Owner", "Priority"):
            self.assertTrue(
                any(field_name in note and "cleared in monday" in note for note in second.notes)
            )

    def test_empty_monday_status_keeps_the_workbook_status(self) -> None:
        """Never flip-flop Status, the one field the importer defaults."""
        first = map_board(self.board, self.base, {}, allow_duplicates=False)
        statusless_task = _task_node()
        statusless_task["column_values"] = [
            _value("date4", "15 Jul 2026", {"date": "2026-07-14", "time": "23:30:00"}),
        ]
        second_board = _fetch_board_with([_combined_first([_epic_node(), statusless_task], None)])
        second = map_board(
            second_board,
            first.snapshot,
            first.map_document,
            allow_duplicates=False,
        )
        rows = {str(row["ID"]): row for row in second.snapshot.tables["tblItems"]}
        self.assertEqual(rows["I-3007"]["Status"], "Working on it")
        self.assertFalse(any("Status" in note and "cleared" in note for note in second.notes))

    def test_rows_deleted_in_the_workbook_stay_deleted(self) -> None:
        """Never resurrect a row the user removed."""
        first = map_board(self.board, self.base, {}, allow_duplicates=False)
        pruned = _without_rows(first.snapshot, {"I-3007"})
        second = map_board(self.board, pruned, first.map_document, allow_duplicates=False)
        self.assertEqual(second.added, 0)
        self.assertTrue(any("I-3007" in entry for entry in second.skipped_deleted))
        row_ids = {str(row["ID"]) for row in second.snapshot.tables["tblItems"]}
        self.assertNotIn("I-3007", row_ids)

    def test_new_subitem_under_deleted_parent_attaches_to_the_container(self) -> None:
        """Fall back to the container instead of pointing at a deleted row."""
        first = map_board(self.board, self.base, {}, allow_duplicates=False)
        pruned = _without_rows(first.snapshot, {"I-3005"})
        epic = _epic_node()
        epic["subitems"] = [
            *epic["subitems"],
            {
                "id": "9102",
                "name": "Signature round",
                "created_at": "2026-07-01T08:00:00Z",
                "updated_at": "2026-07-12T08:00:00Z",
                "board": {"id": SUB_BOARD},
                "parent_item": {"id": "9001"},
                "column_values": [_value("status_s", "In Progress", {"index": 0})],
            },
        ]
        second_board = _fetch_board_with([_combined_first([epic, _task_node()], None)])
        second = map_board(second_board, pruned, first.map_document, allow_duplicates=False)
        rows = {str(row["ID"]): row for row in second.snapshot.tables["tblItems"]}
        new_row = next(row for row in rows.values() if row.get("Title") == "Signature round")
        self.assertEqual(new_row["Parent"], "I-3004")
        self.assertTrue(any("attached to the board container" in note for note in second.notes))

    def test_deleted_container_is_only_recreated_when_needed(self) -> None:
        """Honor container deletion until new items genuinely need a parent."""
        first = map_board(self.board, self.base, {}, allow_duplicates=False)
        pruned = _without_rows(first.snapshot, {"I-3004"})
        unchanged = map_board(self.board, pruned, first.map_document, allow_duplicates=False)
        self.assertEqual(unchanged.added, 0)
        self.assertFalse(any("recreated" in note for note in unchanged.notes))

        extra = _task_node()
        extra["id"] = "9003"
        extra["name"] = "Cutover checklist"
        grown_board = _fetch_board_with([
            _combined_first([_epic_node(), _task_node(), extra], None)
        ])
        grown = map_board(grown_board, pruned, first.map_document, allow_duplicates=False)
        self.assertEqual(grown.added, 2)
        self.assertTrue(any("recreated" in note for note in grown.notes))

    def test_lost_map_with_matching_titles_halts(self) -> None:
        """Refuse a silent duplicate wave when the identifier map is gone."""
        first = map_board(self.board, self.base, {}, allow_duplicates=False)
        with self.assertRaisesRegex(MondayImportError, "Supplier onboarding"):
            map_board(self.board, first.snapshot, {}, allow_duplicates=False)
        forced = map_board(self.board, first.snapshot, {}, allow_duplicates=True)
        self.assertEqual(forced.added, 4)

    def test_unknown_type_labels_fall_back_with_a_note(self) -> None:
        """Default unlevelable monday types instead of breaking the build."""
        node = _task_node()
        node["column_values"] = [_value("text7", "Gremlin", "Gremlin")]
        board = _fetch_board_with([_combined_first([node], None)])
        outcome = map_board(board, self.base, {}, allow_duplicates=False)
        rows = {str(row["ID"]): row for row in outcome.snapshot.tables["tblItems"]}
        self.assertEqual(rows["I-3005"]["Type"], "Task")
        self.assertTrue(any("Gremlin" in note for note in outcome.notes))

    def test_missing_level_one_type_halts(self) -> None:
        """Refuse to import when no type can hold the board container."""
        no_projects = tuple(row for row in self.base.tables["tblTypes"] if row["Level"] != 1)
        snapshot = _snapshot(tables={"tblTypes": no_projects, "tblItems": ()})
        with self.assertRaisesRegex(MondayImportError, "no Level 1 type"):
            map_board(self.board, snapshot, {}, allow_duplicates=False)


class MondayMapDocumentTests(unittest.TestCase):
    """Protect the persisted identifier map and its interruption detector."""

    def test_map_round_trips_and_tolerates_absence(self) -> None:
        """Read back exactly what was written; missing maps start empty."""
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory)
            self.assertEqual(read_map(target, BOARD_ID), {})
            document = {
                "format": 1,
                "board": str(BOARD_ID),
                "board_name": "Delivery board",
                "container": "I-3004",
                "items": {"9001": "I-3005"},
            }
            write_map(target, BOARD_ID, document)
            self.assertEqual(read_map(target, BOARD_ID), document)

    def test_malformed_map_is_rejected(self) -> None:
        """Refuse to guess when the identifier map cannot be trusted."""
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory)
            (target / f"board-{BOARD_ID}.json").write_text("{]", encoding="utf-8")
            with self.assertRaisesRegex(MondayImportError, "malformed"):
                read_map(target, BOARD_ID)

    def test_interrupted_import_is_detected(self) -> None:
        """Halt with recovery steps when a staged map was never promoted."""
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory)
            (target / f"board-{BOARD_ID}.pending.json").write_text("{}", encoding="utf-8")
            with self.assertRaisesRegex(MondayImportError, "interrupted"):
                read_map(target, BOARD_ID)


class MondayWorkbookRoundTripTests(unittest.TestCase):
    """Prove imported rows land in a real built workbook."""

    def test_import_survives_build_and_export(self) -> None:
        """Carry every imported row through injection into a real package."""
        board = _fetch_default_board()
        outcome = map_board(board, _snapshot(), {}, allow_duplicates=False)
        reconciliation = validate_snapshot(outcome.snapshot)
        with tempfile.TemporaryDirectory() as directory:
            workbook = Path(directory) / "imported.xlsx"
            with injected_source(outcome.snapshot, reconciliation):
                pipeline.build_one(workbook, with_vba=False)
            result = export_workbook(workbook)
        self.assertEqual(
            list(result.snapshot.tables["tblItems"]),
            list(outcome.snapshot.tables["tblItems"]),
        )
        self.assertEqual(
            [row["ID"] for row in result.snapshot.tables["tblItems"][len(ITEM_ROWS) :]],
            ["I-3004", "I-3005", "I-3006", "I-3007"],
        )


if __name__ == "__main__":
    unittest.main()
