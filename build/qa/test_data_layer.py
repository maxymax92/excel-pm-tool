"""Regression contracts for the authored-data export, snapshot and injection layer."""

from __future__ import annotations

import json
import tempfile
import unittest
from datetime import date, datetime, time
from pathlib import Path

from .. import pipeline
from ..data.export import (
    ExportError,
    ExportResult,
    _coerce_date,
    _coerce_text,
    export_workbook,
)
from ..data.inject import InjectError, injected_source, validate_snapshot
from ..data.migrate import MigrationError, _require_workbook
from ..data.schema import EXAMPLE_MARKER, SETTINGS_DEFAULTS, schema_fingerprint
from ..data.snapshot import (
    RING_LIMIT,
    Snapshot,
    SnapshotError,
    read_snapshot,
    write_snapshot,
)
from ..spec import config, examples
from ..spec.capacity import DATA_ROWS


def _is_marked(row: dict[str, object]) -> bool:
    return any(isinstance(value, str) and EXAMPLE_MARKER in value for value in row.values())


FIXTURE_SETTINGS = {
    "cfgDueSoonDays": 9,
    "cfgItemIDPrefix": "I-",
    "cfgNextItemID": 3004,
    "cfgRaidIDPrefix": "R-",
    "cfgNextRaidID": 102,
}

ITEM_ROWS = (
    {
        "ID": "I-3001",
        "Title": "Rollout programme",
        "Type": "Project",
        "Status": "In Progress",
        "Delivery Health": "On track",
        "Priority": "P1",
        "Owner": "Ana Ruiz",
        "Start": date(2026, 6, 1),
        "Due": date(2026, 9, 30),
        "Latest Status": "Tracking to plan.",
        "Created": date(2026, 5, 20),
        "Updated": date(2026, 7, 10),
        "LatestUpdateOn": date(2026, 7, 10),
    },
    {
        "ID": "I-3002",
        "Title": "Vendor workstream",
        "Type": "Workstream",
        "Parent": "I-3001",
        "Status": "In Progress",
        "Priority": "P2",
        "Owner": "Ana Ruiz",
        "Start": date(2026, 6, 5),
        "Due": date(2026, 8, 14),
        "Latest Status": "Contracting is underway.",
        "Created": date(2026, 5, 25),
        "Updated": date(2026, 7, 8),
        "InProgressSince": date(2026, 6, 5),
        "LatestUpdateOn": date(2026, 7, 8),
    },
    {
        "ID": "I-3003",
        "Title": "Contract signed",
        "Type": "Task",
        "Parent": "I-3002",
        "Status": "Done",
        "Priority": "P2",
        "Owner": "Ben Cole",
        "Start": date(2026, 6, 5),
        "Due": date(2026, 6, 20),
        "Latest Status": "Signed and archived.",
        "Created": date(2026, 5, 25),
        "Updated": date(2026, 6, 20),
        "InProgressSince": date(2026, 6, 5),
        "DoneDate": date(2026, 6, 20),
        "LatestUpdateOn": date(2026, 6, 20),
    },
)

RAID_ROWS = (
    {
        "RaidID": "R-101",
        "Type": "Risk",
        "Title": "Vendor capacity may slip",
        "Detail": "Contracting depends on one supplier team.",
        "RelatedID": "I-3002",
        "Owner": "Ana Ruiz",
        "Status": "Open",
        "Prob": 3,
        "Impact": 4,
        "Response": "Secure a named backup supplier.",
        "NextReview": date(2026, 7, 20),
        "Raised": date(2026, 7, 1),
        "Updated": date(2026, 7, 10),
    },
)


def _config_tables() -> dict[str, tuple[dict[str, object], ...]]:
    return {
        "tblStatuses": tuple(
            {"Status": status, "IsActive": active, "IsDone": done, "IsCancelled": cancelled}
            for status, active, done, cancelled in config.STATUSES
        ),
        "tblTypes": (
            *({"Type": name, "Level": level} for name, level in config.TYPES),
            {"Type": "Workstream", "Level": 3},
        ),
        "tblPriorities": tuple({"Priority": priority} for priority in config.PRIORITIES),
        "tblTeams": ({"Team": "Alpha"}, {"Team": "Beta"}),
        "tblRaidTypes": tuple(
            {"RaidType": name, "IsAlert": alert, "IsDecision": decision}
            for name, alert, decision in config.RAID_TYPES
        ),
        "tblRaidStatuses": tuple(
            {"RaidStatus": name, "IsClosed": closed} for name, closed in config.RAID_STATUSES
        ),
        "tblSeverity": tuple(
            {"Severity": severity, "MinScore": score} for severity, score in config.SEVERITY
        ),
        "tblDeliveryHealth": tuple({"Delivery Health": value} for value in config.DELIVERY_HEALTH),
        "tblPeople": (
            {"Person": "Ana Ruiz", "Role": "Delivery Lead", "Team": "Alpha"},
            {"Person": "Ben Cole", "Role": "Engineer", "Team": "Beta"},
        ),
    }


def _snapshot(
    *,
    settings: dict[str, object] | None = None,
    tables: dict[str, tuple[dict[str, object], ...]] | None = None,
) -> Snapshot:
    merged_tables = {
        **_config_tables(),
        "tblItems": tuple(dict(row) for row in ITEM_ROWS),
        "tblRAID": tuple(dict(row) for row in RAID_ROWS),
    }
    if tables:
        merged_tables.update(tables)
    return Snapshot(
        schema_fingerprint=schema_fingerprint(),
        workbook_digest="0" * 64,
        exported_at="2026-07-15T00:00:00Z",
        settings=dict(FIXTURE_SETTINGS if settings is None else settings),
        tables=merged_tables,
    )


def _build_exported(directory: str, *, injected: bool) -> ExportResult:
    workbook = Path(directory) / "fixture.xlsx"
    if injected:
        snapshot = _snapshot()
        with injected_source(snapshot, validate_snapshot(snapshot)):
            pipeline.build_one(workbook, with_vba=False)
    else:
        pipeline.build_one(workbook, with_vba=False)
    return export_workbook(workbook)


class DataRoundTripTests(unittest.TestCase):
    """Prove injected rows and settings export back exactly."""

    snapshot: Snapshot
    result: ExportResult

    @classmethod
    def setUpClass(cls) -> None:
        """Build one injected workbook and export it."""
        super().setUpClass()
        directory = tempfile.TemporaryDirectory()
        cls.addClassCleanup(directory.cleanup)
        cls.snapshot = _snapshot()
        cls.result = _build_exported(directory.name, injected=True)

    def test_every_table_round_trips_exactly(self) -> None:
        """Preserve every authored row, value and order through build and export."""
        for table, rows in self.snapshot.tables.items():
            with self.subTest(table=table):
                self.assertEqual(list(self.result.snapshot.tables[table]), list(rows))

    def test_settings_round_trip_as_the_merged_band(self) -> None:
        """Export the full settings band exactly as injected."""
        reconciliation = validate_snapshot(self.snapshot)
        self.assertEqual(self.result.snapshot.settings, reconciliation.settings)

    def test_reconciliation_is_quiet_for_consistent_data(self) -> None:
        """Report no bumps, warnings or empty tables for coherent fixtures."""
        reconciliation = validate_snapshot(self.snapshot)
        self.assertEqual(reconciliation.counter_bumps, {})
        self.assertEqual(reconciliation.warnings, ())
        self.assertEqual(reconciliation.empty_tables, ())
        self.assertTrue(reconciliation.fingerprint_matches)

    def test_export_records_the_current_schema_fingerprint(self) -> None:
        """Stamp snapshots with the schema they were exported under."""
        self.assertEqual(self.result.snapshot.schema_fingerprint, schema_fingerprint())


class ShippedExampleExportTests(unittest.TestCase):
    """Pin the example-row skip rule against the shipped fixtures."""

    result: ExportResult

    @classmethod
    def setUpClass(cls) -> None:
        """Build the default example workbook and export it."""
        super().setUpClass()
        directory = tempfile.TemporaryDirectory()
        cls.addClassCleanup(directory.cleanup)
        cls.result = _build_exported(directory.name, injected=False)

    def test_marked_example_rows_are_skipped(self) -> None:
        """Skip only rows carrying the exact delete-me marker."""
        marked_items = [row for row in examples.ITEMS_EXAMPLES if _is_marked(row)]
        marked_raid = [row for row in examples.RAID_EXAMPLES if _is_marked(row)]
        self.assertTrue(marked_items)
        self.assertEqual(self.result.skipped_examples.get("tblItems"), len(marked_items))
        self.assertEqual(self.result.skipped_examples.get("tblRAID", 0), len(marked_raid))

    def test_unmarked_example_like_rows_are_retained_and_noted(self) -> None:
        """Never drop a row that merely mentions the word example."""
        retained_rows = [row for row in examples.RAID_EXAMPLES if not _is_marked(row)]
        raid_ids = [row.get("RaidID") for row in self.result.snapshot.tables["tblRAID"]]
        self.assertEqual(raid_ids, [str(row["RaidID"]) for row in retained_rows])
        for row in retained_rows:
            if "EXAMPLE" in str(row).upper():
                raid_id = str(row["RaidID"])
                self.assertTrue(any(raid_id in note for note in self.result.notes))

    def test_default_settings_export_exactly(self) -> None:
        """Read the shipped settings band back verbatim."""
        self.assertEqual(self.result.snapshot.settings, SETTINGS_DEFAULTS)


class InjectionGateTests(unittest.TestCase):
    """Halt on data loss or unbuildable snapshots; warn on red-flag values."""

    def test_duplicate_ids_halt(self) -> None:
        """Refuse duplicate keys before any build."""
        rows = (dict(ITEM_ROWS[0]), dict(ITEM_ROWS[0]))
        with self.assertRaisesRegex(InjectError, "duplicate ID values: I-3001"):
            validate_snapshot(_snapshot(tables={"tblItems": rows}))

    def test_orphan_columns_halt(self) -> None:
        """Refuse to drop values from columns outside the schema."""
        row = dict(ITEM_ROWS[0])
        row["Estimate"] = "13d"
        with self.assertRaisesRegex(InjectError, "Estimate"):
            validate_snapshot(_snapshot(tables={"tblItems": (row,)}))

    def test_orphan_settings_halt(self) -> None:
        """Refuse to drop settings outside the schema."""
        settings = dict(FIXTURE_SETTINGS)
        settings["cfgRetired"] = 4
        with self.assertRaisesRegex(InjectError, "cfgRetired"):
            validate_snapshot(_snapshot(settings=settings))

    def test_unknown_item_type_halts(self) -> None:
        """Refuse types the taxonomy cannot level."""
        row = dict(ITEM_ROWS[0])
        row["Type"] = "Ghost"
        with self.assertRaisesRegex(InjectError, "Ghost"):
            validate_snapshot(_snapshot(tables={"tblItems": (row,)}))

    def test_capacity_breach_halts(self) -> None:
        """Refuse more rows than every workbook layer supports."""
        rows = tuple(
            {"ID": f"I-{4000 + index}", "Title": "Bulk", "Type": "Task", "Status": "Backlog"}
            for index in range(DATA_ROWS + 1)
        )
        with self.assertRaisesRegex(InjectError, str(DATA_ROWS)):
            validate_snapshot(_snapshot(tables={"tblItems": rows}))

    def test_stale_counters_are_bumped_and_reported(self) -> None:
        """Keep the VBA ID counters above every existing identifier."""
        settings = dict(FIXTURE_SETTINGS)
        settings["cfgNextItemID"] = 1
        reconciliation = validate_snapshot(_snapshot(settings=settings))
        self.assertEqual(reconciliation.counter_bumps, {"cfgNextItemID": (1, 3004)})

    def test_red_flag_values_warn_without_halting(self) -> None:
        """Report list-breaking values the workbook itself flags in red."""
        row = dict(ITEM_ROWS[0])
        row["Status"] = "Bogus"
        reconciliation = validate_snapshot(_snapshot(tables={"tblItems": (row,)}))
        self.assertTrue(any("Bogus" in warning for warning in reconciliation.warnings))

    def test_missing_type_halts_before_the_build(self) -> None:
        """Refuse rows the outline writer would crash on."""
        row = dict(ITEM_ROWS[0])
        del row["Type"]
        with self.assertRaisesRegex(InjectError, "have no Type: I-3001"):
            validate_snapshot(_snapshot(tables={"tblItems": (row,)}))

    def test_missing_title_halts_before_the_build(self) -> None:
        """Refuse rows the outline writer would crash on."""
        row = dict(ITEM_ROWS[0])
        del row["Title"]
        with self.assertRaisesRegex(InjectError, "have no Title: I-3001"):
            validate_snapshot(_snapshot(tables={"tblItems": (row,)}))


class CellCoercionTests(unittest.TestCase):
    """Pin the export-side cell typing rules."""

    def test_time_of_day_in_a_date_cell_halts(self) -> None:
        """Refuse datetimes that are not pure dates."""
        cell_value = datetime.combine(date(2026, 7, 1), time(10, 30))
        with self.assertRaisesRegex(ExportError, "time of day"):
            _coerce_date("tblItems", "Due", 3, cell_value)

    def test_numeric_text_is_converted_with_a_note(self) -> None:
        """Convert numbers typed into text cells and say so."""
        notes: list[str] = []
        self.assertEqual(_coerce_text("tblItems", "Title", 2, 42, notes), "42")
        self.assertTrue(any("exported as text" in note for note in notes))

    def test_booleans_in_text_cells_halt(self) -> None:
        """Refuse values that cannot be honest text."""
        cell_value: object = True
        with self.assertRaisesRegex(ExportError, "expected text"):
            _coerce_text("tblItems", "Title", 2, cell_value, [])


class SnapshotDocumentTests(unittest.TestCase):
    """Protect the persisted snapshot format and its retention ring."""

    def test_write_and_read_round_trip(self) -> None:
        """Persist typed rows, dates and settings losslessly."""
        snapshot = _snapshot()
        with tempfile.TemporaryDirectory() as directory:
            path = write_snapshot(snapshot, directory=directory)
            loaded = read_snapshot(path)
        self.assertEqual(loaded, snapshot)

    def test_ring_is_pruned(self) -> None:
        """Keep only the newest snapshots."""
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory)
            for index in range(RING_LIMIT + 3):
                stale = target / f"pm-data-20260101T{index:06}-000000000000.json"
                stale.write_text("{}", encoding="utf-8")
            write_snapshot(_snapshot(), directory=directory)
            remaining = list(target.glob("pm-data-*.json"))
        self.assertEqual(len(remaining), RING_LIMIT)

    def test_unsupported_format_is_rejected(self) -> None:
        """Refuse documents from an unknown snapshot format."""
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "pm-data-bad.json"
            path.write_text(json.dumps({"format": 99}), encoding="utf-8")
            with self.assertRaisesRegex(SnapshotError, "unsupported format"):
                read_snapshot(path)

    def test_bad_dates_are_rejected(self) -> None:
        """Refuse rows whose date columns cannot parse."""
        snapshot = _snapshot()
        with tempfile.TemporaryDirectory() as directory:
            path = write_snapshot(snapshot, directory=directory)
            document = json.loads(path.read_text(encoding="utf-8"))
            document["tables"]["tblItems"][0]["Start"] = "not-a-date"
            path.write_text(json.dumps(document), encoding="utf-8")
            with self.assertRaisesRegex(SnapshotError, "ISO date"):
                read_snapshot(path)


class MigrateGuardTests(unittest.TestCase):
    """Refuse migration targets that are missing, foreign or open in Excel."""

    def test_missing_workbook_is_rejected(self) -> None:
        """Fail fast when the workbook path does not exist."""
        with (
            tempfile.TemporaryDirectory() as directory,
            self.assertRaisesRegex(MigrationError, "does not exist"),
        ):
            _require_workbook(Path(directory) / "absent.xlsm")

    def test_non_macro_workbook_is_rejected(self) -> None:
        """Only the macro-enabled artifact is a migration target."""
        with tempfile.TemporaryDirectory() as directory:
            workbook = Path(directory) / "plain.xlsx"
            workbook.touch()
            with self.assertRaisesRegex(MigrationError, "not an .xlsm"):
                _require_workbook(workbook)

    def test_excel_lock_file_is_rejected(self) -> None:
        """Refuse to run while Excel holds the workbook open."""
        with tempfile.TemporaryDirectory() as directory:
            workbook = Path(directory) / "PM_Workbook.xlsm"
            workbook.touch()
            (Path(directory) / "~$PM_Workbook.xlsm").touch()
            with self.assertRaisesRegex(MigrationError, "open in Excel"):
                _require_workbook(workbook)


if __name__ == "__main__":
    unittest.main()
