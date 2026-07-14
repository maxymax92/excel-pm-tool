"""Static source gate for the complete two-module VBA registry."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

from ..paths import VBA_DIR
from ..spec.capacity import DATA_ROWS
from ..spec.items import ITEMS_COLUMNS, RAID_COLUMNS
from ..vba.registry import (
    MODULES,
    PUBLIC_MACROS,
    SOURCE_FILENAMES,
    WORKBOOK_EVENTS,
    ModuleKind,
    VbaModule,
)

if TYPE_CHECKING:
    from collections.abc import Mapping

LOGGER = logging.getLogger(__name__)

PROCEDURE_RE = re.compile(
    r"(?im)^(?P<visibility>Public|Private|Friend) "
    r"(?P<kind>Sub|Function|Property (?:Get|Let|Set)) "
    r"(?P<name>[A-Za-z][A-Za-z0-9_]*)\b"
)
CONSTANT_RE = re.compile(
    r"(?im)^Private Const (?P<name>[A-Z][A-Z0-9_]*) As Long = (?P<value>\d+)\s*$"
)


@dataclass(frozen=True, slots=True)
class Procedure:
    """One parsed VBA procedure declaration."""

    name: str
    visibility: str
    kind: str


def _procedures(source: str) -> tuple[Procedure, ...]:
    return tuple(
        Procedure(match["name"], match["visibility"], match["kind"])
        for match in PROCEDURE_RE.finditer(source)
    )


def _source_inventory_failures() -> list[str]:
    actual = {
        path.name
        for path in VBA_DIR.iterdir()
        if path.is_file() and (path.suffix.lower() == ".bas" or path.name.endswith(".cls.txt"))
    }
    failures = [
        f"VBA source file missing: {filename}" for filename in sorted(SOURCE_FILENAMES - actual)
    ]
    failures.extend(
        f"unexpected VBA source file: {filename}" for filename in sorted(actual - SOURCE_FILENAMES)
    )
    return failures


def _load_sources() -> tuple[dict[str, str], list[str]]:
    failures = _source_inventory_failures()
    sources: dict[str, str] = {}
    for module in MODULES:
        if not module.path.is_file():
            continue
        try:
            source = module.path.read_text(encoding="utf-8")
        except UnicodeError as error:
            failures.append(f"{module.filename} is not valid UTF-8: {error}")
            continue
        if "\x00" in source:
            failures.append(f"{module.filename} contains a NUL character")
        if not source.endswith("\n"):
            failures.append(f"{module.filename} does not end with a newline")
        sources[module.name] = source
    return sources, failures


def _module_header_failures(module: VbaModule, source: str) -> list[str]:
    failures: list[str] = []
    if source.count("Option Explicit") != 1:
        failures.append(f"{module.filename} must contain exactly one Option Explicit")
    private_count = source.count("Option Private Module")
    expected_private_count = 1 if module.private_to_project else 0
    if private_count != expected_private_count:
        failures.append(
            f"{module.filename} has {private_count} Option Private Module directives; "
            f"expected {expected_private_count}"
        )

    attributes = re.findall(r'(?im)^Attribute VB_Name = "([^"]+)"\s*$', source)
    if module.kind is ModuleKind.STANDARD:
        if attributes != [module.name]:
            failures.append(
                f"{module.filename} declares VB_Name {attributes!r}; expected {module.name!r}"
            )
    elif attributes:
        failures.append(
            f"{module.filename} must be a complete document-module body without attributes"
        )
    return failures


def _procedure_surface_failures(module: VbaModule, source: str) -> list[str]:
    procedures = _procedures(source)
    names = [procedure.name for procedure in procedures]
    failures: list[str] = []
    duplicates = sorted({name for name in names if names.count(name) > 1})
    failures.extend(f"{module.filename} duplicates procedure {name}" for name in duplicates)
    friend = [procedure.name for procedure in procedures if procedure.visibility == "Friend"]
    failures.extend(f"{module.filename} uses forbidden Friend procedure {name}" for name in friend)
    public = tuple(procedure.name for procedure in procedures if procedure.visibility == "Public")
    if public != module.public_procedures:
        failures.append(
            f"{module.filename} public procedures are {public!r}; "
            f"expected {module.public_procedures!r}"
        )
    return failures


def _surface_failures(sources: Mapping[str, str]) -> list[str]:
    failures: list[str] = []
    visible_macros = tuple(
        match.group(1)
        for module in MODULES
        if not module.private_to_project and module.kind is ModuleKind.STANDARD
        for match in re.finditer(
            r"(?im)^Public Sub ([A-Za-z][A-Za-z0-9_]*)\(\)\s*$", sources.get(module.name, "")
        )
    )
    if visible_macros != PUBLIC_MACROS:
        failures.append(
            f"visible parameterless macro surface is {visible_macros!r}; expected {PUBLIC_MACROS!r}"
        )
    workbook_procedures = _procedures(sources.get("ThisWorkbook", ""))
    events = tuple(
        procedure.name
        for procedure in workbook_procedures
        if procedure.name.startswith("Workbook_")
    )
    if events != WORKBOOK_EVENTS:
        failures.append(f"workbook events are {events!r}; expected {WORKBOOK_EVENTS!r}")
    public_workbook_procedures = tuple(
        procedure.name for procedure in workbook_procedures if procedure.visibility == "Public"
    )
    if public_workbook_procedures:
        failures.append(f"ThisWorkbook exposes public procedures: {public_workbook_procedures!r}")
    return failures


def _column_index(columns: list[dict[str, object]], name: str) -> int:
    return next(index for index, column in enumerate(columns, start=1) if column["name"] == name)


def _long_constants(source: str) -> dict[str, int]:
    return {match["name"]: int(match["value"]) for match in CONSTANT_RE.finditer(source)}


def _constant_failures(
    module_name: str,
    source: str,
    expected: Mapping[str, int],
) -> list[str]:
    constants = _long_constants(source)
    return [
        f"{module_name} constant {name} is {constants.get(name)!r}; expected {value}"
        for name, value in expected.items()
        if constants.get(name) != value
    ]


def _schema_constant_failures(sources: Mapping[str, str]) -> list[str]:
    event_constants = {
        "DATA_CAPACITY": DATA_ROWS,
        "ITEMS_ID_COLUMN": _column_index(ITEMS_COLUMNS, "ID"),
        "ITEMS_TYPE_COLUMN": _column_index(ITEMS_COLUMNS, "Type"),
        "ITEMS_TITLE_COLUMN": _column_index(ITEMS_COLUMNS, "Title"),
        "ITEMS_STATUS_COLUMN": _column_index(ITEMS_COLUMNS, "Status"),
        "ITEMS_DELIVERY_HEALTH_COLUMN": _column_index(ITEMS_COLUMNS, "Delivery Health"),
        "ITEMS_LATEST_STATUS_COLUMN": _column_index(ITEMS_COLUMNS, "Latest Status"),
        "ITEMS_LAST_INPUT_COLUMN": _column_index(ITEMS_COLUMNS, "BlockedBy"),
        "ITEMS_CREATED_COLUMN": _column_index(ITEMS_COLUMNS, "Created"),
        "ITEMS_UPDATED_COLUMN": _column_index(ITEMS_COLUMNS, "Updated"),
        "ITEMS_ACTIVE_SINCE_COLUMN": _column_index(ITEMS_COLUMNS, "InProgressSince"),
        "ITEMS_DONE_DATE_COLUMN": _column_index(ITEMS_COLUMNS, "DoneDate"),
        "ITEMS_BLOCKED_SINCE_COLUMN": _column_index(ITEMS_COLUMNS, "BlockedSince"),
        "ITEMS_LATEST_UPDATE_COLUMN": _column_index(ITEMS_COLUMNS, "LatestUpdateOn"),
        "RAID_ID_COLUMN": _column_index(RAID_COLUMNS, "RaidID"),
        "RAID_TYPE_COLUMN": _column_index(RAID_COLUMNS, "Type"),
        "RAID_TITLE_COLUMN": _column_index(RAID_COLUMNS, "Title"),
        "RAID_DETAIL_COLUMN": _column_index(RAID_COLUMNS, "Detail"),
        "RAID_RESPONSE_COLUMN": _column_index(RAID_COLUMNS, "Response"),
        "RAID_LAST_INPUT_COLUMN": _column_index(RAID_COLUMNS, "NextReview"),
        "RAID_STATUS_COLUMN": _column_index(RAID_COLUMNS, "Status"),
        "RAID_RAISED_COLUMN": _column_index(RAID_COLUMNS, "Raised"),
        "RAID_CLOSED_COLUMN": _column_index(RAID_COLUMNS, "Closed"),
        "RAID_UPDATED_COLUMN": _column_index(RAID_COLUMNS, "Updated"),
    }
    return [
        *_constant_failures(
            "ThisWorkbook",
            sources.get("ThisWorkbook", ""),
            event_constants,
        ),
        *_constant_failures(
            "PMTool",
            sources.get("PMTool", ""),
            {"ID_COUNTER_MAX": 999999999},
        ),
        *_constant_failures(
            "PMTool",
            sources.get("PMTool", ""),
            {"EXPORT_CHARACTER_MAX": 5000000, "EXPORT_BYTE_MAX": 20000000},
        ),
        *_constant_failures(
            "PMTool",
            sources.get("PMTool", ""),
            {"ITEM_CAPACITY": DATA_ROWS},
        ),
    ]


def _required_tokens(source: str, label: str, tokens: tuple[str, ...]) -> list[str]:
    return [f"{label} contract missing {token}" for token in tokens if token not in source]


def _event_contract_failures(source: str) -> list[str]:
    failures = _required_tokens(
        source,
        "ThisWorkbook event",
        (
            "TableChangeRange",
            "TableIdRange",
            'ListObjects("tblItems")',
            'ListObjects("tblRAID")',
            "touchedRows(touchedCount) = rowIndex",
            "For changedIndex = 1 To touchedCount",
            "statusTouched(rowIndex)",
            "typeTouched(rowIndex)",
            "healthTouched(rowIndex)",
            "latestStatusTouched(rowIndex)",
            "ByVal readStatusRoles As Boolean",
            "ByVal readBlockedRole As Boolean",
            "ByVal readStatusRole As Boolean",
            "ByVal applyStatus As Boolean",
            "ReadItemRowRoles",
            "ReadRaidRowRoles",
            "TryItemStatusRoles",
            "TryRaidStatusIsClosed",
            "TryBlockedDeliveryHealth",
            "TryItemTypeLevel",
            "PMTool.ApplyItemLevelPresentation",
            "ApplyRaidRowPresentation",
            "ApplyNarrativeCellFormat",
            "PMTool.ItemStatusRoles",
            "PMTool.RaidStatusIsClosed",
            "RowCoreHasEnteredData",
            "ITEMS_TYPE_COLUMN).Value2",
            "RAID_TYPE_COLUMN).Value2",
            "TryRestoreEnableEvents",
        ),
    )
    failures.extend(
        f"ThisWorkbook event code scans every supported row: {token}"
        for token in ("For rowIndex = 1 To rowCount", "For rowIndex = rowCount To 1 Step -1")
        if token in source
    )
    failures.extend(
        f"ThisWorkbook event code rejects user data edits: {token}"
        for token in (
            "invalid Items edit",
            "invalid RAID edit",
            "requires a Title",
            "ValidateItemRow",
            "ValidateRaidRow",
        )
        if token in source
    )
    return failures


def _config_contract_failures(source: str) -> list[str]:
    return _required_tokens(
        source,
        "ThisWorkbook Config",
        (
            "configSnapshotReady",
            "CaptureConfigSemantics",
            "ValidateUsedItemStatusSemantics",
            "ValidateUsedRaidStatusSemantics",
            "ValidateUsedDeliveryHealthSemantics",
            "Application.Undo",
            "TryUndoLastUserEdit",
            "failureDescription = failureDescription & _\n"
            '            " The Config edit was undone."',
            "PMTool.ItemStatusRoles",
            "PMTool.RaidStatusIsClosed",
            "PMTool.IsBlockedDeliveryHealth",
        ),
    )


def _export_contract_failures(source: str) -> list[str]:
    return _required_tokens(
        source,
        "PMTool export",
        (
            "On Error GoTo exportFailure",
            "Application.GetSaveAsFilename",
            'ButtonText:="Export"',
            'FileFilter:="Markdown (*.md),*.md"',
            "Private Function NormalizeMacMarkdownPath(",
            "NormalizeMacMarkdownPath(CStr(selectedPath))",
            'Case ".xlsx", ".xlsm", ".xlsb", ".xls"',
            "SelectedMarkdownPath = EnsureMdExtension(CStr(selectedPath))",
            "ValidateItemsHierarchy",
            "ArrayRowHasEnteredData",
            "ValidatedIdentifierText",
            "AddUniqueIdentifier",
            "Private Sub ReadFileBytes",
            "Private Sub WriteFileBytes",
            "Open path For Binary Access Read",
            "Open path For Binary Access Write",
            "originalCaptured",
            "destinationTouched",
            "destinationChanged",
            "destination restoration failed",
            "incomplete destination removal also failed",
            "SanitizeMarkdownControls",
            'ThisWorkbook.Worksheets("Calc").Calculate',
            'ThisWorkbook.Worksheets("Overview").Calculate',
        ),
    ) + [
        f"PMTool export retains unsupported Mac FileDialog code: {token}"
        for token in ("Application.FileDialog", ".SelectedItems")
        if token in source
    ]


def _organise_contract_failures(source: str) -> list[str]:
    failures = _required_tokens(
        source,
        "PMTool organiser",
        (
            "table.Range.Calculate",
            ".SortFields.Add",
            "ClearItemRowOutline sheet",
            ".Group",
            ".OutlineLevel <> level",
            "ApplyItemLevelPresentation sheet",
            "ValidateItemsHierarchy",
            "ArrayRowHasEnteredData",
            "data = table.DataBodyRange.Value2",
            "table.Resize",
            'organiseStage = "verifying the sorted Items rows"',
            'organiseStage = "resizing the Items table"',
            "RestoreOrganiseApplicationState",
            "failureNumber = Err.Number",
            "failureSource = Err.Source",
        ),
    )
    failures.extend(
        f"PMTool organiser uses row-by-row table work: {token}"
        for token in (".ListRows(rowIndex).Delete", "WorksheetFunction.CountIf")
        if token in source
    )
    return failures


def _core_contract_failures(source: str) -> list[str]:
    failures = _required_tokens(
        source,
        "PMTool core",
        (
            'Tbl("tblStatuses")',
            'Tbl("tblRaidStatuses")',
            'Tbl("tblDeliveryHealth")',
            'Tbl("tblTypes")',
            "ItemTypeLevel",
            "ConfiguredBoolean",
            "TableTextRow",
            "TextRangeStats",
            "values.Value2",
            "ValidateItemsHierarchy",
            "AddUniqueIdentifier",
        ),
    )
    if "For Each cell In values.Cells" in source:
        failures.append("PMTool text validation crosses the Excel object model per cell")
    return failures


def _forbidden_source_failures(sources: Mapping[str, str]) -> list[str]:
    forbidden = (
        "On Error " + "Resume Next",
        "Debug" + ".Print",
        "CalculateFullRebuild",
        "CalculateFull",
        "Application.Calculate",
        "For Append",
        '".partial"',
        '".backup"',
        "MacScript",
    )
    failures = [
        f"{module_name} contains forbidden VBA behavior {token}"
        for module_name, source in sources.items()
        for token in forbidden
        if token.lower() in source.lower()
    ]
    type_pattern = re.compile(r"(?i)\bAs\s+(ListObject|Worksheet|Range)\b")
    failures.extend(
        f"{module_name} uses unqualified Excel host type {match.group(1)}"
        for module_name, source in sources.items()
        for match in type_pattern.finditer(source)
    )
    return failures


def source_failures() -> list[str]:
    """Return every full-registry VBA source-contract violation.

    Returns:
        Deterministically ordered source diagnostics.

    """
    sources, failures = _load_sources()
    for module in MODULES:
        source = sources.get(module.name)
        if source is None:
            continue
        failures.extend(_module_header_failures(module, source))
        failures.extend(_procedure_surface_failures(module, source))
    if len(sources) != len(MODULES):
        return failures
    failures.extend(_surface_failures(sources))
    failures.extend(_schema_constant_failures(sources))
    failures.extend(_core_contract_failures(sources["PMTool"]))
    failures.extend(_export_contract_failures(sources["PMTool"]))
    failures.extend(_organise_contract_failures(sources["PMTool"]))
    failures.extend(_config_contract_failures(sources["ThisWorkbook"]))
    failures.extend(_event_contract_failures(sources["ThisWorkbook"]))
    failures.extend(_forbidden_source_failures(sources))
    return failures


def main() -> int:
    """Report VBA source-contract results.

    Returns:
        Zero for success and one for any source violation.

    """
    failures = source_failures()
    if failures:
        LOGGER.error("VBA SOURCE QA FAIL (%s issue(s))", len(failures))
        for failure in failures:
            LOGGER.error("  - %s", failure)
        return 1
    LOGGER.info("VBA SOURCE QA PASS")
    return 0


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    raise SystemExit(main())
