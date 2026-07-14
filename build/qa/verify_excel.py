"""Verify exact authored workbook semantics across a desktop-Excel save.

The verifier operates on a disposable copy with the same workbook suffix. It
compares authored formula expressions, defined-name targets, per-cell effective
data validation, conditional formatting, package-part preservation and embedded
VBA source. Calculation caches, spill footprints, revision identifiers, LET
parser names, self-table qualification and volatile VBA project metadata are
serialization state rather than authored behavior.
"""

from __future__ import annotations

import argparse
import asyncio
import io
import posixpath
import re
import sys
import zipfile
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import TYPE_CHECKING, Protocol

import olefile
from defusedxml import ElementTree as DefusedET
from oletools.olevba import VBA_Parser
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

from ..automation.workspace import excel_workbook_copy
from ..paths import AUTOMATION, ROOT
from ..spec.capacity import CONFIG_ROWS, DATA_ROWS, PLAN_ROWS
from .excel import recalculate

if TYPE_CHECKING:
    from collections.abc import Callable, Iterator

SUCCESS = "REPAIRED-COPY-SAVED"
SUPPORTED_SUFFIXES = {".xlsx", ".xlsm"}
MAX_AUTHORED_ROW = max(CONFIG_ROWS + 2, DATA_ROWS + 5, PLAN_ROWS + 5)
CELL_RANGE = re.compile(
    r"\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?",
    re.IGNORECASE,
)
VOLATILE_VBA_PROJECT_FIELDS = (b"ID=", b"CMG=", b"DPB=", b"GC=")
ORACLE_TIMEOUT_SECONDS = 180
FORMULA_NORMALIZATION_ANCHOR = "LCB524288"


class _SemanticProblem(Enum):
    INVALID_COLUMN = "invalid column label: {!r}"
    UNSUPPORTED_RANGE_REFERENCE = "unsupported sqref token: {!r}"
    REVERSED_SQREF = "reversed sqref range: {!r}"
    SQREF_BOUNDARY = "sqref {!r} exceeds the authored row boundary 1:{}"
    CELL_COORDINATE = "invalid cell coordinate: {!r}"
    TABLE_METADATA = "{} has no table name or range"
    TABLE_RANGE = "{} has unsupported table range {!r}"
    TABLE_RELATIONSHIP = "{} tablePart has no relationship target"
    TABLE_TARGET = "{} refers to unsupported table part {}"
    CELL_NO_COORDINATE = "cell in {} has no coordinate"
    DUPLICATE_FORMULA = "duplicate formula cell: {}!{}"
    SHARED_FORMULA = "{}!{} refers to missing shared formula {!r}"
    MULTIPLE_TABLES = "{}!{} is covered by multiple tables"
    TABLE_NAME = "{} has no table name"
    DUPLICATE_TABLE = "duplicate table name: {}"
    TABLE_COLUMN = "table column in {} has no name"
    DUPLICATE_TABLE_FORMULA = "duplicate {} for {}[{}]"
    DEFINED_NAME = "definedName element has no name"
    DUPLICATE_DEFINED_NAME = "duplicate defined name in {} scope: {}"
    MISSING_SQREF = "{} has no sqref"
    VALIDATION_CHILDREN = "dataValidation {} contains child elements"
    VALIDATION_EMPTY = "dataValidation {} is empty"
    CONDITIONAL_CHILDREN = "conditional-format formula contains child elements in {}"
    CONDITIONAL_EMPTY = "conditional-format formula is empty in {}"
    CONDITIONAL_PRIORITY = "{} cfRule has invalid priority {!r}"
    DUPLICATE_CONDITIONAL_PRIORITY = "{} contains duplicate cfRule priority {}"
    CONDITIONAL_DXF = "{} cfRule has invalid dxfId {!r}"
    MISSING_STYLES = "workbook package has no xl/styles.xml"
    MISSING_DXF = "{} cfRule refers to missing dxfId {}"
    VBA_PROJECT_STREAM = "embedded VBA project has no PROJECT stream"
    VBA_NAMES_STREAM = "embedded VBA project has no PROJECTwm stream"
    DUPLICATE_VBA_MODULE = "duplicate embedded VBA module: {}"
    NO_VBA_MODULES = "embedded VBA project contains no modules"
    CRC_FAILURE = "CRC failure in {}"
    SUFFIX = "repair verification requires .xlsx or .xlsm: {}"


class WorkbookSemanticError(ValueError):
    """Report malformed or changed authored workbook semantics."""

    def __init__(self, problem: _SemanticProblem, *details: object) -> None:
        """Create one stable semantic diagnostic."""
        super().__init__(problem.value.format(*details))


class _VerificationProblem(Enum):
    ORACLE_TIMEOUT = "Excel package oracle exceeded {} seconds"
    ORACLE_TIMEOUT_CLEANUP = "Excel package oracle timed out and cleanup failed: {}: {}"
    ORACLE_EXIT = "Excel package oracle failed for {} (exit {}): {}"
    ORACLE_STDERR = "Excel package oracle wrote to stderr for {}: {}"
    ORACLE_SENTINEL = "Excel package oracle returned {!r}; expected {!r}"


class ExcelVerificationError(RuntimeError):
    """Report a failed Excel preservation operation or its cleanup."""

    def __init__(self, problem: _VerificationProblem, *details: object) -> None:
        """Create one stable operation diagnostic."""
        super().__init__(problem.value.format(*details))


@dataclass(frozen=True, slots=True)
class _AutomationResult:
    """Captured result of one Excel AppleScript process."""

    returncode: int
    stdout: str
    stderr: str


@dataclass(frozen=True, slots=True)
class _ConditionalRule:
    """One conditional-format rule with exact coverage and semantic identity."""

    priority: int
    signature: tuple
    coverage: frozenset[tuple[int, int]]


class XmlElement(Protocol):
    """Structural type for elements returned by the hardened XML parser."""

    tag: str
    attrib: dict[str, str]
    text: str | None

    def __iter__(self) -> Iterator[XmlElement]:
        """Iterate over direct child elements."""

    def get(self, key: str, default: str | None = None) -> str | None:
        """Return one element attribute."""

    def iter(self) -> Iterator[XmlElement]:
        """Iterate over the element and its descendants."""


def _column_number(label: str) -> int:
    result = 0
    for character in label.upper():
        if not "A" <= character <= "Z":
            raise WorkbookSemanticError(_SemanticProblem.INVALID_COLUMN, label)
        result = result * 26 + ord(character) - 64
    return result


def _cells(sqref: str) -> frozenset[tuple[int, int]]:
    """Expand one bounded OOXML range expression into exact cell coverage.

    Returns:
        Every cell covered by the expression as column and row pairs.

    Raises:
        WorkbookSemanticError: If the range expression is invalid or unbounded.

    """
    cells: set[tuple[int, int]] = set()
    for token in sqref.split():
        match = CELL_RANGE.fullmatch(token)
        if match is None:
            raise WorkbookSemanticError(
                _SemanticProblem.UNSUPPORTED_RANGE_REFERENCE,
                token,
            )
        start_col = _column_number(match.group(1))
        start_row = int(match.group(2))
        end_col = _column_number(match.group(3) or match.group(1))
        end_row = int(match.group(4) or match.group(2))
        if start_col > end_col or start_row > end_row:
            raise WorkbookSemanticError(_SemanticProblem.REVERSED_SQREF, token)
        if start_row < 1 or end_row > MAX_AUTHORED_ROW:
            raise WorkbookSemanticError(
                _SemanticProblem.SQREF_BOUNDARY,
                token,
                MAX_AUTHORED_ROW,
            )
        for column in range(start_col, end_col + 1):
            cells.update((column, row) for row in range(start_row, end_row + 1))
    return frozenset(cells)


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _cell_coordinate(coordinate: str) -> tuple[int, int]:
    match = CELL_RANGE.fullmatch(coordinate)
    if match is None or match.group(3) is not None:
        raise WorkbookSemanticError(_SemanticProblem.CELL_COORDINATE, coordinate)
    return _column_number(match.group(1)), int(match.group(2))


def _normalize_current_row_reference(formula: str, table_name: str) -> str:
    qualifier = re.compile(rf"(?i)(?<![A-Z0-9_]){re.escape(table_name)}(?=\[\[#This Row\],)")
    unqualified = qualifier.sub("", formula)
    return re.sub(
        r"\[\[#This Row\],\[([^\[\]]+)\]\]",
        r"[[#This Row],\1]",
        unqualified,
    )


type TableRange = tuple[str, int, int, int, int]


def _table_ranges_by_part(package: zipfile.ZipFile) -> dict[str, TableRange]:
    table_ranges: dict[str, TableRange] = {}
    for part in package.namelist():
        if re.fullmatch(r"xl/tables/table\d+\.xml", part) is None:
            continue
        root = DefusedET.fromstring(package.read(part))
        table_name = root.get("name") or root.get("displayName")
        table_ref = root.get("ref")
        if not table_name or not table_ref:
            raise WorkbookSemanticError(_SemanticProblem.TABLE_METADATA, part)
        match = CELL_RANGE.fullmatch(table_ref)
        if match is None:
            raise WorkbookSemanticError(_SemanticProblem.TABLE_RANGE, part, table_ref)
        table_ranges[part] = (
            table_name,
            _column_number(match.group(1)),
            int(match.group(2)),
            _column_number(match.group(3) or match.group(1)),
            int(match.group(4) or match.group(2)),
        )
    return table_ranges


def _relationships_by_id(package: zipfile.ZipFile, part: str) -> dict[str | None, str | None]:
    relationships_part = f"{posixpath.dirname(part)}/_rels/{posixpath.basename(part)}.rels"
    if relationships_part not in package.namelist():
        return {}
    relationships_root = DefusedET.fromstring(package.read(relationships_part))
    return {
        relationship.get("Id"): relationship.get("Target")
        for relationship in relationships_root.iter()
        if _local_name(relationship.tag) == "Relationship"
    }


def _sheet_table_ranges(
    package: zipfile.ZipFile,
    part: str,
    table_ranges: dict[str, TableRange],
) -> tuple[TableRange, ...]:
    relationships = _relationships_by_id(package, part)
    sheet_root = DefusedET.fromstring(package.read(part))
    ranges: list[TableRange] = []
    for table_part in sheet_root.iter():
        if _local_name(table_part.tag) != "tablePart":
            continue
        relationship_id = next(
            (value for name, value in table_part.attrib.items() if _local_name(name) == "id"),
            None,
        )
        target = relationships.get(relationship_id)
        if not target:
            raise WorkbookSemanticError(_SemanticProblem.TABLE_RELATIONSHIP, part)
        target_part = posixpath.normpath(posixpath.join(posixpath.dirname(part), target)).lstrip(
            "/"
        )
        if target_part not in table_ranges:
            raise WorkbookSemanticError(
                _SemanticProblem.TABLE_TARGET,
                part,
                target_part,
            )
        ranges.append(table_ranges[target_part])
    return tuple(ranges)


def _worksheet_table_ranges(
    package: zipfile.ZipFile,
) -> dict[str, tuple[TableRange, ...]]:
    """Map each worksheet part to the table ranges it owns.

    Returns:
        Worksheet-part keys mapped to table name and boundary tuples.

    """
    table_ranges = _table_ranges_by_part(package)
    result: dict[str, tuple[TableRange, ...]] = {}
    for part in package.namelist():
        if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", part) is None:
            continue
        result[part] = _sheet_table_ranges(package, part, table_ranges)
    return result


def _canonical_element(
    element: XmlElement,
    *,
    excluded_attributes: frozenset[str] = frozenset(),
    excluded_children: frozenset[str] = frozenset(),
) -> tuple:
    """Return a namespace-stable, order-preserving XML semantic signature.

    Returns:
        The element tag, filtered attributes, text and child signatures.

    """
    attributes = tuple(
        sorted(
            (name, value)
            for name, value in element.attrib.items()
            if _local_name(name) not in excluded_attributes
        )
    )
    children = tuple(
        _canonical_element(child)
        for child in element
        if _local_name(child.tag) not in excluded_children
    )
    return element.tag, attributes, (element.text or "").strip(), children


def _formula_cells(root: XmlElement, part: str) -> list[tuple[str, XmlElement]]:
    cells: list[tuple[str, XmlElement]] = []
    for cell in root.iter():
        if _local_name(cell.tag) != "c":
            continue
        coordinate = cell.get("r")
        if not coordinate:
            raise WorkbookSemanticError(_SemanticProblem.CELL_NO_COORDINATE, part)
        formula = next((child for child in cell if _local_name(child.tag) == "f"), None)
        if formula is not None:
            cells.append((coordinate, formula))
    return cells


def _shared_formula_masters(
    formula_cells: list[tuple[str, XmlElement]],
) -> dict[str | None, tuple[str, str]]:
    return {
        formula.get("si"): (coordinate, (formula.text or "").strip())
        for coordinate, formula in formula_cells
        if formula.get("t") == "shared" and (formula.text or "").strip()
    }


def _effective_formula(
    part: str,
    coordinate: str,
    formula: XmlElement,
    shared_masters: dict[str | None, tuple[str, str]],
) -> tuple[str, str]:
    formula_text = (formula.text or "").strip()
    formula_type = formula.get("t", "normal") or "normal"
    if formula_type != "shared":
        return formula_type, formula_text

    shared_index = formula.get("si")
    if shared_index not in shared_masters:
        raise WorkbookSemanticError(
            _SemanticProblem.SHARED_FORMULA,
            part,
            coordinate,
            shared_index,
        )
    origin, master_formula = shared_masters[shared_index]
    translated = _translate_formula(
        master_formula,
        origin=origin,
        destination=coordinate,
    )
    return "normal", translated


def _covering_table(
    part: str,
    coordinate: str,
    table_ranges: tuple[TableRange, ...],
) -> str | None:
    column, row = _cell_coordinate(coordinate)
    matching_tables = [
        table_name
        for table_name, start_column, start_row, end_column, end_row in table_ranges
        if start_column <= column <= end_column and start_row <= row <= end_row
    ]
    if len(matching_tables) > 1:
        raise WorkbookSemanticError(
            _SemanticProblem.MULTIPLE_TABLES,
            part,
            coordinate,
        )
    return matching_tables[0] if matching_tables else None


def _record_worksheet_formulas(
    package: zipfile.ZipFile,
    part: str,
    table_ranges: tuple[TableRange, ...],
    formulas: dict[tuple[str, str], tuple[str, str]],
) -> None:
    root = DefusedET.fromstring(package.read(part))
    formula_cells = _formula_cells(root, part)
    shared_masters = _shared_formula_masters(formula_cells)
    for coordinate, formula in formula_cells:
        key = (part, coordinate)
        if key in formulas:
            raise WorkbookSemanticError(
                _SemanticProblem.DUPLICATE_FORMULA,
                part,
                coordinate,
            )
        formula_type, formula_text = _effective_formula(
            part,
            coordinate,
            formula,
            shared_masters,
        )
        if not formula_text:
            continue
        table_name = _covering_table(part, coordinate, table_ranges)
        if table_name is not None:
            formula_text = _normalize_current_row_reference(formula_text, table_name)
        formulas[key] = (formula_type, formula_text)


def _worksheet_formulas(package: zipfile.ZipFile) -> dict[tuple[str, str], tuple[str, str]]:
    formulas: dict[tuple[str, str], tuple[str, str]] = {}
    worksheet_tables = _worksheet_table_ranges(package)
    for part, table_ranges in worksheet_tables.items():
        _record_worksheet_formulas(package, part, table_ranges, formulas)
    return formulas


def _record_table_formula(
    formulas: dict[tuple[str, str, str], tuple[tuple[tuple[str, str], ...], str]],
    *,
    table_name: str,
    column_name: str,
    child: XmlElement,
) -> None:
    kind = _local_name(child.tag)
    if kind not in {"calculatedColumnFormula", "totalsRowFormula"}:
        return
    key = (table_name, column_name, kind)
    if key in formulas:
        raise WorkbookSemanticError(
            _SemanticProblem.DUPLICATE_TABLE_FORMULA,
            kind,
            table_name,
            column_name,
        )
    formula_text = (child.text or "").strip()
    formulas[key] = (
        tuple(sorted(child.attrib.items())),
        _normalize_current_row_reference(formula_text, table_name),
    )


def _record_table_formulas(
    root: XmlElement,
    table_name: str,
    formulas: dict[tuple[str, str, str], tuple[tuple[tuple[str, str], ...], str]],
) -> None:
    for column in root.iter():
        if _local_name(column.tag) != "tableColumn":
            continue
        column_name = column.get("name")
        if not column_name:
            raise WorkbookSemanticError(_SemanticProblem.TABLE_COLUMN, table_name)
        for child in column:
            _record_table_formula(
                formulas,
                table_name=table_name,
                column_name=column_name,
                child=child,
            )


def _table_formulas(
    package: zipfile.ZipFile,
) -> dict[tuple[str, str, str], tuple[tuple[tuple[str, str], ...], str]]:
    formulas: dict[
        tuple[str, str, str],
        tuple[tuple[tuple[str, str], ...], str],
    ] = {}
    table_names: set[str] = set()
    for part in package.namelist():
        if re.fullmatch(r"xl/tables/table\d+\.xml", part) is None:
            continue
        root = DefusedET.fromstring(package.read(part))
        table_name = root.get("name") or root.get("displayName")
        if not table_name:
            raise WorkbookSemanticError(_SemanticProblem.TABLE_NAME, part)
        if table_name in table_names:
            raise WorkbookSemanticError(_SemanticProblem.DUPLICATE_TABLE, table_name)
        table_names.add(table_name)
        _record_table_formulas(root, table_name, formulas)
    return formulas


def _defined_names(package: zipfile.ZipFile) -> dict[tuple[str, int | None], tuple]:
    root = DefusedET.fromstring(package.read("xl/workbook.xml"))
    names: dict[tuple[str, int | None], tuple] = {}
    for element in root.iter():
        if _local_name(element.tag) != "definedName":
            continue
        name = element.get("name")
        if not name:
            raise WorkbookSemanticError(_SemanticProblem.DEFINED_NAME)
        if name.casefold().startswith("_xlpm."):
            continue
        scope_text = element.get("localSheetId")
        scope = int(scope_text) if scope_text is not None else None
        key = (name, scope)
        if key in names:
            scope_label = "workbook" if scope is None else f"sheet {scope}"
            raise WorkbookSemanticError(
                _SemanticProblem.DUPLICATE_DEFINED_NAME,
                scope_label,
                name,
            )
        names[key] = _canonical_element(
            element,
            excluded_attributes=frozenset({"name", "localSheetId"}),
        )
    return names


def _sqref_text(element: XmlElement, *, context: str) -> str:
    sqref = element.get("sqref")
    if sqref is None:
        sqref_node = next(
            (child for child in element if _local_name(child.tag) == "sqref"),
            None,
        )
        sqref = None if sqref_node is None else sqref_node.text
    if not sqref or not sqref.strip():
        raise WorkbookSemanticError(_SemanticProblem.MISSING_SQREF, context)
    return sqref.strip()


def _sqref_origin(sqref: str) -> str:
    first_token = sqref.split(maxsplit=1)[0]
    match = CELL_RANGE.fullmatch(first_token)
    if match is None:
        raise WorkbookSemanticError(_SemanticProblem.UNSUPPORTED_RANGE_REFERENCE, first_token)
    return f"{match.group(1).upper()}{int(match.group(2))}"


def _translate_formula(formula: str, *, origin: str, destination: str) -> str:
    prefixed = formula if formula.startswith("=") else f"={formula}"
    translated = Translator(prefixed, origin=origin).translate_formula(destination)
    return translated if formula.startswith("=") else translated.removeprefix("=")


def _validation_signature(
    element: XmlElement,
    *,
    origin: str,
    destination: str,
) -> tuple:
    attributes = tuple(
        sorted(
            (name, value)
            for name, value in element.attrib.items()
            if _local_name(name) not in {"sqref", "uid"}
        )
    )
    children: list[tuple] = []
    for child in element:
        child_kind = _local_name(child.tag)
        if child_kind == "sqref":
            continue
        if child_kind in {"formula1", "formula2"}:
            if list(child):
                raise WorkbookSemanticError(
                    _SemanticProblem.VALIDATION_CHILDREN,
                    child_kind,
                )
            formula_text = (child.text or "").strip()
            if not formula_text:
                raise WorkbookSemanticError(
                    _SemanticProblem.VALIDATION_EMPTY,
                    child_kind,
                )
            children.append((
                child.tag,
                tuple(
                    sorted(
                        (name, value)
                        for name, value in child.attrib.items()
                        if _local_name(name) != "uid"
                    )
                ),
                _translate_formula(
                    formula_text,
                    origin=origin,
                    destination=destination,
                ),
                (),
            ))
            continue
        children.append(_canonical_element(child, excluded_attributes=frozenset({"uid"})))
    return element.tag, attributes, (element.text or "").strip(), tuple(children)


def _data_validations(
    package: zipfile.ZipFile,
) -> dict[tuple[str, tuple[int, int]], tuple]:
    result: dict[tuple[str, tuple[int, int]], list[tuple]] = {}
    for part in package.namelist():
        if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", part) is None:
            continue
        root = DefusedET.fromstring(package.read(part))
        for element in root.iter():
            if _local_name(element.tag) != "dataValidation":
                continue
            context = f"dataValidation in {part}"
            sqref = _sqref_text(element, context=context)
            origin = _sqref_origin(sqref)
            for cell in _cells(sqref):
                destination = f"{get_column_letter(cell[0])}{cell[1]}"
                signature = _validation_signature(
                    element,
                    origin=origin,
                    destination=destination,
                )
                result.setdefault((part, cell), []).append(signature)
    return {key: tuple(sorted(signatures, key=repr)) for key, signatures in result.items()}


def _differential_formats(package: zipfile.ZipFile) -> tuple[tuple, ...]:
    if "xl/styles.xml" not in package.namelist():
        raise WorkbookSemanticError(_SemanticProblem.MISSING_STYLES)
    root = DefusedET.fromstring(package.read("xl/styles.xml"))
    dxfs = next((element for element in root.iter() if _local_name(element.tag) == "dxfs"), None)
    if dxfs is None:
        return ()
    return tuple(_canonical_element(child) for child in dxfs)


def _conditional_priority(rule: XmlElement, *, part: str) -> int:
    priority_text = next(
        (value for name, value in rule.attrib.items() if _local_name(name) == "priority"),
        None,
    )
    try:
        priority = int(priority_text or "")
    except ValueError as error:
        raise WorkbookSemanticError(
            _SemanticProblem.CONDITIONAL_PRIORITY,
            part,
            priority_text,
        ) from error
    if priority < 1:
        raise WorkbookSemanticError(
            _SemanticProblem.CONDITIONAL_PRIORITY,
            part,
            priority_text,
        )
    return priority


def _conditional_rule_signature(
    rule: XmlElement,
    *,
    part: str,
    origin: str,
    dxf_signature: tuple | None,
) -> tuple:
    attributes = tuple(
        sorted(
            (name, value)
            for name, value in rule.attrib.items()
            if _local_name(name) not in {"dxfId", "priority"}
        )
    )
    children: list[tuple] = []
    for child in rule:
        if _local_name(child.tag) != "formula":
            children.append(_canonical_element(child))
            continue
        if list(child):
            raise WorkbookSemanticError(
                _SemanticProblem.CONDITIONAL_CHILDREN,
                part,
            )
        formula_text = (child.text or "").strip()
        if not formula_text:
            raise WorkbookSemanticError(
                _SemanticProblem.CONDITIONAL_EMPTY,
                part,
            )
        children.append((
            child.tag,
            tuple(sorted(child.attrib.items())),
            _translate_formula(
                formula_text,
                origin=origin,
                destination=FORMULA_NORMALIZATION_ANCHOR,
            ),
            (),
        ))
    return (
        rule.tag,
        attributes,
        (rule.text or "").strip(),
        tuple(children),
        dxf_signature,
    )


def _conditional_dxf_signature(
    rule: XmlElement,
    *,
    part: str,
    dxfs: tuple[tuple, ...],
) -> tuple | None:
    dxf_text = next(
        (value for name, value in rule.attrib.items() if _local_name(name) == "dxfId"),
        None,
    )
    if dxf_text is None:
        return None
    try:
        dxf_index = int(dxf_text)
    except ValueError as error:
        raise WorkbookSemanticError(
            _SemanticProblem.CONDITIONAL_DXF,
            part,
            dxf_text,
        ) from error
    if dxf_index < 0 or dxf_index >= len(dxfs):
        raise WorkbookSemanticError(
            _SemanticProblem.MISSING_DXF,
            part,
            dxf_index,
        )
    return dxfs[dxf_index]


def _conditional_rules(
    root: XmlElement,
    *,
    part: str,
    dxfs: tuple[tuple, ...],
) -> tuple[_ConditionalRule, ...]:
    rules: list[_ConditionalRule] = []
    priorities: set[int] = set()
    for container in root.iter():
        if _local_name(container.tag) != "conditionalFormatting":
            continue
        context = f"conditionalFormatting in {part}"
        sqref = _sqref_text(container, context=context)
        origin = _sqref_origin(sqref)
        coverage = _cells(sqref)
        for rule in container:
            if _local_name(rule.tag) != "cfRule":
                continue
            priority = _conditional_priority(rule, part=part)
            if priority in priorities:
                raise WorkbookSemanticError(
                    _SemanticProblem.DUPLICATE_CONDITIONAL_PRIORITY,
                    part,
                    priority,
                )
            priorities.add(priority)
            dxf_signature = _conditional_dxf_signature(rule, part=part, dxfs=dxfs)
            signature = _conditional_rule_signature(
                rule,
                part=part,
                origin=origin,
                dxf_signature=dxf_signature,
            )
            rules.append(_ConditionalRule(priority, signature, coverage))
    return tuple(rules)


def _conditional_manifest(rules: tuple[_ConditionalRule, ...]) -> tuple:
    signatures = tuple(sorted({rule.signature for rule in rules}, key=repr))
    signature_ids = {signature: index for index, signature in enumerate(signatures)}
    rule_ids_by_cell: dict[tuple[int, int], list[int]] = {}
    for rule_index, rule in enumerate(rules):
        for cell in rule.coverage:
            rule_ids_by_cell.setdefault(cell, []).append(rule_index)
    grouped: dict[tuple[int, ...], set[tuple[int, int]]] = {}
    for cell, rule_ids in rule_ids_by_cell.items():
        ordered_rule_ids = sorted(rule_ids, key=lambda index: rules[index].priority)
        rule_stack = tuple(signature_ids[rules[index].signature] for index in ordered_rule_ids)
        grouped.setdefault(rule_stack, set()).add(cell)
    coverage_by_stack = tuple(
        sorted(
            ((rule_stack, tuple(sorted(coverage))) for rule_stack, coverage in grouped.items()),
            key=repr,
        )
    )
    return signatures, coverage_by_stack


def _conditional_formats(package: zipfile.ZipFile) -> dict[str, tuple]:
    dxfs = _differential_formats(package)
    result: dict[str, tuple] = {}
    for part in package.namelist():
        if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", part) is None:
            continue
        root = DefusedET.fromstring(package.read(part))
        rules = _conditional_rules(root, part=part, dxfs=dxfs)
        result[part] = _conditional_manifest(rules)
    return result


def _normalize_vba_source(code: str) -> str:
    lines: list[str] = []
    for line in code.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if line.lstrip().startswith("Attribute "):
            continue
        lines.append(line.rstrip(" \t"))
    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines)


def _normalize_vba_project(project: bytes) -> bytes:
    lines = project.replace(b"\r\n", b"\n").replace(b"\r", b"\n").split(b"\n")
    return b"\n".join(
        line.rstrip(b" \t") for line in lines if not line.startswith(VOLATILE_VBA_PROJECT_FIELDS)
    ).rstrip(b"\n")


def _vba_semantics(package: zipfile.ZipFile) -> tuple | None:
    part = "xl/vbaProject.bin"
    if part not in package.namelist():
        return None
    payload = package.read(part)
    project_file = olefile.OleFileIO(io.BytesIO(payload))
    try:
        if not project_file.exists("PROJECT"):
            raise WorkbookSemanticError(_SemanticProblem.VBA_PROJECT_STREAM)
        if not project_file.exists("PROJECTwm"):
            raise WorkbookSemanticError(_SemanticProblem.VBA_NAMES_STREAM)
        project = _normalize_vba_project(project_file.openstream("PROJECT").read())
        project_names = project_file.openstream("PROJECTwm").read()
    finally:
        project_file.close()

    parser = VBA_Parser("vbaProject.bin", data=payload)
    modules: dict[str, str] = {}
    try:
        for _filename, _stream_path, module_name, code in parser.extract_macros():
            if module_name in modules:
                raise WorkbookSemanticError(
                    _SemanticProblem.DUPLICATE_VBA_MODULE,
                    module_name,
                )
            modules[module_name] = _normalize_vba_source(code or "")
    finally:
        parser.close()
    if not modules:
        raise WorkbookSemanticError(_SemanticProblem.NO_VBA_MODULES)
    return project, project_names, tuple(sorted(modules.items()))


def semantic_manifest(path: Path) -> dict[str, object]:
    """Return the exact authored semantics that Excel must preserve.

    Returns:
        A manifest of package parts and authored workbook behavior.

    Raises:
        WorkbookSemanticError: If a package part is malformed or inconsistent.

    """
    with zipfile.ZipFile(path) as package:
        bad_part = package.testzip()
        if bad_part is not None:
            raise WorkbookSemanticError(_SemanticProblem.CRC_FAILURE, bad_part)
        return {
            "parts": frozenset(package.namelist()),
            "worksheet formulas": _worksheet_formulas(package),
            "table formulas": _table_formulas(package),
            "defined names": _defined_names(package),
            "data validations": _data_validations(package),
            "conditional formats": _conditional_formats(package),
            "VBA project": _vba_semantics(package),
        }


def _mapping_issues(label: str, original: dict, saved: dict) -> list[str]:
    issues: list[str] = []
    for key in sorted(set(original) | set(saved), key=repr):
        if key not in saved:
            issues.append(f"{label} MISSING: {key!r}")
        elif key not in original:
            issues.append(f"{label} ADDED: {key!r}")
        elif original[key] != saved[key]:
            issues.append(f"{label} CHANGED: {key!r}")
    return issues


def compare_packages(source: Path, saved_copy: Path) -> list[str]:
    """Return every authored semantic difference between two packages.

    Returns:
        Labelled differences; an empty list means authored semantics match.

    """
    original = semantic_manifest(source)
    saved = semantic_manifest(saved_copy)
    issues = [
        f"PACKAGE PART MISSING: {part}" for part in sorted(original["parts"] - saved["parts"])
    ]
    issues.extend(
        _mapping_issues(
            "WORKSHEET FORMULA",
            original["worksheet formulas"],
            saved["worksheet formulas"],
        )
    )
    issues.extend(
        _mapping_issues("TABLE FORMULA", original["table formulas"], saved["table formulas"])
    )
    issues.extend(
        _mapping_issues("DEFINED NAME", original["defined names"], saved["defined names"])
    )
    issues.extend(
        _mapping_issues(
            "DATA VALIDATION",
            original["data validations"],
            saved["data validations"],
        )
    )
    original_rules = original["conditional formats"]
    saved_rules = saved["conditional formats"]
    issues.extend(
        f"CONDITIONAL FORMATS CHANGED: {part}"
        for part in sorted(set(original_rules) | set(saved_rules))
        if original_rules.get(part) != saved_rules.get(part)
    )
    if original["VBA project"] != saved["VBA project"]:
        issues.append("VBA PROJECT CHANGED")
    return issues


async def _execute_oracle(copy_path: Path) -> _AutomationResult:
    process = await asyncio.create_subprocess_exec(
        "/usr/bin/osascript",
        str(AUTOMATION / "capture_repair.applescript"),
        str(copy_path),
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    try:
        async with asyncio.timeout(ORACLE_TIMEOUT_SECONDS):
            stdout_bytes, stderr_bytes = await process.communicate()
    except TimeoutError as timeout_error:
        try:
            process.kill()
            await process.wait()
        except (ChildProcessError, OSError, ProcessLookupError) as cleanup_error:
            raise ExcelVerificationError(
                _VerificationProblem.ORACLE_TIMEOUT_CLEANUP,
                type(cleanup_error).__name__,
                cleanup_error,
            ) from timeout_error
        raise ExcelVerificationError(
            _VerificationProblem.ORACLE_TIMEOUT,
            ORACLE_TIMEOUT_SECONDS,
        ) from timeout_error
    return _AutomationResult(
        returncode=process.returncode or 0,
        stdout=stdout_bytes.decode("utf-8", errors="strict").strip(),
        stderr=stderr_bytes.decode("utf-8", errors="strict").strip(),
    )


def _run_oracle(copy_path: Path) -> None:
    completed = asyncio.run(_execute_oracle(copy_path))
    if completed.returncode != 0:
        diagnostic = completed.stderr or completed.stdout or "no diagnostic"
        raise ExcelVerificationError(
            _VerificationProblem.ORACLE_EXIT,
            copy_path.name,
            completed.returncode,
            diagnostic,
        )
    if completed.stderr:
        raise ExcelVerificationError(
            _VerificationProblem.ORACLE_STDERR,
            copy_path.name,
            completed.stderr,
        )
    if completed.stdout != SUCCESS:
        raise ExcelVerificationError(
            _VerificationProblem.ORACLE_SENTINEL,
            completed.stdout,
            SUCCESS,
        )


def _verify_disposable(
    source: Path,
    operation: Callable[[Path], None],
) -> list[str]:
    with excel_workbook_copy(source, prefix=f"{source.stem}.repcheck.") as saved_copy:
        operation(saved_copy)
        issues = compare_packages(source, saved_copy)
        extra_files = sorted(
            path.name for path in saved_copy.parent.iterdir() if path != saved_copy
        )
        issues.extend(f"EXCEL GENERATED DIAGNOSTIC FILE: {name}" for name in extra_files)
    return issues


def verify(source: Path) -> list[str]:
    """Save a disposable copy in Excel and compare exact authored semantics.

    Returns:
        Labelled authored-semantic differences.

    """
    return _verify_disposable(source, _run_oracle)


def recalculate_and_compare(source: Path) -> list[str]:
    """Full-rebuild a disposable copy and compare exact authored semantics.

    Returns:
        Labelled authored-semantic differences.

    """
    return _verify_disposable(source, recalculate)


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Verify exact workbook semantics across a desktop-Excel save."
    )
    parser.add_argument(
        "workbooks",
        nargs="*",
        help=".xlsx or .xlsm files; defaults to dist/PM_Workbook.xlsx",
    )
    return parser


def _write_line(message: str) -> None:
    sys.stdout.write(f"{message}\n")


def main(argv: list[str] | None = None) -> int:
    """Run the preservation oracle for every requested workbook.

    Returns:
        Zero when every workbook preserves authored semantics, otherwise one.

    Raises:
        FileNotFoundError: If a requested workbook does not exist.
        WorkbookSemanticError: If a requested file is not an Excel workbook.

    """
    arguments = _parser().parse_args(sys.argv[1:] if argv is None else argv)
    values = arguments.workbooks or ["dist/PM_Workbook.xlsx"]
    rejected = False
    for value in values:
        source = Path(value)
        if not source.is_absolute():
            source = ROOT / source
        source = source.resolve()
        if not source.is_file():
            raise FileNotFoundError(source)
        if source.suffix.lower() not in SUPPORTED_SUFFIXES:
            raise WorkbookSemanticError(_SemanticProblem.SUFFIX, source)

        issues = verify(source)
        if issues:
            rejected = True
            _write_line(f"EXCEL PACKAGE REJECTED — {source.name}: {len(issues)} issue(s)")
            for issue in issues:
                _write_line(f"  - {issue}")
        else:
            _write_line(f"EXCEL PACKAGE PASS — {source.name}: exact authored semantics preserved")
    return int(rejected)


if __name__ == "__main__":
    raise SystemExit(main())
