"""Strict JSON Schema contract for provider-neutral workbook change sets."""

from __future__ import annotations

import json
import math
from datetime import UTC, date, datetime
from typing import TYPE_CHECKING

from jsonschema import Draft202012Validator, FormatChecker

from .diagnostics import ContractError, Diagnostic, DiagnosticPhase

if TYPE_CHECKING:
    from collections.abc import Iterable, Mapping, Sequence

    from jsonschema.exceptions import ValidationError

CONTRACT_NAME = "excel-pm-agent-change-set"
CONTRACT_VERSION = "1.0.0"
SCHEMA_DIALECT = "https://json-schema.org/draft/2020-12/schema"
EXCEL_CELL_LIMIT = 32_767
IJSON_SAFE_INTEGER = 9_007_199_254_740_991
MAX_OPERATIONS = 4_000
OPERATION_PATH_PARTS = 2
SURROGATE_MIN = 0xD800
SURROGATE_MAX = 0xDFFF
XML_BASIC_MIN = 0x20
XML_BASIC_MAX = 0xD7FF
XML_EXTENDED_MIN = 0xE000
XML_EXTENDED_MAX = 0xFFFD
XML_SUPPLEMENTARY_MIN = 0x10000
XML_SUPPLEMENTARY_MAX = 0x10FFFF

_DIGEST_SCHEMA = {"type": "string", "pattern": "^[0-9a-f]{64}$"}
_EXTENSIONS_SCHEMA = {"type": "object"}
_CELL_TEXT_SCHEMA = {"type": "string", "minLength": 1, "maxLength": EXCEL_CELL_LIMIT}
_SHORT_TEXT_SCHEMA = {"type": "string", "minLength": 1, "maxLength": 255}
_DATE_SCHEMA = {"type": "string", "format": "date"}
_UTC_TIMESTAMP_SCHEMA = {
    "type": "string",
    "format": "date-time",
    "pattern": "Z$",
}

ITEM_WRITABLE_FIELDS = (
    "Type",
    "Title",
    "Status",
    "Delivery Health",
    "Priority",
    "Owner",
    "Start",
    "Due",
    "Parent",
    "BlockedBy",
    "Latest Status",
)
RAID_WRITABLE_FIELDS = (
    "Type",
    "Title",
    "Detail",
    "RelatedID",
    "Owner",
    "Status",
    "Prob",
    "Impact",
    "Response",
    "NextReview",
)


def _closed_object(
    properties: Mapping[str, object],
    *,
    required: Sequence[str] = (),
    min_properties: int | None = None,
) -> dict[str, object]:
    schema: dict[str, object] = {
        "type": "object",
        "properties": dict(properties),
        "additionalProperties": False,
    }
    if required:
        schema["required"] = list(required)
    if min_properties is not None:
        schema["minProperties"] = min_properties
    return schema


def _reference_schema() -> dict[str, object]:
    return {
        "oneOf": [
            _closed_object({"workbook_id": _CELL_TEXT_SCHEMA}, required=("workbook_id",)),
            _closed_object({"source": {"$ref": "#/$defs/sourceIdentity"}}, required=("source",)),
            _closed_object({"client_ref": _SHORT_TEXT_SCHEMA}, required=("client_ref",)),
        ]
    }


def _identity_schema() -> dict[str, object]:
    return _closed_object(
        {
            "workbook_id": _CELL_TEXT_SCHEMA,
            "source": {"$ref": "#/$defs/sourceIdentity"},
        },
        min_properties=1,
    )


def _item_set_schema() -> dict[str, object]:
    text_fields = (
        "Type",
        "Title",
        "Status",
        "Delivery Health",
        "Priority",
        "Owner",
        "Latest Status",
    )
    properties: dict[str, object] = dict.fromkeys(text_fields, _CELL_TEXT_SCHEMA)
    properties.update({
        "Start": _DATE_SCHEMA,
        "Due": _DATE_SCHEMA,
        "Parent": {"$ref": "#/$defs/itemReference"},
        "BlockedBy": {
            "type": "array",
            "items": {"$ref": "#/$defs/itemReference"},
            "minItems": 1,
            "maxItems": 2_000,
            "uniqueItems": True,
        },
    })
    return _closed_object(properties)


def _raid_set_schema() -> dict[str, object]:
    text_fields = ("Type", "Title", "Detail", "Owner", "Status", "Response")
    properties: dict[str, object] = dict.fromkeys(text_fields, _CELL_TEXT_SCHEMA)
    properties.update({
        "RelatedID": {"$ref": "#/$defs/itemReference"},
        "Prob": {"type": "integer", "minimum": 1, "maximum": 5},
        "Impact": {"type": "integer", "minimum": 1, "maximum": 5},
        "NextReview": _DATE_SCHEMA,
    })
    return _closed_object(properties)


def _upsert_schema(entity: str, writable_fields: Sequence[str]) -> dict[str, object]:
    return _closed_object(
        {
            "operation_id": _SHORT_TEXT_SCHEMA,
            "op": {"const": "upsert"},
            "entity": {"const": entity},
            "identity": {"$ref": "#/$defs/identity"},
            "client_ref": _SHORT_TEXT_SCHEMA,
            "set": {"$ref": f"#/$defs/{entity}Set"},
            "clear": {
                "type": "array",
                "items": {"enum": list(writable_fields)},
                "uniqueItems": True,
            },
            "extensions": _EXTENSIONS_SCHEMA,
        },
        required=("operation_id", "op", "entity", "identity"),
    )


def _mark_deleted_schema(entity: str) -> dict[str, object]:
    return _closed_object(
        {
            "operation_id": _SHORT_TEXT_SCHEMA,
            "op": {"const": "mark_deleted"},
            "entity": {"const": entity},
            "identity": {"$ref": "#/$defs/identity"},
            "extensions": _EXTENSIONS_SCHEMA,
        },
        required=("operation_id", "op", "entity", "identity"),
    )


def _build_schema() -> dict[str, object]:
    schema = _closed_object(
        {
            "contract": {"const": CONTRACT_NAME},
            "version": {"const": CONTRACT_VERSION},
            "request_id": {"type": "string", "format": "uuid"},
            "created_at": _UTC_TIMESTAMP_SCHEMA,
            "target": {"$ref": "#/$defs/target"},
            "producer": {"$ref": "#/$defs/producer"},
            "source": {"$ref": "#/$defs/provenance"},
            "operations": {
                "type": "array",
                "maxItems": MAX_OPERATIONS,
                "items": {"$ref": "#/$defs/operation"},
            },
            "extensions": _EXTENSIONS_SCHEMA,
        },
        required=(
            "contract",
            "version",
            "request_id",
            "created_at",
            "target",
            "producer",
            "source",
            "operations",
        ),
    )
    schema.update({
        "$schema": SCHEMA_DIALECT,
        "$id": "https://maxymax92.github.io/excel-pm-tool/agent-change-set-1.0.0.schema.json",
        "title": "Excel PM workbook agent change set",
        "$defs": {
            "sourceIdentity": _closed_object(
                {"namespace": _CELL_TEXT_SCHEMA, "record_id": _CELL_TEXT_SCHEMA},
                required=("namespace", "record_id"),
            ),
            "identity": _identity_schema(),
            "itemReference": _reference_schema(),
            "itemSet": _item_set_schema(),
            "raidSet": _raid_set_schema(),
            "target": _closed_object(
                {
                    "workbook_sha256": _DIGEST_SCHEMA,
                    "workbook_schema_fingerprint": _DIGEST_SCHEMA,
                    "build_schema_fingerprint": _DIGEST_SCHEMA,
                },
                required=(
                    "workbook_sha256",
                    "workbook_schema_fingerprint",
                    "build_schema_fingerprint",
                ),
            ),
            "producer": _closed_object(
                {
                    "name": _CELL_TEXT_SCHEMA,
                    "version": _SHORT_TEXT_SCHEMA,
                    "extensions": _EXTENSIONS_SCHEMA,
                },
                required=("name",),
            ),
            "provenance": _closed_object(
                {
                    "description": _CELL_TEXT_SCHEMA,
                    "uri": {"type": "string", "maxLength": 4_096},
                    "retrieved_at": _UTC_TIMESTAMP_SCHEMA,
                    "extensions": _EXTENSIONS_SCHEMA,
                },
                required=("description",),
            ),
            "operation": {
                "oneOf": [
                    _upsert_schema("item", ITEM_WRITABLE_FIELDS),
                    _upsert_schema("raid", RAID_WRITABLE_FIELDS),
                    _mark_deleted_schema("item"),
                    _mark_deleted_schema("raid"),
                ]
            },
        },
    })
    return schema


CHANGESET_SCHEMA = _build_schema()
Draft202012Validator.check_schema(CHANGESET_SCHEMA)
_VALIDATOR = Draft202012Validator(CHANGESET_SCHEMA, format_checker=FormatChecker())


class _DuplicateKeyError(ValueError):
    """Carry the first repeated JSON object key."""


class _NonFiniteError(ValueError):
    """Carry one non-finite JSON number spelling."""


def _diagnostic(
    code: str,
    phase: DiagnosticPhase,
    message: str,
    hint: str,
    *,
    location: tuple[str, str | None] = ("", None),
) -> Diagnostic:
    pointer, operation_id = location
    return Diagnostic(
        code=code,
        severity="error",
        phase=phase,
        pointer=pointer,
        operation_id=operation_id,
        message=message,
        hint=hint,
    )


def _object_without_duplicates(pairs: Iterable[tuple[str, object]]) -> dict[str, object]:
    result: dict[str, object] = {}
    for key, value in pairs:
        if key in result:
            raise _DuplicateKeyError(key)
        result[key] = value
    return result


def _reject_constant(value: str) -> object:
    raise _NonFiniteError(value)


def _escape_pointer(value: object) -> str:
    return str(value).replace("~", "~0").replace("/", "~1")


def _pointer(path: Iterable[object]) -> str:
    parts = [_escape_pointer(value) for value in path]
    return "" if not parts else "/" + "/".join(parts)


def _operation_id(document: object, path: Sequence[object]) -> str | None:
    if len(path) < OPERATION_PATH_PARTS or path[0] != "operations" or not isinstance(path[1], int):
        return None
    if not isinstance(document, dict):
        return None
    operations = document.get("operations")
    if not isinstance(operations, list) or path[1] >= len(operations):
        return None
    operation = operations[path[1]]
    if not isinstance(operation, dict):
        return None
    value = operation.get("operation_id")
    return value if isinstance(value, str) else None


def _leaf_errors(error: ValidationError) -> list[ValidationError]:
    if not error.context:
        return [error]
    leaves: list[ValidationError] = []
    for nested in error.context:
        leaves.extend(_leaf_errors(nested))
    return leaves


def _schema_diagnostics(document: object) -> tuple[Diagnostic, ...]:
    leaves = [leaf for error in _VALIDATOR.iter_errors(document) for leaf in _leaf_errors(error)]
    unique: dict[tuple[object, ...], ValidationError] = {}
    for error in leaves:
        key = (*error.absolute_path, error.message)
        unique[key] = error
    ordered = sorted(unique.values(), key=lambda error: (-len(error.absolute_path), error.message))
    return tuple(
        _diagnostic(
            "contract.schema",
            "schema",
            error.message,
            "Conform the value to the embedded contract schema returned by describe.",
            location=(
                _pointer(error.absolute_path),
                _operation_id(document, tuple(error.absolute_path)),
            ),
        )
        for error in ordered
    )


def _validate_string(value: str, pointer: str) -> None:
    for character in value:
        codepoint = ord(character)
        if SURROGATE_MIN <= codepoint <= SURROGATE_MAX:
            raise ContractError(
                _diagnostic(
                    "json.unpaired_surrogate",
                    "parse",
                    "JSON text contains an unpaired Unicode surrogate.",
                    "Replace it with a valid Unicode scalar value.",
                    location=(pointer, None),
                )
            )
        valid_xml = (
            codepoint in {0x09, 0x0A, 0x0D}
            or XML_BASIC_MIN <= codepoint <= XML_BASIC_MAX
            or XML_EXTENDED_MIN <= codepoint <= XML_EXTENDED_MAX
            or XML_SUPPLEMENTARY_MIN <= codepoint <= XML_SUPPLEMENTARY_MAX
        )
        if not valid_xml:
            raise ContractError(
                _diagnostic(
                    "json.illegal_xml",
                    "parse",
                    f"JSON text contains XML-illegal character U+{codepoint:04X}.",
                    "Remove control characters that Excel cannot store in worksheet XML.",
                    location=(pointer, None),
                )
            )


def _validate_number(value: int | float, pointer: str) -> None:
    if isinstance(value, int):
        if abs(value) > IJSON_SAFE_INTEGER:
            raise ContractError(
                _diagnostic(
                    "json.number_range",
                    "parse",
                    f"Integer {value} is outside the interoperable I-JSON range.",
                    f"Use an integer between {-IJSON_SAFE_INTEGER} and {IJSON_SAFE_INTEGER}.",
                    location=(pointer, None),
                )
            )
        return
    if not math.isfinite(value):
        raise ContractError(
            _diagnostic(
                "json.non_finite",
                "parse",
                "JSON contains a non-finite number.",
                "Use a finite JSON number or encode the external identifier as a string.",
                location=(pointer, None),
            )
        )


def _validate_ijson(value: object, path: tuple[object, ...] = ()) -> None:
    pointer = _pointer(path)
    if isinstance(value, str):
        _validate_string(value, pointer)
        return
    if isinstance(value, bool) or value is None:
        return
    if isinstance(value, (int, float)):
        _validate_number(value, pointer)
        return
    if isinstance(value, list):
        for index, nested in enumerate(value):
            _validate_ijson(nested, (*path, index))
        return
    if isinstance(value, dict):
        for key, nested in value.items():
            _validate_string(key, pointer)
            _validate_ijson(nested, (*path, key))


def _field_conflict_diagnostics(document: dict[str, object]) -> tuple[Diagnostic, ...]:
    operations = document.get("operations")
    if not isinstance(operations, list):
        return ()
    diagnostics: list[Diagnostic] = []
    for index, operation in enumerate(operations):
        if not isinstance(operation, dict) or operation.get("op") != "upsert":
            continue
        set_fields = operation.get("set", {})
        clear_fields = operation.get("clear", [])
        if not isinstance(set_fields, dict) or not isinstance(clear_fields, list):
            continue
        overlap = sorted(set(set_fields) & set(clear_fields))
        if not overlap:
            continue
        operation_id = operation.get("operation_id")
        diagnostics.append(
            _diagnostic(
                "contract.field_conflict",
                "schema",
                f"Fields appear in both set and clear: {', '.join(overlap)}.",
                "Remove each field from either set or clear so its intent is unambiguous.",
                location=(
                    f"/operations/{index}",
                    operation_id if isinstance(operation_id, str) else None,
                ),
            )
        )
    return tuple(diagnostics)


def parse_changeset(payload: bytes) -> dict[str, object]:
    """Parse and validate one complete change-set document.

    Returns:
        The validated JSON object without semantic mutation.

    Raises:
        ContractError: If bytes, JSON, I-JSON rules or the public schema fail.

    """
    try:
        text = payload.decode("utf-8", errors="strict")
    except UnicodeDecodeError as error:
        raise ContractError(
            _diagnostic(
                "json.invalid_utf8",
                "parse",
                f"Change set is not valid UTF-8 at byte {error.start}.",
                "Encode the complete JSON document as UTF-8 without replacement characters.",
            )
        ) from error
    try:
        document = json.loads(
            text,
            object_pairs_hook=_object_without_duplicates,
            parse_constant=_reject_constant,
        )
    except _DuplicateKeyError as error:
        raise ContractError(
            _diagnostic(
                "json.duplicate_key",
                "parse",
                f"JSON object repeats key {error.args[0]!r}.",
                "Keep exactly one occurrence of every object key.",
            )
        ) from error
    except _NonFiniteError as error:
        raise ContractError(
            _diagnostic(
                "json.non_finite",
                "parse",
                f"JSON contains non-finite number {error.args[0]!r}.",
                "Use a finite JSON number or encode the external identifier as a string.",
            )
        ) from error
    except json.JSONDecodeError as error:
        raise ContractError(
            _diagnostic(
                "json.syntax",
                "parse",
                f"Change set is not valid JSON at line {error.lineno}, column {error.colno}.",
                "Correct the JSON syntax and submit one complete document.",
            )
        ) from error

    _validate_ijson(document)
    schema_diagnostics = _schema_diagnostics(document)
    if schema_diagnostics:
        raise ContractError(*schema_diagnostics)
    if not isinstance(document, dict):
        raise ContractError(
            _diagnostic(
                "contract.schema",
                "schema",
                "Change-set root must be an object.",
                "Use the object envelope returned by describe.",
            )
        )
    conflicts = _field_conflict_diagnostics(document)
    if conflicts:
        raise ContractError(*conflicts)
    return document


def _canonical_value(value: object) -> object:
    if isinstance(value, datetime):
        if value.tzinfo is None:
            message = "canonical datetime values must be timezone-aware"
            raise ValueError(message)
        normalized = value.astimezone(UTC).isoformat(timespec="seconds")
        return normalized.replace("+00:00", "Z")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, dict):
        return {str(key): _canonical_value(nested) for key, nested in value.items()}
    if isinstance(value, (list, tuple)):
        return [_canonical_value(nested) for nested in value]
    return value


def canonical_json(value: object) -> bytes:
    """Serialize stable compact UTF-8 JSON for hashing.

    Returns:
        Canonical UTF-8 bytes.

    Raises:
        ValueError: If the value cannot be represented as strict I-JSON.

    """
    normalized = _canonical_value(value)
    try:
        _validate_ijson(normalized)
    except ContractError as error:
        raise ValueError(error.diagnostics[0].message) from error
    try:
        serialized = json.dumps(
            normalized,
            ensure_ascii=False,
            allow_nan=False,
            sort_keys=True,
            separators=(",", ":"),
        )
        return serialized.encode("utf-8", errors="strict")
    except (TypeError, UnicodeEncodeError, ValueError) as error:
        message = "value cannot be represented as canonical I-JSON"
        raise ValueError(message) from error


__all__ = [
    "CHANGESET_SCHEMA",
    "CONTRACT_NAME",
    "CONTRACT_VERSION",
    "ITEM_WRITABLE_FIELDS",
    "RAID_WRITABLE_FIELDS",
    "ContractError",
    "canonical_json",
    "parse_changeset",
]
