"""One-way monday.com board import through the authored-data layer.

The importer reads a board over monday's GraphQL API, maps items and subitems
onto the workbook's Items hierarchy, merges them into an exported snapshot and
publishes through the same injected rebuild as a migration. A per-board
identifier map under ``dist/monday/`` keeps re-imports idempotent: mapped
items update in place, new items receive fresh workbook identifiers, and rows
the user deleted from the workbook stay deleted.

monday owns the fields it supplies on mapped rows: Title, Type, Status,
Priority, Owner, Start, Due and ``Updated`` (which tracks monday's last
activity). Workbook-only fields — Delivery Health, Latest Status, BlockedBy
and the workflow stamps Created, InProgressSince, DoneDate, BlockedSince and
LatestUpdateOn — are never touched on update. Clearing follows monday for
Priority, Owner, Start and Due: when the mapped column comes back empty, the
workbook cell is cleared and reported. Status is the one exception — the
importer defaults it on statusless new items, so an empty monday status keeps
the workbook value instead of flip-flopping it.
"""

from __future__ import annotations

import http.client
import json
import logging
import os
import time
from dataclasses import dataclass, field, replace
from datetime import UTC, date, datetime
from enum import Enum
from pathlib import Path
from typing import TYPE_CHECKING

from ..paths import DIST
from ..pipeline import require_current_vba
from .export import export_workbook
from .inject import validate_snapshot
from .migrate import DEFAULT_WORKBOOK, _log_export, _require_workbook, rebuild_and_publish
from .snapshot import atomic_write_json, write_snapshot

if TYPE_CHECKING:
    from collections.abc import Callable, Iterator

    from .snapshot import Snapshot

LOGGER = logging.getLogger(__name__)

API_HOST = "api.monday.com"
API_PATH = "/v2"
API_VERSION = "2026-07"
AUTH_ENV = "MONDAY_API_TOKEN"
PAGE_SIZE = 250
MAX_ATTEMPTS = 5
_HTTP_OK = 200
_DEFAULT_ITEM_LEVEL = 5
_DEFAULT_SUBITEM_LEVEL = 6
HTTP_TIMEOUT_SECONDS = 60
MAP_FORMAT = 1

_RETRYABLE_STATUS = frozenset({408, 429, 500, 502, 503, 504})
# One hard bound on cursor pagination: far above the workbook's row capacity,
# so it only ever trips on a service that stops advancing its cursor.
MAX_PAGES = 200
_ITEM_FIELDS = (
    "id name created_at updated_at board { id } parent_item { id } column_values { id text value }"
)
_PAGE_ITEMS = "cursor items { FIELDS subitems { FIELDS } }".replace("FIELDS", _ITEM_FIELDS)
_FIRST_PAGE_QUERY = (
    "query ($board: [ID!], $limit: Int!) { boards(ids: $board) { id name "
    "columns { id title type } items_page(limit: $limit) { " + _PAGE_ITEMS + " } } }"
)
_NEXT_PAGE_QUERY = (
    "query ($cursor: String!, $limit: Int!) { next_items_page(cursor: $cursor, limit: $limit) { "
    + _PAGE_ITEMS
    + " } }"
)
_CATALOG_QUERY = "query ($boards: [ID!]) { boards(ids: $boards) { id columns { id title type } } }"

# Workbook field -> case-insensitive monday column titles accepted for it, in
# priority order. A column type listed as the fallback is used when no title
# matches.
_COLUMN_HINTS: dict[str, tuple[tuple[str, ...], str | None]] = {
    "type": (("type",), None),
    "status": (("status",), "status"),
    "priority": (("priority",), None),
    "owner": (("owner", "assignee", "person", "people"), "people"),
    "start": (("start", "start date"), None),
    "due": (("due", "due date", "deadline"), None),
    "timeline": (("timeline", "dates"), "timeline"),
}


class _MondayProblem(Enum):
    MISSING_CREDENTIALS = (
        "no monday.com API token found in the {} environment variable; "
        "create one under your monday profile's Developers section"
    )
    HTTP_STATUS = "monday.com returned HTTP {} for the GraphQL request: {}"
    API_ERRORS = "monday.com reported: {}"
    MALFORMED = "monday.com returned an unexpected response shape: {}"
    RETRY_EXHAUSTED = "monday.com kept failing after {} attempts; last failure: {}"
    PAGE_OVERFLOW = (
        "board {} paginated past {} pages without exhausting its cursor; "
        "the service is not advancing pagination"
    )
    BOARD_MISSING = "board {} was not returned; check the identifier and the token's access"
    NO_LEVEL_ONE_TYPE = "tblTypes has no Level 1 type to hold the imported board; add one in Config"
    DUPLICATE_TITLES = (
        "the identifier map has no record for these monday items, but rows with the same "
        "titles already exist: {}; restore dist/monday/{} or re-run with --allow-duplicates"
    )
    MAP_MALFORMED = "identifier map {} is malformed: {}"
    MAP_PENDING = (
        "a previous import was interrupted between publishing and recording its map; "
        "if {} contains the imported rows, rename {} to {}, otherwise delete it, then re-run"
    )
    VALUE_JSON = "monday.com column {} on item {} holds unparseable JSON: {}"
    VALUE_STAMP = "monday.com item {} has an unparseable {} timestamp: {!r}"


class MondayImportError(RuntimeError):
    """Report an unusable monday.com response, mapping or identifier map."""

    def __init__(self, problem: _MondayProblem, *details: object) -> None:
        """Create an error from a stable diagnostic template."""
        super().__init__(problem.value.format(*details))


@dataclass(frozen=True, kw_only=True, slots=True)
class PostResult:
    """One raw GraphQL exchange outcome."""

    status: int
    retry_after: int | None
    document: dict[str, object] | None


@dataclass(frozen=True, kw_only=True, slots=True)
class ColumnValue:
    """One monday column value on one item."""

    text: str | None
    value: str | None


@dataclass(frozen=True, kw_only=True, slots=True)
class MondayItem:
    """One board item or subitem with its parsed metadata."""

    identifier: str
    name: str
    created: date | None
    updated: date | None
    board_id: str | None
    parent_id: str | None
    values: dict[str, ColumnValue]


@dataclass(frozen=True, kw_only=True, slots=True)
class BoardData:
    """One board's name, per-board column catalogues and flattened item list.

    Subitems live on their own monday board with distinct column identifiers,
    so ``catalogs`` holds one column catalogue per board identifier seen in
    the item list, keyed the way each item's ``board_id`` reports it.
    """

    identifier: str
    name: str
    catalogs: dict[str, dict[str, tuple[str, str]]]
    items: tuple[MondayItem, ...]


@dataclass(frozen=True, kw_only=True, slots=True)
class ImportOutcome:
    """The merged snapshot, identifier map and per-run reporting."""

    snapshot: Snapshot
    map_document: dict[str, object]
    added: int
    updated: int
    unchanged: int
    skipped_deleted: tuple[str, ...]
    notes: tuple[str, ...]


def _default_post(token: str, payload: dict[str, object]) -> PostResult:
    connection = http.client.HTTPSConnection(API_HOST, timeout=HTTP_TIMEOUT_SECONDS)
    try:
        connection.request(
            "POST",
            API_PATH,
            body=json.dumps(payload),
            headers={
                "Authorization": token,
                "API-Version": API_VERSION,
                "Content-Type": "application/json",
            },
        )
        response = connection.getresponse()
        raw = response.read()
        retry_header = response.getheader("Retry-After")
    finally:
        connection.close()
    try:
        document = json.loads(raw)
    except json.JSONDecodeError:
        document = None
    retry_after = int(retry_header) if retry_header and retry_header.isdigit() else None
    return PostResult(status=response.status, retry_after=retry_after, document=document)


def _error_messages(document: dict[str, object] | None) -> list[str]:
    if not isinstance(document, dict):
        return []
    errors = document.get("errors")
    if not isinstance(errors, list):
        return []
    return [str(error.get("message", error)) for error in errors if isinstance(error, dict)]


def _retry_seconds(document: dict[str, object] | None) -> int | None:
    if not isinstance(document, dict):
        return None
    errors = document.get("errors")
    if not isinstance(errors, list):
        return None
    for error in errors:
        if not isinstance(error, dict):
            continue
        extensions = error.get("extensions")
        if isinstance(extensions, dict) and "retry_in_seconds" in extensions:
            seconds = extensions["retry_in_seconds"]
            if isinstance(seconds, (int, float)) and not isinstance(seconds, bool):
                return max(1, int(seconds))
    return None


def _execute(
    payload: dict[str, object],
    *,
    token: str,
    post: Callable[[str, dict[str, object]], PostResult],
    sleeper: Callable[[float], None],
) -> dict[str, object]:
    last_failure = "no request was attempted"
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            result = post(token, payload)
        except (OSError, http.client.HTTPException) as error:
            last_failure = f"{type(error).__name__}: {error}"
            if attempt < MAX_ATTEMPTS:
                sleeper(2**attempt)
            continue
        retry_wait = _retry_seconds(result.document)
        throttled = result.status == _HTTP_OK and retry_wait is not None
        if result.status in _RETRYABLE_STATUS or throttled:
            detail = "; ".join(_error_messages(result.document)) or "throttled"
            last_failure = f"HTTP {result.status} ({detail})"
            if attempt < MAX_ATTEMPTS:
                sleeper(result.retry_after or retry_wait or 2**attempt)
            continue
        if result.status != _HTTP_OK:
            detail = "; ".join(_error_messages(result.document)) or "no error detail"
            raise MondayImportError(_MondayProblem.HTTP_STATUS, result.status, detail)
        messages = _error_messages(result.document)
        if messages:
            raise MondayImportError(_MondayProblem.API_ERRORS, "; ".join(messages))
        data = result.document.get("data") if isinstance(result.document, dict) else None
        if not isinstance(data, dict):
            raise MondayImportError(_MondayProblem.MALFORMED, "no data object in the response")
        return data
    raise MondayImportError(_MondayProblem.RETRY_EXHAUSTED, MAX_ATTEMPTS, last_failure)


def _parse_stamp(item_id: str, field: str, value: object) -> date | None:
    if value in {None, ""}:
        return None
    try:
        stamp = datetime.fromisoformat(str(value))
    except ValueError as error:
        raise MondayImportError(_MondayProblem.VALUE_STAMP, item_id, field, value) from error
    if stamp.tzinfo is None:
        return stamp.date()
    return stamp.astimezone(UTC).date()


def _parse_item(node: dict[str, object], parent_id: str | None) -> MondayItem:
    values: dict[str, ColumnValue] = {}
    column_values = node.get("column_values")
    if isinstance(column_values, list):
        for entry in column_values:
            if not isinstance(entry, dict):
                continue
            text = entry.get("text")
            value = entry.get("value")
            values[str(entry.get("id"))] = ColumnValue(
                text=str(text) if text not in {None, ""} else None,
                value=str(value) if value not in {None, ""} else None,
            )
    parent_node = node.get("parent_item")
    node_parent = (
        str(parent_node["id"])
        if isinstance(parent_node, dict) and parent_node.get("id") is not None
        else parent_id
    )
    board_node = node.get("board")
    board_id = (
        str(board_node["id"])
        if isinstance(board_node, dict) and board_node.get("id") is not None
        else None
    )
    identifier = str(node.get("id"))
    return MondayItem(
        identifier=identifier,
        name=str(node.get("name", "")),
        created=_parse_stamp(identifier, "created_at", node.get("created_at")),
        updated=_parse_stamp(identifier, "updated_at", node.get("updated_at")),
        board_id=board_id,
        parent_id=node_parent,
        values=values,
    )


def _flatten_items(nodes: object) -> Iterator[MondayItem]:
    if not isinstance(nodes, list):
        return
    for node in nodes:
        if not isinstance(node, dict):
            continue
        item = _parse_item(node, None)
        yield item
        subitems = node.get("subitems")
        if isinstance(subitems, list):
            for subitem in subitems:
                if isinstance(subitem, dict):
                    yield _parse_item(subitem, item.identifier)


def fetch_board(
    board_id: int,
    *,
    token: str,
    post: Callable[[str, dict[str, object]], PostResult] | None = None,
    sleeper: Callable[[float], None] | None = None,
) -> BoardData:
    """Fetch one board's metadata and every item through cursor pagination.

    Returns:
        The board name, column catalogue and flattened items.

    Raises:
        MondayImportError: If the API refuses the request or the board is absent.

    """
    active_post = post if post is not None else _default_post
    active_sleeper = sleeper if sleeper is not None else time.sleep

    page = _execute(
        {"query": _FIRST_PAGE_QUERY, "variables": {"board": [str(board_id)], "limit": PAGE_SIZE}},
        token=token,
        post=active_post,
        sleeper=active_sleeper,
    )
    boards = page.get("boards")
    if not isinstance(boards, list) or not boards or not isinstance(boards[0], dict):
        raise MondayImportError(_MondayProblem.BOARD_MISSING, board_id)
    board = boards[0]
    parent_board_id = str(board.get("id", board_id))
    catalogs = {parent_board_id: _catalog(board)}

    items: list[MondayItem] = []
    items_page = board.get("items_page")
    previous_cursor: object = None
    for _page_number in range(MAX_PAGES):
        if not isinstance(items_page, dict):
            break
        items.extend(_flatten_items(items_page.get("items")))
        cursor = items_page.get("cursor")
        if cursor in {None, ""}:
            break
        if cursor == previous_cursor:
            raise MondayImportError(_MondayProblem.PAGE_OVERFLOW, board_id, MAX_PAGES)
        previous_cursor = cursor
        follow_up = _execute(
            {"query": _NEXT_PAGE_QUERY, "variables": {"cursor": str(cursor), "limit": PAGE_SIZE}},
            token=token,
            post=active_post,
            sleeper=active_sleeper,
        )
        items_page = follow_up.get("next_items_page")
    else:
        raise MondayImportError(_MondayProblem.PAGE_OVERFLOW, board_id, MAX_PAGES)

    foreign_boards = sorted({
        item.board_id
        for item in items
        if item.board_id is not None and item.board_id not in catalogs
    })
    if foreign_boards:
        catalog_data = _execute(
            {"query": _CATALOG_QUERY, "variables": {"boards": foreign_boards}},
            token=token,
            post=active_post,
            sleeper=active_sleeper,
        )
        catalog_boards = catalog_data.get("boards")
        if isinstance(catalog_boards, list):
            for entry in catalog_boards:
                if isinstance(entry, dict) and entry.get("id") is not None:
                    catalogs[str(entry["id"])] = _catalog(entry)

    return BoardData(
        identifier=parent_board_id,
        name=str(board.get("name", f"board {board_id}")),
        catalogs=catalogs,
        items=tuple(items),
    )


def _catalog(board_node: dict[str, object]) -> dict[str, tuple[str, str]]:
    columns = board_node.get("columns")
    catalog: dict[str, tuple[str, str]] = {}
    for column in columns if isinstance(columns, list) else []:
        if isinstance(column, dict):
            catalog[str(column.get("id"))] = (
                str(column.get("title", "")),
                str(column.get("type", "")),
            )
    return catalog


def _json_value(item: MondayItem, column_id: str | None, field: str) -> object:
    if column_id is None:
        return None
    holder = item.values.get(column_id)
    if holder is None or holder.value is None:
        return None
    try:
        parsed = json.loads(holder.value)
    except json.JSONDecodeError as error:
        raise MondayImportError(
            _MondayProblem.VALUE_JSON,
            column_id,
            item.identifier,
            holder.value,
        ) from error
    return parsed.get(field) if isinstance(parsed, dict) else None


def _column_date(item: MondayItem, column_id: str | None, field: str) -> date | None:
    raw = _json_value(item, column_id, field)
    if raw in {None, ""}:
        return None
    try:
        return date.fromisoformat(str(raw)[:10])
    except ValueError as error:
        raise MondayImportError(
            _MondayProblem.VALUE_JSON,
            str(column_id),
            item.identifier,
            raw,
        ) from error


def _column_text(item: MondayItem, column_id: str | None) -> str | None:
    if column_id is None:
        return None
    holder = item.values.get(column_id)
    return holder.text if holder is not None else None


def _resolve_catalog(
    catalog: dict[str, tuple[str, str]],
    notes: list[str],
    label: str,
) -> dict[str, str | None]:
    resolved: dict[str, str | None] = {}
    for field_name, (titles, fallback_type) in _COLUMN_HINTS.items():
        chosen: str | None = None
        for wanted in titles:
            for column_id, (title, _column_type) in catalog.items():
                if title.casefold() == wanted:
                    chosen = column_id
                    break
            if chosen is not None:
                break
        if chosen is None and fallback_type is not None:
            chosen = next(
                (
                    column_id
                    for column_id, (_title, column_type) in catalog.items()
                    if column_type == fallback_type
                ),
                None,
            )
        resolved[field_name] = chosen
        if chosen is not None:
            title, _column_type = catalog[chosen]
            notes.append(f"{label}: monday column {title!r} supplies {field_name}")
    return resolved


def _resolve_boards(board: BoardData, notes: list[str]) -> dict[str, dict[str, str | None]]:
    resolutions: dict[str, dict[str, str | None]] = {}
    for board_identifier, catalog in sorted(board.catalogs.items()):
        label = "board" if board_identifier == board.identifier else "subitem board"
        resolutions[board_identifier] = _resolve_catalog(catalog, notes, label)
    return resolutions


@dataclass(frozen=True, kw_only=True, slots=True)
class _Taxonomy:
    """Workbook taxonomy context resolved from one snapshot."""

    container_type: str
    item_type: str
    subitem_type: str
    canonical: dict[str, str]
    first_status: str | None


def _taxonomy(snapshot: Snapshot, notes: list[str]) -> _Taxonomy:
    types = [
        (str(row.get("Type", "")), row.get("Level")) for row in snapshot.tables.get("tblTypes", ())
    ]
    container = next((name for name, level in types if level == 1), None)
    if container is None:
        raise MondayImportError(_MondayProblem.NO_LEVEL_ONE_TYPE)
    deepest = max((level for _name, level in types if isinstance(level, int)), default=1)
    item_type = next((name for name, level in types if level == _DEFAULT_ITEM_LEVEL), None) or next(
        (name for name, level in types if level == deepest),
        container,
    )
    subitem_type = (
        next((name for name, level in types if level == _DEFAULT_SUBITEM_LEVEL), None) or item_type
    )
    statuses = [str(row.get("Status", "")) for row in snapshot.tables.get("tblStatuses", ())]
    notes.append(
        f"defaults: board container {container!r}, items {item_type!r}, subitems {subitem_type!r}"
    )
    return _Taxonomy(
        container_type=container,
        item_type=item_type,
        subitem_type=subitem_type,
        canonical={name.casefold(): name for name, _level in types},
        first_status=statuses[0] if statuses else None,
    )


class _IdAllocator:
    """Sequential workbook identifiers that skip every number in use."""

    def __init__(self, prefix: str, counter: int, in_use: set[str]) -> None:
        self._prefix = prefix
        self._counter = counter
        self._in_use = in_use

    def allocate(self) -> str:
        """Reserve and return the next free workbook identifier.

        Returns:
            The allocated identifier.

        """
        candidate = f"{self._prefix}{self._counter}"
        while candidate in self._in_use:
            self._counter += 1
            candidate = f"{self._prefix}{self._counter}"
        self._counter += 1
        self._in_use.add(candidate)
        return candidate

    @property
    def counter(self) -> int:
        """Return the next number the workbook counter must hold.

        Returns:
            The next unassigned identifier number.

        """
        return self._counter


@dataclass(kw_only=True, slots=True)
class _MergeState:
    """Mutable bookkeeping shared across one board merge."""

    rows: list[dict[str, object]]
    by_id: dict[str, dict[str, object]]
    item_map: dict[str, str]
    allocator: _IdAllocator
    resolutions: dict[str, dict[str, str | None]]
    parent_board: str
    taxonomy: _Taxonomy
    notes: list[str]
    supplied_owners: set[str] = field(default_factory=set)
    added: int = 0
    updated: int = 0
    unchanged: int = 0

    def columns_for(self, item: MondayItem) -> dict[str, str | None]:
        """Return the column resolution for the board one item lives on.

        Returns:
            The field-to-column mapping of the item's own board.

        """
        board_id = item.board_id if item.board_id is not None else self.parent_board
        return self.resolutions.get(board_id, self.resolutions[self.parent_board])


def _supplied_owner(item: MondayItem, state: _MergeState, owner: str) -> str:
    first, _separator, rest = owner.partition(",")
    kept = first.strip()
    dropped = ", ".join(name.strip() for name in rest.split(",") if name.strip())
    if dropped:
        state.notes.append(
            f"{item.name!r}: Owner holds one person; kept {kept!r}, "
            f"co-owner(s) {dropped} not imported"
        )
    state.supplied_owners.add(kept)
    return kept


def _supplied_fields(
    item: MondayItem,
    state: _MergeState,
) -> tuple[dict[str, object], tuple[str, ...]]:
    columns = state.columns_for(item)
    supplied: dict[str, object] = {"Title": item.name}
    cleared: list[str] = []

    def _optional(field_name: str, source_column: str | None, value: object) -> None:
        if value is not None:
            supplied[field_name] = value
        elif source_column is not None:
            cleared.append(field_name)

    status = _column_text(item, columns["status"])
    if status is not None:
        supplied["Status"] = status
    _optional("Priority", columns["priority"], _column_text(item, columns["priority"]))
    owner = _column_text(item, columns["owner"])
    _optional(
        "Owner",
        columns["owner"],
        _supplied_owner(item, state, owner) if owner is not None else None,
    )
    start = _column_date(item, columns["start"], "date") or _column_date(
        item, columns["timeline"], "from"
    )
    _optional("Start", columns["start"] or columns["timeline"], start)
    due = _column_date(item, columns["due"], "date") or _column_date(
        item, columns["timeline"], "to"
    )
    _optional("Due", columns["due"] or columns["timeline"], due)
    if item.updated is not None:
        supplied["Updated"] = item.updated
    type_label = _column_text(item, columns["type"])
    if type_label is not None:
        canonical = state.taxonomy.canonical.get(type_label.casefold())
        if canonical is None:
            state.notes.append(
                f"monday type {type_label!r} is not in tblTypes; the default type was used"
            )
        else:
            supplied["Type"] = canonical
    return supplied, tuple(cleared)


def _new_row(item: MondayItem, state: _MergeState, parent: str) -> dict[str, object]:
    supplied, _cleared = _supplied_fields(item, state)
    default_type = (
        state.taxonomy.subitem_type if item.parent_id is not None else state.taxonomy.item_type
    )
    row: dict[str, object] = {
        "ID": state.allocator.allocate(),
        "Type": supplied.get("Type", default_type),
        "Parent": parent,
    }
    for field_name in ("Title", "Status", "Priority", "Owner", "Start", "Due", "Updated"):
        if field_name in supplied:
            row[field_name] = supplied[field_name]
    if "Status" not in row and state.taxonomy.first_status is not None:
        row["Status"] = state.taxonomy.first_status
    if item.created is not None:
        row["Created"] = item.created
    return row


def _update_row(row: dict[str, object], item: MondayItem, state: _MergeState) -> None:
    supplied, cleared = _supplied_fields(item, state)
    changed = False
    for field_name, value in supplied.items():
        if row.get(field_name) != value:
            row[field_name] = value
            changed = True
    for field_name in cleared:
        if field_name in row:
            del row[field_name]
            changed = True
            state.notes.append(
                f"{item.name!r}: {field_name} was cleared in monday; cleared in the workbook"
            )
    if changed:
        state.updated += 1
    else:
        state.unchanged += 1


def _ensure_container(board: BoardData, state: _MergeState) -> str:
    mapped = state.item_map.get("board")
    if mapped is not None and mapped in state.by_id:
        return mapped
    if mapped is not None:
        state.notes.append(
            f"the board container row {mapped} was deleted in the workbook; "
            "recreated because new monday items need a parent"
        )
    container_id = state.allocator.allocate()
    row: dict[str, object] = {
        "ID": container_id,
        "Title": f"monday.com — {board.name}",
        "Type": state.taxonomy.container_type,
    }
    if state.taxonomy.first_status is not None:
        row["Status"] = state.taxonomy.first_status
    state.rows.append(row)
    state.by_id[container_id] = row
    state.item_map["board"] = container_id
    state.added += 1
    return container_id


def _guard_duplicates(
    board: BoardData,
    state: _MergeState,
    *,
    allow_duplicates: bool,
) -> None:
    if allow_duplicates:
        return
    known = set(state.item_map)
    existing_titles = {
        str(row.get("Title", "")).casefold(): str(row.get("ID", "")) for row in state.rows
    }
    suspects = [
        f"{item.name!r} (workbook {existing_titles[item.name.casefold()]})"
        for item in board.items
        if item.identifier not in known and item.name.casefold() in existing_titles
    ]
    if suspects:
        raise MondayImportError(
            _MondayProblem.DUPLICATE_TITLES,
            ", ".join(sorted(suspects)),
            _map_name(board.identifier),
        )


def _parent_for(item: MondayItem, state: _MergeState, container_id: str) -> str:
    if item.parent_id is None:
        return container_id
    mapped_parent = state.item_map.get(item.parent_id)
    if mapped_parent is not None and mapped_parent in state.by_id:
        return mapped_parent
    if mapped_parent is not None:
        state.notes.append(
            f"{item.name!r}: its monday parent maps to {mapped_parent}, which was deleted "
            "in the workbook; attached to the board container instead"
        )
    return container_id


def _merge_items(board: BoardData, state: _MergeState) -> tuple[str, ...]:
    skipped: list[str] = []
    for item in board.items:
        mapped = state.item_map.get(item.identifier)
        if mapped is not None and mapped in state.by_id:
            _update_row(state.by_id[mapped], item, state)
            continue
        if mapped is not None:
            skipped.append(
                f"{item.name!r} ({mapped} was deleted in the workbook; delete its entry "
                "from the identifier map to re-import it)"
            )
            continue
        container_id = _ensure_container(board, state)
        row = _new_row(item, state, _parent_for(item, state, container_id))
        state.rows.append(row)
        state.by_id[str(row["ID"])] = row
        state.item_map[item.identifier] = str(row["ID"])
        state.added += 1
    return tuple(skipped)


def _appended_people(
    snapshot: Snapshot,
    supplied_owners: set[str],
    notes: list[str],
) -> tuple[dict[str, object], ...]:
    people = list(snapshot.tables.get("tblPeople", ()))
    known = {str(row.get("Person", "")).casefold() for row in people}
    for owner in sorted(supplied_owners):
        if not owner or owner.casefold() in known:
            continue
        known.add(owner.casefold())
        people.append({"Person": owner})
        notes.append(f"added {owner!r} to tblPeople; assign Role and Team in Config")
    return tuple(people)


def map_board(
    board: BoardData,
    snapshot: Snapshot,
    map_document: dict[str, object],
    *,
    allow_duplicates: bool,
) -> ImportOutcome:
    """Merge one fetched board into an exported snapshot.

    Returns:
        The merged snapshot, refreshed identifier map and reporting counts.

    """
    notes: list[str] = []
    state = _merge_state(board, snapshot, map_document, notes)
    _guard_duplicates(board, state, allow_duplicates=allow_duplicates)
    skipped = _merge_items(board, state)
    merged = _merged_snapshot(snapshot, state, notes)
    return ImportOutcome(
        snapshot=merged,
        map_document=_map_document(board, state.item_map),
        added=state.added,
        updated=state.updated,
        unchanged=state.unchanged,
        skipped_deleted=skipped,
        notes=tuple(notes),
    )


def _merge_state(
    board: BoardData,
    snapshot: Snapshot,
    map_document: dict[str, object],
    notes: list[str],
) -> _MergeState:
    items_map_field = map_document.get("items")
    item_map: dict[str, str] = (
        {str(key): str(value) for key, value in items_map_field.items()}
        if isinstance(items_map_field, dict)
        else {}
    )
    container_field = map_document.get("container")
    if isinstance(container_field, str) and container_field:
        item_map.setdefault("board", container_field)

    rows = [dict(row) for row in snapshot.tables.get("tblItems", ())]
    by_id = {str(row.get("ID", "")): row for row in rows}
    prefix = str(snapshot.settings.get("cfgItemIDPrefix", "I-"))
    counter = int(str(snapshot.settings.get("cfgNextItemID", 1)))
    return _MergeState(
        rows=rows,
        by_id=by_id,
        item_map=item_map,
        allocator=_IdAllocator(prefix, counter, set(by_id)),
        resolutions=_resolve_boards(board, notes),
        parent_board=board.identifier,
        taxonomy=_taxonomy(snapshot, notes),
        notes=notes,
    )


def _merged_snapshot(snapshot: Snapshot, state: _MergeState, notes: list[str]) -> Snapshot:
    settings = dict(snapshot.settings)
    settings["cfgNextItemID"] = state.allocator.counter
    tables = dict(snapshot.tables)
    tables["tblItems"] = tuple(state.rows)
    tables["tblPeople"] = _appended_people(snapshot, state.supplied_owners, notes)
    return replace(snapshot, settings=settings, tables=tables)


def _map_document(board: BoardData, item_map: dict[str, str]) -> dict[str, object]:
    return {
        "format": MAP_FORMAT,
        "board": board.identifier,
        "board_name": board.name,
        "container": item_map.get("board"),
        "items": {key: value for key, value in sorted(item_map.items()) if key != "board"},
    }


def _map_name(board_identifier: str) -> str:
    return f"board-{board_identifier}.json"


def _pending_name(board_identifier: str) -> str:
    return f"board-{board_identifier}.pending.json"


def read_map(directory: Path, board_id: int) -> dict[str, object]:
    """Read one board's identifier map, tolerating a missing file.

    Returns:
        The map document, or an empty document when none exists yet.

    Raises:
        MondayImportError: If an existing map cannot be parsed.

    """
    path = directory / _map_name(str(board_id))
    pending = directory / _pending_name(str(board_id))
    if pending.exists():
        raise MondayImportError(
            _MondayProblem.MAP_PENDING,
            "the workbook",
            pending,
            path.name,
        )
    if not path.is_file():
        return {}
    try:
        document = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as error:
        raise MondayImportError(_MondayProblem.MAP_MALFORMED, path.name, error) from error
    if not isinstance(document, dict) or document.get("format") != MAP_FORMAT:
        raise MondayImportError(_MondayProblem.MAP_MALFORMED, path.name, "unsupported format")
    return document


def write_map(directory: Path, board_id: int, document: dict[str, object]) -> Path:
    """Atomically persist one board's identifier map.

    Returns:
        The written map path.

    """
    directory.mkdir(parents=True, exist_ok=True)
    target = directory / _map_name(str(board_id))
    atomic_write_json(target, document)
    return target


def _stage_pending_map(directory: Path, board_id: int, document: dict[str, object]) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    pending = directory / _pending_name(str(board_id))
    atomic_write_json(pending, document)
    return pending


MAP_DIR = DIST / "monday"


def monday_command(
    board_id: int,
    workbook: str | Path = DEFAULT_WORKBOOK,
    *,
    dry_run: bool = False,
    allow_duplicates: bool = False,
    token_env: str = AUTH_ENV,
) -> None:
    """Import one monday.com board into the workbook's Items hierarchy.

    Raises:
        MondayImportError: If no API token is available in the environment.

    """
    source = _require_workbook(Path(workbook))
    token = os.environ.get(token_env)
    if not token:
        raise MondayImportError(_MondayProblem.MISSING_CREDENTIALS, token_env)
    if not dry_run:
        require_current_vba()

    result = export_workbook(source)
    board = fetch_board(board_id, token=token)
    outcome = map_board(
        board,
        result.snapshot,
        read_map(MAP_DIR, board_id),
        allow_duplicates=allow_duplicates,
    )
    reconciliation = validate_snapshot(outcome.snapshot)

    _log_export(result)
    for note in outcome.notes:
        LOGGER.info("%s", note)
    for line in reconciliation.lines():
        LOGGER.info("%s", line)
    for skipped in outcome.skipped_deleted:
        LOGGER.info("left deleted: %s", skipped)
    LOGGER.info(
        "%s: %s added, %s updated, %s unchanged",
        board.name,
        outcome.added,
        outcome.updated,
        outcome.unchanged,
    )
    if dry_run:
        LOGGER.info("dry run: %s was not modified", source)
        return

    pre_snapshot_path = write_snapshot(result.snapshot)
    LOGGER.info("pre-import snapshot: %s", pre_snapshot_path)

    # The refreshed map is staged before publication and promoted after it,
    # so an interruption is always detected by read_map instead of silently
    # duplicating or skipping items on the next run.
    pending = _stage_pending_map(MAP_DIR, board_id, outcome.map_document)
    try:
        backup = rebuild_and_publish(source, outcome.snapshot, reconciliation)
    except BaseException as publish_error:
        try:
            pending.unlink()
        except OSError as cleanup_error:
            publish_error.add_note(
                "the staged identifier map could not be removed either: "
                f"{cleanup_error}; delete {pending} by hand"
            )
        raise
    map_path = MAP_DIR / _map_name(str(board_id))
    pending.replace(map_path)

    post_result = export_workbook(source)
    post_snapshot_path = write_snapshot(post_result.snapshot)
    LOGGER.info("post-import snapshot: %s", post_snapshot_path)
    LOGGER.info("identifier map: %s", map_path)
    LOGGER.info("backup: %s", backup)
    LOGGER.info("imported board %s into %s", board_id, source)
