"""Shared writer machinery for formatting, tables, views, and macro actions."""

from dataclasses import dataclass
from datetime import date
from typing import ClassVar

from xlsxwriter.format import Format
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet

from ..core.design import ALIGNMENT, COLORS, MACRO_ACTIONS, ROWS, TYPOGRAPHY
from ..core.formulas import encode_formula
from ..spec.capacity import DATA_ROWS
from ..spec.items import ColumnSpec

FONT = {
    "font_name": TYPOGRAPHY["body_font"],
    "font_size": TYPOGRAPHY["body"],
    "font_color": COLORS["text"],
}


def insert_macro_action(ws: Worksheet, key: str) -> None:
    """Insert one flat DrawingML action shape.

    XlsxWriter writes textboxes with an empty ``macro`` attribute. The package
    post-processor attaches the macro by matching the stable description in
    ``MACRO_ACTIONS``. This avoids Mac Excel's native-grey Form Control face.
    """
    action = MACRO_ACTIONS[key]
    ws.insert_textbox(
        action["cell"],
        action["caption"],
        {
            "description": action["description"],
            "width": action["width"],
            "height": action["height"],
            "x_offset": action["x_offset"],
            "y_offset": action["y_offset"],
            "object_position": 3,
            "fill": {"color": COLORS["brand"]},
            "line": {"color": COLORS["brand_dark"], "width": 1},
            "font": {
                "name": TYPOGRAPHY["body_font"],
                "size": TYPOGRAPHY["body"],
                "bold": True,
                "color": COLORS["header_fg"],
            },
            "align": {"horizontal": "center", "vertical": "middle"},
        },
    )


# Excel for Mac preserves list validation sourced from inline constants or a
# defined name that refers to a plain cell range. Taxonomy names target Config
# ranges; dynamic names target bounded Calc helpers.
DV_NAME = {
    # Taxonomy ranges are returned by write_config and named by the pipeline.
    "=lstStatus": "=dvStatus",
    "=lstTypes": "=dvTypes",
    "=lstPriorities": "=dvPriorities",
    "=lstTeams": "=dvTeams",
    "=lstRaidTypes": "=dvRaidTypes",
    "=lstRaidStatuses": "=dvRaidStatuses",
    "=lstDeliveryHealth": "=dvDeliveryHealth",
    # Dynamic lists use bounded Calc helper ranges.
    "=lstItemIDs": "=dvItemIDs",
    "=lstPeople": "=dvPeople",
}


def dv_source(src: str) -> str:
    """Normalize a validation-list name to its plain-range source.

    Returns:
        The supported plain-range defined name.

    Raises:
        ValueError: If the requested source has no supported mapping.

    """
    if src in DV_NAME:
        return DV_NAME[src]
    if src in set(DV_NAME.values()) | {"=dvScopeLabels"}:
        return src
    msg = f"unsupported data-validation source: {src!r}"
    raise ValueError(msg)


class Formats:
    """Lazily-created, memoised xlsxwriter formats."""

    BASE: ClassVar[dict[str | None, dict[str, object]]] = {
        None: {},
        "text": {"num_format": "@", **ALIGNMENT["text"]},
        "date": {"num_format": "dd mmm yyyy", **ALIGNMENT["date"]},
        "int": {"num_format": "0", **ALIGNMENT["number"]},
        "duein": {
            "num_format": "+0;-0;0",
            "bg_color": COLORS["formula_bg"],
            "font_color": COLORS["formula_fg"],
            **ALIGNMENT["number"],
        },
        "health": {**ALIGNMENT["control"], "bold": True, "bg_color": COLORS["formula_bg"]},
        "calc": {
            "bg_color": COLORS["formula_bg"],
            "font_color": COLORS["formula_fg"],
            **ALIGNMENT["text"],
        },
        "calcbool": {
            "bg_color": COLORS["formula_bg"],
            "font_color": COLORS["formula_fg"],
            **ALIGNMENT["control"],
        },
        "calcint": {
            "num_format": "0",
            "bg_color": COLORS["formula_bg"],
            "font_color": COLORS["formula_fg"],
            **ALIGNMENT["number"],
        },
        "calcdate": {
            "num_format": "dd mmm yyyy",
            "bg_color": COLORS["formula_bg"],
            "font_color": COLORS["formula_fg"],
            **ALIGNMENT["date"],
        },
    }

    def __init__(self, wb: Workbook) -> None:
        """Initialize a workbook-scoped format cache."""
        self.wb = wb
        self._cache: dict[tuple[str | None, tuple[tuple[str, object], ...]], Format] = {}

    def get(self, key: str | None, **extra: object) -> Format:
        """Return the memoized format for one semantic role.

        Returns:
            The matching workbook format.

        Raises:
            KeyError: If the semantic role is unknown.

        """
        if key not in self.BASE:
            msg = f"unknown workbook format role: {key!r}"
            raise KeyError(msg)
        sig = (key, tuple(sorted(extra.items())))
        if sig not in self._cache:
            props = dict(FONT)
            props.update(self.BASE[key])
            props.update(extra)
            self._cache[sig] = self.wb.add_format(props)
        return self._cache[sig]

    # Named workbook components. Sheet writers should prefer these roles to
    # constructing anonymous formats so the design stays coherent.
    def page_title(self) -> Format:
        """Return the page-title format.

        Returns:
            The page-title workbook format.

        """
        return self.get(
            None,
            font_name=TYPOGRAPHY["display_font"],
            font_size=TYPOGRAPHY["page_title"],
            bold=True,
            font_color=COLORS["header_bg"],
            bg_color=COLORS["brand_tint"],
            **ALIGNMENT["text"],
        )

    def page_subtitle(self) -> Format:
        """Return the page-subtitle format.

        Returns:
            The page-subtitle workbook format.

        """
        return self.get(
            None,
            font_size=TYPOGRAPHY["caption"],
            font_color=COLORS["text_secondary"],
            bg_color=COLORS["brand_tint"],
            **ALIGNMENT["text"],
        )

    def meta(self) -> Format:
        """Return the reporting-metadata format.

        Returns:
            The reporting-metadata workbook format.

        """
        return self.get(
            None,
            font_size=TYPOGRAPHY["caption"],
            bold=True,
            **ALIGNMENT["metadata"],
            bg_color=COLORS["brand_tint"],
            font_color=COLORS["text_secondary"],
        )

    def h2(self) -> Format:
        """Return the secondary-heading format.

        Returns:
            The secondary-heading workbook format.

        """
        return self.get(
            None,
            font_size=TYPOGRAPHY["section"],
            bold=True,
            font_color=COLORS["header_bg"],
            bottom=1,
            bottom_color=COLORS["brand"],
            **ALIGNMENT["text"],
        )

    def label(self) -> Format:
        """Return the field-label format.

        Returns:
            The field-label workbook format.

        """
        return self.get(None, bold=True, font_color=COLORS["text_secondary"], **ALIGNMENT["text"])

    def input_cell(self) -> Format:
        """Return the editable-input format.

        Returns:
            The editable-input workbook format.

        """
        return self.get(
            None,
            bg_color=COLORS["input_bg"],
            border=1,
            border_color=COLORS["border_strong"],
            locked=False,
            **ALIGNMENT["text"],
        )

    def info_banner(self) -> Format:
        """Return the informational-banner format.

        Returns:
            The informational-banner workbook format.

        """
        return self.get(
            None,
            font_size=TYPOGRAPHY["caption"],
            font_color=COLORS["text_secondary"],
            bg_color=COLORS["surface_subtle"],
            **ALIGNMENT["text"],
        )

    def table_header(self) -> Format:
        """Return the primary table-header format.

        Returns:
            The primary table-header workbook format.

        """
        return self.data_table_header()

    def data_table_header(self) -> Format:
        """Return the primary operational-table header.

        Returns:
            The Items and RAID header format.

        """
        return self.get(
            None,
            font_size=TYPOGRAPHY["table_header"],
            bold=True,
            bg_color=COLORS["header_bg"],
            font_color=COLORS["header_fg"],
            border=0,
            text_wrap=True,
            **ALIGNMENT["text"],
        )

    def config_table_header(self) -> Format:
        """Return the lower-emphasis editable-list header.

        Returns:
            The Config-table header format.

        """
        return self.get(
            None,
            font_size=TYPOGRAPHY["table_header"],
            bold=True,
            bg_color=COLORS["surface_subtle"],
            font_color=COLORS["header_bg"],
            bottom=1,
            bottom_color=COLORS["border_strong"],
            text_wrap=True,
            **ALIGNMENT["text"],
        )

    def view_table_header(self) -> Format:
        """Return the derived-view table header.

        Returns:
            The derived-view table-header format.

        """
        return self.get(
            None,
            font_size=TYPOGRAPHY["table_header"],
            bold=True,
            bg_color=COLORS["brand_tint"],
            font_color=COLORS["header_bg"],
            bottom=1,
            bottom_color=COLORS["border_strong"],
            text_wrap=True,
            **ALIGNMENT["text"],
        )

    def panel_title(self) -> Format:
        """Return the Overview panel-title format.

        Returns:
            The Overview panel-title workbook format.

        """
        return self.get(
            None,
            font_size=TYPOGRAPHY["section"],
            bold=True,
            bg_color=COLORS["header_bg"],
            font_color=COLORS["header_fg"],
            **ALIGNMENT["text"],
        )

    def panel_count(self) -> Format:
        """Return the Overview panel-count format.

        Returns:
            The Overview panel-count workbook format.

        """
        return self.get(
            None,
            font_size=TYPOGRAPHY["caption"],
            bold=True,
            align="right",
            bg_color=COLORS["header_bg"],
            font_color=COLORS["header_fg"],
            valign="vcenter",
        )

    def view_body(self, **extra: object) -> Format:
        """Return a derived-view body format with optional overrides.

        Returns:
            The derived-view body workbook format.

        """
        props = {
            "bg_color": COLORS["surface"],
            "font_color": COLORS["text"],
            **ALIGNMENT["panel_text"],
        }
        props.update(extra)
        return self.get(None, **props)

    def checkbox(self, bg_color: str | None = None) -> Format:
        """Return a dedicated native-checkbox format.

        Returns:
            A noncached checkbox workbook format.

        """
        # XlsxWriter marks the supplied format as a checkbox format in place,
        # so every checkbox receives a dedicated Format instance.
        props = dict(
            FONT,
            locked=False,
            bg_color=bg_color or COLORS["input_bg"],
            **ALIGNMENT["control"],
            bottom=1,
            bottom_color=COLORS["border"],
        )
        fmt = self.wb.add_format(props)
        fmt.set_checkbox()
        return fmt


def _column_alignment(column: ColumnSpec) -> dict[str, object]:
    """Resolve one table column to a semantic alignment token.

    Returns:
        The matching XlsxWriter alignment properties.

    """
    if column.get("checkbox") or column["fmt"] in {"health", "calcbool"}:
        return ALIGNMENT["control"]
    if column.get("wrap"):
        return ALIGNMENT["narrative"]
    if column["fmt"] in {"date", "calcdate"}:
        return ALIGNMENT["date"]
    if column["fmt"] in {"int", "duein", "calcint"}:
        return ALIGNMENT["number"]
    return ALIGNMENT["text"]


@dataclass(frozen=True)
class TableSpec:
    """Describe one generated operational table."""

    name: str
    columns: list[ColumnSpec]
    examples: list[dict[str, object]]


def _validate_table_spec(table: TableSpec, fmts: Formats, origin: tuple[int, int]) -> None:
    first_row, first_col = origin
    if not table.columns:
        msg = f"{table.name} has no column specifications"
        raise ValueError(msg)
    if first_row < 0 or first_col < 0:
        msg = f"{table.name} has a negative table origin"
        raise ValueError(msg)
    column_names = [column["name"] for column in table.columns]
    if len(column_names) != len(set(column_names)):
        msg = f"{table.name} has duplicate column names"
        raise ValueError(msg)
    allowed_kinds = {"F", "I", "S", "V"}
    for column in table.columns:
        if column["kind"] not in allowed_kinds:
            msg = f"{table.name}[{column['name']}] has unknown kind {column['kind']!r}"
            raise ValueError(msg)
        if (column["kind"] == "F") != bool(column["formula"]):
            msg = f"{table.name}[{column['name']}] formula/kind contract is inconsistent"
            raise ValueError(msg)
        if column["fmt"] not in fmts.BASE:
            msg = f"{table.name}[{column['name']}] uses unknown format {column['fmt']!r}"
            raise ValueError(msg)
    known = set(column_names)
    for index, example in enumerate(table.examples, start=1):
        unknown = sorted(set(example) - known)
        if unknown:
            msg = f"{table.name} example {index} has unknown fields: {unknown}"
            raise ValueError(msg)


def _table_column_specs(fmts: Formats, columns: list[ColumnSpec]) -> list[dict[str, object]]:
    header_fmt = fmts.table_header()
    column_specs: list[dict[str, object]] = []
    for column in columns:
        is_system = column["kind"] in {"F", "S", "V"}
        fmt_extra = {
            "locked": column["kind"] == "F",
            "bg_color": (COLORS["formula_bg"] if is_system else COLORS["input_bg"]),
            "font_color": (COLORS["formula_fg"] if is_system else COLORS["text"]),
            "bottom": 1,
            "bottom_color": COLORS["border"],
            **_column_alignment(column),
        }
        if column.get("wrap"):
            fmt_extra.update(text_wrap=True)
        spec = {
            "header": column["name"],
            "header_format": header_fmt,
            "format": fmts.get(column["fmt"], **fmt_extra),
        }
        if column["formula"]:
            spec["formula"] = encode_formula(column["formula"], variables=column["vars"])
        column_specs.append(spec)
    return column_specs


def _write_table_examples(
    ws: Worksheet,
    fmts: Formats,
    columns: list[ColumnSpec],
    examples: list[dict[str, object]],
    origin: tuple[int, int],
) -> None:
    first_row, first_col = origin
    for index, example in enumerate(examples):
        row = first_row + 1 + index
        for offset, column in enumerate(columns):
            if column["formula"]:
                continue
            value = example.get(column["name"])
            if column.get("checkbox"):
                ws.insert_checkbox(
                    row,
                    first_col + offset,
                    bool(value),
                    fmts.checkbox(bg_color=COLORS["input_bg"]),
                )
                continue
            is_identity = offset == 0
            is_literal_cue = isinstance(value, str) and "EXAMPLE" in value.upper()
            is_system = column["kind"] in {"S", "V"}
            base_bg = COLORS["formula_bg"] if is_system else COLORS["input_bg"]
            base_fg = COLORS["formula_fg"] if is_system else COLORS["text"]
            cell_format = fmts.get(
                column["fmt"],
                bg_color=COLORS["example_bg"] if is_identity else base_bg,
                locked=False,
                **_column_alignment(column),
                bottom=1,
                bottom_color=COLORS["border"],
                font_color=(COLORS["example_fg"] if is_identity or is_literal_cue else base_fg),
                **({"text_wrap": True} if column.get("wrap") else {}),
            )
            if value is None:
                ws.write_blank(row, first_col + offset, None, cell_format)
            elif isinstance(value, date):
                ws.write_datetime(row, first_col + offset, value, cell_format)
            elif isinstance(value, bool):
                ws.write_boolean(row, first_col + offset, value, cell_format)
            else:
                ws.write(row, first_col + offset, value, cell_format)


def _configure_table_columns(
    ws: Worksheet,
    columns: list[ColumnSpec],
    origin: tuple[int, int],
) -> None:
    first_row, first_col = origin
    for offset, column in enumerate(columns):
        ws.set_column(first_col + offset, first_col + offset, column["width"])
        if not column["dv"]:
            continue
        validation = dict(column["dv"])
        source = validation.get("source")
        if isinstance(source, str) and source.startswith("="):
            validation["source"] = dv_source(source)
        elif isinstance(source, str):
            validation["source"] = [value.strip() for value in source.split(",")]
        validation.setdefault("ignore_blank", True)
        validation.setdefault("show_error", True)
        validation.setdefault("error_type", "stop")
        validation.setdefault("error_title", "Check this value")
        validation.setdefault("error_message", "Use a value allowed by this field's rule.")
        if validation.get("validate") == "list":
            validation.setdefault("show_input", True)
            validation.setdefault("input_title", "Choose a value")
            validation.setdefault("input_message", "Use the dropdown list.")
        ws.data_validation(
            first_row + 1,
            first_col + offset,
            first_row + DATA_ROWS,
            first_col + offset,
            validation,
        )


def write_table(
    ws: Worksheet,
    fmts: Formats,
    table: TableSpec,
    *,
    first_row: int = 0,
    first_col: int = 0,
) -> int:
    """Write one validated Excel data table.

    Returns:
        The final row occupied by the table.

    """
    origin = first_row, first_col
    _validate_table_spec(table, fmts, origin)
    last_row = first_row + max(1, len(table.examples))
    last_col = first_col + len(table.columns) - 1

    ws.add_table(
        first_row,
        first_col,
        last_row,
        last_col,
        {
            "name": table.name,
            "style": "Table Style Light 1",
            "columns": _table_column_specs(fmts, table.columns),
        },
    )
    ws.set_row(first_row, ROWS["table_header"])
    _write_table_examples(ws, fmts, table.columns, table.examples, origin)
    _configure_table_columns(ws, table.columns, origin)
    return last_row


def view_chrome(
    ws: Worksheet,
    fmts: Formats,
    title: str,
    filters: list[tuple[str, str, str | None, object | None]],
    first_spill_row: int,
) -> None:
    """Write shared title, metadata, filters, and freeze panes.

    Each filter contains its label, cell, validation source, and default.
    """
    ws.hide_gridlines(2)
    ws.set_row(0, ROWS["page_title"])
    ws.set_row(1, ROWS["toolbar"])
    ws.merge_range(0, 0, 0, 2, title, fmts.page_title())
    # The current reporting date occupies the metadata zone beside the title.
    metadata_format = fmts.meta()
    ws.merge_range(0, 3, 0, 4, "", metadata_format)
    ws.write_formula(
        0,
        3,
        encode_formula('="As of "&TEXT(TODAY(),"d mmm yyyy")'),
        metadata_format,
    )
    for i, (label, cell, dv_src, default) in enumerate(filters):
        ws.write(1, i * 2, label, fmts.label())
        if default is not None:
            if str(default).startswith("="):
                ws.write_formula(cell, encode_formula(default), fmts.input_cell())
            else:
                ws.write(cell, default, fmts.input_cell())
        else:
            ws.write_blank(cell, None, fmts.input_cell())
        if dv_src:
            src = (
                [s.strip() for s in dv_src.split(",")]
                if not dv_src.startswith("=")
                else dv_source(dv_src)
            )
            blank_hint = (
                "Blank = All; otherwise use the dropdown list."
                if label == "Scope"
                else "Use the dropdown list."
            )
            ws.data_validation(
                cell,
                {
                    "validate": "list",
                    "source": src,
                    "ignore_blank": True,
                    "show_input": True,
                    "input_title": "Choose a value",
                    "input_message": blank_hint,
                    "show_error": True,
                    "error_type": "stop",
                    "error_title": "Choose a listed value",
                    "error_message": "Use the dropdown; free text is not valid.",
                },
            )
    ws.freeze_panes(first_spill_row - 1, 0)
