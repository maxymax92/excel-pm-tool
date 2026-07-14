"""Writers for the Items, RAID, and Config sheets."""

from collections.abc import Sequence
from dataclasses import dataclass
from functools import partial

from xlsxwriter.format import Format
from xlsxwriter.utility import xl_col_to_name
from xlsxwriter.worksheet import Worksheet

from ..core.design import ALIGNMENT, COLORS, HIERARCHY, ROWS
from ..core.formulas import encode_formula
from ..spec import config, examples, items
from ..spec.capacity import CONFIG_ROWS, DATA_ROWS
from .common import Formats, TableSpec, insert_macro_action, write_table


def _data_page(ws: Worksheet) -> None:
    """Make wide input tables readable on screen and when printed/PDFed."""
    ws.set_zoom(100)
    ws.set_default_row(ROWS["data_compact"])
    ws.set_landscape()
    ws.set_paper(9)  # A4
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(0)
    ws.set_margins(left=0.25, right=0.25, top=0.35, bottom=0.35)
    ws.set_row(0, ROWS["page_title"])


def _title_rail(
    ws: Worksheet,
    fmts: Formats,
    title: str,
    subtitle: str,
    end_col: int = 11,
) -> None:
    """Consistent title/meta rail for editable and system sheets."""
    ws.merge_range(0, 0, 0, 2, title, fmts.page_title())
    ws.merge_range(0, 3, 0, end_col, "", fmts.page_subtitle())
    if subtitle.startswith("="):
        ws.write_formula(0, 3, encode_formula(subtitle), fmts.page_subtitle())
    else:
        ws.write(0, 3, subtitle, fmts.page_subtitle())


def _cf_rag_letter(ws: Worksheet, fmts: Formats, rng: str) -> None:
    for letter, bg, fg in (
        ("R", COLORS["rag_r_bg"], COLORS["rag_r_fg"]),
        ("A", COLORS["rag_a_bg"], COLORS["rag_a_fg"]),
        ("G", COLORS["rag_g_bg"], COLORS["rag_g_fg"]),
    ):
        ws.conditional_format(
            rng,
            {
                "type": "cell",
                "criteria": "==",
                "value": f'"{letter}"',
                "format": fmts.get(None, bg_color=bg, font_color=fg, bold=True),
            },
        )


# Title emphasis uses a stepped size ramp for levels 1-3 and standard body text
# for levels 4-6. The workbook VBA applies the ramp when a row's Type is set
# and OrganiseItems reapplies it after hierarchy changes.
LEVEL_FONTS = {level: (role["font_size"], role["bold"]) for level, role in HIERARCHY.items()}
SETTING_BOUNDS = {
    "cfgDueSoonDays": (0, 365),
    "cfgBlockedRedDays": (0, 365),
    "cfgStaleDays": (1, 3650),
    "cfgReportDays": (1, 3650),
    "cfgExecutiveStatusMaxLevel": (1, 6),
    "cfgKeyDateMaxLevel": (2, 6),
    "cfgComingUrgentDays": (0, 3650),
    "cfgComingSoonDays": (1, 3650),
    "cfgComingNearDays": (2, 3650),
    "cfgComingHorizonDays": (3, 3650),
    "cfgAlertSevScore": (1, 25),
    "cfgNextItemID": (1, 999999999),
    "cfgNextRaidID": (1, 999999999),
}
CONFIG_BANDS = {
    "statuses": 4,
    "types": 9,
    "priorities": 12,
    "teams": 14,
    "raid_types": 16,
    "raid_statuses": 20,
    "severity": 23,
    "delivery_health": 26,
    "people": 28,
    "guidance": 32,
}


def _write_item_example_outline(ws: Worksheet, fmts: Formats) -> None:
    ws.outline_settings(
        visible=True,
        symbols_below=False,
        symbols_right=True,
        auto_style=False,
    )
    type_level = dict(config.TYPES)
    if len(type_level) != len(config.TYPES):
        message = "shipped Config types contain duplicate names"
        raise ValueError(message)
    title_column = next(
        index for index, column in enumerate(items.ITEMS_COLUMNS) if column["name"] == "Title"
    )
    for index, example in enumerate(examples.ITEMS_EXAMPLES):
        row = 2 + index
        item_type = example.get("Type")
        if item_type not in type_level:
            message = f"example item uses unknown type: {item_type!r}"
            raise ValueError(message)
        level = type_level[item_type]
        if level not in HIERARCHY:
            message = f"example item type {item_type!r} has invalid level {level}"
            raise ValueError(message)
        role = HIERARCHY[level]
        ws.set_row(row, role["row_height"], None, {"level": level - 1})
        font_size, bold = LEVEL_FONTS[level]
        ws.write(
            row,
            title_column,
            example["Title"],
            fmts.get(
                None,
                bg_color=COLORS["input_bg"],
                locked=False,
                font_color=COLORS["text"],
                text_wrap=True,
                font_size=font_size,
                bold=bold,
                indent=role["indent"],
                bottom=1,
                bottom_color=COLORS["border"],
                **ALIGNMENT["narrative"],
            ),
        )


def _collapse_advanced_columns(
    ws: Worksheet,
    columns: list[items.ColumnSpec],
    first_hidden: int,
) -> None:
    column_count = len(columns)
    for column_index in range(first_hidden, column_count):
        ws.set_column(
            column_index,
            column_index,
            columns[column_index]["width"],
            None,
            {"level": 1, "hidden": True, "collapsed": False},
        )
    ws.set_column(column_count, column_count, None, None, {"collapsed": True})


def _write_items_frame(ws: Worksheet, fmts: Formats, *, is_xlsm: bool) -> None:
    ws.hide_gridlines(2)
    _data_page(ws)
    ws.repeat_rows(1)
    _title_rail(
        ws,
        fmts,
        "Items",
        f"=IF(ROWS(tblItems[ID])>{DATA_ROWS},"
        f'"⚠ Supported capacity exceeded: Items has "&ROWS(tblItems[ID])&'
        f'" rows; views and validation support {DATA_ROWS}","")',
        end_col=8,
    )
    if is_xlsm:
        insert_macro_action(ws, "organise")
    write_table(
        ws,
        fmts,
        TableSpec("tblItems", items.ITEMS_COLUMNS, examples.ITEMS_EXAMPLES),
        first_row=1,
    )
    ws.freeze_panes(2, 3)


def _named_column_letter(columns: dict[str, int], name: str) -> str:
    return xl_col_to_name(columns[name])


def write_items(ws: Worksheet, fmts: Formats, *, is_xlsm: bool = False) -> None:
    """Write the editable Items hierarchy and its validation rules."""
    _write_items_frame(ws, fmts, is_xlsm=is_xlsm)
    _write_item_example_outline(ws, fmts)

    col = {c["name"]: i for i, c in enumerate(items.ITEMS_COLUMNS)}
    column_letter = partial(_named_column_letter, col)

    first = 3  # first table data row (1-based Excel row)
    n = DATA_ROWS + 2  # match validation coverage (rows 3..2002)
    invalid = fmts.get(
        None,
        bg_color=COLORS["rag_r_bg"],
        font_color=COLORS["rag_r_fg"],
        border=1,
        border_color=COLORS["rag_r_fg"],
    )
    attention_border = fmts.wb.add_format({"border": 1, "border_color": COLORS["rag_a_fg"]})
    is_active = f"COUNTIFS(dvStatus,${column_letter('Status')}{first},dvStatusActive,TRUE)>0"
    is_done = f"COUNTIFS(dvStatus,${column_letter('Status')}{first},dvStatusDone,TRUE)>0"
    is_cancelled = f"COUNTIFS(dvStatus,${column_letter('Status')}{first},dvStatusCancelled,TRUE)>0"
    # Empty supported rows stop here. Partially entered rows continue through
    # the validation rules, including the missing-ID check.
    item_cf_ranges = f"A{first}:L{n} S{first}:S{n} X{first}:X{n} Z{first}:Z{n} AC{first}:AH{n}"
    ws.conditional_format(
        f"A{first}:L{n}",
        {
            "type": "formula",
            "criteria": f'=AND($A{first}="",COUNTA($B{first}:$L{first})=0)',
            "multi_range": item_cf_ranges,
            "stop_if_true": True,
        },
    )
    # Conditional formatting exposes structural errors introduced by paste.
    ws.conditional_format(
        f"{column_letter('ID')}{first}:{column_letter('ID')}{n}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("ID")}{first}<>"",COUNTIF('
                f"${column_letter('ID')}${first}:${column_letter('ID')}${n},"
                f"${column_letter('ID')}{first})>1)"
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        f"{column_letter('ID')}{first}:{column_letter('ID')}{n}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("ID")}{first}<>"",OR('
                f"LEFT(${column_letter('ID')}{first},LEN(cfgItemIDPrefix))"
                f"<>cfgItemIDPrefix,LEN(${column_letter('ID')}{first})"
                f"<=LEN(cfgItemIDPrefix),NOT(IFERROR(AND("
                f"VALUE(MID(${column_letter('ID')}{first},LEN(cfgItemIDPrefix)+1,99))>=1,"
                f"INT(VALUE(MID(${column_letter('ID')}{first},LEN(cfgItemIDPrefix)+1,99)))="
                f"VALUE(MID(${column_letter('ID')}{first},LEN(cfgItemIDPrefix)+1,99))),"
                "FALSE))))"
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    populated_inputs = [
        f'${column_letter(column["name"])}{first}<>""'
        for column in items.ITEMS_COLUMNS
        if column["kind"] == "I"
    ]
    ws.conditional_format(
        f"{column_letter('ID')}{first}:{column_letter('ID')}{n}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("ID")}{first}="",OR(' + ",".join(populated_inputs) + "))"
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    for required_name in ("Type", "Title", "Status"):
        ws.conditional_format(
            f"{column_letter(required_name)}{first}:{column_letter(required_name)}{n}",
            {
                "type": "formula",
                "criteria": (
                    f'=AND(${column_letter("ID")}{first}<>"",${column_letter(required_name)}{first}="")'
                ),
                "format": invalid,
                "stop_if_true": True,
            },
        )
    for field_name, validation_name in (
        ("Type", "dvTypes"),
        ("Status", "dvStatus"),
        ("Delivery Health", "dvDeliveryHealth"),
        ("Priority", "dvPriorities"),
        ("Owner", "dvPeople"),
    ):
        ws.conditional_format(
            f"{column_letter(field_name)}{first}:{column_letter(field_name)}{n}",
            {
                "type": "formula",
                "criteria": (
                    f'=AND(${column_letter(field_name)}{first}<>"",'
                    f"COUNTIF({validation_name},"
                    f"${column_letter(field_name)}{first})=0)"
                ),
                "format": invalid,
                "stop_if_true": True,
            },
        )
    ws.conditional_format(
        f"{column_letter('Parent')}{first}:{column_letter('Parent')}{n}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("ID")}{first}<>"",'
                f'${column_letter("Parent")}{first}<>"",OR('
                f"${column_letter('Parent')}{first}=${column_letter('ID')}{first},"
                f"COUNTIF(dvItemIDs,${column_letter('Parent')}{first})=0,"
                f"${column_letter('A2')}{first}=${column_letter('ID')}{first},"
                f"${column_letter('A3')}{first}=${column_letter('ID')}{first},"
                f"${column_letter('A4')}{first}=${column_letter('ID')}{first},"
                f"${column_letter('A5')}{first}=${column_letter('ID')}{first},"
                f"AND(${column_letter('ParentLevel')}{first}>0,"
                f"${column_letter('ParentLevel')}{first}>=${column_letter('Level')}{first})))"
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    for date_name in (
        "Start",
        "Due",
        "Created",
        "Updated",
        "InProgressSince",
        "DoneDate",
        "BlockedSince",
        "LatestUpdateOn",
    ):
        ws.conditional_format(
            f"{column_letter(date_name)}{first}:{column_letter(date_name)}{n}",
            {
                "type": "formula",
                "criteria": (
                    f'=AND(${column_letter("ID")}{first}<>"",'
                    f'${column_letter(date_name)}{first}<>"",OR('
                    f"NOT(ISNUMBER(${column_letter(date_name)}{first})),"
                    f"${column_letter(date_name)}{first}<DATE(2020,1,1)))"
                ),
                "format": invalid,
                "stop_if_true": True,
            },
        )
    for date_name in ("Start", "Due"):
        ws.conditional_format(
            f"{column_letter(date_name)}{first}:{column_letter(date_name)}{n}",
            {
                "type": "formula",
                "criteria": (
                    f"=AND(ISNUMBER(${column_letter('Start')}{first}),"
                    f"ISNUMBER(${column_letter('Due')}{first}),"
                    f"${column_letter('Start')}{first}>${column_letter('Due')}{first})"
                ),
                "format": invalid,
                "stop_if_true": True,
            },
        )
    for stamp_name in ("Created", "Updated"):
        ws.conditional_format(
            f"{column_letter(stamp_name)}{first}:{column_letter(stamp_name)}{n}",
            {
                "type": "formula",
                "criteria": (
                    f'=AND(${column_letter("ID")}{first}<>"",${column_letter(stamp_name)}{first}="")'
                ),
                "format": invalid,
                "stop_if_true": True,
            },
        )
    ws.conditional_format(
        f"{column_letter('InProgressSince')}{first}:{column_letter('InProgressSince')}{n}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("ID")}{first}<>"",{is_active},${column_letter("InProgressSince")}{first}="")'
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        f"{column_letter('DoneDate')}{first}:{column_letter('DoneDate')}{n}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("ID")}{first}<>"",OR('
                f"AND({is_done},NOT({is_cancelled}),"
                f'${column_letter("DoneDate")}{first}=""),'
                f'AND({is_cancelled},${column_letter("DoneDate")}{first}<>"")))'
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        f"{column_letter('BlockedSince')}{first}:{column_letter('BlockedSince')}{n}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("ID")}{first}<>"",OR('
                f"AND(${column_letter('Delivery Health')}{first}="
                f"{items.DIRECT_BLOCKED_HEALTH_FORMULA},"
                f'${column_letter("BlockedSince")}{first}=""),'
                f"AND(${column_letter('Delivery Health')}{first}<>"
                f"{items.DIRECT_BLOCKED_HEALTH_FORMULA},"
                f'${column_letter("BlockedSince")}{first}<>"")))'
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        f"{column_letter('LatestUpdateOn')}{first}:{column_letter('LatestUpdateOn')}{n}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("ID")}{first}<>"",OR('
                f'AND(${column_letter("Latest Status")}{first}<>"",'
                f'${column_letter("LatestUpdateOn")}{first}=""),'
                f'AND(${column_letter("Latest Status")}{first}="",'
                f'${column_letter("LatestUpdateOn")}{first}<>"")))'
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        f"{column_letter('BlockedBy')}{first}:{column_letter('BlockedBy')}{n}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("BlockedBy")}{first}<>"",${column_letter("BlockedRefsValid")}{first}<>TRUE)'
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    # Levels 1-3 receive live bold emphasis; OrganiseItems applies the complete
    # font-size hierarchy.
    for visible_name in ("ID", "Title"):
        ws.conditional_format(
            f"{column_letter(visible_name)}{first}:{column_letter(visible_name)}{n}",
            {
                "type": "formula",
                "criteria": (
                    f'=AND(${column_letter("ID")}{first}<>"",${column_letter("Level")}{first}>=1,${column_letter("Level")}{first}<=3)'
                ),
                "format": fmts.get(None, bold=True),
            },
        )
    # Health RAG letters
    _cf_rag_letter(ws, fmts, f"{column_letter('Health')}{first}:{column_letter('Health')}{n}")
    ws.conditional_format(
        f"{column_letter('Health')}{first}:{column_letter('Health')}{n}",
        {
            "type": "cell",
            "criteria": "==",
            "value": '"\u2013"',
            "format": fmts.get(None, font_color=COLORS["text_muted"]),
        },
    )
    # Overdue dates use the numeric DueIn helper for table-safe formatting.
    ws.conditional_format(
        f"{column_letter('Due')}{first}:{column_letter('Due')}{n}",
        {
            "type": "formula",
            "criteria": (
                f"=AND(ISNUMBER(${column_letter('DueIn')}{first}),${column_letter('DueIn')}{first}<0)"
            ),
            "format": fmts.get(None, font_color=COLORS["rag_r_fg"], bold=True),
        },
    )
    # blocked row: light red ID:Title band
    ws.conditional_format(
        f"{column_letter('ID')}{first}:{column_letter('Title')}{n}",
        {
            "type": "formula",
            "criteria": f"=${column_letter('IsBlocked')}{first}=TRUE",
            "format": fmts.get(None, bg_color=COLORS["rag_r_bg"]),
        },
    )
    # Active work whose narrative stamp is missing or stale receives an amber
    # Latest Status cell. LatestUpdateOn changes only when Latest Status does.
    ws.conditional_format(
        f"{column_letter('Latest Status')}{first}:{column_letter('Latest Status')}{n}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("ID")}{first}<>"",'
                f"{is_active},"
                f'OR(${column_letter("LatestUpdateOn")}{first}="",'
                f"TODAY()-${column_letter('LatestUpdateOn')}{first}>=cfgStaleDays))"
            ),
            "format": fmts.get(None, bg_color=COLORS["rag_a_bg"], font_color=COLORS["rag_a_fg"]),
        },
    )
    # Active work without an Owner receives an amber input border.
    ws.conditional_format(
        f"{column_letter('Owner')}{first}:{column_letter('Owner')}{n}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("ID")}{first}<>"",{is_active},${column_letter("Owner")}{first}="")'
            ),
            "format": attention_border,
        },
    )
    ws.conditional_format(
        f"{column_letter('Delivery Health')}{first}:{column_letter('Delivery Health')}{n}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("ID")}{first}<>"",{is_active},'
                f'${column_letter("Delivery Health")}{first}="")'
            ),
            "format": attention_border,
        },
    )
    ws.conditional_format(
        f"{column_letter('Delivery Health')}{first}:{column_letter('Delivery Health')}{n}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("Delivery Health")}{first}<>"",'
                f"MATCH(${column_letter('Delivery Health')}{first},dvDeliveryHealth,0)>=3)"
            ),
            "format": fmts.get(
                None,
                bg_color=COLORS["rag_r_bg"],
                font_color=COLORS["rag_r_fg"],
                bold=True,
            ),
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        f"{column_letter('Delivery Health')}{first}:{column_letter('Delivery Health')}{n}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("Delivery Health")}{first}<>"",'
                f"MATCH(${column_letter('Delivery Health')}{first},dvDeliveryHealth,0)=2)"
            ),
            "format": fmts.get(
                None,
                bg_color=COLORS["rag_a_bg"],
                font_color=COLORS["rag_a_fg"],
                bold=True,
            ),
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        f"{column_letter('Delivery Health')}{first}:{column_letter('Delivery Health')}{n}",
        {
            "type": "formula",
            "criteria": f"=${column_letter('Delivery Health')}{first}=INDEX(dvDeliveryHealth,1)",
            "format": fmts.get(
                None,
                bg_color=COLORS["rag_g_bg"],
                font_color=COLORS["rag_g_fg"],
                bold=True,
            ),
            "stop_if_true": True,
        },
    )
    # Active work with a missing Due receives an amber Due-cell border.
    ws.conditional_format(
        f"{column_letter('Due')}{first}:{column_letter('Due')}{n}",
        {
            "type": "formula",
            "criteria": f'=AND({is_active},${column_letter("EffDue")}{first}="")',
            "format": attention_border,
        },
    )
    # Show only the 11 core input columns; collapse everything else (advanced
    # inputs + all computed/VBA columns) into ONE hidden outline group so the tab
    # reads as a short input form, with a [+] to reveal the full detail on demand.
    _collapse_advanced_columns(ws, items.ITEMS_COLUMNS, items.N_CORE)
    ws.set_tab_color(COLORS["tab_data"])


def write_raid(ws: Worksheet, fmts: Formats) -> None:
    """Write the editable RAID register and its validation rules."""
    ws.hide_gridlines(2)
    _data_page(ws)
    _title_rail(
        ws,
        fmts,
        "RAID",
        f"=IF(ROWS(tblRAID[RaidID])>{DATA_ROWS},"
        f'"⚠ Supported capacity exceeded: RAID has "&ROWS(tblRAID[RaidID])&'
        f'" rows; views and validation support {DATA_ROWS}",'
        '"Risks, assumptions, issues, dependencies and decisions")',
    )
    ws.repeat_rows(1)
    write_table(
        ws,
        fmts,
        TableSpec("tblRAID", items.RAID_COLUMNS, examples.RAID_EXAMPLES),
        first_row=1,
    )
    ws.freeze_panes(2, 2)
    for r in range(2, 2 + len(examples.RAID_EXAMPLES)):
        ws.set_row(r, ROWS["data_wrapped"])
    col = {c["name"]: i for i, c in enumerate(items.RAID_COLUMNS)}

    def column_letter(name: str) -> str:
        return xl_col_to_name(col[name])

    first = 3
    last = DATA_ROWS + 2
    invalid = fmts.get(
        None,
        bg_color=COLORS["rag_r_bg"],
        font_color=COLORS["rag_r_fg"],
        border=1,
        border_color=COLORS["rag_r_fg"],
    )
    attention_border = fmts.wb.add_format({"border": 1, "border_color": COLORS["rag_a_fg"]})
    not_closed = f"COUNTIFS(dvRaidStatuses,${column_letter('Status')}{first},dvRaidClosed,TRUE)=0"
    is_closed = f"COUNTIFS(dvRaidStatuses,${column_letter('Status')}{first},dvRaidClosed,TRUE)>0"
    is_alert = f"COUNTIFS(dvRaidTypes,${column_letter('Type')}{first},dvRaidAlert,TRUE)>0"
    is_decision = f"COUNTIFS(dvRaidTypes,${column_letter('Type')}{first},dvRaidDecision,TRUE)>0"
    # Severity is calculated, so the blank-row test uses only editable fields.
    # Partially entered rows still run every validation rule.
    raid_cf_ranges = f"A{first}:C{last} E{first}:Q{last}"
    ws.conditional_format(
        f"A{first}:C{last}",
        {
            "type": "formula",
            "criteria": (f'=AND($A{first}="",COUNTA($B{first}:$I{first},$K{first}:$L{first})=0)'),
            "multi_range": raid_cf_ranges,
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        f"{column_letter('RaidID')}{first}:{column_letter('RaidID')}{last}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("RaidID")}{first}<>"",COUNTIF('
                f"${column_letter('RaidID')}${first}:${column_letter('RaidID')}${last},"
                f"${column_letter('RaidID')}{first})>1)"
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        f"{column_letter('RaidID')}{first}:{column_letter('RaidID')}{last}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("RaidID")}{first}<>"",OR('
                f"LEFT(${column_letter('RaidID')}{first},LEN(cfgRaidIDPrefix))"
                f"<>cfgRaidIDPrefix,LEN(${column_letter('RaidID')}{first})"
                f"<=LEN(cfgRaidIDPrefix),NOT(IFERROR(AND("
                f"VALUE(MID(${column_letter('RaidID')}{first},LEN(cfgRaidIDPrefix)+1,99))>=1,"
                f"INT(VALUE(MID(${column_letter('RaidID')}{first},LEN(cfgRaidIDPrefix)+1,99)))="
                f"VALUE(MID(${column_letter('RaidID')}{first},LEN(cfgRaidIDPrefix)+1,99))),"
                "FALSE))))"
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    populated_inputs = [
        f'${column_letter(column["name"])}{first}<>""'
        for column in items.RAID_COLUMNS
        if column["kind"] == "I"
    ]
    ws.conditional_format(
        f"{column_letter('RaidID')}{first}:{column_letter('RaidID')}{last}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("RaidID")}{first}="",OR(' + ",".join(populated_inputs) + "))"
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    for required_name in ("Type", "Title", "Status"):
        ws.conditional_format(
            f"{column_letter(required_name)}{first}:{column_letter(required_name)}{last}",
            {
                "type": "formula",
                "criteria": (
                    f'=AND(${column_letter("RaidID")}{first}<>"",${column_letter(required_name)}{first}="")'
                ),
                "format": invalid,
                "stop_if_true": True,
            },
        )
    for field_name, validation_name in (
        ("Type", "dvRaidTypes"),
        ("Status", "dvRaidStatuses"),
        ("Owner", "dvPeople"),
    ):
        ws.conditional_format(
            f"{column_letter(field_name)}{first}:{column_letter(field_name)}{last}",
            {
                "type": "formula",
                "criteria": (
                    f'=AND(${column_letter(field_name)}{first}<>"",'
                    f"COUNTIF({validation_name},"
                    f"${column_letter(field_name)}{first})=0)"
                ),
                "format": invalid,
                "stop_if_true": True,
            },
        )
    ws.conditional_format(
        f"{column_letter('RelatedID')}{first}:{column_letter('RelatedID')}{last}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("RelatedID")}{first}<>"",COUNTIF(dvItemIDs,${column_letter("RelatedID")}{first})=0)'
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    for score_name in ("Prob", "Impact"):
        cell = f"${column_letter(score_name)}{first}"
        ws.conditional_format(
            f"{column_letter(score_name)}{first}:{column_letter(score_name)}{last}",
            {
                "type": "formula",
                "criteria": (
                    f'=AND(${column_letter("RaidID")}{first}<>"",OR('
                    f'AND({is_alert},{cell}=""),'
                    f'AND({cell}<>"",OR(NOT(ISNUMBER({cell})),'
                    f"{cell}<1,{cell}>5,INT({cell})<>{cell}))))"
                ),
                "format": invalid,
                "stop_if_true": True,
            },
        )
    ws.conditional_format(
        f"{column_letter('Score')}{first}:{column_letter('Score')}{last}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("RaidID")}{first}<>"",OR('
                f"AND(ISNUMBER(${column_letter('Prob')}{first}),"
                f"ISNUMBER(${column_letter('Impact')}{first}),"
                f"${column_letter('Score')}{first}<>${column_letter('Prob')}{first}*"
                f"${column_letter('Impact')}{first}),AND(OR("
                f'${column_letter("Prob")}{first}="",${column_letter("Impact")}{first}=""),'
                f'${column_letter("Score")}{first}<>"")))'
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        f"{column_letter('Severity')}{first}:{column_letter('Severity')}{last}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("RaidID")}{first}<>"",'
                f'${column_letter("Score")}{first}<>"",OR('
                f"ISERROR(${column_letter('Severity')}{first}),"
                f'${column_letter("Severity")}{first}="",IFERROR(COUNTIF('
                f"dvSeverity,${column_letter('Severity')}{first})=0,TRUE)))"
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    for date_name in ("NextReview", "Raised", "Closed", "Updated"):
        ws.conditional_format(
            f"{column_letter(date_name)}{first}:{column_letter(date_name)}{last}",
            {
                "type": "formula",
                "criteria": (
                    f'=AND(${column_letter("RaidID")}{first}<>"",'
                    f'${column_letter(date_name)}{first}<>"",OR('
                    f"NOT(ISNUMBER(${column_letter(date_name)}{first})),"
                    f"${column_letter(date_name)}{first}<DATE(2020,1,1)))"
                ),
                "format": invalid,
                "stop_if_true": True,
            },
        )
    for stamp_name in ("Raised", "Updated"):
        ws.conditional_format(
            f"{column_letter(stamp_name)}{first}:{column_letter(stamp_name)}{last}",
            {
                "type": "formula",
                "criteria": (
                    f'=AND(${column_letter("RaidID")}{first}<>"",${column_letter(stamp_name)}{first}="")'
                ),
                "format": invalid,
                "stop_if_true": True,
            },
        )
    ws.conditional_format(
        f"{column_letter('Closed')}{first}:{column_letter('Closed')}{last}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("RaidID")}{first}<>"",OR('
                f'AND({is_closed},${column_letter("Closed")}{first}=""),'
                f'AND({not_closed},${column_letter("Closed")}{first}<>"")))'
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        f"{column_letter('NextReview')}{first}:{column_letter('NextReview')}{last}",
        {
            "type": "formula",
            "criteria": (
                f"=AND(ISNUMBER(${column_letter('Raised')}{first}),"
                f"ISNUMBER(${column_letter('NextReview')}{first}),"
                f"${column_letter('NextReview')}{first}<${column_letter('Raised')}{first})"
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    # Severity colours by POSITION in tblSeverity (top two CONFIG_BANDS red, bottom
    # green, the rest amber) — renaming CONFIG_BANDS in Config keeps colours truthful.
    sev = f"{column_letter('Severity')}{first}:{column_letter('Severity')}{last}"
    sc = f"${column_letter('Severity')}{first}"
    ws.conditional_format(
        sev,
        {
            "type": "formula",
            "criteria": (
                f'=AND(COUNTA(dvSeverity)>=2,{sc}<>"",'
                f"COUNTIF(dvSeverity,{sc})>0,MATCH({sc},dvSeverity,0)"
                f">=MAX(2,COUNTA(dvSeverity)-1))"
            ),
            "format": fmts.get(None, bg_color=COLORS["rag_r_bg"], font_color=COLORS["rag_r_fg"]),
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        sev,
        {
            "type": "formula",
            "criteria": f'=AND({sc}<>"",{sc}=INDEX(dvSeverity,1))',
            "format": fmts.get(None, bg_color=COLORS["rag_g_bg"], font_color=COLORS["rag_g_fg"]),
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        sev,
        {
            "type": "formula",
            "criteria": (
                f'=AND({sc}<>"",COUNTIF(dvSeverity,{sc})>0,'
                f"{sc}<>INDEX(dvSeverity,1),NOT(AND("
                f"COUNTA(dvSeverity)>=2,MATCH({sc},dvSeverity,0)"
                f">=MAX(2,COUNTA(dvSeverity)-1))))"
            ),
            "format": fmts.get(None, bg_color=COLORS["rag_a_bg"], font_color=COLORS["rag_a_fg"]),
            "stop_if_true": True,
        },
    )
    # Open RAID status is defined by the Config IsClosed role.
    ws.conditional_format(
        f"{column_letter('NextReview')}{first}:{column_letter('NextReview')}{last}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("NextReview")}{first}<>"",'
                f"${column_letter('NextReview')}{first}<TODAY(),{not_closed})"
            ),
            "format": fmts.get(
                None, bg_color=COLORS["rag_r_bg"], font_color=COLORS["rag_r_fg"], bold=True
            ),
        },
    )
    # Open alert/decision records must stay reviewable and owned. These are
    # warning cues, while structural omissions above remain red invalid data.
    ws.conditional_format(
        f"{column_letter('NextReview')}{first}:{column_letter('NextReview')}{last}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("RaidID")}{first}<>"",{not_closed},'
                f"OR({is_alert},{is_decision}),"
                f'${column_letter("NextReview")}{first}="")'
            ),
            "format": fmts.get(
                None, bg_color=COLORS["rag_a_bg"], font_color=COLORS["rag_a_fg"], bold=True
            ),
        },
    )
    ws.conditional_format(
        f"{column_letter('Owner')}{first}:{column_letter('Owner')}{last}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("RaidID")}{first}<>"",{not_closed},'
                f"OR({is_alert},{is_decision}),"
                f'${column_letter("Owner")}{first}="")'
            ),
            "format": attention_border,
        },
    )
    ws.conditional_format(
        f"{column_letter('Response')}{first}:{column_letter('Response')}{last}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("RaidID")}{first}<>"",{not_closed},'
                f'{is_alert},${column_letter("Score")}{first}<>"",'
                f"${column_letter('Score')}{first}>=cfgAlertSevScore,"
                f'${column_letter("Response")}{first}="")'
            ),
            "format": fmts.get(None, bg_color=COLORS["rag_a_bg"], font_color=COLORS["rag_a_fg"]),
        },
    )
    ws.conditional_format(
        f"{column_letter('RelatedID')}{first}:{column_letter('RelatedID')}{last}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${column_letter("Title")}{first}<>"",{not_closed},${column_letter("RelatedID")}{first}="")'
            ),
            "format": fmts.get(
                None,
                bg_color=COLORS["rag_r_bg"],
                font_color=COLORS["rag_r_fg"],
                border=1,
                border_color=COLORS["rag_r_fg"],
            ),
        },
    )
    for name in ("Title", "Detail", "Response"):
        j = col[name]
        # Width only: the table column format already supplies wrapping and
        # row treatment. A column-default border/fill would continue through
        # unused rows and create a lined empty grid.
        ws.set_column(j, j, items.RAID_COLUMNS[j]["width"])
    _collapse_advanced_columns(ws, items.RAID_COLUMNS, items.N_RAID_CORE)
    ws.set_tab_color(COLORS["tab_data"])


@dataclass(frozen=True)
class _ConfigListSpec:
    name: str
    column: int
    headers: tuple[str, ...]
    rows: Sequence[object]
    widths: tuple[int, ...]
    validation_names: dict[int, str]


@dataclass(frozen=True)
class _ConfigListContext:
    worksheet: Worksheet
    formats: Formats
    base_row: int
    unlocked: Format
    unlocked_integer: Format
    header: Format
    ranges: dict[str, str]


def _normalize_config_rows(
    context: _ConfigListContext,
    spec: _ConfigListSpec,
) -> list[tuple[object, ...]]:
    if not spec.headers:
        message = f"{spec.name} has no columns"
        raise ValueError(message)
    if len(spec.headers) != len(spec.widths):
        message = f"{spec.name} columns and widths have different lengths"
        raise ValueError(message)
    if len(spec.rows) > CONFIG_ROWS:
        message = f"{spec.name} has {len(spec.rows)} rows; capacity is {CONFIG_ROWS}"
        raise ValueError(message)
    normalized: list[tuple[object, ...]] = []
    for index, row in enumerate(spec.rows, start=1):
        values = row if isinstance(row, tuple) else (row,)
        if len(values) != len(spec.headers):
            message = (
                f"{spec.name} row {index} has {len(values)} values; expected {len(spec.headers)}"
            )
            raise ValueError(message)
        normalized.append(values)
    for offset, validation_name in spec.validation_names.items():
        if not 0 <= offset < len(spec.headers):
            message = f"{spec.name} has invalid DV column offset {offset}"
            raise ValueError(message)
        if validation_name in context.ranges:
            message = f"duplicate Config DV name: {validation_name}"
            raise ValueError(message)
    return normalized


def _write_config_list(context: _ConfigListContext, spec: _ConfigListSpec) -> None:
    normalized = _normalize_config_rows(context, spec)
    worksheet = context.worksheet
    body_rows = max(1, len(normalized))
    worksheet.add_table(
        context.base_row,
        spec.column,
        context.base_row + body_rows,
        spec.column + len(spec.headers) - 1,
        {
            "name": spec.name,
            "style": "Table Style Light 1",
            "columns": [
                {"header": header, "header_format": context.header} for header in spec.headers
            ],
        },
    )
    if not normalized:
        for offset in range(len(spec.headers)):
            worksheet.write_blank(
                context.base_row + 1,
                spec.column + offset,
                None,
                context.unlocked,
            )
    for index, values in enumerate(normalized):
        row = context.base_row + 1 + index
        worksheet.set_row(row, ROWS["data_compact"])
        for offset, value in enumerate(values):
            if isinstance(value, bool):
                worksheet.insert_checkbox(
                    row,
                    spec.column + offset,
                    value,
                    context.formats.checkbox(COLORS["input_bg"]),
                )
            elif isinstance(value, (int, float)):
                worksheet.write(row, spec.column + offset, value, context.unlocked_integer)
            else:
                worksheet.write(row, spec.column + offset, value, context.unlocked)
    for offset, width in enumerate(spec.widths):
        worksheet.set_column(spec.column + offset, spec.column + offset, width)
    for offset, validation_name in spec.validation_names.items():
        column_letter = xl_col_to_name(spec.column + offset)
        context.ranges[validation_name] = (
            f"Config!${column_letter}${context.base_row + 2}:"
            f"${column_letter}${context.base_row + 1 + CONFIG_ROWS}"
        )


def _write_config_settings(
    ws: Worksheet,
    fmts: Formats,
    base: int,
) -> dict[str, int]:
    # Settings band A:C. Named cells are created from the values in column B.
    settings_header = fmts.config_table_header()
    ws.write(base, 0, "Setting", settings_header)
    ws.write(base, 1, "Value", settings_header)
    ws.write(base, 2, "What it does", settings_header)
    ws.set_column(0, 0, 26)
    ws.set_column(1, 1, 12)
    ws.set_column(2, 2, 42)
    setting_rows = {}
    for i, (name, val, desc) in enumerate(config.SETTINGS):
        r = base + 1 + i
        setting_rows[name] = r
        ws.set_row(r, ROWS["data_compact"])
        ws.write(r, 0, name.replace("cfg", ""), fmts.label())
        fmt = (
            fmts.get(
                "int",
                bg_color=COLORS["input_bg"],
                border=1,
                border_color=COLORS["border_strong"],
                locked=False,
                **ALIGNMENT["number"],
            )
            if isinstance(val, (int, float)) and not isinstance(val, bool)
            else fmts.input_cell()
        )
        if val is None:
            ws.write_blank(r, 1, None, fmt)
        else:
            ws.write(r, 1, val, fmt)
        ws.write(
            r,
            2,
            desc,
            fmts.get(
                None,
                font_color=COLORS["text_muted"],
                text_wrap=True,
                bottom=1,
                bottom_color=COLORS["border"],
                **ALIGNMENT["text"],
            ),
        )
        if name in SETTING_BOUNDS:
            minimum, maximum = SETTING_BOUNDS[name]
            ws.data_validation(
                r,
                1,
                r,
                1,
                {
                    "validate": "integer",
                    "criteria": "between",
                    "minimum": minimum,
                    "maximum": maximum,
                    "ignore_blank": False,
                    "show_input": True,
                    "input_title": "Allowed value",
                    "input_message": f"Enter a whole number from {minimum} to {maximum}.",
                    "show_error": True,
                    "error_type": "stop",
                    "error_title": "Value outside the allowed range",
                    "error_message": f"Use a whole number from {minimum} to {maximum}.",
                },
            )
        elif name in {"cfgItemIDPrefix", "cfgRaidIDPrefix"}:
            cell = f"B{r + 1}"
            ws.data_validation(
                r,
                1,
                r,
                1,
                {
                    "validate": "custom",
                    "value": (
                        f"=AND(LEN({cell})>=1,LEN({cell})<=8,"
                        f'TRIM({cell})={cell},ISERROR(SEARCH(",",{cell})),'
                        f"ISERROR(SEARCH(CHAR(9),{cell})),"
                        f"ISERROR(SEARCH(CHAR(10),{cell})),"
                        f"ISERROR(SEARCH(CHAR(13),{cell})))"
                    ),
                    "ignore_blank": False,
                    "show_input": True,
                    "input_title": "ID prefix",
                    "input_message": (
                        "Use 1-8 trimmed characters with no comma, "
                        "tab or line break, for example I-."
                    ),
                    "show_error": True,
                    "error_type": "stop",
                    "error_title": "Use a valid prefix",
                    "error_message": (
                        "Use 1-8 trimmed characters with no comma, tab or line break."
                    ),
                },
            )

    # Conditional formatting mirrors Settings constraints for pasted values.
    settings_invalid = fmts.get(
        None,
        bg_color=COLORS["rag_r_bg"],
        font_color=COLORS["rag_r_fg"],
        border=1,
        border_color=COLORS["rag_r_fg"],
    )
    for name, (minimum, maximum) in SETTING_BOUNDS.items():
        r = setting_rows[name]
        cell = f"B{r + 1}"
        ws.conditional_format(
            r,
            1,
            r,
            1,
            {
                "type": "formula",
                "criteria": (
                    f"=OR(NOT(ISNUMBER({cell})),INT({cell})<>{cell},"
                    f"{cell}<{minimum},{cell}>{maximum})"
                ),
                "format": settings_invalid,
                "stop_if_true": True,
            },
        )
    for name in ("cfgItemIDPrefix", "cfgRaidIDPrefix"):
        r = setting_rows[name]
        cell = f"B{r + 1}"
        ws.conditional_format(
            r,
            1,
            r,
            1,
            {
                "type": "formula",
                "criteria": (
                    f"=OR(LEN({cell})<1,LEN({cell})>8,"
                    f'TRIM({cell})<>{cell},NOT(ISERROR(SEARCH(",",{cell}))),'
                    f"NOT(ISERROR(SEARCH(CHAR(9),{cell}))),"
                    f"NOT(ISERROR(SEARCH(CHAR(10),{cell}))),"
                    f"NOT(ISERROR(SEARCH(CHAR(13),{cell}))))"
                ),
                "format": settings_invalid,
                "stop_if_true": True,
            },
        )

    # The four Coming Up thresholds form one strictly increasing scale. A
    # relationship error highlights the complete scale.
    urgent_row = setting_rows["cfgComingUrgentDays"]
    horizon_row = setting_rows["cfgComingHorizonDays"]
    ws.conditional_format(
        urgent_row,
        1,
        horizon_row,
        1,
        {
            "type": "formula",
            "criteria": (
                "=OR(cfgComingUrgentDays>=cfgComingSoonDays,"
                "cfgComingSoonDays>=cfgComingNearDays,"
                "cfgComingNearDays>=cfgComingHorizonDays)"
            ),
            "format": settings_invalid,
            "stop_if_true": True,
        },
    )

    return setting_rows


def _write_config_lists(
    ws: Worksheet,
    fmts: Formats,
    base: int,
) -> dict[str, str]:
    # List tables run in parallel with Settings. Every list is editable and
    # consumer (dropdowns, level logic, status roles, RAID roles, severity
    # CONFIG_BANDS) reads it live from the Config taxonomy and role tables.
    unl = fmts.get(
        None,
        locked=False,
        bg_color=COLORS["input_bg"],
        bottom=1,
        bottom_color=COLORS["border"],
        **ALIGNMENT["text"],
    )
    unl_int = fmts.get(
        "int",
        locked=False,
        bg_color=COLORS["input_bg"],
        bottom=1,
        bottom_color=COLORS["border"],
        **ALIGNMENT["number"],
    )
    list_ranges = {}  # dv-name -> "Config!$A$r1:$A$r2"
    header_fmt = fmts.config_table_header()
    list_context = _ConfigListContext(
        ws,
        fmts,
        base,
        unl,
        unl_int,
        header_fmt,
        list_ranges,
    )

    # A:C Settings | D gutter | E:H Statuses | I gutter | J:K Types | ...
    # Each table band has independent widths and one narrow gutter.
    for gutter in (3, 8, 11, 13, 15, 19, 22, 25, 27, 31):
        ws.set_column(gutter, gutter, 2)

    list_specs = (
        _ConfigListSpec(
            "tblStatuses",
            CONFIG_BANDS["statuses"],
            ("Status", "IsActive", "IsDone", "IsCancelled"),
            config.STATUSES,
            (16, 12, 11, 16),
            {0: "dvStatus", 1: "dvStatusActive", 2: "dvStatusDone", 3: "dvStatusCancelled"},
        ),
        _ConfigListSpec(
            "tblTypes",
            CONFIG_BANDS["types"],
            ("Type", "Level"),
            config.TYPES,
            (15, 8),
            {0: "dvTypes", 1: "dvTypeLevels"},
        ),
        _ConfigListSpec(
            "tblPriorities",
            CONFIG_BANDS["priorities"],
            ("Priority",),
            [(priority,) for priority in config.PRIORITIES],
            (11,),
            {0: "dvPriorities"},
        ),
        _ConfigListSpec(
            "tblTeams",
            CONFIG_BANDS["teams"],
            ("Team",),
            [(team,) for team in config.TEAMS],
            (12,),
            {0: "dvTeams"},
        ),
        _ConfigListSpec(
            "tblRaidTypes",
            CONFIG_BANDS["raid_types"],
            ("RaidType", "IsAlert", "IsDecision"),
            config.RAID_TYPES,
            (15, 11, 13),
            {0: "dvRaidTypes", 1: "dvRaidAlert", 2: "dvRaidDecision"},
        ),
        _ConfigListSpec(
            "tblRaidStatuses",
            CONFIG_BANDS["raid_statuses"],
            ("RaidStatus", "IsClosed"),
            config.RAID_STATUSES,
            (16, 11),
            {0: "dvRaidStatuses", 1: "dvRaidClosed"},
        ),
        _ConfigListSpec(
            "tblSeverity",
            CONFIG_BANDS["severity"],
            ("Severity", "MinScore"),
            config.SEVERITY,
            (12, 12),
            {0: "dvSeverity", 1: "dvSeverityMinScore"},
        ),
        _ConfigListSpec(
            "tblDeliveryHealth",
            CONFIG_BANDS["delivery_health"],
            ("Delivery Health",),
            [(value,) for value in config.DELIVERY_HEALTH],
            (19,),
            {0: "dvDeliveryHealth"},
        ),
        _ConfigListSpec(
            "tblPeople",
            CONFIG_BANDS["people"],
            ("Person", "Role", "Team"),
            [
                (person.get("Person", ""), person.get("Role", ""), person.get("Team", ""))
                for person in examples.PEOPLE_EXAMPLES
            ],
            (14, 17, 13),
            {},
        ),
    )
    for list_spec in list_specs:
        _write_config_list(list_context, list_spec)

    return list_ranges


def _write_config_entry_validations(ws: Worksheet, base: int) -> None:
    # Config is executable data. Validation governs entry and red conditional
    # rules expose invalid pasted hierarchy and severity values.
    dv_stop = {"ignore_blank": True, "show_input": True, "show_error": True, "error_type": "stop"}
    type_level_col = CONFIG_BANDS["types"] + 1
    severity_score_col = CONFIG_BANDS["severity"] + 1
    people_team_col = CONFIG_BANDS["people"] + 2
    ws.data_validation(
        base + 1,
        type_level_col,
        base + CONFIG_ROWS,
        type_level_col,
        {
            **dv_stop,
            "validate": "integer",
            "criteria": "between",
            "minimum": 1,
            "maximum": 6,
            "input_title": "Hierarchy level",
            "input_message": "Choose a whole-number level from 1 to 6.",
            "error_title": "Level must be 1-6",
            "error_message": "Enter a whole number from 1 to 6.",
        },
    )
    ws.data_validation(
        base + 1,
        severity_score_col,
        base + CONFIG_ROWS,
        severity_score_col,
        {
            **dv_stop,
            "validate": "integer",
            "criteria": "between",
            "minimum": 1,
            "maximum": 25,
            "input_title": "Minimum score",
            "input_message": (
                "Score = Probability \u00d7 Impact (1-25). "
                "Bands must increase down the table from 1."
            ),
            "error_title": "Score must be 1-25",
            "error_message": "Enter a whole number from 1 to 25.",
        },
    )
    ws.data_validation(
        base + 1,
        people_team_col,
        base + CONFIG_ROWS,
        people_team_col,
        {
            "validate": "list",
            "source": "=dvTeams",
            **dv_stop,
            "input_title": "Team",
            "input_message": "Choose a team from Config.",
            "error_title": "Choose a listed team",
            "error_message": "Use the dropdown list.",
        },
    )


def _write_config_band_rules(ws: Worksheet, fmts: Formats, base: int) -> None:
    type_level_col = CONFIG_BANDS["types"] + 1
    severity_score_col = CONFIG_BANDS["severity"] + 1
    invalid = fmts.get(
        None,
        bg_color=COLORS["rag_r_bg"],
        font_color=COLORS["rag_r_fg"],
        border=1,
        border_color=COLORS["rag_r_fg"],
    )
    start_xl = base + 2
    config_end_xl = base + 1 + CONFIG_ROWS
    type_id_col = xl_col_to_name(CONFIG_BANDS["types"])
    type_lvl_col = xl_col_to_name(type_level_col)
    ws.conditional_format(
        f"{type_lvl_col}{start_xl}:{type_lvl_col}{config_end_xl}",
        {
            "type": "formula",
            "criteria": (
                f'=OR(AND(${type_id_col}{start_xl}="",'
                f'${type_lvl_col}{start_xl}<>""),AND('
                f'${type_id_col}{start_xl}<>"",OR('
                f"NOT(ISNUMBER(${type_lvl_col}{start_xl})),"
                f"INT(${type_lvl_col}{start_xl})"
                f"<>${type_lvl_col}{start_xl},"
                f"${type_lvl_col}{start_xl}<1,"
                f"${type_lvl_col}{start_xl}>6)))"
            ),
            "format": invalid,
        },
    )
    sev_id_col = xl_col_to_name(CONFIG_BANDS["severity"])
    sev_score_col = xl_col_to_name(severity_score_col)
    ws.conditional_format(
        f"{sev_score_col}{start_xl}:{sev_score_col}{config_end_xl}",
        {
            "type": "formula",
            "criteria": (
                f'=OR(AND(${sev_id_col}{start_xl}="",'
                f'${sev_score_col}{start_xl}<>""),AND('
                f'${sev_id_col}{start_xl}<>"",OR('
                f'${sev_score_col}{start_xl}="",'
                f"NOT(ISNUMBER(${sev_score_col}{start_xl})),"
                f"INT(${sev_score_col}{start_xl})"
                f"<>${sev_score_col}{start_xl},"
                f"${sev_score_col}{start_xl}<1,"
                f"${sev_score_col}{start_xl}>25,"
                f"AND(ROW()={start_xl},${sev_score_col}{start_xl}<>1),"
                f"AND(ROW()>{start_xl},OR("
                f'${sev_id_col}{start_xl - 1}="",'
                f"${sev_score_col}{start_xl}<="
                f"${sev_score_col}{start_xl - 1})))))"
            ),
            "format": invalid,
        },
    )
    status_name_col = xl_col_to_name(CONFIG_BANDS["statuses"])
    status_active_col = xl_col_to_name(CONFIG_BANDS["statuses"] + 1)
    status_done_col = xl_col_to_name(CONFIG_BANDS["statuses"] + 2)
    status_cancel_col = xl_col_to_name(CONFIG_BANDS["statuses"] + 3)
    ws.conditional_format(
        f"{status_active_col}{start_xl}:{status_cancel_col}{config_end_xl}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${status_name_col}{start_xl}<>"",OR('
                f"AND(${status_active_col}{start_xl}=TRUE,"
                f"${status_done_col}{start_xl}=TRUE),"
                f"AND(${status_cancel_col}{start_xl}=TRUE,"
                f"${status_done_col}{start_xl}<>TRUE)))"
            ),
            "format": invalid,
        },
    )


def _write_config_role_rules(ws: Worksheet, fmts: Formats, base: int) -> None:
    invalid = fmts.get(
        None,
        bg_color=COLORS["rag_r_bg"],
        font_color=COLORS["rag_r_fg"],
        border=1,
        border_color=COLORS["rag_r_fg"],
    )
    start_xl = base + 2
    config_end_xl = base + 1 + CONFIG_ROWS
    # Every role flag is a true Boolean. Native checkboxes enforce this during
    # normal editing; these rules catch pasted numbers/text and orphan TRUE
    # flags whose identity label is blank.
    boolean_groups = (
        (
            CONFIG_BANDS["statuses"],
            (
                CONFIG_BANDS["statuses"] + 1,
                CONFIG_BANDS["statuses"] + 2,
                CONFIG_BANDS["statuses"] + 3,
            ),
        ),
        (
            CONFIG_BANDS["raid_types"],
            (CONFIG_BANDS["raid_types"] + 1, CONFIG_BANDS["raid_types"] + 2),
        ),
        (CONFIG_BANDS["raid_statuses"], (CONFIG_BANDS["raid_statuses"] + 1,)),
    )
    for identity_col, flag_cols in boolean_groups:
        identity = xl_col_to_name(identity_col)
        for flag_col in flag_cols:
            flag = xl_col_to_name(flag_col)
            ws.conditional_format(
                f"{flag}{start_xl}:{flag}{config_end_xl}",
                {
                    "type": "formula",
                    "criteria": (
                        f'=OR(AND(${identity}{start_xl}="",'
                        f"${flag}{start_xl}=TRUE),AND("
                        f'${identity}{start_xl}<>"",'
                        f"NOT(ISLOGICAL(${flag}{start_xl}))))"
                    ),
                    "format": invalid,
                    "stop_if_true": True,
                },
            )

    # People rows keep Role and Team paired, and Team values resolve through
    # tblTeams after direct entry or paste.
    person_col = xl_col_to_name(CONFIG_BANDS["people"])
    role_col = xl_col_to_name(CONFIG_BANDS["people"] + 1)
    team_col = xl_col_to_name(CONFIG_BANDS["people"] + 2)
    people_end_xl = config_end_xl
    ws.conditional_format(
        f"{person_col}{start_xl}:{team_col}{people_end_xl}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${person_col}{start_xl}="",OR('
                f'${role_col}{start_xl}<>"",'
                f'${team_col}{start_xl}<>""))'
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        f"{team_col}{start_xl}:{team_col}{people_end_xl}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${person_col}{start_xl}<>"",'
                f'${team_col}{start_xl}<>"",'
                f"COUNTIF(dvTeams,${team_col}{start_xl})=0)"
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )


def _write_config_integrity_rules(ws: Worksheet, fmts: Formats, base: int) -> None:
    invalid = fmts.get(
        None,
        bg_color=COLORS["rag_r_bg"],
        font_color=COLORS["rag_r_fg"],
        border=1,
        border_color=COLORS["rag_r_fg"],
    )
    start_xl = base + 2
    config_end_xl = base + 1 + CONFIG_ROWS
    # Duplicate labels make lookups and validation lists ambiguous. Flag them
    # across every Config identity column while still permitting blank spare
    # rows for table growth.
    duplicate_specs = (
        xl_col_to_name(CONFIG_BANDS["statuses"]),
        xl_col_to_name(CONFIG_BANDS["types"]),
        xl_col_to_name(CONFIG_BANDS["priorities"]),
        xl_col_to_name(CONFIG_BANDS["teams"]),
        xl_col_to_name(CONFIG_BANDS["raid_types"]),
        xl_col_to_name(CONFIG_BANDS["raid_statuses"]),
        xl_col_to_name(CONFIG_BANDS["severity"]),
        xl_col_to_name(CONFIG_BANDS["delivery_health"]),
        xl_col_to_name(CONFIG_BANDS["people"]),
    )
    for col_letter in duplicate_specs:
        ws.conditional_format(
            f"{col_letter}{start_xl}:{col_letter}{config_end_xl}",
            {
                "type": "formula",
                "criteria": (
                    f'=AND({col_letter}{start_xl}<>"",COUNTIF('
                    f"${col_letter}${start_xl}:${col_letter}${config_end_xl},"
                    f"{col_letter}{start_xl})>1)"
                ),
                "format": invalid,
            },
        )

    # Delivery Health labels are contiguous. A blank before a later label is
    # invalid because row order defines severity and the final nonblank label
    # defines the direct-blocked state.
    health_col = xl_col_to_name(CONFIG_BANDS["delivery_health"])
    ws.conditional_format(
        f"{health_col}{start_xl}:{health_col}{config_end_xl - 1}",
        {
            "type": "formula",
            "criteria": (
                f'=AND(${health_col}{start_xl}="",'
                f"COUNTA(${health_col}{start_xl + 1}:${health_col}${config_end_xl})>0)"
            ),
            "format": invalid,
            "stop_if_true": True,
        },
    )

    # Config roles are structural dependencies for workbook behavior.
    # Flag the owning header when a list loses the minimum semantics needed by
    # formulas, stamping or views. This creates one clear error surface rather
    # than coloring the blank capacity rows reserved for growth.
    required_roles = (
        (
            CONFIG_BANDS["statuses"],
            CONFIG_BANDS["statuses"] + 3,
            "=OR(COUNTA(dvStatus)=0,COUNTIF(dvStatusActive,TRUE)=0,COUNTIF(dvStatusDone,TRUE)=0)",
        ),
        (
            CONFIG_BANDS["types"],
            CONFIG_BANDS["types"] + 1,
            "=OR(COUNTA(dvTypes)=0,COUNTIF(dvTypeLevels,1)=0)",
        ),
        (CONFIG_BANDS["priorities"], CONFIG_BANDS["priorities"], "=COUNTA(dvPriorities)=0"),
        (
            CONFIG_BANDS["raid_types"],
            CONFIG_BANDS["raid_types"] + 2,
            "=OR(COUNTA(dvRaidTypes)=0,COUNTIF(dvRaidAlert,TRUE)=0,COUNTIF(dvRaidDecision,TRUE)=0)",
        ),
        (
            CONFIG_BANDS["raid_statuses"],
            CONFIG_BANDS["raid_statuses"] + 1,
            "=OR(COUNTA(dvRaidStatuses)=0,COUNTIF(dvRaidClosed,TRUE)=0)",
        ),
        (
            CONFIG_BANDS["severity"],
            CONFIG_BANDS["severity"] + 1,
            "=OR(COUNTA(dvSeverity)=0,COUNT(dvSeverityMinScore)"
            "<>COUNTA(dvSeverity),COUNTIF(dvSeverityMinScore,1)=0)",
        ),
        (
            CONFIG_BANDS["delivery_health"],
            CONFIG_BANDS["delivery_health"],
            "=COUNTA(dvDeliveryHealth)<4",
        ),
    )
    for first_col, last_col, formula in required_roles:
        ws.conditional_format(
            base,
            first_col,
            base,
            last_col,
            {
                "type": "formula",
                "criteria": formula,
                "format": invalid,
                "stop_if_true": True,
            },
        )


def _write_config_guidance(ws: Worksheet, fmts: Formats, base: int) -> None:
    # Guidance is a final parallel band rather than a second vertical section.
    guide_start = CONFIG_BANDS["guidance"]
    guide_end = guide_start + 5
    ws.set_column(guide_start, guide_end, 9)
    ws.merge_range(base, guide_start, base, guide_end, "Editing rules", fmts.config_table_header())
    guide_text = (
        "OVERVIEW TRUTH\n"
        + "\n".join("• " + t for t in config.OVERVIEW_RULES)
        + "\n\nLIST BEHAVIOUR\n"
        "• Row order ranks Priority and Delivery Health.\n"
        "• Delivery Health row 1 is best; row 2 is amber; rows 3+ are red; "
        "the final row is the direct-blocked state.\n"
        "• RAID Probability and Impact use 1-5. Score = Probability \u00d7 Impact (1-25).\n"
        "• Severity is the highest band whose MinScore is no greater than Score.\n"
        "• Type levels 1-6 drive hierarchy and Plan depth.\n"
        "• Status and RAID flags drive stamping and views.\n\n"
        "ENTERING DATES\n"
        "Type 13 Jul 2026. Ctrl+; inserts today. Excel for Mac "
        "validates dates but does not show an in-cell calendar."
    )
    guide_last = base + max(len(config.TYPES), len(config.SETTINGS))
    ws.merge_range(
        base + 1,
        guide_start,
        guide_last,
        guide_end,
        guide_text,
        fmts.get(
            None,
            bg_color=COLORS["surface_subtle"],
            font_color=COLORS["text_secondary"],
            text_wrap=True,
            **ALIGNMENT["narrative"],
        ),
    )


def write_config(ws: Worksheet, fmts: Formats) -> tuple[dict[str, int], dict[str, str]]:
    """Write Config settings, taxonomy tables, roles, people, and guidance.

    Returns:
        Setting-row indexes and plain validation-list ranges.

    """
    ws.hide_gridlines(2)
    _data_page(ws)
    ws.set_tab_color(COLORS["tab_sys"])
    _title_rail(
        ws,
        fmts,
        "Config",
        f"=IF(MAX(ROWS(tblStatuses[Status]),ROWS(tblTypes[Type]),"
        "ROWS(tblPriorities[Priority]),ROWS(tblTeams[Team]),"
        "ROWS(tblRaidTypes[RaidType]),ROWS(tblRaidStatuses[RaidStatus]),"
        "ROWS(tblSeverity[Severity]),ROWS(tblDeliveryHealth[Delivery Health]),"
        f"ROWS(tblPeople[Person]))>{CONFIG_ROWS},"
        f'"⚠ A Config list exceeds the supported {CONFIG_ROWS}-row capacity",'
        '"Settings, taxonomy, workflow roles and people")',
        end_col=37,
    )

    # Every Config component starts on the same row and owns its own column
    # band. A single narrow, unformatted gutter separates adjacent CONFIG_BANDS.
    # This creates a horizontal control library with clean space below each
    # table's populated body.
    base = 2  # Excel row 3: all band headers
    ws.set_row(base, ROWS["table_header"])
    ws.repeat_rows(0, base)

    setting_rows = _write_config_settings(ws, fmts, base)
    list_ranges = _write_config_lists(ws, fmts, base)
    _write_config_entry_validations(ws, base)
    _write_config_band_rules(ws, fmts, base)
    _write_config_role_rules(ws, fmts, base)
    _write_config_integrity_rules(ws, fmts, base)
    _write_config_guidance(ws, fmts, base)
    return setting_rows, list_ranges
