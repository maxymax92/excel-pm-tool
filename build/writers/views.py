"""View sheets: Overview, Plan.

Every spill is bounded and registered in ``layout.REGISTRY``. View sheets are
protected with the Plan controls unlocked. Editable table sheets remain
unprotected so table expansion and row outlines work normally.
Scope predicates use the boolean-algebra pattern ((sel="All")+(col=sel)>0)*...
"""

from xlsxwriter.format import Format
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet

from ..core.design import ALIGNMENT, COLORS, ROWS, TYPOGRAPHY
from ..core.formulas import encode_formula
from ..core.layout import REGISTRY, SpillZone
from ..spec.capacity import PLAN_ROWS, PLAN_WEEKS
from ..spec.items import DIRECT_BLOCKED_HEALTH_FORMULA
from .common import Formats, insert_macro_action, view_chrome

MAX_SPILL_DIMENSIONS = 2


def _dyn(
    ws: Worksheet,
    cell: str,
    formula: str,
    *values: object,
    **options: object,
) -> None:
    if not values or len(values) > MAX_SPILL_DIMENSIONS:
        message = "a dynamic spill requires rows and accepts at most one column count"
        raise TypeError(message)
    rows = int(values[0])
    cols = int(values[1]) if len(values) == MAX_SPILL_DIMENSIONS else 1
    variables = options.get("variables", ())
    fmt = options.get("fmt")
    label = str(options.get("label", ""))
    unknown_options = options.keys() - {"variables", "fmt", "label"}
    if unknown_options:
        message = f"unsupported dynamic-spill options: {sorted(unknown_options)}"
        raise TypeError(message)
    letters = "".join(ch for ch in cell if ch.isalpha())
    r = int(cell[len(letters) :]) - 1
    c = 0
    for ch in letters:
        c = c * 26 + (ord(ch) - 64)
    c -= 1
    REGISTRY.reserve(
        ws.get_name(),
        SpillZone(r, c, r + rows - 1, c + cols - 1, label or cell),
    )
    ws.write_dynamic_array_formula(
        r,
        c,
        r,
        c,
        encode_formula(formula, variables=variables),
        fmt,
    )


# ------------------------------------------------------------ Overview ----
# Side-by-side panel geometry: every panel owns its own columns, separated by
# one thin gutter column, so each panel retains independent geometry.
# Each tuple contains its starting column, widths and headers. Panel content is
# derived from Items, RAID and Config.
PANELS = {
    "scopes": (
        0,
        [26, 14, 9, 12],
        ["Item", "Delivery Health", "Owner", "Due"],
    ),
    "risks": (
        5,
        [9, 26, 10, 9, 12, 20],
        ["Type", "Description", "Severity", "Owner", "Next review", "Latest Status"],
    ),
    "keydates": (12, [28, 12, 16], ["Milestones / Decisions / Deadlines", "Date", "Scope"]),
    "recent": (16, [24, 9, 9, 12, 16], ["Completed work", "Type", "Owner", "Completed", "Scope"]),
}


def _write_overview_frame(ws: Worksheet, fmts: Formats) -> tuple[Format, Format]:
    ws.set_tab_color(COLORS["tab_view"])
    ws.hide_gridlines(2)
    ws.set_zoom(100)
    ws.set_landscape()
    ws.set_paper(9)  # A4
    ws.fit_to_pages(1, 1)
    ws.center_horizontally()
    ws.set_margins(left=0.25, right=0.25, top=0.35, bottom=0.35)
    ws.print_area("A1:U7")

    body = fmts.view_body(text_wrap=True)
    body_anchor = fmts.view_body(text_wrap=True, bold=True)
    body_date = fmts.view_body(text_wrap=True, **ALIGNMENT["panel_date"])
    section = fmts.panel_title()
    section_right = fmts.panel_count()
    table_hdr = fmts.view_table_header()

    for col0, widths, headers in PANELS.values():
        for j, w in enumerate(widths):
            column_fmt = (
                body_date
                if headers[j] in {"Due", "Date", "Completed", "Next review"}
                else body_anchor
                if j == 0
                else body
            )
            ws.set_column(col0 + j, col0 + j, w, column_fmt)
        for j, h in enumerate(headers):
            ws.write(1, col0 + j, h, table_hdr)
        if col0 > 0:
            ws.set_column(col0 - 1, col0 - 1, 3)  # intentional white gutter
    # E/L/P remain visual gutters, but their five body cells hold the numeric
    # date paired with each visible TEXT date to their left. A hidden number
    # format gives date-state rules locale-independent numeric inputs inside
    # the existing visual gutters.
    hidden_date = fmts.get(None, num_format=";;;")
    for gutter in ("E:E", "L:L", "P:P"):
        ws.set_column(gutter, 3, hidden_date)

    def section_bar(key: str, label: str, count_formula: str | None = None) -> None:
        col0, widths, _ = PANELS[key]
        last = col0 + len(widths) - 1
        for c in range(col0, last + 1):
            ws.write_blank(0, c, None, section)
        ws.write(0, col0, label, section)
        if count_formula:
            ws.write_formula(0, last, encode_formula(count_formula), section_right)

    ws.set_row(0, ROWS["panel_title"])
    ws.set_row(1, ROWS["panel_header"])
    for r in range(2, 7):
        ws.set_row(r, ROWS["panel_body"])

    section_bar(
        "scopes", "Executive Status Summary", '=IF(Calc!$AK$2>5,"Showing 5 of "&Calc!$AK$2,"")'
    )
    section_bar("risks", "Top RAID", '=IF(Calc!$AK$3>5,"Showing 5 of "&Calc!$AK$3,"")')
    section_bar("keydates", "Coming up", '=IF(Calc!$AK$4>5,"Showing 5 of "&Calc!$AK$4,"")')
    section_bar("recent", "Recent progress", '=IF(Calc!$AK$5>5,"Showing 5 of "&Calc!$AK$5,"")')
    return body, body_anchor


def write_overview(ws: Worksheet, fmts: Formats, *, is_xlsm: bool = False) -> None:
    """Write the four side-by-side stakeholder panels.

    Every value is derived from Items, RAID and Config. Counts live on Calc AK;
    every displayed date is emitted as TEXT so a column can never show ###;
    the compact scopes panel stays frozen while scrolling right. Spill anchors
    take the bold first-column format so the first visible row matches the
    column emphasis below it.
    """
    _, body_anchor = _write_overview_frame(ws, fmts)

    executive_status = (
        f"=LET(directblocked,{DIRECT_BLOCKED_HEALTH_FORMULA},"
        'pred,(tblItems[ID]<>"")*(tblItems[Level]>=1)*'
        "(((tblItems[Level]<=cfgExecutiveStatusMaxLevel)+"
        "(tblItems[Delivery Health]=directblocked))>0)*"
        "(1-ISNUMBER(XMATCH(tblItems[Status],lstDoneStatus)))*"
        "(1-ISNUMBER(XMATCH(tblItems[Status],lstCancelledStatus))),"
        'ids,FILTER(tblItems[ID],pred,""),'
        "lv,FILTER(tblItems[Level],pred,0),"
        'typ,FILTER(tblItems[Type],pred,""),'
        'ttl,FILTER(tblItems[Title],pred,""),'
        'own,FILTER(tblItems[Owner],pred,""),'
        'dhealth,FILTER(IF(tblItems[Delivery Health]="","Not set",'
        'tblItems[Delivery Health]),pred,""),'
        "hrank,IFNA(XMATCH(dhealth,dvDeliveryHealth),0),"
        "due,FILTER(tblItems[Due],pred,0),"
        'wk,FILTER(tblItems[WbsKey],pred,""),'
        't,HSTACK(REPT("\u203a ",IF(lv>1,lv-1,0))&ids&" · "&typ&" · "&ttl,'
        'dhealth,IF(own="","Owner not set",own),'
        'IF(due=0,"Due date not set",TEXT(due,"d mmm yyyy")),due),'
        'empty,AND(ROWS(ids)=1,INDEX(ids,1)=""),'
        'IF(empty,HSTACK("No open items at configured levels.","","","",""),'
        "LET(s,SORTBY(t,hrank,-1,lv,1,wk,1),"
        'IF(ROWS(s)<=5,s,FILTER(s,SEQUENCE(ROWS(s))<=5,"")))))'
    )
    _dyn(
        ws,
        "A3",
        executive_status,
        5,
        5,
        variables=(
            "directblocked",
            "pred",
            "ids",
            "lv",
            "typ",
            "ttl",
            "own",
            "dhealth",
            "hrank",
            "due",
            "wk",
            "t",
            "empty",
            "s",
        ),
        fmt=body_anchor,
        label="overview-projects",
    )

    risks = (
        '=LET(pred,(tblRAID[Title]<>"")*'
        "(1-ISNUMBER(XMATCH(tblRAID[Status],lstClosedRaid)))*"
        "ISNUMBER(XMATCH(tblRAID[Type],lstAlertRaid))*"
        '(tblRAID[Score]<>"")*(tblRAID[Score]>=cfgAlertSevScore),'
        "t,FILTER(HSTACK(tblRAID[Type],tblRAID[Title],"
        'IF(tblRAID[Severity]="","Not set",tblRAID[Severity]),'
        'IF(tblRAID[Owner]="","Owner not set",tblRAID[Owner]),'
        'IF(tblRAID[NextReview]="","Review date not set",'
        'TEXT(tblRAID[NextReview],"d mmm yyyy")),'
        'IF(tblRAID[Response]="","No response recorded",'
        'IF(LEN(tblRAID[Response])>72,LEFT(tblRAID[Response],69)&"…",'
        'tblRAID[Response])),tblRAID[NextReview]),pred,""),'
        "rk,FILTER(tblRAID[Score],pred,0),"
        'dk,FILTER(IF(tblRAID[NextReview]="",DATE(9999,12,31),tblRAID[NextReview]),'
        'pred,DATE(9999,12,31)),empty,AND(ROWS(t)*COLUMNS(t)=1,INDEX(t,1,1)=""),'
        'IF(empty,HSTACK("—","No high-severity open RAID.","","","","",""),'
        "LET(s,SORTBY(t,rk,-1,dk,1),"
        'IF(ROWS(s)<=5,s,FILTER(s,SEQUENCE(ROWS(s))<=5,"")))))'
    )
    _dyn(
        ws,
        "F3",
        risks,
        5,
        7,
        variables=("pred", "t", "rk", "dk", "empty", "s"),
        fmt=body_anchor,
        label="overview-risks",
    )

    keydate = "(tblItems[IsPoint]=TRUE)*(tblItems[Level]>=2)*(tblItems[Level]<=cfgKeyDateMaxLevel)"
    not_cancelled = "(1-ISNUMBER(XMATCH(tblItems[Status],lstCancelledStatus)))"
    raid_open = "(1-ISNUMBER(XMATCH(tblRAID[Status],lstClosedRaid)))"
    raid_decision = "ISNUMBER(XMATCH(tblRAID[Type],lstDecisionRaid))"
    milestones = (
        f'=LET(ip,{keydate}*(tblItems[Due]<>"")*(tblItems[Due]>=TODAY())*'
        "(1-ISNUMBER(XMATCH(tblItems[Status],lstDoneStatus)))*"
        f"{not_cancelled},"
        'it,FILTER(HSTACK(tblItems[Type]&" · "&tblItems[Title],'
        'TEXT(tblItems[Due],"d mmm yyyy"),'
        'IF(tblItems[Scope]="","Unscoped",XLOOKUP(tblItems[Scope],'
        'tblItems[ID],tblItems[Title],"Unscoped")),tblItems[Due]),ip,""),'
        "ik,FILTER(tblItems[Due],ip,0),"
        'ie,AND(ROWS(it)*COLUMNS(it)=1,INDEX(it,1,1)=""),'
        f'dp,(tblRAID[Title]<>"")*{raid_open}*{raid_decision}*'
        '(tblRAID[NextReview]<>"")*(tblRAID[NextReview]>=TODAY()),'
        'dt,FILTER(HSTACK(tblRAID[Type]&" · "&tblRAID[Title],'
        'TEXT(tblRAID[NextReview],"d mmm yyyy"),'
        'IF(tblRAID[Scope]="","Unscoped",XLOOKUP(tblRAID[Scope],'
        'tblItems[ID],tblItems[Title],"Unscoped")),tblRAID[NextReview]),dp,""),'
        "dk,FILTER(tblRAID[NextReview],dp,0),"
        'de,AND(ROWS(dt)*COLUMNS(dt)=1,INDEX(dt,1,1)=""),'
        "t,IF(ie,dt,IF(de,it,VSTACK(it,dt))),"
        "k,IF(ie,dk,IF(de,ik,VSTACK(ik,dk))),"
        "IF(AND(ie,de),HSTACK("
        '"No upcoming milestones, decisions or deadlines.","","",""),'
        "LET(s,SORTBY(t,k,1),"
        'IF(ROWS(s)<=5,s,FILTER(s,SEQUENCE(ROWS(s))<=5,"")))))'
    )
    _dyn(
        ws,
        "M3",
        milestones,
        5,
        4,
        variables=("ip", "it", "ik", "ie", "dp", "dt", "dk", "de", "t", "k", "s"),
        fmt=body_anchor,
        label="overview-milestones",
    )

    pri_rank = "IFNA(XMATCH(tblItems[Priority],dvPriorities),99)"
    completed = (
        '=LET(pred,(tblItems[ID]<>"")*(tblItems[DoneDate]<>"")*'
        f"(tblItems[DoneDate]>=TODAY()-cfgReportDays)*{not_cancelled},"
        "t,FILTER(HSTACK(tblItems[Title],tblItems[Type],"
        'IF(tblItems[Owner]="","Owner not set",tblItems[Owner]),'
        'TEXT(tblItems[DoneDate],"d mmm yyyy"),'
        'IF(tblItems[Scope]="","Unscoped",XLOOKUP(tblItems[Scope],'
        'tblItems[ID],tblItems[Title],"Unscoped"))),pred,""),'
        "tk,FILTER(IF(tblItems[Level]=0,9,tblItems[Level]),pred,9),"
        f"pk,FILTER({pri_rank},pred,99),"
        "dk,FILTER(tblItems[DoneDate],pred,0),empty,AND(ROWS(t)*COLUMNS(t)=1,"
        'INDEX(t,1,1)=""),IF(empty,HSTACK("No dated completions in this period.",'
        '"","","",""),LET(s,SORTBY(t,tk,1,pk,1,dk,-1),'
        'IF(ROWS(s)<=5,s,FILTER(s,SEQUENCE(ROWS(s))<=5,"")))))'
    )
    _dyn(
        ws,
        "Q3",
        completed,
        5,
        5,
        variables=("pred", "t", "tk", "pk", "dk", "empty", "s"),
        fmt=body_anchor,
        label="overview-completed",
    )

    # frozen scopes panel: scrolling right keeps the anchor context
    ws.freeze_panes(2, 5)

    # Delivery Health and severity pair text with semantic color. Delivery
    # Health row 1 is green, row 2 amber and rows 3 onward red.
    ws.conditional_format(
        "B3:B7",
        {
            "type": "formula",
            "criteria": (
                '=AND($B3<>"",COUNTIF(dvDeliveryHealth,$B3)>0,MATCH($B3,dvDeliveryHealth,0)>=3)'
            ),
            "format": fmts.get(
                None, bg_color=COLORS["rag_r_bg"], font_color=COLORS["rag_r_fg"], bold=True
            ),
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        "B3:B7",
        {
            "type": "formula",
            "criteria": '=AND($B3<>"",$B3=INDEX(dvDeliveryHealth,1))',
            "format": fmts.get(
                None, bg_color=COLORS["rag_g_bg"], font_color=COLORS["rag_g_fg"], bold=True
            ),
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        "B3:B7",
        {
            "type": "formula",
            "criteria": '=AND($B3<>"",$B3=INDEX(dvDeliveryHealth,2))',
            "format": fmts.get(
                None, bg_color=COLORS["rag_a_bg"], font_color=COLORS["rag_a_fg"], bold=True
            ),
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        "H3:H7",
        {
            "type": "formula",
            "criteria": (
                '=AND(COUNTA(dvSeverity)>=2,$H3<>"",'
                "COUNTIF(dvSeverity,$H3)>0,MATCH($H3,dvSeverity,0)"
                ">=MAX(2,COUNTA(dvSeverity)-1))"
            ),
            "format": fmts.get(
                None, bg_color=COLORS["rag_r_bg"], font_color=COLORS["rag_r_fg"], bold=True
            ),
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        "H3:H7",
        {
            "type": "formula",
            "criteria": '=AND($H3<>"",$H3=INDEX(dvSeverity,1))',
            "format": fmts.get(
                None, bg_color=COLORS["rag_g_bg"], font_color=COLORS["rag_g_fg"], bold=True
            ),
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        "H3:H7",
        {
            "type": "formula",
            "criteria": (
                '=AND($H3<>"",COUNTIF(dvSeverity,$H3)>0,'
                "$H3<>INDEX(dvSeverity,1),NOT(AND("
                "COUNTA(dvSeverity)>=2,MATCH($H3,dvSeverity,0)"
                ">=MAX(2,COUNTA(dvSeverity)-1))))"
            ),
            "format": fmts.get(
                None, bg_color=COLORS["rag_a_bg"], font_color=COLORS["rag_a_fg"], bold=True
            ),
            "stop_if_true": True,
        },
    )

    # Scope/RAID date exceptions use the hidden numeric values in the adjacent
    # gutters. Exact visible dates remain the primary cue; colour and weight
    # only accelerate scanning.
    ws.conditional_format(
        "D3:D7",
        {
            "type": "formula",
            "criteria": "=AND(ISNUMBER($E3),$E3>0,$E3<TODAY())",
            "format": fmts.get(
                None, bg_color=COLORS["rag_r_bg"], font_color=COLORS["rag_r_fg"], bold=True
            ),
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        "D3:D7",
        {
            "type": "formula",
            "criteria": ("=AND(ISNUMBER($E3),$E3>=TODAY(),$E3<=TODAY()+cfgComingNearDays)"),
            "format": fmts.get(
                None, bg_color=COLORS["rag_a_bg"], font_color=COLORS["rag_a_fg"], bold=True
            ),
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        "J3:J7",
        {
            "type": "formula",
            "criteria": "=AND(ISNUMBER($L3),$L3>0,$L3<TODAY())",
            "format": fmts.get(
                None, bg_color=COLORS["rag_r_bg"], font_color=COLORS["rag_r_fg"], bold=True
            ),
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        "J3:J7",
        {
            "type": "formula",
            "criteria": '=AND($G3<>"",OR($L3="",$L3=0))',
            "format": fmts.get(
                None, bg_color=COLORS["rag_a_bg"], font_color=COLORS["rag_a_fg"], bold=True
            ),
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        "J3:J7",
        {
            "type": "formula",
            "criteria": ("=AND(ISNUMBER($L3),$L3>=TODAY(),$L3<=TODAY()+cfgComingSoonDays)"),
            "format": fmts.get(
                None, bg_color=COLORS["rag_a_bg"], font_color=COLORS["rag_a_fg"], bold=True
            ),
            "stop_if_true": True,
        },
    )
    # Coming up contains future dated events only: Due-only key dates plus open
    # RAID decisions. The visible date is TEXT for width safety; P carries its
    # paired numeric date invisibly so urgency rules use locale-independent
    # values. Exact date and increasing weight provide the additional cue.
    ws.conditional_format(
        "N3:N7",
        {
            "type": "formula",
            "criteria": ("=AND(ISNUMBER($P3),$P3>=TODAY(),$P3<=TODAY()+cfgComingUrgentDays)"),
            "format": fmts.get(
                None,
                bg_color=COLORS["brand_dark"],
                font_color=COLORS["header_fg"],
                bold=True,
                bottom=2,
                bottom_color=COLORS["danger_strong"],
            ),
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        "N3:N7",
        {
            "type": "formula",
            "criteria": (
                "=AND(ISNUMBER($P3),$P3>TODAY()+cfgComingUrgentDays,$P3<=TODAY()+cfgComingSoonDays)"
            ),
            "format": fmts.get(
                None, bg_color=COLORS["brand"], font_color=COLORS["header_fg"], bold=True
            ),
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        "N3:N7",
        {
            "type": "formula",
            "criteria": (
                "=AND(ISNUMBER($P3),$P3>TODAY()+cfgComingSoonDays,$P3<=TODAY()+cfgComingNearDays)"
            ),
            "format": fmts.get(
                None, bg_color=COLORS["info_bg"], font_color=COLORS["info_fg"], bold=True
            ),
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        "N3:N7",
        {
            "type": "formula",
            "criteria": (
                "=AND(ISNUMBER($P3),$P3>TODAY()+cfgComingNearDays,"
                "$P3<=TODAY()+cfgComingHorizonDays)"
            ),
            "format": fmts.get(
                None, bg_color=COLORS["info_soft_bg"], font_color=COLORS["info_soft_fg"]
            ),
            "stop_if_true": True,
        },
    )

    # One-click export of the panels + Items + RAID as Markdown. Row 32 places
    # the action toward the bottom of the initial maximised Overview viewport,
    # while keeping it comfortably outside the A1:U7 print area.
    if is_xlsm:
        ws.set_row(31, ROWS["toolbar"])
        insert_macro_action(ws, "export")

    # The protected sheet presents derived output.
    ws.protect()


# ---------------------------------------------------------------- Plan ----
def _write_plan_controls(ws: Worksheet, fmts: Formats) -> None:
    ws.set_tab_color(COLORS["tab_view"])
    view_chrome(
        ws, fmts, "Plan - schedule & key dates", [("Scope", "B2", "=dvScopeLabels", "All")], 6
    )
    ws.set_zoom(100)
    ws.set_default_row(ROWS["data_compact"])

    # Two optional scope slots extend the primary Scope filter to three
    # concurrent Level-1 scopes. Blank slots select nothing extra.
    for slot in ("C2", "C3"):
        ws.write_blank(slot, None, fmts.input_cell())
        ws.data_validation(
            slot,
            {
                "validate": "list",
                "source": "=dvScopeLabels",
                "ignore_blank": True,
                "show_input": True,
                "input_title": "Additional scope",
                "input_message": "Blank = unused; otherwise choose "
                "another scope to show alongside.",
                "show_error": True,
                "error_type": "stop",
                "error_title": "Choose a listed value",
                "error_message": "Use the dropdown; free text is not valid.",
            },
        )

    ws.write("A3", "Depth", fmts.label())
    depth_input = fmts.get(
        "int",
        bg_color=COLORS["input_bg"],
        border=1,
        border_color=COLORS["border_strong"],
        locked=False,
        **ALIGNMENT["text"],
    )
    ws.write("B3", 6, depth_input)
    ws.data_validation(
        "B3",
        {
            "validate": "list",
            "source": ["1", "2", "3", "4", "5", "6"],
            "ignore_blank": True,
            "show_input": True,
            "input_title": "Visible depth",
            "input_message": "Blank = all six levels; otherwise choose 1-6.",
            "show_error": True,
            "error_type": "stop",
            "error_title": "Choose a level 1-6",
            "error_message": "Use the dropdown; free text is not valid.",
        },
    )
    win_label = fmts.get(None, bold=True, font_color=COLORS["text_secondary"], **ALIGNMENT["text"])
    win_input = fmts.get(
        "date",
        locked=False,
        bg_color=COLORS["input_bg"],
        border=1,
        border_color=COLORS["border_strong"],
        **ALIGNMENT["date"],
    )
    ws.write("D2", "From", win_label)
    ws.write_blank("E2", None, win_input)
    ws.write("D3", "To", win_label)
    ws.write_blank("E3", None, win_input)
    for cell, edge in (("E2", "start"), ("E3", "end")):
        ws.data_validation(
            cell,
            {
                "validate": "date",
                "criteria": ">=",
                "value": "DATE(2020,1,1)",
                "input_title": "Window " + edge,
                "input_message": "Blank = automatic (the shown items' range). "
                "Use 13 Jul 2026; Ctrl+; inserts today.",
                "show_input": True,
                "ignore_blank": True,
                "show_error": True,
                "error_type": "stop",
                "error_title": "Use a valid date",
                "error_message": "Enter a real date from 2020 onwards, or clear "
                "the cell for automatic.",
            },
        )
    invalid_input = fmts.get(
        None,
        bg_color=COLORS["rag_r_bg"],
        font_color=COLORS["rag_r_fg"],
        border=2,
        border_color=COLORS["rag_r_fg"],
    )
    for target, criteria in (
        ("E2:E3", "=AND(ISNUMBER($E$2),ISNUMBER($E$3),$E$2>$E$3)"),
        ("E2:E3", '=AND(E2<>"",OR(NOT(ISNUMBER(E2)),E2<DATE(2020,1,1)))'),
        ("B3", '=AND($B$3<>"",OR(NOT(ISNUMBER($B$3)),$B$3<1,$B$3>6,INT($B$3)<>$B$3))'),
        ("B2", '=AND($B$2<>"",$B$2<>"All",COUNTIF(dvScopeLabels,$B$2)=0)'),
        ("C2", '=AND($C$2<>"",$C$2<>"All",COUNTIF(dvScopeLabels,$C$2)=0)'),
        ("C3", '=AND($C$3<>"",$C$3<>"All",COUNTIF(dvScopeLabels,$C$3)=0)'),
    ):
        ws.conditional_format(
            target,
            {
                "type": "formula",
                "criteria": criteria,
                "format": invalid_input,
                "stop_if_true": True,
            },
        )


def _write_plan_axis_frame(
    ws: Worksheet,
    fmts: Formats,
) -> tuple[Format, Format, Format, Format, Format]:
    for cell_range, text, background, foreground in (
        ("F1:I1", "✓ Done", COLORS["bar_done_bg"], COLORS["bar_done_fg"]),
        ("J1:N1", "● In progress", COLORS["bar_active_bg"], COLORS["bar_active_fg"]),
        ("O1:R1", "— Planned", COLORS["bar_plan_bg"], COLORS["bar_plan_fg"]),
        ("S1:V1", "! Overdue", COLORS["bar_over_bg"], COLORS["bar_over_fg"]),
        ("W1:Z1", "\u00d7 Cancelled", COLORS["bar_cancel_bg"], COLORS["bar_cancel_fg"]),
    ):
        ws.merge_range(
            cell_range,
            text,
            fmts.get(
                None,
                align="center",
                font_size=TYPOGRAPHY["caption"],
                bold=True,
                bg_color=background,
                font_color=foreground,
                valign="vcenter",
            ),
        )
    ws.merge_range(
        "AA1:AE1",
        "◆ Key date",
        fmts.get(
            None,
            align="center",
            font_size=TYPOGRAPHY["caption"],
            bold=True,
            bg_color=COLORS["brand_tint"],
            font_color=COLORS["pt_next"],
            valign="vcenter",
        ),
    )
    ws.merge_range(
        "AF1:AI1",
        "│ Today",
        fmts.get(
            None,
            align="center",
            font_size=TYPOGRAPHY["caption"],
            bold=True,
            bg_color=COLORS["surface_subtle"],
            font_color=COLORS["today"],
            valign="vcenter",
        ),
    )

    left_text = fmts.view_body(**ALIGNMENT["text"])
    left_date = fmts.view_body(num_format="dd mmm yyyy", **ALIGNMENT["date"])
    ws.set_column("A:A", 10, left_text)
    ws.set_column("B:B", 44, left_text)
    ws.set_column("C:C", 11, left_text)
    ws.set_column("D:E", 12, left_date)
    grid_fmt = fmts.view_body(align="center", valign="vcenter")
    axis_month_fmt = fmts.get(
        None,
        font_size=TYPOGRAPHY["caption"],
        bold=True,
        align="left",
        valign="vcenter",
        bg_color=COLORS["brand_tint"],
        font_color=COLORS["header_bg"],
        bottom=1,
        bottom_color=COLORS["border"],
    )
    axis_week_fmt = fmts.get(
        None,
        num_format="d",
        font_size=TYPOGRAPHY["caption"],
        align="center",
        valign="vcenter",
        bg_color=COLORS["header_bg"],
        font_color=COLORS["header_fg"],
    )
    ws.set_column(5, 56, 3.5, grid_fmt)
    ws.set_row(3, ROWS["axis_month"], axis_month_fmt)
    ws.set_row(4, ROWS["axis_week"], axis_week_fmt)
    ws.freeze_panes(5, 5)
    for column, header in enumerate(("ID", "Item", "Owner", "Start", "Due")):
        ws.write(4, column, header, fmts.data_table_header())
    ws.write_formula(
        "BG2",
        encode_formula('=IF(OR(B2="",B2="All"),"All",LEFT(B2,FIND(" · ",B2)-1))'),
        fmts.get("calc"),
    )
    # The extra scope slots resolve to an em-dash sentinel that matches no
    # Scope value, so blank or malformed slots never widen the view.
    for helper, slot in (("BH2", "C2"), ("BI2", "C3")):
        ws.write_formula(
            helper,
            encode_formula(
                f'=IFERROR(IF({slot}="","(unused)",IF({slot}="All","All",'
                f'LEFT({slot},FIND(" · ",{slot})-1))),"(unused)")'
            ),
            fmts.get("calc"),
        )
    ws.set_column("BG:BI", 14, fmts.get("calc"), {"hidden": True})
    return left_text, left_date, grid_fmt, axis_month_fmt, axis_week_fmt


def write_plan(wb: Workbook, ws: Worksheet, fmts: Formats) -> None:
    """Write the WBS-ordered schedule and weekly timeline.

    Filters stack in the frozen pane (Scope/Depth in B2/B3, From/To window in
    E2/E3 — blank window = the shown rows' own range padded a week each
    side), so the grid starts immediately at F with no dead column. Rows are
    dynamic spills; per-row treatments arrive via conditional formatting keyed
    off hidden helper spills (BH level, BI category). CF cannot change font
    SIZE — level emphasis here is bold + banding; the true per-level type
    sizes live on the Items sheet. The sheet is protected: only the four
    filter cells are editable.
    """
    _write_plan_controls(ws, fmts)
    left_text, left_date, grid_fmt, axis_month_fmt, axis_week_fmt = _write_plan_axis_frame(ws, fmts)

    # Scope and depth determine row inclusion. Schedule completeness only
    # changes the displayed dates and timeline mark; it never hides an item.
    pred = (
        '(tblItems[ID]<>"")*'
        '(((selPScopeID="All")+(selPScopeID2="All")+(selPScopeID3="All")+'
        "(tblItems[Scope]=selPScopeID)+(tblItems[Scope]=selPScopeID2)+"
        "(tblItems[Scope]=selPScopeID3))>0)*"
        "(tblItems[Level]<=dp)"
    )
    _dyn(
        ws,
        "A6",
        f'=LET(dp,IF(selPDepth="",6,VALUE(selPDepth)),'
        f'k,FILTER(tblItems[ID],{pred},""),'
        f'w,FILTER(tblItems[WbsKey],{pred},""),'
        'IF(ROWS(k)=1,IF(INDEX(k,1,1)="","— none —",k),'
        f"LET(s,SORTBY(k,w),TAKE(s,{PLAN_ROWS}))))",
        PLAN_ROWS,
        variables=("dp", "k", "w", "s"),
        fmt=left_text,
        label="plan-ids",
    )
    guard = '=IF(A6#="","",IF(A6#="— none —","",{}))'
    # Elementwise IF preserves one indent value per spilled row.
    _dyn(
        ws,
        "B6",
        "=LET(ids,A6#,lv,XLOOKUP(ids,tblItems[ID],tblItems[Level],1),"
        "t,fnItemLookup(ids,tblItems[Title]),"
        'IF((ids="")+(ids="— none —"),"",'
        'REPT("   ",IF(lv<1,0,lv-1))&t))',
        PLAN_ROWS,
        variables=("ids", "lv", "t"),
        fmt=left_text,
        label="plan-title",
    )
    _dyn(
        ws,
        "C6",
        guard.format("fnItemLookup(A6#,tblItems[Owner])"),
        PLAN_ROWS,
        fmt=left_text,
        label="plan-owner",
    )
    _dyn(
        ws,
        "D6",
        guard.format("fnItemLookup(A6#,tblItems[Start])"),
        PLAN_ROWS,
        fmt=left_date,
        label="plan-start",
    )
    _dyn(
        ws,
        "E6",
        guard.format("fnItemLookup(A6#,tblItems[Due])"),
        PLAN_ROWS,
        fmt=left_date,
        label="plan-due",
    )
    _dyn(
        ws,
        "BH6",
        guard.format("XLOOKUP(A6#,tblItems[ID],tblItems[Level],0)"),
        PLAN_ROWS,
        fmt=fmts.get("calc"),
        label="plan-level",
    )
    _dyn(
        ws,
        "BI6",
        "=LET(ids,A6#,"
        'stt,XLOOKUP(ids,tblItems[ID],tblItems[Status],""),'
        "ed,XLOOKUP(ids,tblItems[ID],tblItems[Due],0),"
        "cn,ISNUMBER(XMATCH(stt,lstCancelledStatus)),"
        "dn,ISNUMBER(XMATCH(stt,lstDoneStatus)),"
        "ac,ISNUMBER(XMATCH(stt,lstActiveStatus)),"
        'IF((ids="")+(ids="— none —"),"",'
        'IF(cn,"C",IF(dn,"D",IF((ed<>0)*(ed<TODAY()),"O",'
        'IF(ac,"A","P"))))))',
        PLAN_ROWS,
        variables=("ids", "stt", "ed", "cn", "dn", "ac"),
        fmt=fmts.get("calc"),
        label="plan-cat",
    )

    # The Monday axis uses From/To when present, otherwise the visible date
    # range with weekly padding. An empty view centers on the current period.
    _dyn(
        ws,
        "F5",
        "=LET(ids,A6#,big,DATE(9999,12,31),"
        'es0,XLOOKUP(ids,tblItems[ID],tblItems[Start],""),'
        'du0,XLOOKUP(ids,tblItems[ID],tblItems[Due],""),'
        'starts,IF((es0="")+(es0=0),IF((du0="")+(du0=0),big,du0),es0),'
        'ends,IF((du0="")+(du0=0),IF((es0="")+(es0=0),0,es0),du0),'
        "lo0,MIN(starts),hi0,MAX(ends),"
        'lo,IF(selPFrom<>"",selPFrom,IF(lo0>=big,TODAY()-28,lo0-7)),'
        'hi,IF(selPTo<>"",selPTo,IF(hi0=0,TODAY()+56,hi0+7)),'
        "a,lo-WEEKDAY(lo,3),"
        f"n,MAX(4,MIN({PLAN_WEEKS},ROUNDUP((hi-a+1)/7,0))),"
        "SEQUENCE(1,n,a,7))",
        1,
        PLAN_WEEKS,
        variables=(
            "ids",
            "big",
            "es0",
            "du0",
            "starts",
            "ends",
            "lo0",
            "hi0",
            "lo",
            "hi",
            "a",
            "n",
        ),
        fmt=axis_week_fmt,
        label="plan-axis",
    )
    _dyn(
        ws,
        "F4",
        "=LET(wk,F5#,s,SEQUENCE(1,COLUMNS(wk)),"
        'IF((s=1)+(MONTH(wk)<>MONTH(wk-7)),TEXT(wk,"mmm"),""))',
        1,
        PLAN_WEEKS,
        variables=("wk", "s"),
        fmt=axis_month_fmt,
        label="plan-months",
    )
    _dyn(
        ws,
        "F6",
        "=LET(ids,A6#,wk,F5#,ct,BI6#,"
        'es0,XLOOKUP(ids,tblItems[ID],tblItems[Start],""),'
        'es,IF(es0=0,"",es0),'
        "pt,XLOOKUP(ids,tblItems[ID],tblItems[IsPoint],FALSE),"
        'du0,XLOOKUP(ids,tblItems[ID],tblItems[Due],""),'
        'du,IF(du0=0,"",du0),'
        'glyph,IF(ct="D","✓",IF(ct="A","●",IF(ct="O","!",'
        'IF(ct="C","\u00d7","—")))),'
        'IF((ids="")+(ids="— none —"),"",'
        'IF(pt=TRUE,IF((du<>"")*(du>=wk)*(du<wk+7),"◆",""),'
        'IF((es<>"")*(du<>"")*(wk+6>=es)*(wk<=du),glyph,""))))',
        PLAN_ROWS,
        PLAN_WEEKS,
        variables=(
            "ids",
            "wk",
            "ct",
            "es0",
            "es",
            "pt",
            "du0",
            "du",
            "glyph",
        ),
        fmt=grid_fmt,
        label="plan-grid",
    )

    # The status rail reports invalid controls, capacity and unscheduled work.
    # Its dynamic calculation lives in an unmerged system cell because Excel
    # requires dynamic-array formula records to remain outside merged ranges.
    status_formula = encode_formula(
        '=LET(scopeBad,OR(AND(selPScope<>"",selPScope<>"All",'
        "COUNTIF(dvScopeLabels,selPScope)=0),"
        'AND(selPScope2<>"",selPScope2<>"All",'
        "COUNTIF(dvScopeLabels,selPScope2)=0),"
        'AND(selPScope3<>"",selPScope3<>"All",'
        "COUNTIF(dvScopeLabels,selPScope3)=0)),"
        'depthBad,AND(selPDepth<>"",OR(NOT(ISNUMBER(selPDepth)),'
        "selPDepth<1,selPDepth>6,INT(selPDepth)<>selPDepth)),"
        'fromBad,AND(selPFrom<>"",OR(NOT(ISNUMBER(selPFrom)),'
        "selPFrom<DATE(2020,1,1))),"
        'toBad,AND(selPTo<>"",OR(NOT(ISNUMBER(selPTo)),'
        "selPTo<DATE(2020,1,1))),"
        "orderBad,AND(ISNUMBER(selPFrom),ISNUMBER(selPTo),selPFrom>selPTo),"
        'dp,IF(depthBad,6,IF(selPDepth="",6,VALUE(selPDepth))),'
        'scopeID,IF(scopeBad,"All",IF(OR(selPScope="",selPScope="All"),'
        '"All",selPScopeID)),'
        'scope2ID,IF(scopeBad,"(unused)",selPScopeID2),'
        'scope3ID,IF(scopeBad,"(unused)",selPScopeID3),'
        'rowsok,(tblItems[ID]<>"")*'
        '(((scopeID="All")+(scope2ID="All")+(scope3ID="All")+'
        "(tblItems[Scope]=scopeID)+(tblItems[Scope]=scope2ID)+"
        "(tblItems[Scope]=scope3ID))>0)*"
        "(tblItems[Level]<=dp),"
        "eligible,SUMPRODUCT(rowsok),"
        'ids,FILTER(tblItems[ID],rowsok,""),big,DATE(9999,12,31),'
        'es0,XLOOKUP(ids,tblItems[ID],tblItems[Start],""),'
        'du0,XLOOKUP(ids,tblItems[ID],tblItems[Due],""),'
        'starts,IF((es0="")+(es0=0),IF((du0="")+(du0=0),big,du0),es0),'
        'ends,IF((du0="")+(du0=0),IF((es0="")+(es0=0),0,es0),du0),'
        "lo0,MIN(starts),hi0,MAX(ends),"
        "lo,IF(fromBad,IF(lo0>=big,TODAY()-28,lo0-7),"
        'IF(selPFrom<>"",selPFrom,IF(lo0>=big,TODAY()-28,lo0-7))),'
        "hi,IF(toBad,IF(hi0=0,TODAY()+56,hi0+7),"
        'IF(selPTo<>"",selPTo,IF(hi0=0,TODAY()+56,hi0+7))),'
        "a,lo-WEEKDAY(lo,3),span,ROUNDUP((hi-a+1)/7,0),"
        "n,SUMPRODUCT(rowsok*"
        "(1-(tblItems[IsPoint]=TRUE))*"
        '(1-((tblItems[Start]<>"")*(tblItems[Due]<>"")))),'
        'IF(scopeBad,"⚠ Scope is not in the Config-driven list",'
        'IF(depthBad,"⚠ Depth must be a whole number from 1 to 6",'
        'IF(OR(fromBad,toBad),"⚠ From/To must be real dates from 2020 onwards",'
        'IF(orderBad,"⚠ From is after To — correct the reporting window",'
        f'IF(eligible>{PLAN_ROWS},"⚠ "&eligible-{PLAN_ROWS}&'
        '" items exceed the supported Plan capacity",'
        f'IF(span>{PLAN_WEEKS},"⚠ The requested schedule spans "&span&'
        f'" weeks; Plan shows the first {PLAN_WEEKS}",'
        'IF(n=0,"","⚠ "&n&IF(n=1," item has"," items have")&'
        '" incomplete schedule dates — shown without a timeline mark; add '
        'Start+Due, or Due only for a key date"'
        "))))))))",
        variables=(
            "scopeBad",
            "depthBad",
            "fromBad",
            "toBad",
            "orderBad",
            "dp",
            "scopeID",
            "scope2ID",
            "scope3ID",
            "rowsok",
            "eligible",
            "ids",
            "big",
            "es0",
            "du0",
            "starts",
            "ends",
            "lo0",
            "hi0",
            "lo",
            "hi",
            "a",
            "span",
            "n",
        ),
    )
    ws.write_formula("BG3", status_formula, fmts.get("calc"))
    status_format = fmts.get(
        None,
        font_size=TYPOGRAPHY["caption"],
        font_color=COLORS["text_secondary"],
        valign="vcenter",
    )
    ws.merge_range(
        "F2:R2",
        "",
        status_format,
    )
    ws.write_formula("F2", "=BG3", status_format)
    ws.conditional_format(
        "F2:R2",
        {
            "type": "formula",
            "criteria": '=$F$2<>""',
            "format": fmts.get(None, bg_color=COLORS["rag_a_bg"], font_color=COLORS["rag_a_fg"]),
        },
    )

    # Key dates use one consistent navy diamond; bar rules apply state fills to
    # interval glyphs.
    # Border-only current-week and month rulers evaluate across both states.
    plan_last_row = PLAN_ROWS + 5
    grid = f"F6:BE{plan_last_row}"
    # Retain the full supported range without asking Excel to evaluate the
    # timeline rule stack for unused spill rows.
    ws.conditional_format(
        grid,
        {
            "type": "formula",
            "criteria": '=OR($A6="",$A6="— none —")',
            "stop_if_true": True,
        },
    )
    ws.conditional_format(
        grid,
        {
            "type": "formula",
            "criteria": '=F6="◆"',
            "format": fmts.get(None, font_color=COLORS["pt_next"], bold=True),
        },
    )
    for cat, bg, fg in [
        ("D", COLORS["bar_done_bg"], COLORS["bar_done_fg"]),
        ("A", COLORS["bar_active_bg"], COLORS["bar_active_fg"]),
        ("O", COLORS["bar_over_bg"], COLORS["bar_over_fg"]),
        ("P", COLORS["bar_plan_bg"], COLORS["bar_plan_fg"]),
        ("C", COLORS["bar_cancel_bg"], COLORS["bar_cancel_fg"]),
    ]:
        # White top/bottom borders inset each bar within its cell so
        # adjacent rows read as separate bars rather than one block.
        ws.conditional_format(
            grid,
            {
                "type": "formula",
                "criteria": f'=AND(F6<>"",F6<>"◆",$BI6="{cat}")',
                "format": fmts.get(
                    None,
                    bg_color=bg,
                    font_color=fg,
                    top=2,
                    top_color=COLORS["surface"],
                    bottom=2,
                    bottom_color=COLORS["surface"],
                ),
            },
        )
    # Border-only differential formats are created directly so they carry only
    # the timeline ruler and preserve the navy axis text color.
    today_rule = wb.add_format({"left": 2, "left_color": COLORS["today"]})
    month_rule = wb.add_format({"left": 1, "left_color": COLORS["bar_grey"]})
    ws.conditional_format(
        "F4:BE5",
        {"type": "formula", "criteria": "=AND(F$5<=TODAY(),TODAY()<F$5+7)", "format": today_rule},
    )
    ws.conditional_format(
        grid,
        {
            "type": "formula",
            "criteria": '=AND($A6<>"",F$5<=TODAY(),TODAY()<F$5+7)',
            "format": today_rule,
        },
    )
    ws.conditional_format(
        "F4:BE5",
        {
            "type": "formula",
            "criteria": (
                "=AND(COLUMN(F$5)>COLUMN($F$5),"
                "MONTH(F$5)<>MONTH(F$5-7),"
                "NOT(AND(F$5<=TODAY(),TODAY()<F$5+7)))"
            ),
            "format": month_rule,
        },
    )
    ws.conditional_format(
        grid,
        {
            "type": "formula",
            "criteria": (
                '=AND($A6<>"",COLUMN(F$5)>COLUMN($F$5),'
                "MONTH(F$5)<>MONTH(F$5-7),"
                "NOT(AND(F$5<=TODAY(),TODAY()<F$5+7)))"
            ),
            "format": month_rule,
        },
    )

    # -- hierarchy and scan rhythm. Bar rules above keep priority; these fills
    # affect the identity area and empty timeline cells only.
    ws.conditional_format(
        f"A6:E{plan_last_row}",
        {
            "type": "formula",
            "criteria": '=AND($A6<>"",$A6<>"— none —",$BH6=1)',
            "format": fmts.get(None, bold=True, bg_color=COLORS["surface_subtle"]),
        },
    )
    ws.conditional_format(
        grid,
        {
            "type": "formula",
            "criteria": '=AND($A6<>"",$A6<>"— none —",$BH6=1,F6="")',
            "format": fmts.get(None, bg_color=COLORS["surface_subtle"]),
        },
    )
    ws.conditional_format(
        f"A6:E{plan_last_row}",
        {
            "type": "formula",
            "criteria": '=AND($A6<>"",$A6<>"— none —",$BH6>=1,$BH6<=3)',
            "format": fmts.get(None, bold=True),
        },
    )
    ws.conditional_format(
        f"A6:E{plan_last_row}",
        {
            "type": "formula",
            "criteria": '=AND($A6<>"",ISEVEN(ROW()))',
            "format": fmts.get(None, bg_color=COLORS["canvas"]),
        },
    )
    ws.conditional_format(
        grid,
        {
            "type": "formula",
            "criteria": '=AND($A6<>"",F6="",ISEVEN(ROW()))',
            "format": fmts.get(None, bg_color=COLORS["canvas"]),
        },
    )

    # print parity with Overview: landscape A4, all weeks on one page width
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(0, 4)
    ws.print_area(f"A1:BE{plan_last_row}")
    ws.set_margins(left=0.25, right=0.25, top=0.35, bottom=0.35)

    # computed view: everything locked except Scope / Depth / From / To
    ws.protect()
