"""Calc sheet: shared validation spills and bounded Overview totals.

Everything semantic resolves through Config: Plan scopes and scope-level
selectors = level-1 types, open/closed and cancelled through status flags,
RAID panels through the IsAlert/IsDecision flags, and severity through
tblSeverity. No taxonomy string is hardcoded here.
"""

from xlsxwriter.utility import xl_cell_to_rowcol
from xlsxwriter.worksheet import Worksheet

from ..core.design import COLORS
from ..core.formulas import encode_formula
from ..core.layout import REGISTRY, SpillZone
from ..spec.capacity import CONFIG_ROWS, DATA_ROWS
from ..spec.items import DIRECT_BLOCKED_HEALTH_FORMULA
from .common import Formats


def write_calc(ws: Worksheet, fmts: Formats) -> None:
    """Write the protected calculation and shared-list worksheet."""
    ws.set_tab_color(COLORS["tab_sys"])
    ws.hide_gridlines(2)
    ws.set_zoom(100)
    grey = fmts.get("calc")
    hdr = fmts.h2()

    heads = {
        "A": "ItemIDs",
        "B": "People",
        "D": "ActiveStatus",
        "E": "DoneStatus",
        "F": "CancelledStatus",
        "G": "ClosedRaid",
        "H": "AlertRaidTypes",
        "I": "DecisionRaidTypes",
        "K": "ScopeIDs",
        "L": "ScopeLabels",
        "M": "ScopeSelectors",
    }
    for c, h in heads.items():
        ws.write(f"{c}1", h, hdr)
        ws.set_column(f"{c}:{c}", 14, grey)

    def dyn(cell: str, formula: str, rows: int, **options: object) -> None:
        cols = int(options.get("cols", 1))
        variables = options.get("variables", ())
        label = str(options.get("label", ""))
        unknown_options = options.keys() - {"cols", "variables", "label"}
        if unknown_options:
            message = f"unsupported dynamic-array options: {sorted(unknown_options)}"
            raise TypeError(message)
        r, c_ = xl_cell_to_rowcol(cell)
        REGISTRY.reserve("Calc", SpillZone(r, c_, r + rows - 1, c_ + cols - 1, label or cell))
        ws.write_dynamic_array_formula(
            r,
            c_,
            r,
            c_,
            encode_formula(formula, variables=variables),
            grey,
        )

    # Shared dropdown sources use bounded-range defined names; formula predicates
    # use the spilled ANCHORARRAY names built over these columns.
    dyn("A2", '=SORT(FILTER(tblItems[ID],tblItems[ID]<>"",""))', DATA_ROWS)
    dyn("B2", '=SORT(FILTER(tblPeople[Person],tblPeople[Person]<>"",""))', CONFIG_ROWS)
    dyn("D2", '=FILTER(tblStatuses[Status],tblStatuses[IsActive]=TRUE,"")', CONFIG_ROWS)
    dyn("E2", '=FILTER(tblStatuses[Status],tblStatuses[IsDone]=TRUE,"")', CONFIG_ROWS)
    dyn("F2", '=FILTER(tblStatuses[Status],tblStatuses[IsCancelled]=TRUE,"")', CONFIG_ROWS)
    dyn("G2", '=FILTER(tblRaidStatuses[RaidStatus],tblRaidStatuses[IsClosed]=TRUE,"")', CONFIG_ROWS)
    dyn("H2", '=FILTER(tblRaidTypes[RaidType],tblRaidTypes[IsAlert]=TRUE,"")', CONFIG_ROWS)
    dyn("I2", '=FILTER(tblRaidTypes[RaidType],tblRaidTypes[IsDecision]=TRUE,"")', CONFIG_ROWS)
    # Scopes = open items whose Type is configured Level 1.
    dyn(
        "K2",
        "=LET(lv,XLOOKUP(tblItems[Type],tblTypes[Type],tblTypes[Level],0),"
        'ids,FILTER(tblItems[ID],(lv=1)*(tblItems[ID]<>"")'
        '*(1-ISNUMBER(XMATCH(tblItems[Status],lstDoneStatus))),""),'
        "SORT(UNIQUE(ids)))",
        DATA_ROWS,
        variables=("lv", "ids"),
    )
    dyn(
        "L2",
        '=IF(K2#="","",K2#&" · "&XLOOKUP(K2#,tblItems[ID],tblItems[Title],"Untitled"))',
        DATA_ROWS,
    )
    dyn("M2", '=VSTACK("All",FILTER(L2#,L2#<>"",""))', DATA_ROWS + 1)

    # Open-RAID predicate pieces shared by the bounded Overview totals.
    raid_open = "(1-ISNUMBER(XMATCH(tblRAID[Status],lstClosedRaid)))"
    raid_alert = "ISNUMBER(XMATCH(tblRAID[Type],lstAlertRaid))"
    raid_decision = "ISNUMBER(XMATCH(tblRAID[Type],lstDecisionRaid))"

    # Overview support block. Totals for the "Showing N of M" disclosures live
    # here, leaving the view surface dedicated to its four panels.
    ws.write("AK1", "OverviewCounts", hdr)
    keydate = "(tblItems[IsPoint]=TRUE)*(tblItems[Level]>=2)*(tblItems[Level]<=cfgKeyDateMaxLevel)"
    not_cancelled = "(1-ISNUMBER(XMATCH(tblItems[Status],lstCancelledStatus)))"
    summaries = {
        "AK2": (
            f"=LET(directblocked,{DIRECT_BLOCKED_HEALTH_FORMULA},"
            'SUMPRODUCT((tblItems[ID]<>"")*(tblItems[Level]>=1)*'
            "(((tblItems[Level]<=cfgExecutiveStatusMaxLevel)+"
            "(tblItems[Delivery Health]=directblocked))>0)*"
            "(1-ISNUMBER(XMATCH(tblItems[Status],lstDoneStatus)))*"
            "(1-ISNUMBER(XMATCH(tblItems[Status],lstCancelledStatus)))))",
            ("directblocked",),
        ),
        "AK3": (
            f'=SUMPRODUCT((tblRAID[Title]<>"")*{raid_open}*{raid_alert}*'
            '(tblRAID[Score]<>"")*(tblRAID[Score]>=cfgAlertSevScore))',
            (),
        ),
        "AK4": (
            f'=SUMPRODUCT({keydate}*(tblItems[Due]<>"")*'
            "(tblItems[Due]>=TODAY())*"
            "(1-ISNUMBER(XMATCH(tblItems[Status],lstDoneStatus)))*"
            f"{not_cancelled})+"
            f'SUMPRODUCT((tblRAID[Title]<>"")*{raid_open}*{raid_decision}*'
            '(tblRAID[NextReview]<>"")*(tblRAID[NextReview]>=TODAY()))',
            (),
        ),
        "AK5": (
            f'=SUMPRODUCT((tblItems[ID]<>"")*(tblItems[DoneDate]<>"")*'
            f"(tblItems[DoneDate]>=TODAY()-cfgReportDays)*{not_cancelled})",
            (),
        ),
    }
    for cell, (formula, variables) in summaries.items():
        ws.write_formula(cell, encode_formula(formula, variables=variables), grey)

    ws.write("AM1", "Calc - internal helper ranges; nothing here is typed by hand", fmts.label())
    ws.hide()
    ws.protect()
