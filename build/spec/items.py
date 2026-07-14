"""tblItems column spec and supporting table specs.

Each column: (name, kind, fmt, width, dv, formula, vars)
  kind: I=input, F=formula (consistent calculated column), V=VBA-stamped value
  fmt:  key into the format factory (writers/common.py)
  dv:   data-validation spec dict or None (applied through DATA_ROWS)
"""

type ColumnSpec = dict[str, object]


def _c(name: str, kind: str, *values: object, **overrides: object) -> ColumnSpec:
    """Build one column specification from the compact table declaration.

    Returns:
        A normalized table-column declaration.

    Raises:
        TypeError: If options are duplicated, unknown, or excessive.

    """
    option_names = ("fmt", "width", "dv", "formula", "vars", "checkbox", "wrap")
    if len(values) > len(option_names):
        message = f"column {name!r} has {len(values)} positional options; expected at most 7"
        raise TypeError(message)
    options = dict(zip(option_names, values, strict=False))
    duplicate_options = options.keys() & overrides.keys()
    if duplicate_options:
        message = f"column {name!r} repeats options: {sorted(duplicate_options)}"
        raise TypeError(message)
    options.update(overrides)
    unknown_options = options.keys() - set(option_names)
    if unknown_options:
        message = f"column {name!r} has unknown options: {sorted(unknown_options)}"
        raise TypeError(message)

    return {
        "name": name,
        "kind": kind,
        "fmt": options.get("fmt"),
        "width": options.get("width", 10),
        "dv": options.get("dv"),
        "formula": options.get("formula"),
        "vars": options.get("vars", ()),
        "checkbox": options.get("checkbox", False),
        "wrap": options.get("wrap", False),
    }


def list_validation(source: str) -> dict[str, str]:
    """Return a list-validation specification for one supported source.

    Returns:
        The XlsxWriter validation settings.

    """
    return {"validate": "list", "source": source}


DATE_OK = {
    "validate": "date",
    "criteria": ">=",
    "value": "DATE(2020,1,1)",
    "input_title": "Enter a date",
    "input_message": "Use 13 Jul 2026. Ctrl+; inserts today.",
    "show_input": True,
    "show_error": True,
    "error_type": "stop",
    "error_title": "Use a valid date",
    "error_message": "Enter a real date from 2020 onwards.",
}


# The final configured nonblank Delivery Health value is the direct-blocked state.
DIRECT_BLOCKED_HEALTH_FORMULA = 'LOOKUP(2,1/(dvDeliveryHealth<>""),dvDeliveryHealth)'


def raid_rating_validation(title: str, low_label: str, high_label: str) -> dict[str, object]:
    """Return the shared 1-5 RAID rating contract with an in-cell explanation.

    Returns:
        The XlsxWriter validation settings.

    """
    return {
        "validate": "integer",
        "criteria": "between",
        "minimum": 1,
        "maximum": 5,
        "show_input": True,
        "input_title": f"{title} (1-5)",
        "input_message": (
            f"1 = {low_label}; 5 = {high_label}. Score = Probability \u00d7 Impact (1-25)."
        ),
        "show_error": True,
        "error_type": "stop",
        "error_title": f"{title} must be 1-5",
        "error_message": "Enter a whole number from 1 to 5.",
    }


ITEMS_COLUMNS = [
    _c("ID", "V", "text", 9),
    _c("Title", "I", None, 36, wrap=True),
    _c("Type", "I", None, 12, list_validation("=lstTypes")),
    _c("Parent", "I", "text", 11, list_validation("=lstItemIDs")),
    _c("Status", "I", None, 13, list_validation("=lstStatus")),
    _c("Delivery Health", "I", None, 16, list_validation("=lstDeliveryHealth")),
    _c("Priority", "I", None, 10, list_validation("=lstPriorities")),
    _c("Owner", "I", None, 12, list_validation("=lstPeople")),
    _c("Start", "I", "date", 12, DATE_OK),
    _c("Due", "I", "date", 12, DATE_OK),
    _c("BlockedBy", "I", "text", 12),
    # Free-text delivery narrative. VBA stamps LatestUpdateOn when this cell is
    # edited; open narratives older than cfgStaleDays receive attention styling.
    _c("Latest Status", "I", None, 30, wrap=True),
    # ---- formula layer -------------------------------------------------
    _c(
        "ParentTitle",
        "F",
        "calc",
        16,
        formula='=IF([@Parent]="","",fnItemLookup([@Parent],tblItems[Title]))',
    ),
    _c(
        "ParentLevel",
        "F",
        "calcint",
        6,
        formula='=IF([@Parent]="",0,IFNA(INDEX(tblItems[Level],'
        "MATCH([@Parent],tblItems[ID],0)),0))",
    ),
    _c(
        "A2",
        "F",
        "calc",
        4,
        formula='=IF([@Parent]="","",fnItemLookup([@Parent],tblItems[Parent]))',
    ),
    _c("A3", "F", "calc", 4, formula='=IF([@A2]="","",fnItemLookup([@A2],tblItems[Parent]))'),
    _c("A4", "F", "calc", 4, formula='=IF([@A3]="","",fnItemLookup([@A3],tblItems[Parent]))'),
    _c("A5", "F", "calc", 4, formula='=IF([@A4]="","",fnItemLookup([@A4],tblItems[Parent]))'),
    # Level = the Type's configured level in tblTypes (1-6; 0 = blank/unknown
    # type). Drives indentation, emphasis and the Plan/Items depth controls.
    _c("Level", "F", "calcint", 5, formula='=IF([@ID]="",0,fnTypeLevel([@Type]))'),
    # Scope is the nearest Level-1 ancestor, including the row itself at Level 1.
    _c("Scope", "F", "calc", 9, formula="=fnAncestorAtLevel([@ID],1,0)"),
    _c("Children", "F", "calcint", 7, formula='=IF([@ID]="",0,COUNTIF(tblItems[Parent],[@ID]))'),
    _c("WaitingOn", "F", "calc", 12, formula="=fnDepOpen([@BlockedBy])"),
    _c("BlockedRefsValid", "F", "calcbool", 8, formula="=fnRefsValid([@BlockedBy],[@ID])"),
    _c(
        "IsBlocked",
        "F",
        "calcbool",
        8,
        formula=(
            '=OR(AND([@[Delivery Health]]<>"",[@[Delivery Health]]='
            f'{DIRECT_BLOCKED_HEALTH_FORMULA}),[@WaitingOn]<>"")'
        ),
    ),
    _c(
        "DueIn", "F", "duein", 6, formula='=IF(OR([@Due]="",fnIsDone([@Status])),"",[@Due]-TODAY())'
    ),
    _c(
        "Health",
        "F",
        "health",
        6,
        formula="=fnHealthRAG([@Status],[@Due],[@IsBlocked],[@BlockedSince],[@Updated])",
    ),
    _c(
        "AgeDays",
        "F",
        "calcint",
        7,
        formula='=IF([@Created]="","",fnBizDays([@Created],'
        'IF([@DoneDate]="",TODAY(),[@DoneDate])))',
    ),
    _c(
        "CycleDays",
        "F",
        "calcint",
        7,
        formula='=IF([@InProgressSince]="","",fnBizDays([@InProgressSince],'
        'IF([@DoneDate]="",TODAY(),[@DoneDate])))',
    ),
    # ---- VBA-stamped values --------------------------------------------
    _c("Created", "V", "date", 12, DATE_OK),
    _c("Updated", "V", "date", 12, DATE_OK),
    _c("InProgressSince", "V", "date", 12, DATE_OK),
    _c("DoneDate", "V", "date", 12, DATE_OK),
    _c("BlockedSince", "V", "date", 12, DATE_OK),
    _c("LatestUpdateOn", "V", "date", 12, DATE_OK),
    # ---- schedule envelope (Plan reads these) ---------------------------
    # EffStart/EffDue: the item's own dates widened to the envelope of every
    # descendant's dates (min Start / max Due over the Parent + A2..A5 chain).
    # Parent bars reflect the complete descendant schedule.
    _c(
        "EffStart",
        "F",
        "calcdate",
        10,
        formula='=IF([@ID]="","",IF([@Children]=0,'
        'IF([@Start]="","",[@Start]),LET(b,DATE(9999,12,31),'
        's1,MINIFS(tblItems[Start],tblItems[Parent],[@ID],tblItems[Start],">0"),'
        's2,MINIFS(tblItems[Start],tblItems[A2],[@ID],tblItems[Start],">0"),'
        's3,MINIFS(tblItems[Start],tblItems[A3],[@ID],tblItems[Start],">0"),'
        's4,MINIFS(tblItems[Start],tblItems[A4],[@ID],tblItems[Start],">0"),'
        's5,MINIFS(tblItems[Start],tblItems[A5],[@ID],tblItems[Start],">0"),'
        'own,IF([@Start]="",b,[@Start]),'
        "m,MIN(IF(s1=0,b,s1),IF(s2=0,b,s2),IF(s3=0,b,s3),"
        'IF(s4=0,b,s4),IF(s5=0,b,s5),own),IF(m=b,"",m))))',
        vars=("b", "s1", "s2", "s3", "s4", "s5", "own", "m"),
    ),
    _c(
        "EffDue",
        "F",
        "calcdate",
        10,
        formula='=IF([@ID]="","",IF([@Children]=0,'
        'IF([@Due]="","",[@Due]),LET(m,MAX('
        "MAXIFS(tblItems[Due],tblItems[Parent],[@ID]),"
        "MAXIFS(tblItems[Due],tblItems[A2],[@ID]),"
        "MAXIFS(tblItems[Due],tblItems[A3],[@ID]),"
        "MAXIFS(tblItems[Due],tblItems[A4],[@ID]),"
        "MAXIFS(tblItems[Due],tblItems[A5],[@ID]),"
        'N([@Due])),IF(m=0,"",m))))',
        vars=("m",),
    ),
    # A key date has a Due and a blank Start. It renders as a Plan diamond and
    # feeds Overview through the cfgKeyDateMaxLevel boundary.
    _c("IsPoint", "F", "calcbool", 8, formula='=IF([@ID]="",FALSE,AND([@Due]<>"",[@Start]=""))'),
    # Sortable hierarchy path: ascending sort = WBS order (parents first,
    # children grouped, siblings by date). Plan sorts by it; the VBA
    # OrganiseItems button sorts the physical Items rows by it. Blank rows use
    # a high alphabetic sentinel because Mac Excel collates punctuation before
    # the digit-only hierarchy keys in a table sort.
    _c(
        "WbsKey",
        "F",
        "calc",
        14,
        formula='=IF([@ID]="",REPT("Z",50),fnWbsKey([@ID],0))',
    ),
]

# Core columns form the everyday input surface. Parent identifies the branch
# used for Scope, ordering and schedule envelopes; calculated and automation
# fields occupy one collapsed detail group to the right.
CORE_VISIBLE = [
    "ID",
    "Type",
    "Title",
    "Parent",
    "Priority",
    "Start",
    "Status",
    "Due",
    "Delivery Health",
    "Latest Status",
    "Owner",
]


def _core_first(cols: list[ColumnSpec], core: list[str]) -> list[ColumnSpec]:
    column_names = [column["name"] for column in cols]
    if len(column_names) != len(set(column_names)):
        msg = "column specifications contain duplicate names"
        raise ValueError(msg)
    if len(core) != len(set(core)):
        msg = "core-visible column list contains duplicate names"
        raise ValueError(msg)
    missing = [name for name in core if name not in column_names]
    if missing:
        msg = f"core-visible columns are missing from the table: {missing}"
        raise ValueError(msg)
    core_names = set(core)
    by_name = {column["name"]: column for column in cols}
    ordered = [by_name[name] for name in core]
    ordered += [column for column in cols if column["name"] not in core_names]
    return ordered


ITEMS_COLUMNS = _core_first(ITEMS_COLUMNS, CORE_VISIBLE)
N_CORE = len(CORE_VISIBLE)

RAID_COLUMNS = [
    _c("RaidID", "V", "text", 9),
    _c("Type", "I", None, 12, list_validation("=lstRaidTypes")),
    _c("Title", "I", None, 32, wrap=True),
    _c("Detail", "I", None, 38, wrap=True),
    _c("RelatedID", "I", "text", 11, list_validation("=lstItemIDs")),
    _c("Owner", "I", None, 12, list_validation("=lstPeople")),
    _c("Status", "I", None, 12, list_validation("=lstRaidStatuses")),
    _c(
        "Prob",
        "I",
        "int",
        9,
        raid_rating_validation("Probability", "unlikely", "very likely"),
    ),
    _c(
        "Impact",
        "I",
        "int",
        9,
        raid_rating_validation("Impact", "low", "severe"),
    ),
    # Severity is the highest ascending tblSeverity band whose MinScore is at
    # or below Score. INDEX/MATCH stores a standard calculated-column formula.
    _c(
        "Severity",
        "F",
        "calc",
        10,
        formula='=IF([@Score]="","",'
        "INDEX(tblSeverity[Severity],"
        "MATCH([@Score],tblSeverity[MinScore],1)))",
    ),
    _c("Response", "I", None, 32, wrap=True),
    _c("NextReview", "I", "date", 14, DATE_OK),
    _c("Score", "F", "calcint", 6, formula='=IF(OR([@Prob]="",[@Impact]=""),"",[@Prob]*[@Impact])'),
    _c(
        "Scope",
        "F",
        "calc",
        9,
        formula='=IF([@RelatedID]="","",'
        "IFNA(INDEX(tblItems[Scope],"
        'MATCH([@RelatedID],tblItems[ID],0)),""))',
    ),
    _c("Raised", "V", "date", 12, DATE_OK),
    _c("Closed", "V", "date", 12, DATE_OK),
    _c("Updated", "V", "date", 12, DATE_OK),
]

RAID_CORE_VISIBLE = [
    "RaidID",
    "Type",
    "Title",
    "Detail",
    "RelatedID",
    "Owner",
    "Status",
    "Prob",
    "Impact",
    "Severity",
    "Response",
    "NextReview",
]
RAID_COLUMNS = _core_first(RAID_COLUMNS, RAID_CORE_VISIBLE)
N_RAID_CORE = len(RAID_CORE_VISIBLE)
