"""Named LAMBDA library — 12 functions.

Each entry: (name, params, body, extra_vars).
Bodies are written in clean Excel syntax; formulas.encode_lambda() handles all
_xlfn./_xlpm. encoding. All variable names are lowercase by convention
(table columns are TitleCase, so substitution can't collide).
"""

LAMBDAS = [
    # Blank and missing target values normalize to an empty string for
    # ancestor-chain and lookup consumers.
    (
        "fnItemLookup",
        ("id", "col"),
        'LET(v,IFNA(XLOOKUP(id,tblItems[ID],col),""),IF(v=0,"",v))',
        ("v",),
    ),
    # A type's configured level resolves through tblTypes; blank or unknown
    # values resolve to zero.
    (
        "fnTypeLevel",
        ("typ",),
        'IF(typ="",0,LET(v,IFNA(XLOOKUP(typ,tblTypes[Type],tblTypes[Level]),0),N(v)))',
        ("v",),
    ),
    # Parent-chain walk to the nearest ancestor or self at a configured level.
    # The eight-step bound resolves invalid chains to an empty result; callers
    # seed the required recursion depth with zero.
    (
        "fnAncestorAtLevel",
        ("id", "lvl", "d"),
        'IF(OR(id="",d>8),"",'
        "IF(fnTypeLevel(fnItemLookup(id,tblItems[Type]))=lvl,id,"
        "fnAncestorAtLevel(fnItemLookup(id,tblItems[Parent]),lvl,d+1)))",
        (),
    ),
    # Sibling ordering date for the WBS key: effective start, else effective
    # due, else a far-future sentinel so undated siblings sort last.
    (
        "fnSortDate",
        ("id",),
        "LET(s,fnItemLookup(id,tblItems[EffStart]),"
        "du,fnItemLookup(id,tblItems[EffDue]),"
        'IF(s<>"",s,IF(du<>"",du,99999999)))',
        ("s", "du"),
    ),
    # Sortable hierarchy path: parent's key and this item's padded segment
    # (date serial 8 digits + table position 5 digits). A plain ascending sort
    # of this key IS the WBS order: parents first, children grouped under
    # them, siblings by date then entry order. Depth-capped like the walkers.
    (
        "fnWbsKey",
        ("id", "d"),
        'IF(OR(id="",d>8),"",'
        "fnWbsKey(fnItemLookup(id,tblItems[Parent]),d+1)&"
        'TEXT(fnSortDate(id),"00000000")&'
        'TEXT(IFNA(XMATCH(id,tblItems[ID]),0),"00000"))',
        (),
    ),
    ("fnBizDays", ("d1", "d2"), 'IF(OR(d1="",d2=""),"",NETWORKDAYS(d1,d2))', ()),
    (
        "fnHealthRAG",
        ("status", "due", "blocked", "bsince", "updated"),
        'IF(fnIsDone(status),"\u2013",'
        'IF(OR(AND(due<>"",due<TODAY()),'
        'AND(blocked=TRUE,bsince<>"",TODAY()-bsince>=cfgBlockedRedDays)),"R",'
        'IF(OR(blocked=TRUE,AND(due<>"",due-TODAY()<=cfgDueSoonDays),'
        'AND(updated<>"",TODAY()-updated>=cfgStaleDays)),"A","G")))',
        (),
    ),
    # CSV of predecessor IDs becomes a comma list of outstanding references.
    # Unknown identifiers remain visible as outstanding references.
    (
        "fnDepOpen",
        ("csv",),
        'IF(csv="","",LET(parts,MAP(TEXTSPLIT(csv,",",,TRUE),'
        'LAMBDA(p,LET(pid,TRIM(p),IF(pid="","",'
        'IF(fnIsDone(fnItemLookup(pid,tblItems[Status])),"",pid))))),'
        'TEXTJOIN(", ",TRUE,parts)))',
        ("parts", "p", "pid"),
    ),
    # One scalar validity flag for the BlockedBy CSV. TEXTSPLIT runs inside the
    # named LAMBDA and returns a scalar result to the table column.
    (
        "fnRefsValid",
        ("csv", "selfid"),
        'IF(csv="",TRUE,LET(parts,TRIM(TEXTSPLIT(csv,",",,TRUE)),'
        'AND(parts<>"",parts<>selfid,ISNUMBER(XMATCH(parts,tblItems[ID])))))',
        ("parts",),
    ),
    ("fnIsDone", ("s",), "ISNUMBER(XMATCH(s,lstDoneStatus))", ()),
    ("fnIsActive", ("s",), "ISNUMBER(XMATCH(s,lstActiveStatus))", ()),
    ("fnIsCancelled", ("s",), "ISNUMBER(XMATCH(s,lstCancelledStatus))", ()),
]
