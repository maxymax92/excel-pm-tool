"""Config-sheet settings, workflow lists and operational guidance.

Type levels, status roles, RAID roles, severity bands and ID prefixes live in
editable Config tables and named cells. ``tblSeverity`` is ordered from low to
high, ``tblDeliveryHealth`` from best to worst, and status and priority rows
use display order. The final Delivery Health row is the direct-blocked state.
"""

# name, default value, comment (settings block, Config zone A)
SETTINGS = [
    ("cfgDueSoonDays", 5, "Due within N days -> amber"),
    ("cfgBlockedRedDays", 3, "Blocked >= N days -> red"),
    ("cfgStaleDays", 7, "No update for >= N days -> stale"),
    ("cfgReportDays", 14, "Overview look-back window (days)"),
    (
        "cfgExecutiveStatusMaxLevel",
        1,
        "Executive Status Summary shows open items at Levels 1..N; "
        "directly blocked deeper items also appear; each row keeps its own health",
    ),
    ("cfgKeyDateMaxLevel", 4, "Key dates show on Overview at levels 2..N"),
    ("cfgComingUrgentDays", 3, "Coming Up darkest band ends at N days"),
    ("cfgComingSoonDays", 7, "Coming Up strong band ends at N days"),
    ("cfgComingNearDays", 30, "Coming Up standard band ends at N days"),
    ("cfgComingHorizonDays", 60, "Coming Up light band ends at N days"),
    ("cfgAlertSevScore", 9, "RAID score >= N is a high-severity alert"),
    ("cfgItemIDPrefix", "I-", "Auto item ID prefix (VBA)"),
    ("cfgRaidIDPrefix", "R-", "Auto RAID ID prefix (VBA)"),
    ("cfgNextItemID", 1005, "Next item number (VBA-incremented)"),
    ("cfgNextRaidID", 3, "Next RAID number (VBA-incremented)"),
]

STATUSES = [
    # Status, IsActive, IsDone, IsCancelled, IsDeleted. Row order is display order.
    # IsActive stamps InProgressSince and drives active-work signals;
    # IsDone stamps DoneDate and ends attention; IsCancelled marks a done-flagged
    # status that must NOT count as delivered work.
    ("Backlog", False, False, False, False),
    ("Ready", False, False, False, False),
    ("In Progress", True, False, False, False),
    ("Review", True, False, False, False),
    ("Done", False, True, False, False),
    ("Cancelled", False, True, True, False),
    ("Deleted", False, True, True, True),
]

# (Type, Level). Levels drive hierarchy indentation, emphasis and the Plan /
# Items expand-collapse depth. Edit freely in the workbook: the taxonomy is a
# Config table, and every formula resolves a type's level by lookup.
TYPES = [
    ("Project", 1),
    ("Product", 1),
    ("Release", 2),
    ("Initiative", 2),
    ("Phase", 3),
    ("Team", 4),
    ("Feature", 4),
    ("Epic", 4),
    ("Deliverable", 4),
    ("Task", 5),
    ("Test Case", 5),
    ("Story", 5),
    ("Sub Task", 6),
    ("Bug", 6),
]

PRIORITIES = ["P0", "P1", "P2", "P3", "P4"]  # row order = rank (P0 highest)
TEAMS = ["Core"]

# (RaidType, IsAlert, IsDecision). IsAlert types require severity scoring and
# can feed Top RAID; IsDecision types feed Coming Up when their open rows have
# a future NextReview.
RAID_TYPES = [
    ("Risk", True, False),
    ("Assumption", False, False),
    ("Issue", True, False),
    ("Dependency", False, False),
    ("Decision", False, True),
]

# (RaidStatus, IsClosed, IsDeleted). IsClosed stamps Closed and drops the record from
# every open-RAID view.
RAID_STATUSES = [
    ("Open", False, False),
    ("Monitoring", False, False),
    ("Closed", True, False),
    ("Deleted", True, True),
]

# (Severity, MinScore) — ascending severity, ascending MinScore. Severity of a
# RAID row = the highest band whose MinScore <= Probability x Impact.
SEVERITY = [("Low", 1), ("Medium", 4), ("High", 9), ("Critical", 16)]

DELIVERY_HEALTH = ["On track", "At risk", "Off track", "Blocked"]

OVERVIEW_RULES = [
    "Executive Status Summary shows open items from Level 1 through "
    "ExecutiveStatusMaxLevel plus directly blocked deeper items; each row shows only its own "
    "Delivery Health.",
    "Plan Scope choices remain Level-1 Project/Product rows; changing the "
    "Executive Status Summary depth does not redefine Scope.",
    "Work: set Parent, Priority, Start, Status, Due and Owner.",
    "Delivery Health runs best-to-worst; its final Config option is the direct "
    "blocked state. BlockedBy remains available for dependency blockers.",
    "Key dates: give an item a Due date and no Start - it shows as a diamond "
    "on Plan and feeds the Overview outlook.",
    "RAID: RelatedID and Owner are optional; every open row needs NextReview. Alert types "
    "need Probability and Impact; non-alert rating cells reject entry.",
    "Top RAID shows open alert types only when Score is at or above "
    "AlertSevScore (High/Critical in the shipped bands).",
    "Coming Up merges future Due-only key dates with open decisions that have "
    "a future NextReview date; its four urgency thresholds come from Settings.",
    "Keep each item's Latest Status fresh - active items flag amber after "
    "StaleDays, and scope freshness is derived automatically.",
]
