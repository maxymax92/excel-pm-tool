"""Example rows (identity-marked, with 'EXAMPLE — delete this row' copy).

Items ships 4 rows demonstrating the levelled hierarchy: a Project (L1) with a
Release key date (L2, Due only -> diamond), an Epic (L4) and a completed Task
(L5). Rows are listed in WBS order so the build-time outline levels match what
OrganiseItems would produce. Supporting tables ship 1-2 rows so every Overview
panel has a useful example. Dates are fixed literals (July 2026) so builds are
deterministic. Values are keyed by column name; F columns are computed, never
seeded.
"""

from datetime import date

EX_NOTE = "EXAMPLE — delete this row"

ITEMS_EXAMPLES = [
    {
        "ID": "I-1001",
        "Title": "Example Project",
        "Type": "Project",
        "Status": "In Progress",
        "Delivery Health": "At risk",
        "Priority": "P2",
        "Owner": "Max",
        "Start": date(2026, 6, 15),
        "Due": date(2026, 10, 30),
        "Latest Status": EX_NOTE,
        "Created": date(2026, 6, 1),
        "Updated": date(2026, 7, 12),
        "LatestUpdateOn": date(2026, 7, 12),
    },
    {
        "ID": "I-1002",
        "Title": "Example epic",
        "Type": "Epic",
        "Parent": "I-1001",
        "Status": "In Progress",
        "Delivery Health": "On track",
        "Priority": "P2",
        "Owner": "Max",
        "Start": date(2026, 6, 15),
        "Due": date(2026, 8, 14),
        "Latest Status": EX_NOTE,
        "Created": date(2026, 6, 1),
        "Updated": date(2026, 7, 12),
        "InProgressSince": date(2026, 6, 15),
        "LatestUpdateOn": date(2026, 7, 12),
    },
    {
        "ID": "I-1003",
        "Title": "Example completed deliverable",
        "Type": "Task",
        "Parent": "I-1002",
        "Status": "Done",
        "Priority": "P2",
        "Owner": "Max",
        "Start": date(2026, 6, 15),
        "Due": date(2026, 7, 3),
        "Latest Status": EX_NOTE,
        "Created": date(2026, 6, 1),
        "Updated": date(2026, 7, 3),
        "InProgressSince": date(2026, 6, 15),
        "DoneDate": date(2026, 7, 3),
        "LatestUpdateOn": date(2026, 7, 3),
    },
    {
        "ID": "I-1004",
        "Title": "Example release",
        "Type": "Release",
        "Parent": "I-1001",
        "Status": "Ready",
        "Priority": "P1",
        "Owner": "Max",
        "Due": date(2026, 7, 20),
        "Latest Status": EX_NOTE,
        "Created": date(2026, 6, 20),
        "Updated": date(2026, 7, 12),
        "LatestUpdateOn": date(2026, 7, 12),
    },
]

PEOPLE_EXAMPLES = [
    {"Person": "Max", "Role": "Technical PM", "Team": "Core"},
]

RAID_EXAMPLES = [
    {
        "RaidID": "R-001",
        "Type": "Risk",
        "Title": "Example risk — " + EX_NOTE,
        "Detail": "Key dependency may slip; mitigation owner assigned.",
        "RelatedID": "I-1001",
        "Owner": "Max",
        "Status": "Open",
        "Prob": 3,
        "Impact": 4,
        "Response": "Mitigate: weekly checkpoint with supplier",
        "NextReview": date(2026, 7, 20),
        "Raised": date(2026, 7, 1),
        "Updated": date(2026, 7, 6),
    },
    {
        "RaidID": "R-002",
        "Type": "Decision",
        "Title": "Approve example launch scope",
        "Detail": "Sponsor decision required before the release date.",
        "RelatedID": "I-1001",
        "Owner": "Max",
        "Status": "Open",
        "Response": "Confirm scope at the steering meeting",
        "NextReview": date(2026, 7, 15),
        "Raised": date(2026, 7, 6),
        "Updated": date(2026, 7, 6),
    },
]
