# Project-management domain model

## Operating model

The workbook is a hierarchy-led delivery tracker. It combines work, schedule, delivery health, narrative status and RAID information without duplicating records between views.

It is deliberately a small shared layer for delivery across teams, organisations, companies, regions or separate software systems. Teams can keep detailed work in their own tools. The workbook standardises the minimum common data needed to coordinate the whole delivery, brief senior stakeholders and create one status file when direct system access is unavailable.

`tblItems` is the canonical work register. `tblRAID` is the canonical risk and decision register. Config tables define the language and role behavior used by formulas and macros.

## Hierarchy

Each item has a `Type` and an optional `Parent` identifier. `tblTypes` maps every type to a level from 1 to 6.

| Level | Shipped types | Purpose |
|---:|---|---|
| 1 | Project, Product | Scope boundary |
| 2 | Release, Initiative | Major delivery outcome |
| 3 | Phase | Delivery stage |
| 4 | Team, Feature, Epic, Deliverable | Coherent work package |
| 5 | Task, Test Case, Story | Executable work |
| 6 | Sub Task, Bug | Atomic work |

`Scope` is the nearest Level-1 ancestor. `WbsKey` provides deterministic parent-first ordering. Parent chains also drive indentation, outline groups, schedule envelopes and scope reporting.

Valid hierarchy requires:

- a unique well-formed ID;
- a configured type;
- a parent for Levels 2 through 6;
- a parent at a lower numeric level;
- an acyclic ancestor chain within six levels.

## Work fields

The everyday Items surface is:

| Field | Meaning |
|---|---|
| ID | VBA-assigned identifier |
| Type | Configured hierarchy type |
| Title | Work name |
| Parent | Immediate parent ID |
| Priority | Config-ranked P0 to P4 |
| Start, Due | Schedule dates |
| Status | Configured workflow state |
| Delivery Health | Direct operating state: On track, At risk, Off track or Blocked |
| Latest Status | Current delivery narrative |
| Owner | Accountable person |

`BlockedBy` can reference one or more item IDs for dependency blockers. `WaitingOn`, `BlockedRefsValid` and `IsBlocked` are derived. `IsBlocked` is true when Delivery Health is Blocked or an open dependency remains. `EffStart` and `EffDue` widen a parent’s own dates to the full descendant envelope.

## Status roles and lifecycle dates

`tblStatuses` carries `IsActive`, `IsDone` and `IsCancelled` flags. The labels can change while their behavior remains stable.

- Entering Title on a new item assigns ID, Created and Updated.
- Every material edit refreshes Updated.
- Entry into an active role sets InProgressSince when empty.
- Entry into a delivered done role sets DoneDate when empty.
- A cancelled role clears DoneDate.
- Selecting Blocked in Delivery Health sets BlockedSince; choosing another state clears it.
- Editing Latest Status sets LatestUpdateOn; clearing the narrative clears the stamp.

Health combines configured status roles, due dates, blocked duration and update freshness.

## Schedule and key dates

- Start plus Due creates a scheduled interval.
- Due with a blank Start creates a key date.
- Parent intervals use the effective descendant envelope.
- Plan filters by Level-1 Scope, visible Depth and an optional date window.
- The automatic window covers the displayed schedule with one week of padding.
- The timeline is weekly and displays up to 52 columns.

## RAID model

RAID records contain Type, Title, Detail, RelatedID, Owner, Status, probability, impact, Response and NextReview.

- Probability and Impact each use a whole-number scale from 1 to 5.
- Score is `Prob × Impact`, producing a range from 1 to 25.
- Severity is the highest configured band whose minimum score is satisfied.
- The shipped bands are Low 1-3, Medium 4-8, High 9-15 and Critical 16-25. A 5 × 5 record scores 25 and is therefore Critical.
- Scope is inherited from the related item.
- `IsAlert` types can appear in Top RAID.
- `IsDecision` types can appear in Coming Up when NextReview is current or future.
- `IsClosed` statuses remove records from open views and stamp Closed.

Top RAID requires an open alert type and a score at or above `cfgAlertSevScore`. The shipped value of 9 corresponds to High and Critical.

## Overview panels

| Panel | Inclusion and order |
|---|---|
| Executive Status Summary | Open items from Level 1 through `cfgExecutiveStatusMaxLevel`; lowest Delivery Health across the item and its open descendants, then level and WBS order |
| Top RAID | Open alert types meeting the configured severity threshold; score descending then review date |
| Coming Up | Key dates within Levels 2 through `cfgKeyDateMaxLevel`, plus open decisions, when their date is today or later; date ascending |
| Recent progress | Delivered items completed within `cfgReportDays`; level, priority and completion date order |

Each panel displays up to five records and shows the total when more records qualify.

## Configurable settings

Settings govern due-soon, blocked, stale and reporting windows; Overview level bounds; Coming Up urgency thresholds; RAID alert score; and item/RAID identifier prefixes and counters. `ExecutiveStatusMaxLevel` is the visible Overview depth from 1 through 6. Each displayed row rolls up the lowest Delivery Health among itself and its open descendants. Ordered thresholds are strictly increasing and validated at entry and after paste.
