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

`Scope` is the nearest Level-1 ancestor. `WbsKey` provides deterministic parent-first ordering. Parent chains drive structural metadata such as ParentTitle, ParentLevel, ancestor path, Level, Scope, Children, indentation and outline groups. They do not supply operational values to another row.

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

Every operational field belongs to its own item. Start, Due, Status, Delivery Health, Owner, Priority, Latest Status and lifecycle dates are never inherited from an ancestor or descendant. A blank value remains blank in reporting views instead of being replaced by another hierarchy row's value.

`BlockedBy` can explicitly reference one or more item IDs for dependency blockers. `WaitingOn`, `BlockedRefsValid` and `IsBlocked` are derived from that relationship. `IsBlocked` is true when the item's own Delivery Health is Blocked or an explicitly referenced open dependency remains.

`Source` and `Source ID` form an optional paired identity for agent updates. They are system-managed, exact strings and are not part of the everyday input surface. Workbook IDs remain the relationship keys inside Excel; source identities make repeated external records portable across agent runs.

## Status roles and lifecycle dates

`tblStatuses` carries `IsActive`, `IsDone`, `IsCancelled` and `IsDeleted` flags. The labels can change while their behavior remains stable. Exactly one row has the deletion role; the shipped row is `Deleted`, with inactive, done and cancelled behavior.

- Picking a Type on a new item assigns its ID; the first entry on the row stamps Created and Updated.
- Every material edit refreshes Updated.
- Edits are never rejected or undone; a value that Config does not define skips its role stamp and is flagged by the invalid-input formatting.
- Entry into an active role sets InProgressSince when empty.
- Entry into a delivered done role sets DoneDate when empty.
- A cancelled role clears DoneDate.
- Selecting Blocked in Delivery Health sets BlockedSince; choosing another state clears it.
- Editing Latest Status sets LatestUpdateOn; clearing the narrative clears the stamp.
- An explicit agent `mark_deleted` operation changes only Status to the deletion-role row. The item, descendants and relationships remain present, and nothing cascades.

Health combines configured status roles, due dates, blocked duration and update freshness.

A Deleted row retains its workbook and source identities as visible history. A later `upsert` for that source identity creates a fresh row and workbook ID. Only one non-Deleted row may carry a given source identity.

## Schedule and key dates

- An item's own Start plus Due creates its scheduled interval.
- An item's own Due with a blank Start creates a key date, rendered as a diamond in Plan.
- Plan shows every populated item within the selected Scope and Depth, including rows with no dates or only Start.
- A missing direct date stays blank. Rows without both direct dates draw no interval bar, and rows without their own Due-only key date draw no key-date glyph.
- Descendant dates never widen, populate or otherwise change an ancestor's dates, timeline marks, overdue classification or reporting values.
- Plan filters by up to three Level-1 scopes, visible Depth and an optional date window.
- The automatic window covers the directly dated displayed work with one week of padding.
- The timeline is weekly and displays up to 52 columns.

## RAID model

RAID records contain Type, Title, Detail, RelatedID, Owner, Status, probability, impact, Response and NextReview.

- RelatedID is optional. When present it must match an Items ID and supplies Scope; a blank record remains unscoped without breaking views, export or migration.
- Owner is optional. A nonblank value must come from the Config people list.
- Only types with the Config `IsAlert` role require Probability and Impact. The shipped alert types are Risk and Issue; Assumption, Dependency and Decision are non-alert. Non-alert scoring cells are shown as not applicable, normal Probability or Impact entry is rejected, and pasted ratings are flagged red.
- Probability and Impact each use a whole-number scale from 1 to 5.
- Score is `Prob × Impact`, producing a range from 1 to 25.
- Severity is the highest configured band whose minimum score is satisfied.
- The shipped bands are Low 1-3, Medium 4-8, High 9-15 and Critical 16-25. A 5 × 5 record scores 25 and is therefore Critical.
- Scope is resolved from the item explicitly selected in RelatedID.
- `IsAlert` types can appear in Top RAID.
- `IsDecision` types can appear in Coming Up when NextReview is current or future.
- `IsClosed` statuses remove records from open views and stamp Closed.
- `IsDeleted` identifies the one explicit deletion status; it is also closed and retains the RAID row as visible history.
- Every open RAID row without NextReview receives an amber attention cue.

Top RAID requires an open alert type and a score at or above `cfgAlertSevScore`. The shipped value of 9 corresponds to High and Critical.

## Overview panels

| Panel | Inclusion and order |
|---|---|
| Executive Status Summary | Open items through `cfgExecutiveStatusMaxLevel`, plus directly blocked items at deeper levels; each row uses only that item’s own Delivery Health, Owner and Due, then sorts by health, level and WBS order |
| Top RAID | Open alert types meeting the configured severity threshold; score descending then review date |
| Coming Up | Key dates within Levels 2 through `cfgKeyDateMaxLevel`, plus open decisions, when their date is today or later; date ascending |
| Recent progress | Delivered items completed within `cfgReportDays`; level, priority and completion date order |

Each panel displays up to five records and shows the total when more records qualify.

## Configurable settings

Settings govern due-soon, blocked, stale and reporting windows; Overview level bounds; Coming Up urgency thresholds; RAID alert score; and item/RAID identifier prefixes and counters. `ExecutiveStatusMaxLevel` is the normal visible Overview depth from 1 through 6; directly blocked deeper rows still appear as themselves. Every displayed row uses its own operational values and never inherits an ancestor's or descendant's dates, status, health, owner, priority or narrative. Ordered thresholds are strictly increasing and validated at entry and after paste.
