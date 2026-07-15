# Workbook user guide

## Start here

Use this workbook when delivery spans teams, organisations, companies or regions and no single project, product or programme management system is available to everyone involved. Detailed team work can stay in its existing tools. Record the shared commitments, dates, status, risks and decisions needed to coordinate the delivery as a whole.

Open `PM_Workbook.xlsm` in Microsoft Excel 365 desktop and enable macros. The workbook opens maximized on Overview.

The normal workflow is:

1. Maintain people, workflow lists and settings in Config.
2. Add and update work in Items.
3. Add and update risks and decisions in RAID.
4. Use Overview for the stakeholder brief.
5. Use Plan for schedule review.
6. Export a Markdown snapshot for reporting or AI-assisted work when a live system connection is unavailable.

## Config

Config is arranged as parallel bands separated by narrow blank columns:

- Settings
- Statuses and their active/done/cancelled/deleted roles
- Types and hierarchy levels
- Priorities
- Teams
- RAID types and alert/decision roles
- RAID statuses and closed/deleted roles
- Severity bands
- Delivery Health order
- People
- Operating guidance

Keep list values unique and preserve the intended rank order. Red cells indicate invalid values or relationships that require correction.

## Items

Start a new row by picking a Type. The workbook assigns the next ID the moment the Type is chosen, applies that level's row height, indentation and title emphasis, and stamps Created and Updated from the first entry on the row. Add the Title, Status and the rest in any order you like.

For Levels 2 through 6, choose a Parent. Set Start and Due for scheduled work. Set Due alone for a key date. Set Delivery Health on any active row, use Latest Status for the current narrative and Owner for accountability.

Each row owns those operational values. A child or descendant never supplies Start, Due, Status, Delivery Health, Owner, Priority, Latest Status or lifecycle dates to its parent. Hierarchy still derives structural information such as Scope, ancestors, children, indentation and parent-first WBS grouping. BlockedBy remains an explicit dependency reference rather than inherited status.

Blocking has two routes:

- choose **Blocked** in Delivery Health for a direct blocker;
- reveal the advanced group and populate BlockedBy with item IDs separated by commas.

Click **Organise rows** after hierarchy changes. The action validates the table, sorts it in parent-first WBS order, applies indentation and font hierarchy, rebuilds the expandable row groups and confirms how many rows were organised.

The advanced group contains derivations, event stamps and the system-managed `Source` / `Source ID` pair. Open it to check how a value was derived or inspect lifecycle dates. Leave source identity cells to the agent data bridge; they keep repeated updates tied to the same external record.

## RAID

Choose a RAID Type. The workbook assigns the RaidID and stamps Raised as soon as the row has data. Enter Title, Detail and Status, then set NextReview; every open row without a review date is amber. RelatedID is optional: select an Items row when the RAID record should resolve Scope from that explicit relationship, or leave it blank for an unscoped RAID record. Owner is also optional, while any nonblank RelatedID or Owner must match its Config-backed list.

Config alert types require Probability and Impact. In the shipped setup these are Risk and Issue. Both rating cells accept whole numbers from 1 to 5 and show the scale when selected; Severity is red until the scoring pair is complete. Assumption, Dependency and Decision do not need scoring: their Probability, Impact and Severity cells are grey, and stop validation prevents normal entry in Probability or Impact. Pasted ratings are highlighted red and do not produce a Score. Score is `Probability × Impact`, so the range is 1 to 25. Severity uses the highest Config band whose MinScore is no greater than the score. The shipped bands are Low 1-3, Medium 4-8, High 9-15 and Critical 16-25; a 5 × 5 record scores 25, so its Severity is Critical.

Open alert types with scores at or above the configured threshold appear in Top RAID. Open decision types with a current or future NextReview date appear in Coming Up.

RAID's collapsed system group also carries `Source` and `Source ID`. The normal `Deleted` status closes a record while keeping it visible for review.

## Overview

Overview is a one-page status brief with four panels. Every panel is derived from Items, RAID and Config and shows up to five records.

Executive Status Summary normally includes open items through the configured maximum level, with directly blocked deeper items included as exceptions. Every row shows only that item's own Delivery Health, Owner and Due. A child’s On track, At risk, Off track, Blocked, owner or schedule never changes the values shown for an ancestor.

Coming Up contains key dates from Level 2 through the configured maximum, plus open decision reviews, when their date is today or later. It uses four Config-defined urgency bands, and the exact date remains visible in every band. Overdue dates stay visible as exceptions in Items and RAID rather than appearing in Coming Up.

Click **Export to Markdown** near the lower-left of the visible sheet and choose a destination. The workbook saves the four panels plus compact Items and RAID registers as `PM_Status_yyyy-mm-dd.md`. Fully blank capacity rows are ignored; a partially entered source row stops the export with its table row and missing field. Export creates only the UTF-8 Markdown file and does not create logs, sidecars, history files or persistent caches.

The Markdown file is a point-in-time handoff for a report or any tool that cannot connect to the underlying systems. It is not a live sync; export it again after the workbook changes.

## Plan

Use the controls at the top:

- Scope: all work or one Level-1 scope, plus two optional slots beside it to show up to three scopes together
- Depth: levels 1 through 6
- From: optional window start
- To: optional window end

Blank From and To use the displayed work’s own date range with a week of padding. The legend shares the top row with the title.

Plan shows every populated item within the selected Scope and Depth, including undated and Start-only rows. Its Start, Due, status and schedule marks always come from that item's own row. Blank dates remain blank. A row with both Start and Due draws an interval bar; a row with its own Due and a blank Start is a key date and draws `◆`; an undated or Start-only row draws no bar or key-date glyph. Dates or status on descendants never alter an ancestor's Plan row.

Plan marks states with both color and glyph:

- `✓` done
- `●` in progress
- `—` planned
- `!` overdue
- `×` cancelled
- `◆` key date
- vertical rule for the current week

The status rail reports invalid controls, capacity limits and visible work lacking complete schedule dates.

## Upgrades and agent updates

A newer workbook version never asks you to re-enter data. With the workbook closed in Excel, `python -m build.data migrate` re-renders it from the current source with your rows, Config lists and settings injected, keeps the replaced file in `dist/backups/` and a JSON snapshot of your data in `dist/snapshots/`, and prints exactly what changed. The command reports every skipped example row, defaulted setting and value the workbook flags in red - nothing is dropped or corrected silently.

For agent-assisted updates, run `describe`, give the agent the source material, and have it produce a strict change set. `plan` shows the exact creates, updates, explicit Deleted transitions, no-ops, warnings and field diffs without writing the workbook. Review the plan token and expiry, then run `apply` with that exact token. Source omission never deletes anything. `mark_deleted` changes only the named row's status, and a later reappearance creates a fresh workbook row while retaining the Deleted row as history. Follow the [agent data bridge guide](agent-data-bridge.md) for the complete workflow and paste-ready prompt.

## Error handling

Red input or Config cells indicate invalid pasted or entered data. Correct the highlighted value before relying on derived views. Items and RAID never reject or undo an edit: a value the workbook does not recognise simply skips its lifecycle stamp and stays highlighted until you fix it. Organise rows and the Markdown export still check the data they act on and report exactly what needs completing. An Excel repair report, formula error or failed macro is a release issue.
