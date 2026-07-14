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
- Statuses and their active/done/cancelled roles
- Types and hierarchy levels
- Priorities
- Teams
- RAID types and alert/decision roles
- RAID statuses and closed role
- Severity bands
- Delivery Health order
- People
- Operating guidance

Keep list values unique and preserve the intended rank order. Red cells indicate invalid values or relationships that require correction.

## Items

Start a new row by entering Title. The workbook assigns the next ID, Created and Updated dates. Then complete Type and Status.

For Levels 2 through 6, choose a Parent. Set Start and Due for scheduled work. Set Due alone for a key date. Set Delivery Health on any active row, use Latest Status for the current narrative and Owner for accountability.

Blocking has two routes:

- choose **Blocked** in Delivery Health for a direct blocker;
- reveal the advanced group and populate BlockedBy with item IDs separated by commas.

Click **Organise rows** after hierarchy changes. The action validates the table, sorts it in parent-first WBS order, applies indentation and font hierarchy, rebuilds the expandable row groups and confirms how many rows were organised.

The advanced group contains derivations and event stamps. Open it to check how a value was derived or inspect lifecycle dates.

## RAID

Choose a RAID Type, enter Title and Detail, link RelatedID, set Owner and Status, then complete Probability and Impact. Both cells accept whole numbers from 1 to 5 and show the scale when selected.

Score is `Probability × Impact`, so the range is 1 to 25. Severity uses the highest Config band whose MinScore is no greater than the score. The shipped bands are Low 1-3, Medium 4-8, High 9-15 and Critical 16-25; a 5 × 5 record scores 25, so its Severity is Critical.

Open alert types with scores at or above the configured threshold appear in Top RAID. Open decision types with a current or future NextReview date appear in Coming Up.

## Overview

Overview is a one-page status brief with four panels. Every panel is derived from Items, RAID and Config and shows up to five records.

Coming Up contains key dates from Level 2 through the configured maximum, plus open decision reviews, when their date is today or later. It uses four Config-defined urgency bands, and the exact date remains visible in every band. Overdue dates stay visible as exceptions in Items and RAID rather than appearing in Coming Up.

Click **Export to Markdown** near the lower-left of the visible sheet and choose a destination. The workbook saves the four panels plus compact Items and RAID registers as `PM_Status_yyyy-mm-dd.md`. Fully blank capacity rows are ignored; a partially entered source row stops the export with its table row and missing field. Export creates only the UTF-8 Markdown file and does not create logs, sidecars, history files or persistent caches.

The Markdown file is a point-in-time handoff for a report or any tool that cannot connect to the underlying systems. It is not a live sync; export it again after the workbook changes.

## Plan

Use the four controls at the top:

- Scope: all work or one Level-1 scope
- Depth: levels 1 through 6
- From: optional window start
- To: optional window end

Blank From and To use the displayed work’s own date range with a week of padding. The legend shares the top row with the title.

Plan marks states with both color and glyph:

- `✓` done
- `●` in progress
- `—` planned
- `!` overdue
- `×` cancelled
- `◆` key date
- vertical rule for the current week

The status rail reports invalid controls, capacity limits and work lacking usable schedule dates.

## Error handling

Red input or Config cells indicate invalid pasted or entered data. Correct the highlighted value before relying on derived views. VBA validation errors appear as explicit messages and prevent lifecycle stamping until the value is corrected. An Excel repair report, formula error or failed macro is a release issue.
