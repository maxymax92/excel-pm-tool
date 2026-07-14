# Workbook design system

`build/core/design.py` is the implementation authority for visual tokens. Worksheet writers consume semantic colors, typography, alignment, hierarchy and size roles from that module.

## Design principles

1. Information hierarchy leads the page.
2. Editable, derived and exception states are visually distinct.
3. Semantic color is compact and paired with a non-color cue.
4. Alignment follows data type and interaction role.
5. Tables use quiet horizontal rhythm and clear headers.
6. Views are bounded, scannable and print-ready.
7. Repeated components share one token and geometry system.
8. Invalid data remains visible until corrected.

## Color

### Foundation

| Token | Value | Role |
|---|---:|---|
| `white` | `#FFFFFF` | Primary surface |
| `canvas` | `#F5F7FA` | Alternate row and background surface |
| `surface_subtle` | `#EEF2F6` | Quiet hierarchy and cancelled state |
| `surface_editable` | `#F8FBFE` | Editable cell surface |
| `surface_derived` | `#F1F4F7` | Calculated and system surface |
| `ink` | `#172B4D` | Strong header and title color |
| `text` | `#24364B` | Body text |
| `text_secondary` | `#4A5E73` | Supporting text |
| `text_muted` | `#5E7184` | Metadata and guidance |
| `border` | `#C8D2DE` | Quiet dividers |
| `border_strong` | `#71869D` | Input boundaries |

### Brand and meaning

| Token | Value | Role |
|---|---:|---|
| `brand` | `#0F6CBD` | Primary interaction and active schedule |
| `brand_dark` | `#0B4A6F` | Most urgent upcoming date |
| `brand_tint` | `#E8F3FC` | Informational and hierarchy tint |
| `teal` | `#0E6655` | Completed schedule state |
| `slate` | `#66788A` | Planned schedule state |
| `success_bg` / `success_fg` | `#E6F4EA` / `#137333` | Positive semantic pair |
| `warning_bg` / `warning_fg` | `#FFF4CE` / `#7A4A00` | Attention semantic pair |
| `danger_bg` / `danger_fg` | `#FDE7E9` / `#A4262C` | Error semantic pair |
| `danger_strong` | `#C42B1C` | Overdue schedule and urgent underline |
| `today` | `#5B6573` | Current-week rule |

Example rows use the blue informational pair. Yellow is reserved for attention states.

## Typography

| Role | Typeface | Size | Weight |
|---|---|---:|---|
| Page title | Aptos Display | 18 pt | Bold |
| Section | Aptos | 11 pt | Bold |
| Body | Aptos | 10 pt | Regular |
| Table header | Aptos | 9 pt | Bold |
| Caption and metadata | Aptos | 9 pt | Contextual |

Hierarchy titles use a stepped ramp:

| Level | Size | Weight | Indent | Row height |
|---:|---:|---|---:|---:|
| 1 | 12 pt | Bold | 0 | 30 pt |
| 2 | 11 pt | Bold | 1 | 28 pt |
| 3 | 10.5 pt | Bold | 2 | 26 pt |
| 4 | 10 pt | Regular | 3 | 24 pt |
| 5 | 10 pt | Regular | 4 | 24 pt |
| 6 | 10 pt | Regular | 5 | 24 pt |

## Alignment

| Content | Horizontal | Vertical |
|---|---|---|
| Short text and identifiers | Left | Center |
| Wrapped narrative | Left | Top |
| Numbers | Right | Center |
| Dates | Right | Center |
| Panel narrative | Left | Top |
| Panel dates | Right | Top |
| Checkboxes, controls and axis labels | Center | Center |
| Metadata | Right | Center |

Headers follow the alignment of their content when comparison benefits, while compact table headers remain left anchored for fast column scanning.

## Spacing and borders

- Core spacing follows 4, 8, 12, 16, 24 and 32 units.
- Row roles are 24 pt for compact data, 34 pt for table headers, 44 pt for wrapped RAID rows and 48 pt for Overview panel rows.
- Table bodies use a bottom divider in `border`.
- Input cells use `surface_editable` with `border_strong` where an explicit control boundary is required.
- Calculated cells use `surface_derived` and secondary text.
- Blank worksheet space has gridlines hidden and carries no ruled table formatting.
- Panel and Config gutters are plain white spacing columns.

## Component patterns

### Title rail

Editable and system sheets use a 32 pt first row. The title occupies the left segment and current operational information occupies the remainder. Items places the **Organise rows** action at J1 in macro builds.

### Table header

Table headers use `ink` fill, white 9 pt bold text and native filter controls. Widths reserve sufficient space for the label and filter affordance.

### Editable table body

Editable cells use a near-white blue surface. Narrative fields wrap and align to the top. Calculated and event-stamped fields sit in one collapsed group to the right of each table’s core columns. Invalid pasted values use the danger pair and a strong border.

### Overview panel

Each panel has:

- a 26 pt navy title row;
- a 30 pt pale-blue header row;
- five 48 pt body rows;
- a truncation count in the title row when more than five records qualify;
- explicit empty-state text.

### Macro action

Macro actions are flat DrawingML shapes with brand fill, white text and an accessible description. The stable description key attaches the macro name.

## Sheet specifications

### Overview

| Panel | Columns | Headers |
|---|---|---|
| Executive Status Summary | A:D | Item - Top Level + Lowest Health, Delivery Health, Owner, Due |
| Top RAID | F:K | Type, Description, Severity, Owner, Next review, Latest Status |
| Coming up | M:O | Milestones / Decisions / Deadlines, Date, Scope |
| Recent progress | Q:U | Completed work, Type, Owner, Completed, Scope |

E, L and P are 3-unit gutters and hold protected numeric date mirrors using the `;;;` number format. The visible date remains text. Freeze panes at F3 keep the Scopes panel and both header rows visible. Print area is A1:U7. **Export to Markdown** is at A32, within the opening viewport and below the print area.

### Plan

- Title and legend: row 1
- Scope: B2, with optional additional scopes in C2 and C3
- Depth: B3
- From: E2
- To: E3
- Identity header: A5:E5
- Identity data: A6:E2005
- Weekly axis: F5:BE5
- Schedule grid: F6:BE2005
- Hidden helpers: BG:BI
- Freeze panes: F6
- Print area: A1:BE2005

Legend states are Done, In progress, Planned, Overdue, Cancelled, Key date and Today. Schedule bars use a full-row glyph in every filled week, inset by white top and bottom borders so adjacent rows read as separate bars. Month boundaries and the current week use border-only rules so bar fills remain intact.

### Items

- Row 1: title and action rail
- Row 2: `tblItems` header
- Rows 3 onward: table body
- A:K core surface: ID, Type, Title, Parent, Priority, Start, Status, Due, Delivery Health, Latest Status, Owner
- L:AL: collapsed advanced, calculated and stamped fields
- Freeze panes: D3

Hierarchy styling applies to ID and Title. OrganiseItems applies the full font-size, indent, row-height and outline treatment.

Conditional states on the core surface follow one contract:

- Red marks invalid entered data or a missing universally required value.
- Amber fill marks a stale or missing Latest Status on active work.
- Amber borders mark missing active-work inputs: Owner, Due and Delivery Health on any active row.
- A blank Parent is neutral. A nonblank Parent is red only when it is unknown, self-referential, circular or not above the child in the configured hierarchy.
- Delivery Health uses configured row order: the first state is green, the second amber and every later state red. The final state is the direct blocked role.

### RAID

- Row 1: title rail
- Row 2: `tblRAID` header
- A:L core surface: RaidID, Type, Title, Detail, RelatedID, Owner, Status, Prob, Impact, Severity, Response, NextReview
- M:Q: collapsed calculated and stamped fields
- Freeze panes: C3

Wrapped records use 44 pt rows with top-aligned narrative.

Probability and Impact accept whole numbers from 1 to 5. Score is `Probability × Impact`, giving a range of 1 to 25. Severity is the highest configured band whose MinScore is no greater than Score.

### Config

All bands start at row 3:

| Band | Columns |
|---|---|
| Settings | A:C |
| Statuses | E:H |
| Types | J:K |
| Priorities | M |
| Teams | O |
| RAID types | Q:S |
| RAID statuses | U:V |
| Severity | X:Y |
| Delivery Health | AA |
| People | AC:AE |
| Guidance | AG:AL |

D, I, L, N, P, T, W, Z, AB and AF are two-unit gutters. Config list values use editable surfaces, native checkboxes for Boolean roles and red paste-safety rules for invalid values.

### Calc

Calc uses derived/system formatting, protected cells and registered bounded spill zones. The sheet is hidden in the release.

## Semantic states

| State | Treatment | Additional cue |
|---|---|---|
| Valid positive | Success pair | Status or delivery-health text |
| Attention | Warning pair | Exact date, label or border |
| Invalid or overdue | Danger pair | Exact value, error text or `!` |
| Active schedule | Brand fill | `●` in each filled week |
| Planned schedule | Slate fill | `—` in each filled week |
| Done schedule | Teal fill | `✓` in each filled week |
| Cancelled schedule | Subtle neutral fill | `×` in each filled week |
| Key date | Point color | `◆` |
| Current week | Neutral rule | `│ Today` legend |

Coming Up urgency uses four ordered ranges:

- through `cfgComingUrgentDays`: dark brand, white bold text and danger underline;
- through `cfgComingSoonDays`: brand, white bold text;
- through `cfgComingNearDays`: information pair with bold text;
- through `cfgComingHorizonDays`: soft information pair.

## Interaction and protection

- Overview and Plan are protected view surfaces.
- Plan’s four controls are unlocked and use stop-style validation.
- Items, RAID and Config support table growth and hierarchy controls.
- Checkboxes use native 365 cells.
- Input prompts explain expected values at the point of entry.
- Conditional formatting exposes invalid data introduced through paste.

## Accessibility

- Normal text/background pairs meet the 4.5:1 contrast target.
- Color states include labels, glyphs, dates, borders or font weight.
- Text remains at 9 pt or larger.
- Wrapped narrative begins at the top-left of its cell.
- Visible dates render as readable text or fixed-width date formats.
- Shapes have descriptive alternative text.
- Sheets open at 100% zoom.

## Design QA

`build.qa.design` verifies theme fonts, explicit colors, layout geometry, row roles, alignments, macro action shapes, conditional-format coverage and protected view surfaces. Structural and live-Excel QA verify the corresponding workbook behavior.
