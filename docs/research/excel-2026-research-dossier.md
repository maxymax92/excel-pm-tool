# Excel 365 workbook research synthesis

## Purpose

This synthesis records the evidence supporting the workbook’s current architecture, interaction model, role in a wider delivery toolset and release controls. Product capabilities and library behavior are version-sensitive; the linked primary sources remain the authority.

## Findings applied to the workbook

### One record store, multiple views

Smartsheet and Airtable expose grid, schedule, card and summary experiences as projections over the same records. The workbook applies that pattern through `tblItems`, `tblRAID`, a Calc layer, and the Overview and Plan views.

### One shared view when the tools cannot be joined

Practitioners keep making the same distinction: teams need detailed task management, while the person responsible across projects needs a much smaller common view. That matters when teams, clients or partner organisations use different systems, instances, permissions or status language. The workbook keeps the common commitments, dates, health, narrative and RAID without forcing every contributing team into another tool.

### Bounded stakeholder surfaces

Operational dashboards work best when each component answers one question and preserves a route to the underlying records. Overview therefore uses four bounded record panels with explicit truncation counts and Items/RAID as the detail surface.

### Modern formulas as a calculation layer

Excel 365 dynamic arrays, `LET`, `LAMBDA`, `FILTER`, `SORTBY`, `HSTACK`, `VSTACK`, `MAP` and related functions support a formula-native calculation layer. Reusable rules live in named functions, and every spill has a registered capacity and an explicit empty state.

### VBA for event facts

Formulas recalculate current state; they cannot preserve the date of a transition. Workbook events therefore assign IDs and stamp lifecycle dates. Formulas derive every repeatable calculation from those facts.

### Desktop-safe validation

Desktop testing establishes named plain-cell ranges as the stable source for dynamic list validation in the current release environment. Calc produces the lists, workbook names expose bounded ranges, and input cells reference those names.

### Layered governance

The workbook combines clear prompts, stop-style validation, paste-safety conditional formatting, explicit VBA validation and release QA. This keeps routine entry simple while exposing structural corruption immediately.

### Deterministic artifact generation

XlsxWriter is the sole workbook author. openpyxl inspects structure, oletools verifies VBA, and desktop Excel recalculates and checks the final package. Temporary artifacts and atomic replacement keep failed builds away from release paths.

### Restrained visual hierarchy

A professional workbook benefits from semantic color, quiet surfaces, consistent typography, clear alignment, bounded components and hidden gridlines outside tables. The design system uses Aptos, navy hierarchy, one blue brand, small semantic fills and accessible glyphs.

## Research notes

- [Formula and file authoring](excel-2026/excel-formulas-and-authoring.md)
- [Automation and governance](excel-2026/automation-and-governance.md)
- [Dashboard and workbook craft](excel-2026/dashboard-and-workbook-craft.md)
- [Project-tool patterns](excel-2026/pm-tool-patterns.md)
- [Source register](excel-2026/source-register.md)
- [Visual reference catalog](excel-2026/visual-reference-catalog.md)
