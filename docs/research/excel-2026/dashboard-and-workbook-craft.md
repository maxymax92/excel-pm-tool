# Dashboard and workbook craft research

## Information hierarchy

Effective spreadsheet interfaces use a small number of consistent layers:

1. title and purpose;
2. controls or section labels;
3. column headers;
4. records and exceptions;
5. supporting metadata.

Overview applies this as four adjacent record panels. Plan applies it as one title/legend rail, a compact control block, a frozen identity table and a synchronized schedule grid.

## Bounded panels

A panel answers one stakeholder question and preserves the underlying evidence. The current four questions are:

- Which scopes or configured hierarchy levels require attention?
- Which open RAID records carry the highest severity?
- Which key dates and decisions are approaching?
- What work completed recently?

Each panel has a fixed five-record body, an explicit empty state and a count when qualifying records exceed the visible limit.

Overview is a briefing surface, not a copy of every team system. It keeps the common status, exceptions and dates that a senior stakeholder needs, then leaves detailed operating data in Items, RAID or the contributing team's own tool.

## Schedule construction

The Plan follows a proven Gantt pattern:

- identity columns remain frozen on the left;
- one weekly date axis drives every timeline cell;
- every populated item within the selected Scope and Depth remains visible, including undated and Start-only rows;
- each row's own Start and Due define its interval overlap;
- a row with its own Due and blank Start renders as a key-date point;
- blank dates remain blank, and undated or Start-only rows render no bar or point;
- hierarchy rows never borrow descendant dates, status or other operational values;
- conditional formatting applies semantic state;
- the current week and month changes use independent border rules.

The exact state glyph repeats in each filled week, which improves scanning and preserves meaning in grayscale.

## Table craft

- Hide worksheet gridlines and create structure only where data exists.
- Use a strong header and quiet row dividers.
- Size columns by role rather than by incidental sample text.
- Top-align wrapped narrative.
- Keep numeric and date values right aligned.
- Reserve yellow and red for meaningful exceptions.
- Use near-white blue for editable surfaces and neutral gray for derived fields.
- Freeze identifiers and headers at the working edge.

## Calculation discipline

The Formula Modeling World Cup and Excel World Championship demonstrate the value of transparent, deterministic formulas and quickly verifiable outputs. The workbook applies that discipline through named functions, bounded spills, explicit errors and scenario-based QA.

References:

- [Financial Modeling World Cup](https://fmworldcup.com/)
- [Microsoft Excel World Championship rules](https://excel-esports.com/rules/)
- [Microsoft MVP Blog: Excel World Championship](https://techcommunity.microsoft.com/blog/mvp-blog/from-espn-to-the-spreadsheet-arena-how-excel-mvps-powered-the-microsoft-excel-wo/4497642)

## Workbook-specific application

- Aptos and a restrained Office-native palette create continuity with current Excel.
- The Overview print area contains only the four panels.
- The export action sits below the report surface and within the initial maximized viewport.
- Plan dates are horizontal and column-sized to display their day number.
- Items and RAID expose compact operating columns and one advanced outline group.
- Config is a horizontal control library with consistent gutters.

Selected visual references are cataloged in [visual-reference-catalog.md](visual-reference-catalog.md).
