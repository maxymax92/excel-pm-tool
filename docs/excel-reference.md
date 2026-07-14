# Excel implementation reference

## Runtime and release verification

- Microsoft Excel 365 desktop
- Macro-enabled Open XML workbook (`.xlsm`)
- Modern dynamic arrays and named LAMBDAs
- Native in-cell checkboxes for Boolean inputs
- A4 landscape printing at 100% worksheet zoom

The release test uses desktop Excel for calculation, repair detection and compatibility checks.

The current release test environment is Microsoft Excel 365 Personal for Mac, version 16.110. The workbook includes Windows-specific VBA branches for destination selection and path separators, but Windows is not yet part of the release test.

## Workbook architecture

The workbook separates three responsibilities:

```text
Data: Items, RAID, Config
Calc: named functions, helper arrays and bounded spills
View: Overview and Plan
```

Input tables own editable facts. Calc owns repeatable derivations. Views display bounded projections. VBA owns irreversible event facts such as creation and transition dates.

## Formula authoring

- `build/core/formulas.py` is the formula encoder.
- Formula source uses uppercase function tokens.
- Reusable business rules live in named `fn*` LAMBDAs from `build/spec/lambdas.py`.
- Dynamic arrays use `write_dynamic_array_formula()` and reserve their complete destination through `build/core/layout.py`.
- Each spill has a supported maximum from `build/spec/capacity.py`.
- Predictable empty results receive explicit empty-state output.
- Unexpected formula errors remain visible and fail QA.
- Volatile and deprecated lookup patterns are rejected by the formula linter.

Modern functions can require OOXML prefixes. The encoder manages `_xlfn.`, `_xlws.`, `_xlpm.` and spill-anchor storage consistently. New functions require an encoder classification and focused QA coverage.

`WbsKey` values are digit-only hierarchy paths. Blank Items rows use a high alphabetic text sentinel so Excel for Mac sorts unused capacity after every populated hierarchy row before `OrganiseItems` resizes the table.

## Tables and calculated columns

- `tblItems`, `tblRAID` and every Config list are Excel Tables.
- Table names and column names are stable interfaces used by formulas and VBA.
- Formula columns use one consistent calculated-column formula.
- Config list tables always contain at least one body row so structured references remain valid.
- Editable sheets remain unprotected so table expansion and outline controls work.

## Data validation

List validation points to workbook names over bounded plain-cell ranges. The underlying dynamic lists are calculated on `Calc`.

This pattern is used because the current desktop-Excel release environment reliably preserves named plain-range validation sources across open/save cycles. Validation rules use stop-style errors and clear input guidance. Matching conditional-format rules expose invalid pasted values.

## Conditional formatting

Rule order follows this precedence:

1. Invalid structure or value
2. Missing required lifecycle data
3. Semantic exceptions such as overdue, blocked or stale
4. Mutually exclusive status, delivery-health, severity or urgency states
5. Hierarchy and scan rhythm
6. Base formatting

Each semantic color has a text, date, glyph, border or weight cue. Rules use Config roles and ranks. Overview date comparisons read numeric mirrors in hidden gutter cells while displaying dates as text. Plan point rules, bar fills and border-only timeline rulers remain independent.

## Date semantics

- Excel serial dates are stored as numeric values.
- Visible operational dates use `dd mmm yyyy`.
- Overview emits date text to guarantee stable panel rendering.
- `TODAY()` is used for relative status and window logic.
- VBA writes Created, Updated, InProgressSince, DoneDate, BlockedSince, LatestUpdateOn, Raised and Closed.
- A Due-only item is a point event; an item with Start and Due is a scheduled interval.

## OOXML packaging

`build/core/package_style.py` applies the Office theme, window styling and DrawingML macro actions through a temporary package. Publication uses atomic replacement. Package parsing and rewriting use strict UTF-8 and preserve the original exception when cleanup also reports a problem.

The QA twin and macro release share formulas, layout and styling. The macro release additionally contains `xl/vbaProject.bin` and the two authorized DrawingML actions.

## Protection

- Overview, Plan and Calc are protected without a password.
- Plan unlocks Scope, Depth, From and To.
- Items, RAID and Config are unprotected operating surfaces.
- Calc is hidden.

## Calculation verification

openpyxl inspects formula text and workbook structure. LibreOffice provides a disposable first-pass recalculation. Excel opens, recalculates and saves disposable release copies, and its repair output is a release blocker.

Primary technical references are listed in [research/excel-2026/source-register.md](research/excel-2026/source-register.md).
