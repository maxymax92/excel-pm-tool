# Formula and workbook authoring research

## Modern formula stack

The workbook uses Excel 365’s formula language as a bounded calculation engine:

- `LET` names intermediate values and avoids repeated work.
- `LAMBDA` creates reusable workbook functions.
- `FILTER`, `SORT`, `SORTBY` and `UNIQUE` create live projections.
- `TAKE`, `DROP`, `CHOOSECOLS`, `HSTACK` and `VSTACK` shape results.
- `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL` and `MAKEARRAY` support array-wise logic.
- `SEQUENCE` creates the Plan axis and bounded iteration sets.

Each calculation has one responsibility, a declared maximum output area and a deliberate empty state.

Primary references:

- [Microsoft: LAMBDA](https://support.microsoft.com/en-us/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67)
- [Microsoft: GROUPBY](https://support.microsoft.com/en-us/office/groupby-function-5e08ae8c-6800-4b72-b623-c41773611505)
- [Microsoft: PIVOTBY](https://support.microsoft.com/en-us/office/pivotby-function-de86516a-90ad-4ced-8522-3a25fac389cf)

## Formula encoding

Excel stores several modern functions with compatibility namespaces. XlsxWriter documents `_xlfn.`, `_xlws.`, `_xlpm.` and `ANCHORARRAY` requirements.

The workbook routes formulas through `build.core.formulas`:

- `F()` classifies and prefixes supported functions, rewrites supported spill anchors and validates formula tokens.
- `LAM()` writes named LAMBDAs with encoded parameters.
- formula source uses uppercase function names and simple relative spill anchors;
- new functions receive an explicit classification and QA case.

References:

- [XlsxWriter: formulas](https://xlsxwriter.readthedocs.io/working_with_formulas.html)
- [XlsxWriter: LAMBDA example](https://xlsxwriter.readthedocs.io/example_lambda.html)

## Dynamic list validation

The reliable target-platform pattern is:

1. Calculate a clean list on Calc.
2. Reserve a bounded plain-cell range.
3. Define a workbook name over that range.
4. Point data validation to the workbook name.

This keeps the list live while preserving validation through open/save cycles in the current desktop-Excel release environment.

## Event facts and formulas

Created, Updated, InProgressSince, DoneDate, BlockedSince, LatestUpdateOn, Raised and Closed represent events. VBA writes those values. Formulas calculate current health, elapsed time, severity, structural hierarchy metadata, explicit dependency results, RAID scope and reporting views. Operational values remain row-local: formulas do not replace a blank Start, Due, Status, Delivery Health, Owner, Priority, narrative or lifecycle value with one from an ancestor or descendant. Plan still projects every in-scope/depth item; direct Start plus Due produces an interval, direct Due with blank Start produces a key-date diamond, and missing dates remain blank with no bar.

## Writer and reader roles

XlsxWriter creates the workbook and its modern formulas. openpyxl reads package structure and formula text for QA. Desktop Excel calculates the release and confirms the package opens without repairs.

Reference:

- [openpyxl: formula handling](https://openpyxl.readthedocs.io/en/latest/simple_formulae.html)

## Formula acceptance checklist

1. Confirm availability in Excel 365 desktop and the current pinned release build.
2. Put reusable rules in a named function.
3. Route stored formula text through the encoder.
4. Register the complete spill area.
5. Bound the result to the shared capacity contract.
6. Provide a clear empty state.
7. Add structural and scenario assertions.
8. Recalculate the workbook in desktop Excel.
