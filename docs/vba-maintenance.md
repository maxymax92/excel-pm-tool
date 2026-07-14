# VBA maintenance

## Source and binary

The authoritative source files are:

- `build/vba/PMTool.bas`
- `build/vba/ThisWorkbook.cls.txt`

The build embeds `build/vba/vbaProject.bin`. `build.qa.verify_vba` requires compiled-cache streams, validates the exact component inventory and independently extracts both modules for exact comparison after normalizing package metadata.

## Refresh procedure

Close every workbook in desktop Excel, then run:

```bash
.venv/bin/python -m build.automation.refresh_vba
```

The command:

1. validates both complete source files and the exact eight-component host project;
2. copies the current release to an isolated disposable workbook;
3. replaces `PMTool` and `ThisWorkbook` through pyOpenVBA;
4. rejects duplicate ZIP members and any change outside `xl/vbaProject.bin`;
5. independently verifies the source-only project against both repository files;
6. opens the exact disposable path through Excel's object model, performs a full rebuild, saves and closes it;
7. requires regenerated compiled-cache streams and exact extracted source;
8. atomically publishes `build/vba/vbaProject.bin`; and
9. removes the disposable working directory.

Any source, package, Excel, compilation, publication or cleanup error stops the refresh. The existing compiled binary remains in place unless every pre-publication check passes.

Run the independent gate after refresh:

```bash
.venv/bin/python -m build.qa.verify_vba
```

The release build starts only when both commands pass.

## Live smoke tests

Use a fresh disposable copy of the generated `PM_Workbook.xlsm`. Confirm its SHA-256 matches the release before opening it, perform the checks without saving test data back into the release, and verify:

### Visual inspection

- Overview, Plan, Items, RAID, Config and the hidden-state behavior of Calc match the current design system.
- Text is readable and aligned by data type; headers, filters, controls, conditional states, outlines and the Plan grid are not clipped or displaced.
- Empty capacity has no decorative ruled rows, unexplained lines or conditional fill.

### Workbook open

- The embedded VBA project compiles without an error.
- The workbook window maximizes.
- Overview is active.
- The workbook opens without a repair notice.

### Item events

- A populated new row receives a well-formed unique ID, Created and Updated.
- The Config counter advances after successful assignment.
- Active status entry stamps InProgressSince once.
- Delivered done entry stamps DoneDate once.
- Cancelled status clears DoneDate.
- Selecting Blocked in Delivery Health sets BlockedSince; choosing another state clears it.
- Latest Status edits stamp LatestUpdateOn and clearing the narrative clears the stamp.
- Invalid IDs, duplicate IDs, invalid roles, invalid dates and malformed bulk edits are rejected before any stamps change.

### RAID events

- A populated new row receives a well-formed unique RaidID, Raised and Updated.
- Closing stamps Closed and reopening clears it.
- Invalid IDs, roles and dates are rejected before any stamps change.

### OrganiseItems

- The action validates the complete Items hierarchy.
- Rows sort by WbsKey.
- Outline levels, indentation, font sizes and row heights match Levels 1 through 6.
- Expand and collapse controls work after the action completes.
- Fully blank table rows are ignored; partially entered rows fail with the exact row and missing field.
- Success reports the number of organised rows; failures report the stage, VBA error number, source and description.

### ExportMarkdown

- Excel's standard Save As picker selects the complete destination path without saving the workbook. Mac uses its native Export button and normalizes the workbook suffix that Excel may add to the selected path; Windows applies the Markdown file filter. The resulting filename is `PM_Status_yyyy-mm-dd.md`.
- The selected destination is the only filesystem location used by the export.
- The output is valid UTF-8.
- Calculation is limited to the Items and RAID tables plus Calc and Overview.
- Overview panels, Items and RAID render as valid Markdown tables.
- Fully blank table capacity is ignored; partially entered rows and duplicate identifiers fail with their source row or identifier.
- Pipes, backslashes, line breaks and control characters are escaped consistently.
- An existing destination is held in memory while the replacement is written.
- Export creates exactly one `.md` file and no sidecar, log, history or cache files.
- Cancel leaves existing files unchanged.
- An ordinary write failure restores the previous destination, or removes an incomplete new destination, and reports the VBA error number, source and description.

### Performance

- Opening completes without repeated readiness retries.
- Tab switching and ordinary cell selection do not invoke VBA work.
- One Items or RAID edit processes only the touched rows and the fields that drive the relevant stamps.
- The macro-enabled workbook passes `build.qa.performance`.

## Bound live evidence

The release orchestrator creates a JSON template after every automated gate passes:

```bash
.venv/bin/python -m build.qa.release prepare --evidence-template /tmp/pm-workbook-macro-evidence.json
```

The prepare phase returns status 2 because live interaction evidence is required. In the template:

- Keep the prefilled `source_sha256`, `release_workbook_sha256`, `schema` and `excel_version` unchanged.
- Set `tested_at` to a timezone-aware ISO-8601 time and identify the tester.
- Set every named check to the exact string `PASS` only after completing it.
- Run the Markdown export into a dedicated directory that is empty before the test. Record `directory_entries_before` as `[]` and `directory_entries_after` as the single generated filename.
- Record the absolute `.md` path and its lowercase SHA-256 digest. On macOS, `shasum -a 256 /absolute/path/to/file.md` prints that digest.
- Leave the export directory unchanged until final verification. The gate reads the file as strict UTF-8, checks every required section and table, verifies its digest, and rejects any unrecorded directory entry.

The evidence is valid for 24 hours and only for the exact source and release workbook digests. Store it outside the repository; it is a release input, not product state.

## Release gate

```bash
.venv/bin/python -m build.qa.release final --macro-evidence /tmp/pm-workbook-macro-evidence.json
```

The final phase validates the evidence before and after rerunning the complete non-mutating matrix. A zero exit status is required before publication or merge.
