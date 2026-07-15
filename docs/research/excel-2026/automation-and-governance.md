# Automation and governance research

## Automation boundary

The workbook uses formulas for derivable state and VBA for workbook-local events:

```text
edit -> validate complete target -> assign identifier -> stamp event facts -> recalculate views
```

This aligns with Microsoft’s platform distinction: VBA supports embedded desktop workbook events, while Office Scripts run through explicit script execution or Power Automate flows.

Reference:

- [Microsoft Learn: Office Scripts and VBA](https://learn.microsoft.com/en-us/office/dev/scripts/resources/vba-differences)

## Transition semantics

Lifecycle dates represent entry into a state:

- an active-role status sets InProgressSince when empty;
- a delivered done-role status sets DoneDate when empty;
- a cancelled role clears DoneDate;
- selecting Blocked in Delivery Health sets BlockedSince and choosing another state clears it;
- Latest Status edits set LatestUpdateOn and clearing the narrative removes the stamp;
- a closed RAID role sets Closed and reopening clears it.

The handler validates the complete edit target before writing any identifier or stamp. Multi-cell paste is therefore one accepted or rejected operation.

## Identifier generation

Item and RAID identifiers use Config prefixes and counters. The generator validates prefix shape, counter bounds, existing identifier shape, duplicates and collisions. It increments the counter only after assigning a unique identifier.

## Application-state safety

Event handlers capture the current event state, disable recursive events for the controlled update, and restore state in a single exit path. Errors retain the original failure and include any state-restoration diagnostic.

User-facing macro actions validate all required tables, columns, roles and ranges before mutation. They operate on `ThisWorkbook` and named Excel objects.

## Mac file access and unattended QA

Office for Mac is sandboxed. Microsoft's `GrantAccessToMultipleFiles` API deliberately presents a user permission prompt; it can consolidate prompts and remember access to the approved files, but it cannot silently approve a new random path. The automated release harness therefore does not use it as an unattended-access bypass.

Every workbook that desktop Excel opens for build, recalculation, semantic preservation, scenario, performance or VBA-refresh work is an exact disposable copy. The preferred root is `~/Library/Containers/com.microsoft.Excel/Data/Documents/PMWorkbookAutomation`; Microsoft documents the enclosing Excel `Data/Documents` container for Mac add-in sideloading, and Excel can open this private location without an external-file grant prompt. If macOS denies the host process permission to create a child there, automation emits a warning and creates one unique mode-700 directory directly beneath `/private/tmp`, which Excel can also open through the existing automation route. Each operation owns one isolated directory, and its cleanup path deletes that directory. Other private-workspace errors still halt immediately; if fallback creation also fails, the diagnostic preserves the private failure and reports the fallback failure.

The AppleScript adapters call Excel's `open workbook` object-model command with an HFS path derived from the absolute POSIX path. They do not use Finder, `open -a`, accessibility clicks or restored-session state.

References:

- [Microsoft Learn: Request access to multiple files](https://learn.microsoft.com/en-us/office/vba/office-mac/grantaccesstomultiplefiles)
- [Microsoft Learn: Sideload Office Add-ins on Mac for testing](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)

## Markdown export

ExportMarkdown:

- uses Excel's standard `GetSaveAsFilename` picker for a destination path; the Mac branch avoids `Application.FileDialog`, which returns no folder-picker object in the current 16.110 release environment, and replaces any workbook suffix added by the native picker with `.md`;
- reads the four Overview panel ranges and the current Items and RAID tables;
- calculates only the Items and RAID tables plus the Calc and Overview sheets needed by the export;
- escapes Markdown-sensitive text and normalizes cell values;
- writes UTF-8 directly to the selected destination;
- produces one point-in-time file for senior reporting or any tool that cannot reach the underlying systems;
- keeps prior destination bytes in memory during replacement;
- restores the previous destination on failure;
- creates no sidecar, log, history or cache files;
- reports cancellation and errors explicitly.

## Governance layers

| Layer | Mechanism |
|---|---|
| Guidance | Labels, validation prompts and Config descriptions |
| Normal entry | Stop-style validation and native checkboxes |
| Paste safety | Conditional formatting over the full supported range |
| Event safety | VBA preflight validation and transactional updates |
| Build safety | Source hygiene, VBA verification and atomic publication |
| Release safety | Structural, design, scenario and desktop-Excel QA |

## Release evidence

A releasable workbook has:

- source and compiled VBA in agreement;
- a successful deterministic build;
- one release-only full rebuild in desktop Excel, followed by removal of the package's full-calculation-on-open flag;
- passing structural, design, empty-state, formula, abuse and Overview scenarios;
- a clean desktop Excel recalculate/save cycle;
- passing macro interaction tests;
- a representative populated workbook that renders correctly.

Project-tool workflow references are listed in [source-register.md](source-register.md).

## VBA source publication

`build.automation.refresh_vba` is the only supported source-to-binary route. It uses pyOpenVBA to replace both complete modules in a disposable package, rejects any non-VBA ZIP change, verifies the source-only project independently, then has desktop Excel regenerate the compiled caches before atomic publication. The build does not require programmatic access to the VBE object model.
