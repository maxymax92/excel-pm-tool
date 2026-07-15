# Toolchain and release capabilities

## Authoring stack

| Component | Version | Responsibility |
|---|---:|---|
| Python | 3.12 | Package entry points, build orchestration and QA |
| XlsxWriter | 3.2.9 | Workbook creation, tables, formulas, validation, conditional formatting, checkboxes and VBA embedding |
| openpyxl | 3.1.5 | Structural and presentation inspection |
| pyOpenVBA | 3.0.1 | Source-controlled VBA replacement inside a disposable Office package |
| oletools | 0.60.2 | VBA source extraction and comparison |
| olefile | 0.47 | OLE container inspection |
| Ruff | 0.15.7 | Formatting and broad static-analysis gate |
| Microsoft Excel desktop | 365 Personal for Mac 16.110, current release test | Calculation, repair detection, VBA compilation and interaction testing |

Runtime and development dependencies are pinned in `pyproject.toml` and `uv.lock` and installed with `uv sync --frozen`.

The workbook runtime is Microsoft Excel 365 desktop. The listed Mac version is the current release-test environment; Windows is not yet part of that test.

## Repository capabilities

- `python -m build` builds, verifies and publishes every release destination together, with rollback on failure.
- `build.core.formulas` encodes modern Excel formulas and named LAMBDAs.
- `build.core.layout` prevents spill zones from overlapping.
- `build.core.package_style` applies the Office theme and DrawingML macro actions transactionally.
- `build.writers` composes the six worksheets from schema and design tokens.
- `build.qa.release` is the fail-fast release entry point. Its prepare phase builds and verifies the packages; its final phase requires fresh live-macro evidence bound to the exact source and workbook digests.
- `build.qa` verifies source hygiene, formulas, tables, names, validation, conditional formatting, package structure, VBA, calculated results and interactive latency.
- `build.scenarios.ship_demo` produces a representative populated workbook for visual and behavioral review.
- `python -m build.data export` and `python -m build.data migrate` capture authored rows, Config lists and settings into a bounded snapshot ring and re-render a populated workbook onto the current structure through the standard build, recalculation, semantic-preservation and rollback-capable publication path.
- `python -m build.data monday` imports a monday.com board into the Items hierarchy over the pinned GraphQL API with cursor pagination and bounded rate-limit retries; per-board identifier maps under `dist/monday/` keep re-imports idempotent.
- `build.automation.refresh_vba` replaces the complete source-controlled modules in a disposable workbook, proves package isolation, compiles in desktop Excel and atomically publishes the verified binary.
- `build.automation.workspace` keeps every Excel-facing disposable copy inside Excel's private macOS Documents container, so the automated build, recalculation, repair-detection, performance and VBA-refresh paths do not request access to a new external temporary file on every run.
- `build.automation` also contains direct Excel object-model recalculation, repair-detection and performance scripts; they do not launch workbooks through Finder or scripted UI clicks.
- `ExportMarkdown` produces one status file for senior reporting or any downstream tool when no live connection to the underlying systems is available.

## Verification layers

1. Source validation: pinned Ruff format and lint gates, Python compilation and repository hygiene.
2. VBA validation: exact module inventory, source-only replacement isolation, desktop-Excel compilation caches and independent compiled-source extraction after package-metadata normalization.
3. Build isolation: untouched authored baselines and same-suffix calculated copies in a temporary staging area.
4. Excel semantic preservation: exact worksheet and table formula text, defined-name targets, data-validation rules and coverage, effective per-cell conditional-format formulas, styles and precedence, package parts and embedded VBA. Conditional-format ranges may be split or coalesced only when every affected cell retains the same behavior.
5. Package validation: ZIP/OOXML integrity, theme, shapes, tables, names, formulas, validation and conditional formatting for `.xlsx` and `.xlsm`.
6. Formula scenarios: empty, representative and adversarial workbook states.
7. Performance testing: open, tab-switch, selection and edit latency against hard thresholds.
8. Live interaction evidence: workbook open, identifier and lifecycle events, RAID events, hierarchy organization, buttons and Markdown export including UTF-8 content and absence of sidecars.
9. Publication: all destinations and the root `.xlsx` cleanup commit together or roll back to their original bytes.

## Primary tool references

- Microsoft documents that [`Application.CalculateFullRebuild`](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculatefullrebuild) recalculates all open workbooks and rebuilds dependencies. The automation therefore requires Excel to have no other workbooks open.
- Microsoft documents that VBA performance caches are optional reader data in [`[MS-OVBA]`](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/c66b58a6-f8ba-4141-9382-0612abce9926). Source replacement invalidates those caches; desktop Excel must regenerate them before the binary is publishable.
- [pyOpenVBA](https://github.com/WilliamSmithEdward/pyopenvba) performs the package-local source replacement. Independent oletools extraction and desktop-Excel compilation remain mandatory release gates.
- Excel workbook persistence follows the documented [`Workbook.Save`](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.save), [`Workbook.SaveAs`](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.saveas) and [`Workbook.Close`](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.close) behavior.
- Ruff installation and configuration follow the official [installation](https://docs.astral.sh/ruff/installation/) and [configuration](https://docs.astral.sh/ruff/configuration/) references. The rule selection is explicit, preview-enabled and formatter-compatible.
- The pinned development group follows uv's [dependency-group model](https://docs.astral.sh/uv/concepts/projects/dependencies/#development-dependencies).
- Package and publication checks use Python's documented [`zipfile`](https://docs.python.org/3/library/zipfile.html), [`tempfile`](https://docs.python.org/3/library/tempfile.html), [`subprocess.run`](https://docs.python.org/3/library/subprocess.html#subprocess.run) and atomic [`os.replace`](https://docs.python.org/3/library/os.html#os.replace) APIs.

Version-sensitive Excel research is indexed under `docs/research/excel-2026/`.

## Static-analysis boundaries

The Ruff gate covers annotations, exceptions, complexity, error messages, boolean interfaces, refactoring, security, performance, documentation, commented code, task markers and Python correctness. `E501` is active alongside the formatter.

The selection intentionally omits:

- `CPY`, because this repository has no approved copyright owner or mandatory source-header text.
- Formatter-conflicting `COM`, `Q`, `ISC001`, `ISC002`, `W191`, `E111`, `E114`, `E117`, `D206` and `D300` rules. Pycodestyle alternatives `D203` and `D213` are excluded in favour of `D211` and `D212`.
- Framework-only rule families because the build is a Python workbook generator, not a web or notebook framework.
- `TID` and `SLF`, because import boundaries and internal access are governed by the repository architecture and focused QA rather than a generic linter convention.

There are no Ruff ignore lists, per-file exemptions or inline suppression markers.
