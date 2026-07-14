# Excel Project Management Workbook

This repository builds one lean, macro-enabled Excel 365 project tracker for when there is no single project, product or programme management tool that everyone involved can use. The generated release is [`PM_Workbook.xlsm`](PM_Workbook.xlsm).

That might be because the work spans teams, organisations, companies or regions, with each group using a different tool - or a different instance of the same one. `Overview` gives senior leaders one clean report without asking them to log into any of them. `Export to Markdown` gives you one status file to hand to any tool that cannot connect to those systems either.

Enter work in `Items`, record risks and decisions in `RAID`, and manage the workbook rules in `Config`. `Overview`, `Plan` and the protected `Calc` layer are derived from that source data.

## See it in action

The mock-ups use the populated demo workbook, so the hierarchy, dates and exceptions show how the finished system behaves.

<p align="center">
  <a href="docs/assets/readme/plan.png">
    <img src="docs/assets/readme/plan.png" alt="Plan view showing a six-level project schedule, status legend and weekly Gantt timeline" width="100%">
  </a>
</p>

<table>
  <tr>
    <td width="50%" valign="top">
      <a href="docs/assets/readme/overview.png">
        <img src="docs/assets/readme/overview.png" alt="Overview showing scopes, top RAID, coming dates and recent progress" width="100%">
      </a>
      <br><strong>Overview</strong> - a compact stakeholder readout built from live workbook data.
    </td>
    <td width="50%" valign="top">
      <a href="docs/assets/readme/items-raid.png">
        <img src="docs/assets/readme/items-raid.png" alt="Items hierarchy and RAID register shown together" width="100%">
      </a>
      <br><strong>Items & RAID</strong> - structured delivery data, hierarchy controls and clear exception highlighting.
    </td>
  </tr>
</table>

## What the workbook does

| Sheet | Purpose |
|---|---|
| Overview | Shows scopes, top RAID, upcoming dates and recent progress in four stakeholder panels |
| Plan | Shows a filterable six-level schedule with weekly bars, key dates and the current week |
| Items | Holds the editable hierarchy, workflow, ownership, dates, delivery health and latest status |
| RAID | Holds editable risks, assumptions, issues, dependencies and decisions |
| Config | Owns settings, taxonomy, workflow roles, severity bands and people |
| Calc | Supplies protected calculations, lookups and dynamic view data |

Its VBA assigns IDs, stamps workflow dates, organises the item hierarchy, fits the opening window to the screen and exports one clean UTF-8 Markdown file without runtime logs or sidecars.

## Build from source

The build needs Python 3.12 and the dependencies locked in `uv.lock`.

```bash
uv sync --frozen
.venv/bin/python -m build
```

The release pipeline requires the compiled VBA to match both source modules exactly. It builds untouched `.xlsx` and `.xlsm` baselines, full-rebuilds disposable copies in desktop Excel, and compares worksheet and table formulas, defined names, validation, conditional formatting, package parts and the VBA project before publication. Excel must have no other workbooks open because `CalculateFullRebuild` operates across the Excel application.

After either VBA source file changes, refresh the compiled project without opening the Visual Basic Editor:

```bash
.venv/bin/python -m build.automation.refresh_vba
```

The command replaces both complete modules in a disposable workbook, proves that no non-VBA package member changed, has desktop Excel compile and save the project, verifies the compiled caches and exact source, then atomically publishes `build/vba/vbaProject.bin`.

All destinations are published as one rollback-capable transaction:

- `dist/PM_Workbook.xlsx` - formula-only QA copy
- `dist/PM_Workbook.xlsm` - macro-enabled release
- `PM_Workbook.xlsm` - promoted release copy

Desktop Excel must be installed and the release command must be permitted to control it. Do not maintain the generated workbook by editing it directly. Change the source in `build/`, rebuild it and verify the result.

## Verify a release

The release gate has two explicit phases. First, run every automated check and create a template bound to the exact source tree and release workbook:

```bash
.venv/bin/python -m build.qa.release prepare --evidence-template /tmp/pm-workbook-macro-evidence.json
```

`prepare` returns status 2 after the automated gates pass because a release is deliberately blocked at that point. Use a disposable copy of the exact generated `.xlsm` to complete the live checks in [VBA maintenance](docs/vba-maintenance.md), fill every template check with `PASS`, and record the generated Markdown evidence. Then run:

```bash
.venv/bin/python -m build.qa.release final --macro-evidence /tmp/pm-workbook-macro-evidence.json
```

`final` accepts evidence only for the current source digest and exact release workbook, requires it to be less than 24 hours old, verifies the UTF-8 Markdown and export-directory contents, and reruns the non-mutating release matrix. Only its zero exit status is a shippable result. The matrix includes strict Ruff checks, source hygiene, compiled VBA matching, structural and design QA for both formats, formula scenarios, desktop-Excel save preservation, interaction performance and the populated demonstration workbook.

## Repository map

```text
build/
  automation/   VBA refresh, Excel recalculation, repair detection and performance measurement
  core/         design tokens, formula encoding, layout and OOXML styling
  qa/           structural, design, scenario, VBA and live-Excel checks
  scenarios/    representative release-data builder
  spec/         workbook schema, capacities, formulas and fixtures
  vba/          authoritative VBA source and compiled project
  writers/      worksheet composers
  pipeline.py   verified build and rollback-safe publication pipeline
docs/
  research/     source-backed Excel and product-pattern research
  *.md          user, design, domain, implementation and maintenance guides
PM_Workbook.xlsm
```

## Documentation

- [User guide](docs/user-guide.md)
- [Architecture and domain model](docs/pm-domain-reference.md)
- [Design system](docs/design-system.md)
- [Excel implementation reference](docs/excel-reference.md)
- [VBA maintenance](docs/vba-maintenance.md)
- [Tooling and release capabilities](docs/capabilities.md)
- [Research index](docs/research/excel-2026/README.md)
