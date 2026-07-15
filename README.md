# Excel Project Management Workbook

<p align="center">
  <strong>One shared delivery view when the work lives across teams, companies and tools.</strong>
</p>

<p align="center">
  <a href="https://github.com/maxymax92/excel-pm-tool/releases/latest"><img src="https://img.shields.io/github/v/release/maxymax92/excel-pm-tool?style=flat-square&amp;label=release" alt="Latest release"></a>
  <img src="https://img.shields.io/badge/Excel-365%20desktop-217346?style=flat-square&amp;logo=microsoftexcel&amp;logoColor=white" alt="Excel 365 desktop">
  <img src="https://img.shields.io/badge/Python-3.12-3776AB?style=flat-square&amp;logo=python&amp;logoColor=white" alt="Python 3.12 build">
</p>

<p align="center">
  <a href="https://github.com/maxymax92/excel-pm-tool/releases/latest/download/PM_Workbook.xlsm"><strong>Download the latest workbook</strong></a>
  &nbsp;&middot;&nbsp;
  <a href="docs/user-guide.md">User guide</a>
  &nbsp;&middot;&nbsp;
  <a href="docs/agent-data-bridge.md">Agent workflow</a>
</p>

When delivery spans teams, organisations, suppliers or regions, another project system rarely solves the reporting problem. This workbook gives everyone one compact coordination layer without asking every team to abandon the tools they already use.

Work lives in `Items` and `RAID`. `Config` owns the rules. `Overview` gives leaders an exception-led brief, while `Plan` gives delivery teams a six-level weekly schedule. The `Export to Markdown` button writes one clean `PM_Status_yyyy-mm-dd.md` handoff for reporting or an AI tool.

This is a shared record, not live synchronisation between the systems around it. Store and govern the working file through the document controls your organisation already trusts.

Under the workbook is a deterministic Python build, source-controlled VBA, desktop-Excel verification, durable data migration and a review-before-apply agent workflow. The interface stays Excel; the repository treats the workbook as a versioned, generated release.

> **What changed in version 1.2.0:** provider-neutral agent changes, durable source identities, explicit Deleted history, direct per-item reporting and stronger transactional publication. [Read the release notes](https://github.com/maxymax92/excel-pm-tool/releases/tag/v1.2.0).

## See it in action

These views use the populated demonstration workbook, so the hierarchy, dates and exceptions show how the finished system behaves.

<p align="center">
  <a href="docs/assets/readme/overview.png">
    <img src="docs/assets/readme/overview.png" alt="Overview showing scope health, top RAID, coming dates and recent progress" width="100%">
  </a>
  <br><strong>Overview</strong> - one compact stakeholder brief built from the live delivery record.
</p>

<table>
  <tr>
    <td width="50%" valign="top">
      <a href="docs/assets/readme/plan.png">
        <img src="docs/assets/readme/plan.png" alt="Plan view showing a six-level project schedule, status legend and weekly timeline" width="100%">
      </a>
      <br><strong>Plan</strong> - a filterable 52-week schedule with intervals, key dates and the current week.
    </td>
    <td width="50%" valign="top">
      <a href="docs/assets/readme/items-raid.png">
        <img src="docs/assets/readme/items-raid.png" alt="Items hierarchy and RAID register shown together" width="100%">
      </a>
      <br><strong>Items &amp; RAID</strong> - the shared operational record, with hierarchy controls and clear exception cues.
    </td>
  </tr>
</table>

## What you get

| Capability | What it gives you |
|---|---|
| One shared delivery model | Commitments, dates, status, ownership, health, risks and decisions in one portable file |
| Executive reporting | Scope health, top RAID, coming dates and recent progress without another system login |
| A real delivery hierarchy | Six configurable levels, parent-first WBS order, up to three scopes and a 52-week Plan |
| Bounded, visible capacity | 2,000 Item rows, 2,000 RAID rows and 500 rows per Config list, with explicit capacity warnings |
| Item-owned reporting values | Every item keeps its own owner, health, status, dates and narrative; hierarchy never substitutes another row's values |
| Config-driven behaviour | Types, status roles, RAID rules, severity bands, people, thresholds, prefixes and counters stay editable in Excel |
| Clean handoffs | One UTF-8 Markdown snapshot with no runtime logs, caches or sidecars |
| Safe upgrades | Authored rows and Config values survive structural rebuilds through snapshots, semantic comparison and rollback |
| Controlled agent updates | Source-specific interpretation stays outside Excel; the workbook accepts only a strict change set with a read-only plan and reviewed token |

## Get started

1. [Download `PM_Workbook.xlsm`](https://github.com/maxymax92/excel-pm-tool/releases/latest/download/PM_Workbook.xlsm).
2. Open it in Microsoft Excel 365 desktop and enable macros.
3. Review or delete the clearly marked example rows.
4. Maintain the shared rules and people in `Config`.
5. Record work in `Items` and risks or decisions in `RAID`.
6. Use `Overview` for the stakeholder brief and `Plan` for schedule review.

Macros assign IDs, stamp lifecycle dates, organise the hierarchy and export Markdown. Those actions do not run when macros are disabled.

The current release is verified on Microsoft Excel 365 Personal for Mac 16.111. The VBA includes Windows-specific path and destination branches, but Windows is not yet part of the release test. Normal workbook use needs only the downloaded file; the Python workflows below are for maintainers and agents working from a source checkout.

## Six sheets, clear responsibilities

| Sheet | Responsibility |
|---|---|
| Overview | Protected, derived stakeholder brief with scope health, top RAID, coming dates and recent progress |
| Plan | Protected, derived six-level schedule with scope, depth and date-window controls |
| Items | Editable work hierarchy, workflow, ownership, dates, delivery health and latest status |
| RAID | Editable risks, assumptions, issues, dependencies and decisions |
| Config | Editable settings, taxonomy, workflow roles, severity bands and people |
| Calc | Protected, hidden calculation and bounded-spill layer |

The VBA project assigns IDs, stamps lifecycle dates, applies hierarchy presentation, organises Items into WBS order, fits the opening window and exports the Markdown status file. Formula and conditional-formatting rules keep derived views current and flag incomplete or contradictory data without silently changing what the user entered.

## Review agent changes before they reach Excel

The repository ships a provider-neutral change-set boundary, not connectors or a model SDK. An agent may interpret an API, MCP tool, file or pasted text using whatever access it has. Only the strict change set crosses the workbook-mutation boundary, and a read-only plan shows the exact result before anything is applied.

```bash
.venv/bin/python -m build.data describe [workbook] [--output FILE]
.venv/bin/python -m build.data plan CHANGESET|- [workbook] [--output FILE]
.venv/bin/python -m build.data apply CHANGESET|- [workbook] --approve PLAN_TOKEN [--output FILE]
```

- `describe` returns the contract, current Config choices, records and exact workbook digests.
- `plan` returns creates, updates, explicit Deleted transitions, no-ops, field diffs, warnings and errors without writing anything.
- `apply` reparses and replans the same bytes, verifies the reviewed token and workbook digest, then uses the normal snapshot, rebuild, Excel and publication pipeline.

Only `Items` and `RAID` input fields are writable. IDs, formulas, lifecycle stamps, Config and People remain workbook-owned. Source omission never means deletion. `mark_deleted` changes one row's status and keeps the row as visible history; a later reappearance receives a fresh workbook ID. A true no-op creates no snapshot, backup or build artifact.

Imported text stays literal even when it begins with `=`, `+`, `-` or `@`, or looks like a URL. The token guards the reviewed plan against a stale workbook; it is not a user-authentication system. The complete contract, worked example and paste-ready prompt are in the [agent data bridge guide](docs/agent-data-bridge.md).

## Keep data across rebuilds

A generated workbook is a disposable rendering of two durable inputs: the source under `build/` and a JSON snapshot of everything a user authored. Structural upgrades never edit a populated workbook in place.

```bash
.venv/bin/python -m build.data export [workbook]
.venv/bin/python -m build.data migrate [workbook]
```

`export` captures Items, RAID, Config lists, people, settings and VBA-stamped values into the bounded ring in `dist/snapshots/`. `migrate` rebuilds the workbook with that data injected, recalculates it in desktop Excel, compares the intended authored snapshot with the rebuilt result and publishes it transactionally. The replaced workbook is retained in `dist/backups/`.

Out-of-schema tables, columns or settings, duplicate or missing IDs, capacity breaches and types the taxonomy cannot level stop before the build with an actionable diagnostic. Existing values that the workbook flags red migrate unchanged and are reported as warnings.

## Built like software

This is not a hand-maintained template. The release pipeline:

- generates every sheet, table, formula, name, validation rule, format and theme from Python source;
- embeds a lean VBA project and verifies the compiled binary against both source modules;
- disables automatic formula and URL conversion for imported text;
- full-rebuilds disposable copies in desktop Excel and treats repair output as a release blocker;
- compares formulas, names, validation, conditional formatting, package parts and VBA before publication;
- verifies empty, representative and adversarial workbook states, plus live interactions and latency;
- publishes every destination and required pre-change backup as one rollback-capable transaction;
- binds final live evidence to the exact source and workbook digests.

The implementation details and verification layers are documented in [Toolchain and release capabilities](docs/capabilities.md).

## Build from source

The build needs Python 3.12 and the dependencies locked in `uv.lock`.

```bash
uv sync --frozen
.venv/bin/python -m build
```

All release destinations publish together or return to their original bytes:

- `dist/PM_Workbook.xlsx` - formula-only QA copy
- `dist/PM_Workbook.xlsm` - macro-enabled release
- `PM_Workbook.xlsm` - promoted release copy

After either VBA source file changes, refresh the compiled project without opening the Visual Basic Editor:

```bash
.venv/bin/python -m build.automation.refresh_vba
```

The refresh replaces both complete modules in a disposable workbook, proves that no non-VBA package member changed, has Excel compile and save the project, verifies the result and atomically publishes `build/vba/vbaProject.bin`.

Do not maintain the generated workbook by editing the release artifact. Change the source in `build/`, rebuild it and verify the result.

## Verify a release

The release gate has two explicit phases. `prepare` runs the automated matrix and creates an evidence template bound to the exact source tree and workbook:

```bash
.venv/bin/python -m build.qa.release prepare --evidence-template /tmp/pm-workbook-macro-evidence.json
```

The deliberate status 2 means the automated checks passed and the release is waiting for live Excel evidence. Complete the checks in [VBA maintenance](docs/vba-maintenance.md) on a disposable copy of that exact workbook, then run:

```bash
.venv/bin/python -m build.qa.release final --macro-evidence /tmp/pm-workbook-macro-evidence.json
```

`final` accepts only fresh evidence for the current source and workbook digests, verifies the Markdown export and reruns the non-mutating release matrix. Only its zero exit status is a release pass. Excel must have no other workbooks open during rebuild checks because `CalculateFullRebuild` works across the Excel application.

## Repository map

```text
build/
  automation/   VBA refresh, Excel recalculation, repair detection and performance measurement
  core/         design tokens, formula encoding, layout and OOXML styling
  data/         snapshots, provider-neutral change sets, merge, apply and migration
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
- [Agent data bridge](docs/agent-data-bridge.md)
- [Tooling and release capabilities](docs/capabilities.md)
- [Research index](docs/research/excel-2026/README.md)
