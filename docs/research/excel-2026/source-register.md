# Research source register

Primary sources establish product and library behavior. Practitioner sources support construction and visual exploration. Version-sensitive statements are checked against the live source and the installed software before implementation.

## Excel formulas and file authoring

| Topic | Source | Use |
|---|---|---|
| LAMBDA | [Microsoft Support](https://support.microsoft.com/en-us/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67) | Named workbook functions |
| GROUPBY | [Microsoft Support](https://support.microsoft.com/en-us/office/groupby-function-5e08ae8c-6800-4b72-b623-c41773611505) | Formula-native grouped aggregation |
| PIVOTBY | [Microsoft Support](https://support.microsoft.com/en-us/office/pivotby-function-de86516a-90ad-4ced-8522-3a25fac389cf) | Formula-native cross-tab aggregation |
| Future-function compatibility | [Microsoft Support](https://support.microsoft.com/en-us/office/issue-an-xlfn-prefix-is-displayed-in-front-of-a-formula-882f1ef7-68fb-4fcd-8d54-9fbb77fd5025) | `_xlfn` behavior and compatibility diagnostics |
| Dynamic and future formulas | [XlsxWriter](https://xlsxwriter.readthedocs.io/working_with_formulas.html) | Formula storage, array writing and spill references |
| Named LAMBDA encoding | [XlsxWriter](https://xlsxwriter.readthedocs.io/example_lambda.html) | `_xlfn.LAMBDA` and `_xlpm.` parameters |
| Formula inspection | [openpyxl](https://openpyxl.readthedocs.io/en/latest/simple_formulae.html) | Reader behavior for stored formulas |
| Macro-enabled workbooks and events | [Microsoft Support](https://support.microsoft.com/en-us/office/run-a-macro-in-excel-5e855fd2-02d1-45f5-90a3-50e645fe3155) | `.xlsm`, VBA and `Workbook_Open` behavior on Windows and Mac desktop Excel |
| Native in-cell checkboxes | [Microsoft Support](https://support.microsoft.com/en-us/office/using-check-boxes-in-excel-da85546d-c110-49b8-b633-9cebadcaf8d4) | Boolean checkbox behavior in Microsoft 365 Excel |

## Excel automation

| Topic | Source | Use |
|---|---|---|
| VBA and Office Scripts | [Microsoft Learn](https://learn.microsoft.com/en-us/office/dev/scripts/resources/vba-differences) | Platform and event-model boundaries |
| Scripts in Power Automate | [Microsoft Learn](https://learn.microsoft.com/en-us/office/dev/scripts/tutorials/excel-power-automate-manual) | Flow execution and stable workbook identifiers |
| VBA project storage | [Microsoft MS-OVBA specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/c66b58a6-f8ba-4141-9382-0612abce9926) | Source modules, performance-cache invalidation and host interoperability |
| VBA package writing | [pyOpenVBA](https://github.com/WilliamSmithEdward/pyopenvba) | Complete source replacement without VBE or COM automation |
| Mac Office file access | [Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/office-mac/grantaccesstomultiplefiles) | Sandboxing, explicit access prompts and persisted grants |
| Excel Mac container | [Microsoft Learn](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac) | Documented Excel `Data/Documents` sandbox path used by disposable release automation |
| Excel save-path picker | [Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.application.getsaveasfilename) | Cross-platform destination selection without saving the workbook; includes the Macintosh-only button label |
| Formula modules | [Microsoft Advanced Formula Environment](https://github.com/microsoft/advanced-formula-environment) | Named-function organization patterns |

## Project tools and workflow design

| Topic | Source | Use |
|---|---|---|
| Shared sheet views | [Smartsheet](https://help.smartsheet.com/articles/765715-grid-gantt-calendar-and-card-views) | Operational and schedule projections over shared rows |
| Dashboard components | [Smartsheet](https://help.smartsheet.com/articles/518558-widget-types-for-smartsheet-dashboards) | Component roles and report-style detail |
| Workflow blocks | [Smartsheet](https://help.smartsheet.com/articles/2479061-automate-processes-end-to-end-with-visual-workflows) | Trigger, condition and action grammar |
| Trigger behavior | [Smartsheet](https://help.smartsheet.com/articles/2479236-trigger-blocks-when-your-workflow-is-executed) | Event timing and deterministic conditions |
| Interface layouts | [Airtable](https://support.airtable.com/docs/getting-started-with-airtable-interface-designer) | Data, appearance, action and view-setting separation |
| Automation pipeline | [Airtable](https://support.airtable.com/docs/getting-started-with-airtable-automations) | Trigger, action, branch and test model |
| Trigger catalog | [Airtable](https://support.airtable.com/docs/automation-triggers) | Current trigger behavior |
| Dashboard and governance patterns | [Wrike](https://help.wrike.com/hc/en-us/articles/35673266706578-What-s-New-in-Wrike-May-2026) | Drill-down, required fields and generated output |

## Practitioner and competition references

| Source | Use |
|---|---|
| [Excel University: spill-range dropdowns](https://www.excel-university.com/create-dependent-drop-downs-with-spill-ranges/) | Dynamic validation pattern exploration |
| [Ablebits: dependent dropdowns](https://www.ablebits.com/office-addins-blog/dependent-dropdown-list-multiple-rows-excel/) | Per-row validation patterns |
| [ICAEW: GROUPBY and PIVOTBY](https://www.icaew.com/technical/technology/excel-community/excel-community-articles/2024/excels-groupby-and-pivotby-functions) | Aggregation and chart/table caveats |
| [Chandoo: project dashboard](https://chandoo.org/wp/interactive-project-dashboard-with-excel/) | Gantt and workbook construction examples |
| [Chandoo: dashboard gallery](https://chandoo.org/wp/excel-dashboards/examples/) | Visual composition examples |
| [Financial Modeling World Cup](https://fmworldcup.com/) | Transparent formula construction and verifiable outputs |
| [Excel World Championship rules](https://excel-esports.com/rules/) | Objective workbook-case grading |
| [Microsoft MVP Blog: championship retrospective](https://techcommunity.microsoft.com/blog/mvp-blog/from-espn-to-the-spreadsheet-arena-how-excel-mvps-powered-the-microsoft-excel-wo/4497642) | Current competition context |

## Practitioner evidence: one view across different tools

| Source | Use |
|---|---|
| [r/projectmanagement: organising and tracking multiple projects, March 2026](https://www.reddit.com/r/projectmanagement/comments/1s1v9hj/organising_and_tracking_multiple_projects/) | First-hand discussion of separate team and portfolio layers, mixed Microsoft tools and login or training barriers |
| [r/projectmanagement: tool advice, May 2026](https://www.reddit.com/r/projectmanagement/comments/1tdznqe/tool_advice/) | Multiple systems, incompatible status language, manual translation and reporting-staleness trade-offs |
| [r/projectmanagement: choosing PM software, September 2025](https://www.reddit.com/r/projectmanagement/comments/1n7adzm/how_many_project_management_tools_did_you_try/) | Corporate tool constraints and the gap between executive reporting and day-to-day project operation |
| [r/consulting: multiple client projects, May 2024 with replies through 2026](https://www.reddit.com/r/consulting/comments/1crab18/project_management_tools_for_multiple_client/) | Client-owned tool stacks and the need for a cross-client bridge without forced migration |
| [ProjectManagement.com: orchestrating multiple projects, May 2026](https://www.projectmanagement.com/discussion-topic/239152/which-tools-have-you-used-to-orchestrate-multiple-projects-in-parallel-) | Named practitioners on deliberately combining delivery, portfolio, reporting and documentation tools |
| [ProjectManagement.com: moving from Excel to enterprise PPM, February 2026](https://www.projectmanagement.com/discussion-topic/235391/seeking-recommendations--enterprise-project-management-tool--portfolio---timeline-focus-) | Named practitioners on the point where manual consolidation outgrows Excel and a fuller platform becomes the right move |

## Verification notes

- Microsoft and library documentation governs technical behavior.
- Product plans, interface labels and limits are checked at the time of use.
- XlsxWriter and openpyxl behavior is verified against the versions declared in `pyproject.toml`.
- Visual examples support design judgment and carry no functional authority.
- Desktop Excel testing resolves package, formula and interaction questions for the release target.
