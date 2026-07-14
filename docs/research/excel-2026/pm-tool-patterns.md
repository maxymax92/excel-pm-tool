# Project-tool interaction patterns

## Shared-record architecture

Smartsheet, Airtable and Wrike support multiple experiences over canonical records. The durable pattern is:

```text
canonical records
  -> operational editor
  -> schedule projection
  -> stakeholder summary
  -> deterministic export
```

The workbook maps this to Items, Plan, Overview and ExportMarkdown. RAID is a dedicated register linked to the same scope hierarchy.

## When teams cannot share one tool

Practitioners keep describing two different jobs: detailed work management inside each team, and a small common view across projects. A March 2026 [r/projectmanagement discussion](https://www.reddit.com/r/projectmanagement/comments/1s1v9hj/organising_and_tracking_multiple_projects/) describes teams using spreadsheets, Planner, Lists and documents while needing cross-project visibility without another login or training burden. A May 2026 [tool-advice discussion](https://www.reddit.com/r/projectmanagement/comments/1tdznqe/tool_advice/) highlights the extra translation and stale reporting created when Jira, GitHub and Microsoft tools use different fields and status language.

Named practitioners in a May 2026 [ProjectManagement.com discussion](https://www.projectmanagement.com/discussion-topic/239152/which-tools-have-you-used-to-orchestrate-multiple-projects-in-parallel-) describe organisations deliberately combining delivery, portfolio, reporting and documentation tools instead of forcing one platform to do everything. Consultants describe the same problem when [each client brings its own stack](https://www.reddit.com/r/consulting/comments/1crab18/project_management_tools_for_multiple_client/).

The workbook covers the common layer. It keeps the fields needed for coordination, senior reporting and one status file to pass on when systems cannot be connected. Detailed team records stay in their existing tools. This is a fallback for access and integration gaps, not a claim that Excel should replace a connected project-management platform.

## Smartsheet

Smartsheet documents Grid, Gantt, Calendar, Card, Board, Timeline and Table views over shared sheet data. Gantt combines an identity grid with a time-scaled schedule, which directly supports the Plan layout.

Its dashboard components also clarify the difference between metrics, reports, timelines, shortcuts and explanatory content. Overview uses report-style record panels because stakeholders need the actual exceptions and dates behind each summary.

References:

- [Smartsheet: sheet views](https://help.smartsheet.com/articles/765715-grid-gantt-calendar-and-card-views)
- [Smartsheet: dashboard widgets](https://help.smartsheet.com/articles/518558-widget-types-for-smartsheet-dashboards)

## Airtable

Airtable Interface Designer separates data scope, appearance, permitted actions and view settings. Plan mirrors that grammar through Scope, Depth, From and To. Items owns the full hierarchy and editable data.

Airtable automation also emphasizes explicit triggers, conditional branches, testing and run evidence. The workbook expresses those ideas as deterministic event preflight, visible errors and release QA.

References:

- [Airtable: Interface Designer](https://support.airtable.com/docs/getting-started-with-airtable-interface-designer)
- [Airtable: automations](https://support.airtable.com/docs/getting-started-with-airtable-automations)

## Wrike

Wrike’s dashboard drill-down, required-field and generated-document patterns reinforce three current workbook choices:

- every summary has a route to the underlying record;
- governance is progressive and tied to workflow meaning;
- exports are generated from canonical data.

Reference:

- [Wrike product updates](https://help.wrike.com/hc/en-us/articles/35673266706578-What-s-New-in-Wrike-May-2026)

## Applied interaction rules

1. One editable record store per domain.
2. Views project records and retain stable identifiers.
3. A panel answers one bounded question.
4. A summary shows its qualifying records or an explicit truncation count.
5. Controls sit next to the view they affect.
6. Hierarchy lives in the operational data and outline controls.
7. Automation validates the complete transition before mutation.
8. Export reads canonical records and derived panels.
9. The shared layer stays smaller than the team systems beneath it.
