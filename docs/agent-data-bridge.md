# Agent data bridge

## Purpose

The agent data bridge is the local mutation boundary for `Items` and `RAID`. An agent may
read an API, MCP tool, file, message or pasted text in any way it can justify. The workbook
accepts only one strict provider-neutral JSON change set.

The workflow separates interpretation from mutation:

1. `describe` reads the workbook contract, Config choices, current records and exact target
   digests.
2. The agent normalises its source into `upsert` and `mark_deleted` operations.
3. `plan` validates identities, relationships, lifecycle behavior and the complete merged
   workbook state without writing anything.
4. A person reviews the exact diffs, warnings, effective time and plan token.
5. `apply` reparses and replans the same JSON, checks the reviewed token and publishes only
   when the workbook still matches the approved digest.

The bridge does not contain a connector, model SDK, credential store, mapping cache or
background sync. Source omission never changes a workbook row.

## Commands

Run commands from the repository root. Each command defaults to `PM_Workbook.xlsm`.

```bash
.venv/bin/python -m build.data describe [workbook] [--output FILE]
.venv/bin/python -m build.data plan CHANGESET|- [workbook] [--output FILE]
.venv/bin/python -m build.data apply CHANGESET|- [workbook] --approve PLAN_TOKEN [--output FILE]
```

`-` reads one complete UTF-8 JSON document from stdin. Without `--output`, the JSON result is
written to stdout. With `--output`, it is written atomically to that file. Operational build
messages use stderr. An output file cannot alias the workbook or the change-set input.

| Exit | Meaning |
|---:|---|
| 0 | Success, including a reviewed no-change result |
| 2 | Invalid command, JSON, contract, workbook data or operation |
| 3 | Stale target, changed workbook or mismatched approval token |
| 4 | Build, desktop-Excel, verification, backup, publication or cleanup failure |

`export` and `migrate` remain available for snapshot and structure-upgrade work.

## Describe

Start every run from a fresh description:

```bash
.venv/bin/python -m build.data describe PM_Workbook.xlsm --output /tmp/pm-describe.json
```

The result contains:

- the complete JSON Schema Draft 2020-12 contract at version `1.0.0`;
- the workbook SHA-256, observed workbook-schema fingerprint and current build-schema
  fingerprint;
- the effective local date and supported capacities;
- every writable Item and RAID field with its ownership classification;
- current Config settings, types, statuses, people and other choices; and
- ordered current Item and RAID records, including workbook IDs and `Source` / `Source ID`.

Copy the three values under `target` into the change set exactly. A source identity shown by
`describe` is durable and should be reused on later runs.

## Change-set rules

The root object and every core nested object are closed. Unknown fields are errors. The
optional `extensions` objects carry metadata only and cannot request workbook writes.

### Identity

An operation identifies a row by:

- `{"workbook_id": "I-1001"}` or `{"workbook_id": "R-1"}`;
- `{"source": {"namespace": "api:portfolio", "record_id": "opaque-42"}}`; or
- both values, to attach a source identity to one existing manual row.

Workbook IDs match case-insensitively, as Excel does. Source namespace and record ID are exact
case-sensitive strings. A new row requires a source identity. For pasted material without a
natural key, create a stable synthetic namespace and record ID, then reuse the pair returned by
later `describe` calls.

`client_ref` names a row created in the same batch. Item `Parent`, Item `BlockedBy` and RAID
`RelatedID` accept a workbook ID, source identity or `client_ref` reference.

### Writes and clearing

- A field under `set` is written.
- A field named in `clear` is blanked.
- An omitted field is preserved.
- JSON `null` is never a clearing instruction.

The bridge writes Item and RAID input fields only. Workbook IDs, formulas, Config, People and
lifecycle stamps are workbook-owned. `Source` and `Source ID` are assigned only through the
operation identity.

Formula-looking text and URLs are stored literally. Values beginning with `=`, `+`, `-` or `@`
do not become formulas, and URL-shaped values do not become automatic hyperlinks.

### Deletion

`mark_deleted` changes only the target row's Status to the Config row carrying `IsDeleted`.
Nothing cascades and no physical row is erased. Repeating the operation is a no-op.

A deleted row remains visible with its workbook and source identities. If the same source
identity appears in a later `upsert`, the bridge creates a new row and workbook ID while keeping
the deleted row as history. At most one non-deleted row may carry a source identity.

## Complete change-set example

The digest strings below satisfy the JSON schema but are illustrative. Replace the entire
`target` object with the exact values from the current `describe` result before planning.

```json
{
  "contract": "excel-pm-agent-change-set",
  "version": "1.0.0",
  "request_id": "3e5f7e30-7df5-4f16-8ce8-8725a02f3d51",
  "created_at": "2026-07-15T10:00:00Z",
  "target": {
    "workbook_sha256": "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
    "workbook_schema_fingerprint": "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb",
    "build_schema_fingerprint": "cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc"
  },
  "producer": {
    "name": "delivery-status-agent",
    "version": "1.0",
    "extensions": {}
  },
  "source": {
    "description": "Weekly delivery status supplied in a meeting note",
    "uri": "clipboard:weekly-status",
    "retrieved_at": "2026-07-15T09:55:00Z",
    "extensions": {}
  },
  "operations": [
    {
      "operation_id": "create-project",
      "op": "upsert",
      "entity": "item",
      "identity": {
        "source": {
          "namespace": "paste:weekly-status",
          "record_id": "project-001"
        }
      },
      "client_ref": "project",
      "set": {
        "Type": "Project",
        "Title": "Delivery programme",
        "Status": "In Progress",
        "Delivery Health": "On track",
        "Start": "2026-07-01",
        "Due": "2026-12-18",
        "Latest Status": "Mobilisation is complete and delivery is tracking to plan."
      },
      "clear": [],
      "extensions": {}
    },
    {
      "operation_id": "create-task",
      "op": "upsert",
      "entity": "item",
      "identity": {
        "source": {
          "namespace": "paste:weekly-status",
          "record_id": "task-001"
        }
      },
      "set": {
        "Type": "Task",
        "Title": "Confirm supplier capacity",
        "Parent": {"client_ref": "project"},
        "Status": "Ready",
        "Priority": "P1",
        "Due": "2026-07-24"
      },
      "clear": [],
      "extensions": {}
    },
    {
      "operation_id": "create-risk",
      "op": "upsert",
      "entity": "raid",
      "identity": {
        "source": {
          "namespace": "paste:weekly-status",
          "record_id": "risk-001"
        }
      },
      "set": {
        "Type": "Risk",
        "Title": "Supplier capacity may constrain mobilisation",
        "Detail": "One specialist team is shared with another programme.",
        "RelatedID": {"client_ref": "project"},
        "Status": "Open",
        "Prob": 3,
        "Impact": 4,
        "Response": "Confirm named backup capacity before mobilisation.",
        "NextReview": "2026-07-22"
      },
      "clear": [],
      "extensions": {}
    }
  ],
  "extensions": {}
}
```

## Plan, revise and apply

Write the JSON to a file, then plan it:

```bash
.venv/bin/python -m build.data plan /tmp/pm-changes.json PM_Workbook.xlsm \
  --output /tmp/pm-plan.json
```

A valid plan reports every create, update, deletion transition, no-op and field-level
before/after value. It also reports current workbook warnings and any structural adjustment that
the rebuild would make. Any error blocks the whole batch and omits `plan_token`.

Correct the change set and rerun `plan` until it is valid. Review:

- the operation summary and every field diff;
- warnings and Config reconciliation adjustments;
- the effective local date, timezone, UTC offset and `expires_at`; and
- the exact `plan_token`.

The token expires at local midnight because lifecycle stamps use the effective local date. Apply
the same change-set bytes only after approval:

```bash
.venv/bin/python -m build.data apply /tmp/pm-changes.json PM_Workbook.xlsm \
  --approve "PLAN_TOKEN_FROM_REVIEW" --output /tmp/pm-apply.json
```

Apply replans first. A changed JSON document, workbook digest, schema fingerprint, intended
authored state, warning set, timezone boundary or effective date produces a conflict. A true
no-op creates no snapshot, backup, build or publication file. A real change writes a pre-change
snapshot, rebuilds from source, runs desktop-Excel calculation and semantic verification, then
publishes the exact pre-change backup and calculated workbook as one rollback-capable
transaction.

## Paste-ready agent prompt

```text
Update the Excel project-management workbook at <absolute workbook path> from <source or task>.

Use the local provider-neutral bridge only. First run:
python -m build.data describe "<absolute workbook path>" --output /tmp/pm-describe.json

Read the embedded schema, Config choices, current IDs, Source identities and all three target
digests. Obtain and interpret the source data, then write one UTF-8 contract version 1.0.0 change
set to /tmp/pm-changes.json. Use only upsert and explicit mark_deleted operations. Preserve
omitted fields, use clear for intentional blanks, give every new row a durable source namespace
and record ID, and use structured references for Parent, BlockedBy and RelatedID. Do not write
IDs, lifecycle stamps, formulas, Config, People, Source or Source ID directly.

Run:
python -m build.data plan /tmp/pm-changes.json "<absolute workbook path>" --output /tmp/pm-plan.json

If the plan has errors, correct the change set and plan again. When it is valid, show me a concise
review containing creates, updates, deletion transitions, no-ops, every material field diff,
warnings, structural adjustments, effective time, expiry and the exact plan token. Stop and wait
for my explicit approval. Do not infer approval from this prompt.

After I approve that exact token, run:
python -m build.data apply /tmp/pm-changes.json "<absolute workbook path>" --approve "<approved token>" --output /tmp/pm-apply.json

Report the apply exit code and JSON result. Do not edit the workbook directly and do not create
connector state, mapping files, logs or background work.
```
