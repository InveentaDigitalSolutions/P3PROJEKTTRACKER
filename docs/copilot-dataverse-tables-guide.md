# GitHub Copilot Guide — Provisioning Dataverse Tables via Script

Use this file as a prompt and reference when asking GitHub Copilot (in any project)
to generate a Node.js script that creates Dataverse tables, columns, relationships,
and seed data through the Dataverse Web API.

---

## How to Use This File

1. Open GitHub Copilot Chat in your project.
2. Type: `Using the guide below, generate a Dataverse provisioning script for my tables:`
3. Paste the **Copilot Prompt** section (and fill in your own tables).
4. Use the **Follow-Up Prompts** for extras like solution assignment and seed data.

---

## Copilot Prompt

```
You are working on a Power Platform project that needs Dataverse tables provisioned
via the Dataverse Web API (REST). Follow these rules exactly.

## Environment
- Dataverse org URL: read from DATAVERSE_URL environment variable
- Auth: Bearer token — try Azure CLI first:
    az account get-access-token --resource <DATAVERSE_URL> --query accessToken -o tsv
  If that fails, fall back to MSAL device-code flow using the public client app ID:
    51f81489-12ee-4a9e-aaae-a2591f45987d  (Power Apps first-party)
- All API calls go to: <DATAVERSE_URL>/api/data/v9.2/

## Table Naming Conventions
- Use a consistent publisher prefix for all names, e.g. `xyz_`
- Schema name (PascalCase):  Xyz_Project
- Logical name (lowercase):  xyz_project
- Collection name:           xyz_projects
- Primary name column:       xyz_name  (string, required)

## Supported Column Types
| Type      | Dataverse AttributeType         | Key Properties                           |
|-----------|---------------------------------|------------------------------------------|
| string    | StringAttributeMetadata         | MaxLength: 100–4000                      |
| text/memo | MemoAttributeMetadata           | MaxLength: 2000–10000                    |
| int       | IntegerAttributeMetadata        | MinValue / MaxValue                      |
| decimal   | DecimalAttributeMetadata        | Precision (0–4)                          |
| datetime  | DateTimeAttributeMetadata       | DateTimeBehavior: "UserLocal"            |
| boolean   | BooleanAttributeMetadata        | TrueOption / FalseOption labels + values |
| choice    | PicklistAttributeMetadata       | Local option set with Value (int) + Label|
| lookup    | OneToManyRelationship           | Must be created AFTER both tables exist  |

## API Endpoints Reference
| Action              | Method | Endpoint                                                          |
|---------------------|--------|-------------------------------------------------------------------|
| Create table        | POST   | /EntityDefinitions                                                |
| Create column       | POST   | /EntityDefinitions(LogicalName='<table>')/Attributes              |
| Create lookup       | POST   | /RelationshipDefinitions                                          |
| Check table exists  | GET    | /EntityDefinitions(LogicalName='<table>')?$select=LogicalName     |
| Check column exists | GET    | /EntityDefinitions(LogicalName='<table>')/Attributes(LogicalName='<col>') |
| Get table MetadataId| GET    | /EntityDefinitions(LogicalName='<table>')?$select=MetadataId      |
| Publish all         | POST   | /PublishAllXml  (body: {})                                        |

## Script Requirements
- Write as a single ES module file (.mjs), Node.js 18+
- Use native `fetch` only — no axios or node-fetch
- Use `execSync` from `child_process` to attempt Azure CLI token first
- Accept DATAVERSE_URL from process.env (exit with clear error if missing)
- Every create operation must check if the resource already exists first (idempotent)
- Phase 1: Create all tables + all non-lookup columns
- Phase 2: Create all lookup relationships (after all tables from Phase 1 exist)
- Call POST /PublishAllXml at the end of Phase 2
- Log each step clearly: ✓ created, ⏭ already exists, ✗ error

## Schema to Provision
<!-- REPLACE THIS SECTION WITH YOUR OWN TABLES -->

Tables:
  - Xyz_Project  (primary column: xyz_name "Project Name")
    Columns:
      - xyz_status       choice    [Active=100000000, Inactive=100000001]
      - xyz_startdate    datetime
      - xyz_enddate      datetime
      - xyz_budget       decimal   precision:2
      - xyz_description  text      maxLength:2000

  - Xyz_Task  (primary column: xyz_name "Task Name")
    Columns:
      - xyz_duedate      datetime
      - xyz_priority     choice    [High=100000000, Medium=100000001, Low=100000002]
      - xyz_completed    boolean   trueLabel:"Yes" falseLabel:"No"
      - xyz_project      lookup    → xyz_project   (Phase 2)

## Required Output
Generate a single self-contained .mjs script that:
1. Reads DATAVERSE_URL from env (exit with error if missing)
2. Resolves a Dataverse Bearer token (Azure CLI → MSAL device-code fallback)
3. Phase 1: Creates each table idempotently, then each non-lookup column idempotently
4. Phase 2: Creates each lookup relationship idempotently
5. Publishes all customizations (POST /PublishAllXml)
6. Prints a summary at the end: X created, Y skipped, Z failed
```

---

## Follow-Up Prompts

### Add tables to a Power Platform solution

```
After creating each table, retrieve its MetadataId:
  GET /api/data/v9.2/EntityDefinitions(LogicalName='<table>')?$select=MetadataId

Then add it to the solution:
  POST /api/data/v9.2/AddSolutionComponent
  Body: {
    ComponentType: 1,
    ComponentId: "<MetadataId>",
    SolutionUniqueName: "<yourSolutionUniqueName>",
    AddRequiredComponents: true,
    DoNotIncludeSubcomponents: false
  }
```

### Add seed / demo data after provisioning

```
After Phase 2, add a seedData() function that:
- POSTs sample records to each table's collection endpoint
  e.g. POST /api/data/v9.2/xyz_projects
- Before inserting, checks if the record already exists:
  GET /api/data/v9.2/xyz_projects?$filter=xyz_name eq '<name>'&$select=xyz_projectid
- Skips insertion if a matching record is found
- Logs ✓ seeded or ⏭ already exists for each record
```

### Add --dry-run support

```
Add a --dry-run CLI flag. When present, log every operation that would be performed
(including the full request body) but skip all POST/PATCH calls. Useful for
reviewing the schema before applying it to the environment.
```

---

## Running the Script

```bash
# Set your Dataverse org URL
export DATAVERSE_URL="https://<your-org>.crm.dynamics.com"

# Optional: provide a pre-fetched token to skip interactive auth
export DATAVERSE_TOKEN="<bearer-token>"

# Apply schema
node scripts/dataverse/apply-schema.mjs

# Dry run (review without making changes)
node scripts/dataverse/apply-schema.mjs --dry-run
```

---

## Common Errors & Fixes

| Error | Cause | Fix |
|-------|-------|-----|
| `0x80040217` entity not found | Wrong logical name | Use singular lowercase: `xyz_project` not `xyz_projects` |
| `0x80060888` bad query syntax | OData `$` chars not escaped in shell | Prefix as `\$filter` / `\$select` in shell strings |
| `409 AppLeaseActive` | App locked by another user session | Close all editor tabs, wait 15 min |
| `401 Unauthorized` | Wrong token resource | Run token command with `--resource <DATAVERSE_URL>` |
| `404 on RelationshipDefinitions` | Target table doesn't exist yet | Ensure Phase 1 completes before Phase 2 |

---

## Reference Links

- [Dataverse Web API overview](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/overview)
- [Create/update entity definitions](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/create-update-entity-definitions-using-web-api)
- [Code Apps ALM guide](https://learn.microsoft.com/en-us/power-apps/developer/code-apps/how-to/alm)
- [Power Platform CLI reference](https://learn.microsoft.com/en-us/power-platform/developer/cli/reference/code)
- [AddSolutionComponent action](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/reference/entities/solutioncomponent)
