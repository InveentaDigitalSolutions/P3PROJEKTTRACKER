# Connecting to Dataverse & Managing Tables with Claude Code

This guide documents how we connected to a Microsoft Dataverse environment using
a personal Microsoft (Entra ID) account and used Claude Code to create, edit,
and delete tables — all from the terminal, with no portal clicks required.

---

## Prerequisites

| Requirement | Details |
|---|---|
| **Node.js** | v18+ (for native `fetch` support) |
| **Azure CLI** | `az` — used for quick token acquisition |
| **MSAL package** | `@azure/msal-node` — fallback when Azure CLI isn't signed in |
| **Dataverse environment** | An org URL like `https://orgname.crm.dynamics.com` |
| **Security role** | Your Microsoft account needs *System Customizer* (or higher) in the target environment |

---

## Step 1 — Set Environment Variables

```bash
# Your Dataverse org URL (no trailing slash)
export DATAVERSE_URL="https://<your-org>.crm.dynamics.com"
```

Optionally, if you already have a Bearer token:

```bash
export DATAVERSE_TOKEN="eyJ0eXAi..."
```

---

## Step 2 — Authenticate

Authentication is handled automatically by the shared `scripts/dataverse/auth.mjs`
module. It tries three strategies in order:

### Strategy 1 — `DATAVERSE_TOKEN` env var

If you've set `DATAVERSE_TOKEN`, it is used immediately. Useful for CI or when you
already grabbed a token from another tool.

### Strategy 2 — Azure CLI (`az`)

If you're signed in to the Azure CLI, the script runs:

```bash
az account get-access-token --resource https://orgname.crm.dynamics.com --query accessToken -o tsv
```

This silently returns a token scoped to your Dataverse environment — no browser
interaction needed.

### Strategy 3 — MSAL Device-Code Flow

If the Azure CLI isn't available or isn't signed in, the script falls back to
Microsoft's device-code flow:

1. The script prints a URL and a one-time code to the terminal.
2. You open the URL in a browser, sign in with your Microsoft account, and enter the code.
3. The token is cached on disk (`/tmp/pth-msal-cache.json`) so subsequent runs
   authenticate silently — you only need to complete the browser flow once.

The client ID used is `51f81489-12ee-4a9e-aaae-a2591f45987d` (Power Apps first-party
public client), with authority `https://login.microsoftonline.com/organizations`.

---

## Step 3 — Define Your Schema

All tables are defined in a single JSON file:

```
dataverse/ppm.dataverse.schema.json
```

This schema declares:

- **Publisher prefix**: `pth`
- **Solution**: `PTH_PortfolioManagement`
- **6 tables**: Project, Milestone, Activity, Resource, Assignment, PPMSetting

Each table definition includes columns with their types (string, memo, choice,
decimal, datetime, boolean, lookup), constraints, and choice option values.

### Table naming conventions

| Concept | Pattern | Example |
|---|---|---|
| Schema name | PascalCase with prefix | `pth_Project` |
| Logical name | lowercase with prefix | `pth_project` |
| Collection (API) | plural lowercase | `pth_projects` |
| Primary column | `pth_name` | String, required |

---

## Step 4 — Create Tables (Apply Schema)

Run the provisioning script:

```bash
node scripts/dataverse/apply-schema.mjs
```

Or preview first without making changes:

```bash
node scripts/dataverse/apply-schema.mjs --dry-run
```

### What the script does

1. **Phase 1 — Tables & columns**: Creates each table and its non-lookup columns.
   Every operation is idempotent — it checks whether the table/column already exists
   before creating it.

2. **Phase 2 — Relationships**: Creates lookup columns (e.g., Milestone → Project,
   Activity → Milestone) after all target tables exist.

3. **Publish**: Calls `POST /api/data/v9.2/PublishAllXml` to make customizations
   visible in the environment.

Each step logs its result:
- `✓ created` — new resource created
- `⏭ already exists` — skipped (idempotent)
- `✗ error` — something went wrong (with details)

### API endpoints used

| Operation | Method | Endpoint |
|---|---|---|
| Create table | POST | `/api/data/v9.2/EntityDefinitions` |
| Create column | POST | `/api/data/v9.2/EntityDefinitions(LogicalName='<table>')/Attributes` |
| Create relationship | POST | `/api/data/v9.2/RelationshipDefinitions` |
| Check if table exists | GET | `/EntityDefinitions(LogicalName='<table>')?$select=LogicalName` |
| Publish customizations | POST | `/api/data/v9.2/PublishAllXml` |

---

## Step 5 — Seed Demo Data

After the schema is in place, populate sample records:

```bash
node scripts/dataverse/seed-data.mjs
```

This creates resources, projects, milestones, activities, and assignments with
realistic sample data. Lookups are bound using OData syntax:

```json
{ "pth_Project@odata.bind": "/pth_projects(<project-id>)" }
```

---

## Step 6 — CRUD from the App (Runtime)

Once tables exist, the React app performs CRUD via auto-generated service classes
in `src/generated/services/`. These use the `@microsoft/power-apps/data` SDK.

### Example: Projects

```typescript
import { Pth_projectsService } from './generated/services/Pth_projectsService';

// Create
await Pth_projectsService.create({ pth_name: 'New Project', pth_category: 100000000 });

// Read all
const all = await Pth_projectsService.getAll();

// Read one
const project = await Pth_projectsService.get('<guid>');

// Update
await Pth_projectsService.update('<guid>', { pth_name: 'Updated Name' });

// Delete
await Pth_projectsService.delete('<guid>');
```

Every entity (projects, milestones, activities, resources, assignments, ppmsettings)
has an identical service with `create`, `get`, `getAll`, `update`, and `delete`.

---

## Step 7 — Inspect & Troubleshoot

Additional utility scripts are available:

| Script | Purpose |
|---|---|
| `scripts/dataverse/_check-tables.mjs` | Verify tables exist in the environment |
| `scripts/dataverse/inspect-tables.mjs` | Dump column metadata for debugging |
| `scripts/dataverse/add-columns.mjs` | Add new columns to existing tables |
| `scripts/dataverse/fix-sites.mjs` | One-off migration fix for site column |

---

## Common Errors

| Error | Cause | Fix |
|---|---|---|
| `0x80040217` entity not found | Wrong logical name | Use singular lowercase: `pth_project` not `pth_projects` |
| `0x80060888` bad query syntax | Unescaped `$` in shell | Prefix as `\$filter` in shell strings |
| `401 Unauthorized` | Token expired or wrong resource | Re-run `az login` or delete `/tmp/pth-msal-cache.json` to re-auth |
| `404 on RelationshipDefinitions` | Target table doesn't exist yet | Ensure Phase 1 completes before Phase 2 |

---

## Summary

The entire Dataverse schema lifecycle — authentication, table creation, column
provisioning, relationship wiring, seed data, and runtime CRUD — is handled
through scripts and generated services, driven by a single JSON schema file.
No portal interaction is required beyond the initial device-code sign-in.
