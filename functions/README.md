# PTH Notifications — Azure Functions + Microsoft Graph

Automated notification functions for the Project Tracker Hub.
Queries Dataverse for project/activity/milestone data and sends
emails via the Microsoft Graph API.

---

## Functions

| Function | Schedule | Purpose |
|---|---|---|
| **dailyOverdueScan** | Every day 07:00 UTC | Emails a report of activities past their end date that aren't DONE |
| **milestoneReminder** | Every day 07:30 UTC | Emails reminders for milestones due within the warning threshold |
| **weeklyDigest** | Every Monday 08:00 UTC | Full portfolio status digest with KPIs, progress, and highlights |

---

## Prerequisites

### 1. Entra ID App Registration

Create an app registration in the Azure portal with these **Application permissions**:

| API | Permission | Type |
|---|---|---|
| Dynamics CRM | `user_impersonation` | Application |
| Microsoft Graph | `Mail.Send` | Application |
| Microsoft Graph | `User.Read.All` | Application (optional, for user lookup) |

> Grant admin consent after adding the permissions.

### 2. App Settings / Environment Variables

| Setting | Description |
|---|---|
| `DATAVERSE_URL` | Your Dataverse org URL, e.g. `https://org2d99840c.crm.dynamics.com` |
| `AZURE_TENANT_ID` | Entra ID tenant ID |
| `AZURE_CLIENT_ID` | App registration client ID |
| `AZURE_CLIENT_SECRET` | App registration client secret |
| `NOTIFICATION_SENDER` | Mailbox to send from (e.g. `no-reply@inveenta.com`). Must be a real mailbox or shared mailbox. |
| `NOTIFICATION_OVERDUE_RECIPIENTS` | Comma-separated emails for overdue + milestone reminders |
| `NOTIFICATION_DIGEST_RECIPIENTS` | Comma-separated emails for the weekly digest |
| `WARNING_DAYS_THRESHOLD` | Days before a due date triggers a warning (default: `7`) |

---

## Local Development

```bash
cd functions

# Install dependencies
npm install

# Copy and fill in local settings
cp local.settings.json local.settings.json  # edit placeholders

# Build TypeScript
npm run build

# Run locally (requires Azure Functions Core Tools v4)
func start
```

Install Azure Functions Core Tools if needed:
```bash
brew install azure-functions-core-tools@4
```

---

## Deploy to Azure

### Option A: Azure CLI

```bash
# Create a Function App (Node.js 20, Consumption plan)
az functionapp create \
  --resource-group <rg-name> \
  --name pth-notifications \
  --consumption-plan-location westeurope \
  --runtime node \
  --runtime-version 20 \
  --functions-version 4 \
  --storage-account <storage-name>

# Configure app settings
az functionapp config appsettings set \
  --name pth-notifications \
  --resource-group <rg-name> \
  --settings \
    DATAVERSE_URL="https://org2d99840c.crm.dynamics.com" \
    AZURE_TENANT_ID="<tenant-id>" \
    AZURE_CLIENT_ID="<client-id>" \
    AZURE_CLIENT_SECRET="<client-secret>" \
    NOTIFICATION_SENDER="no-reply@inveenta.com" \
    NOTIFICATION_OVERDUE_RECIPIENTS="santiago.garciaruiz@inveenta.com" \
    NOTIFICATION_DIGEST_RECIPIENTS="santiago.garciaruiz@inveenta.com" \
    WARNING_DAYS_THRESHOLD="7"

# Deploy
cd functions
func azure functionapp publish pth-notifications
```

### Option B: VS Code Azure Functions Extension

1. Install the **Azure Functions** VS Code extension
2. Sign in to Azure
3. Right-click the `functions/` folder → **Deploy to Function App**
4. Configure app settings in the Azure portal

---

## Architecture

```
┌─────────────────────┐    OData/REST     ┌──────────────┐
│  Azure Functions     │ ◄──────────────── │  Dataverse   │
│  (Timer triggers)    │                   │  (6 tables)  │
│                      │                   └──────────────┘
│  dailyOverdueScan    │
│  milestoneReminder   │    Graph API      ┌──────────────┐
│  weeklyDigest        │ ─────────────────►│  Outlook /   │
│                      │                   │  Teams       │
└─────────────────────┘                   └──────────────┘
```

All three functions use `ClientSecretCredential` from `@azure/identity`
for token acquisition — one scope for Dataverse reads, another for
Graph API mail sends.
