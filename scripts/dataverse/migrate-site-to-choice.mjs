/**
 * migrate-site-to-choice.mjs
 *
 * 1. Creates a new choice column pth_SiteLocation on pth_projects
 *    with options: 100000000 = TCA (Toluca), 100000001 = SLP (San Luis Potosí)
 * 2. Reads every project row and maps old pth_site string → new choice value
 * 3. Patches each record with the new column
 *
 * Usage:
 *   DATAVERSE_URL=https://orgname.crm.dynamics.com node scripts/dataverse/migrate-site-to-choice.mjs
 */
import { resolveToken } from './auth.mjs'
import process from 'node:process'

const dataverseUrl = process.env.DATAVERSE_URL
if (!dataverseUrl) { console.error('Missing DATAVERSE_URL'); process.exit(1) }

const token = await resolveToken(dataverseUrl)
if (!token) { console.error('Could not obtain token'); process.exit(1) }

const API = `${dataverseUrl.replace(/\/+$/, '')}/api/data/v9.2`
const headers = {
  Authorization: `Bearer ${token}`,
  'Content-Type': 'application/json',
  Accept: 'application/json',
  'OData-MaxVersion': '4.0',
  'OData-Version': '4.0',
}

function label(text) {
  return {
    '@odata.type': 'Microsoft.Dynamics.CRM.Label',
    LocalizedLabels: [
      { '@odata.type': 'Microsoft.Dynamics.CRM.LocalizedLabel', Label: text, LanguageCode: 1033 },
    ],
  }
}

// ── Step 1: Create the choice column ────────────────────────────────────

console.log('Step 1: Creating pth_SiteLocation choice column…')

const columnBody = {
  '@odata.type': 'Microsoft.Dynamics.CRM.PicklistAttributeMetadata',
  SchemaName: 'pth_SiteLocation',
  DisplayName: label('Site Location'),
  Description: label('Plant site location (Toluca or San Luis Potosí)'),
  RequiredLevel: { Value: 'ApplicationRequired' },
  OptionSet: {
    IsGlobal: false,
    OptionSetType: 'Picklist',
    Options: [
      { Value: 100000000, Label: label('TCA') },
      { Value: 100000001, Label: label('SLP') },
    ],
  },
}

const createRes = await fetch(
  `${API}/EntityDefinitions(LogicalName='pth_project')/Attributes`,
  { method: 'POST', headers, body: JSON.stringify(columnBody) },
)

if (createRes.ok || createRes.status === 204) {
  console.log('  ✓ pth_SiteLocation column created')
} else {
  const errText = await createRes.text()
  if (errText.includes('already exists') || errText.includes('DuplicateAttributeSchemaName')) {
    console.log('  ✓ pth_SiteLocation already exists — skipping creation')
  } else {
    console.error(`  ✗ Failed (${createRes.status}):`, errText)
    process.exit(1)
  }
}

// ── Step 2: Migrate data ────────────────────────────────────────────────

console.log('\nStep 2: Migrating existing records…')

/** Map old string values → choice integer */
const STRING_TO_CHOICE = {
  tca: 100000000,
  slp: 100000001,
  toluca: 100000000,
  san_luis_potosi: 100000001,
  tip: 100000000,
  slpp: 100000001,
  stuttgart: 100000000,
  turin: 100000001,
  madrid: 100000000,
  wroclaw: 100000001,
}

const listRes = await fetch(
  `${API}/pth_projects?$select=pth_projectid,pth_projectname,pth_site`,
  { headers },
)
if (!listRes.ok) {
  console.error('Failed to list projects:', listRes.status, await listRes.text())
  process.exit(1)
}

const { value: projects } = await listRes.json()
console.log(`  Found ${projects.length} project(s)\n`)

let updated = 0
for (const p of projects) {
  const raw = (p.pth_site ?? '').toString().trim().toLowerCase().replace(/\s+/g, '_')
  const choiceVal = STRING_TO_CHOICE[raw]

  if (choiceVal == null) {
    console.warn(`  ? "${p.pth_projectname}" — unknown site "${p.pth_site}", skipping`)
    continue
  }

  const choiceLabel = choiceVal === 100000000 ? 'TCA' : 'SLP'
  console.log(`  → "${p.pth_projectname}": "${p.pth_site}" → ${choiceLabel} (${choiceVal})`)

  const patchRes = await fetch(`${API}/pth_projects(${p.pth_projectid})`, {
    method: 'PATCH',
    headers,
    body: JSON.stringify({ pth_sitelocation: choiceVal }),
  })

  if (patchRes.ok || patchRes.status === 204) {
    updated++
  } else {
    console.error(`    ✗ PATCH failed (${patchRes.status}):`, await patchRes.text())
  }
}

console.log(`\n✅ Migration complete — ${updated}/${projects.length} record(s) updated`)
