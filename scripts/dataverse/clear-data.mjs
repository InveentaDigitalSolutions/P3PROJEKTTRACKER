/**
 * clear-data.mjs — Delete all rows from the PTH data tables (FK-safe order).
 *
 * Deletes in dependency order so lookups never block a delete:
 *   Assignments → Activities → Milestones → Resources → Projects
 * PPMSettings (config) is intentionally NOT touched.
 *
 * Usage:
 *   DATAVERSE_URL=https://orgname.crm.dynamics.com node scripts/dataverse/clear-data.mjs --dry-run
 *   DATAVERSE_URL=https://orgname.crm.dynamics.com node scripts/dataverse/clear-data.mjs
 */
import { resolveToken } from './auth.mjs'
import process from 'node:process'

const dataverseUrl = process.env.DATAVERSE_URL
if (!dataverseUrl) {
  console.error('Missing DATAVERSE_URL env var')
  process.exit(1)
}

const dryRun = process.argv.includes('--dry-run')

const token = await resolveToken(dataverseUrl)
if (!token) {
  console.error('Could not obtain a Dataverse access token.')
  process.exit(1)
}

const API = `${dataverseUrl.replace(/\/+$/, '')}/api/data/v9.2`

const headers = {
  Authorization: `Bearer ${token}`,
  'Content-Type': 'application/json; charset=utf-8',
  Accept: 'application/json',
  'OData-MaxVersion': '4.0',
  'OData-Version': '4.0',
}

// entity set → primary-key attribute, in FK-safe delete order
const TABLES = [
  { set: 'pth_assignments', key: 'pth_assignmentid', label: 'Assignments' },
  { set: 'pth_activities', key: 'pth_activityid', label: 'Activities' },
  { set: 'pth_milestones', key: 'pth_milestoneid', label: 'Milestones' },
  { set: 'pth_projects', key: 'pth_projectid', label: 'Projects' },
  // pth_resources intentionally excluded — keeps the resource/people list.
  // pth_ppmsettings intentionally excluded — keeps the app's config row.
]

async function fetchAllIds(set, key) {
  const ids = []
  let url = `${API}/${set}?$select=${key}&$top=5000`
  while (url) {
    const res = await fetch(url, { headers })
    if (!res.ok) throw new Error(`GET ${set} failed (${res.status}): ${await res.text()}`)
    const json = await res.json()
    for (const row of json.value) ids.push(row[key])
    url = json['@odata.nextLink'] || null
  }
  return ids
}

async function del(set, id) {
  const res = await fetch(`${API}/${set}(${id})`, { method: 'DELETE', headers })
  if (!res.ok && res.status !== 404) {
    throw new Error(`DELETE ${set}(${id}) failed (${res.status}): ${await res.text()}`)
  }
}

async function run() {
  console.log(dryRun ? '— DRY RUN (no deletes) —\n' : '— DELETING DATA ROWS —\n')
  let grandTotal = 0

  for (const t of TABLES) {
    const ids = await fetchAllIds(t.set, t.key)
    grandTotal += ids.length
    if (dryRun) {
      console.log(`  ${t.label.padEnd(12)} ${ids.length} rows`)
      continue
    }
    process.stdout.write(`  ${t.label.padEnd(12)} deleting ${ids.length}…`)
    let done = 0
    for (const id of ids) {
      await del(t.set, id)
      done++
      if (done % 25 === 0) process.stdout.write(` ${done}`)
    }
    console.log(` ✓`)
  }

  console.log(`\n${dryRun ? 'Would delete' : 'Deleted'} ${grandTotal} rows across ${TABLES.length} tables.`)
  console.log('PPMSettings left untouched.')
}

run().catch((err) => {
  console.error(err.message)
  process.exit(1)
})
