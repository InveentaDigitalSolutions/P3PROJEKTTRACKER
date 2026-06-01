/**
 * import-resources.mjs — One-time import of the real resource roster.
 *
 *  1. Adds 4 columns to pth_resource: pth_Area (choice), pth_Location
 *     (multi-select choice), pth_OrgCode (string), pth_Manager (self lookup).
 *  2. Publishes customizations.
 *  3. Deletes existing (demo) resource rows.
 *  4. Creates managers, then associates linked to their manager.
 *
 * Reads normalized data from data/_resources.json (produced from the Excel).
 *
 * Usage:
 *   DATAVERSE_URL=https://org91869c7b.crm4.dynamics.com node scripts/dataverse/import-resources.mjs [--skip-schema]
 */
import { resolveToken } from './auth.mjs'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import process from 'node:process'

const dvUrl = process.env.DATAVERSE_URL || 'https://org91869c7b.crm4.dynamics.com'
const skipSchema = process.argv.includes('--skip-schema')
const token = await resolveToken(dvUrl)
if (!token) { console.error('No token'); process.exit(1) }
const API = `${dvUrl.replace(/\/+$/, '')}/api/data/v9.2`
const ENTITY = 'pth_resource'   // metadata logical name (singular)
const SET = 'pth_resources'      // entity set name (plural)

const headers = {
  Authorization: `Bearer ${token}`,
  'Content-Type': 'application/json; charset=utf-8',
  Accept: 'application/json',
  'OData-MaxVersion': '4.0',
  'OData-Version': '4.0',
}

const label = (text) => ({
  '@odata.type': 'Microsoft.Dynamics.CRM.Label',
  LocalizedLabels: [{ '@odata.type': 'Microsoft.Dynamics.CRM.LocalizedLabel', Label: text, LanguageCode: 1033 }],
})

const AREA = { PPS: 100000000, ENG: 100000001, QMM: 100000002, LOD: 100000003, LOP: 100000004, LOP4: 100000005, LOP5: 100000006, MSE: 100000007, CTG: 100000008, PUQ1: 100000009, PUQ2: 100000010 }
const LOC = { SlpP: 100000000, TlP: 100000001 }

async function req(path, method, body) {
  const res = await fetch(`${API}/${path}`, { method, headers, body: body ? JSON.stringify(body) : undefined })
  if (!res.ok && res.status !== 204) throw new Error(`${method} ${path} (${res.status}): ${(await res.text()).slice(0, 400)}`)
  const txt = await res.text()
  return txt ? JSON.parse(txt) : null
}
/** POST a row and return the created record (with its GUID). */
async function create(body) {
  const res = await fetch(`${API}/${SET}`, {
    method: 'POST',
    headers: { ...headers, Prefer: 'return=representation' },
    body: JSON.stringify(body),
  })
  if (!res.ok) throw new Error(`POST ${SET} (${res.status}): ${(await res.text()).slice(0, 400)}`)
  return res.json()
}
async function exists(path) {
  const res = await fetch(`${API}/${path}`, { headers })
  return res.ok
}

async function ensureSchema() {
  const optionSet = (opts) => ({ IsGlobal: false, OptionSetType: 'Picklist', Options: opts })

  const attrs = [
    {
      logical: 'pth_area',
      payload: {
        '@odata.type': '#Microsoft.Dynamics.CRM.PicklistAttributeMetadata',
        SchemaName: 'pth_Area', DisplayName: label('Area'), Description: label('Functional area / department'),
        RequiredLevel: { Value: 'None' },
        OptionSet: optionSet(Object.entries(AREA).map(([k, v]) => ({ Value: v, Label: label(k) }))),
      },
    },
    {
      logical: 'pth_location',
      payload: {
        '@odata.type': '#Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata',
        SchemaName: 'pth_Location', DisplayName: label('Location'), Description: label('Site location(s)'),
        RequiredLevel: { Value: 'None' },
        OptionSet: optionSet(Object.entries(LOC).map(([k, v]) => ({ Value: v, Label: label(k) }))),
      },
    },
    {
      logical: 'pth_orgcode',
      payload: {
        '@odata.type': '#Microsoft.Dynamics.CRM.StringAttributeMetadata',
        SchemaName: 'pth_OrgCode', DisplayName: label('Org Code'), Description: label('Organizational / cost-center code'),
        RequiredLevel: { Value: 'None' }, MaxLength: 100, FormatName: { Value: 'Text' },
      },
    },
  ]

  for (const a of attrs) {
    if (await exists(`EntityDefinitions(LogicalName='${ENTITY}')/Attributes(LogicalName='${a.logical}')?$select=LogicalName`)) {
      console.log(`  · ${a.logical} already exists`); continue
    }
    await req(`EntityDefinitions(LogicalName='${ENTITY}')/Attributes`, 'POST', a.payload)
    console.log(`  ✓ created ${a.logical}`)
  }

  // Self-referential manager lookup via OneToMany relationship
  if (await exists(`EntityDefinitions(LogicalName='${ENTITY}')/Attributes(LogicalName='pth_manager')?$select=LogicalName`)) {
    console.log('  · pth_manager lookup already exists')
  } else {
    await req('RelationshipDefinitions', 'POST', {
      '@odata.type': 'Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata',
      SchemaName: `pth_Manager_${ENTITY}_${ENTITY}`,
      ReferencedEntity: ENTITY, ReferencingEntity: ENTITY,
      CascadeConfiguration: { Assign: 'NoCascade', Delete: 'RemoveLink', Merge: 'NoCascade', Reparent: 'NoCascade', Share: 'NoCascade', Unshare: 'NoCascade' },
      Lookup: { SchemaName: 'pth_Manager', DisplayName: label('Manager'), Description: label('Reporting manager'), RequiredLevel: { Value: 'None' } },
    })
    console.log('  ✓ created pth_manager lookup')
  }

  console.log('  publishing…')
  await req('PublishAllXml', 'POST', {})
  await new Promise((r) => setTimeout(r, 12000))
  console.log('  published.')
}

async function deleteAllResources() {
  let n = 0
  let url = `${SET}?$select=pth_resourceid&$top=5000`
  const ids = []
  while (url) {
    const res = await fetch(`${API}/${url}`, { headers })
    const json = await res.json()
    for (const r of json.value) ids.push(r.pth_resourceid)
    url = json['@odata.nextLink'] ? json['@odata.nextLink'].replace(`${API}/`, '') : null
  }
  for (const id of ids) {
    await fetch(`${API}/${SET}(${id})`, { method: 'DELETE', headers })
    n++
  }
  console.log(`  deleted ${n} existing resource rows`)
}

async function run() {
  const data = JSON.parse(readFileSync(resolve(import.meta.dirname, '../../data/_resources.json'), 'utf8'))

  if (!skipSchema) { console.log('Schema:'); await ensureSchema() }
  else console.log('Skipping schema step (--skip-schema)')

  console.log('Clearing existing resources…'); await deleteAllResources()

  let seq = 0
  const pad = (n) => String(n).padStart(4, '0')

  console.log('Creating managers…')
  const mgrId = {}   // manager name → resource GUID
  for (const m of data.managers) {
    const body = {
      pth_personidexternal: `RES-${pad(++seq)}`,
      pth_name: m.name,
      pth_role: 'Manager',
      pth_department: m.area,
      pth_weeklycapacityhours: 40,
      pth_area: AREA[m.area],
      pth_location: m.locs.map((l) => LOC[l]).join(','),
      pth_orgcode: m.orgcode || '',
    }
    const row = await create(body)
    mgrId[m.name] = row?.pth_resourceid
  }
  console.log(`  ${data.managers.length} managers`)

  console.log('Creating associates…')
  let linked = 0, unlinked = 0
  for (const a of data.associates) {
    const body = {
      pth_personidexternal: `RES-${pad(++seq)}`,
      pth_name: a.name,
      pth_role: 'Associate',
      pth_department: a.area,
      pth_weeklycapacityhours: 40,
      pth_area: AREA[a.area],
      pth_location: a.locs.map((l) => LOC[l]).join(','),
    }
    const mid = mgrId[a.manager]
    if (mid) { body['pth_Manager@odata.bind'] = `/${SET}(${mid})`; linked++ } else { unlinked++ }
    await create(body)
  }
  console.log(`  ${data.associates.length} associates (${linked} linked to a manager, ${unlinked} unlinked)`)
  console.log('\n✅ Import complete.')
}

run().catch((e) => { console.error(e.message); process.exit(1) })
