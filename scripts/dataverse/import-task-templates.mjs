/**
 * import-task-templates.mjs — Create the pth_tasktemplate table and seed it
 * from the "Timing Templates" workbook (parsed to data/_task_templates.json).
 *
 * Table holds the editable per-project-type task standard:
 *   Name (=task) · Project Type · Sequence · Responsible · Leadtime (weeks) ·
 *   Inputs · Workload %
 *
 * Usage:
 *   DATAVERSE_URL=https://org91869c7b.crm4.dynamics.com node scripts/dataverse/import-task-templates.mjs [--seed-only]
 */
import { resolveToken } from './auth.mjs'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import process from 'node:process'

const dvUrl = process.env.DATAVERSE_URL || 'https://org91869c7b.crm4.dynamics.com'
const seedOnly = process.argv.includes('--seed-only')
const token = await resolveToken(dvUrl)
if (!token) { console.error('No token'); process.exit(1) }
const API = `${dvUrl.replace(/\/+$/, '')}/api/data/v9.2`
const ENTITY = 'pth_tasktemplate'

const headers = {
  Authorization: `Bearer ${token}`,
  'Content-Type': 'application/json; charset=utf-8',
  Accept: 'application/json',
  'OData-MaxVersion': '4.0', 'OData-Version': '4.0',
}
const label = (t) => ({ '@odata.type': 'Microsoft.Dynamics.CRM.Label', LocalizedLabels: [{ '@odata.type': 'Microsoft.Dynamics.CRM.LocalizedLabel', Label: t, LanguageCode: 1033 }] })

async function req(path, method, body) {
  const res = await fetch(`${API}/${path}`, { method, headers, body: body ? JSON.stringify(body) : undefined })
  if (!res.ok && res.status !== 204) throw new Error(`${method} ${path} (${res.status}): ${(await res.text()).slice(0, 400)}`)
  const txt = await res.text(); return txt ? JSON.parse(txt) : null
}
async function exists(path) { return (await fetch(`${API}/${path}`, { headers })).ok }

async function ensureTable() {
  if (await exists(`EntityDefinitions(LogicalName='${ENTITY}')?$select=LogicalName`)) {
    console.log('  · table already exists'); return
  }
  console.log('  creating table…')
  await req('EntityDefinitions', 'POST', {
    SchemaName: 'pth_TaskTemplate',
    DisplayName: label('Task Template'),
    DisplayCollectionName: label('Task Templates'),
    Description: label('Editable standard task per project type'),
    OwnershipType: 'UserOwned', IsActivity: false, HasActivities: false, HasNotes: false,
    PrimaryNameAttribute: 'pth_name',
    Attributes: [{
      '@odata.type': 'Microsoft.Dynamics.CRM.StringAttributeMetadata',
      SchemaName: 'pth_Name', LogicalName: 'pth_name', IsPrimaryName: true,
      DisplayName: label('Task'), Description: label('Task name'),
      MaxLength: 400, FormatName: { Value: 'Text' }, RequiredLevel: { Value: 'ApplicationRequired' },
    }],
  })
  // wait for provisioning
  for (let i = 0; i < 10; i++) {
    await new Promise((r) => setTimeout(r, 3000))
    if (await exists(`EntityDefinitions(LogicalName='${ENTITY}')?$select=LogicalName`)) break
  }
  console.log('  table created')
}

async function ensureColumns() {
  const cols = [
    { logical: 'pth_projecttype', p: { '@odata.type': '#Microsoft.Dynamics.CRM.StringAttributeMetadata', SchemaName: 'pth_ProjectType', DisplayName: label('Project Type'), Description: label('Template project type'), MaxLength: 100, FormatName: { Value: 'Text' }, RequiredLevel: { Value: 'ApplicationRequired' } } },
    { logical: 'pth_sequence', p: { '@odata.type': '#Microsoft.Dynamics.CRM.IntegerAttributeMetadata', SchemaName: 'pth_Sequence', DisplayName: label('Sequence'), Description: label('Order within the template'), MinValue: 0, MaxValue: 1000, RequiredLevel: { Value: 'None' } } },
    { logical: 'pth_responsible', p: { '@odata.type': '#Microsoft.Dynamics.CRM.StringAttributeMetadata', SchemaName: 'pth_Responsible', DisplayName: label('Responsible'), Description: label('Responsible role/area code(s)'), MaxLength: 200, FormatName: { Value: 'Text' }, RequiredLevel: { Value: 'None' } } },
    { logical: 'pth_leadtimeweeks', p: { '@odata.type': '#Microsoft.Dynamics.CRM.IntegerAttributeMetadata', SchemaName: 'pth_LeadtimeWeeks', DisplayName: label('Leadtime (weeks)'), Description: label('Lead time in calendar weeks'), MinValue: 0, MaxValue: 520, RequiredLevel: { Value: 'None' } } },
    { logical: 'pth_inputs', p: { '@odata.type': '#Microsoft.Dynamics.CRM.MemoAttributeMetadata', SchemaName: 'pth_Inputs', DisplayName: label('Inputs'), Description: label('Required inputs / deliverables'), MaxLength: 2000, Format: 'Text', RequiredLevel: { Value: 'None' } } },
    { logical: 'pth_workloadpct', p: { '@odata.type': '#Microsoft.Dynamics.CRM.IntegerAttributeMetadata', SchemaName: 'pth_WorkloadPct', DisplayName: label('Workload %'), Description: label('Effort estimate (% / hrs)'), MinValue: 0, MaxValue: 100000, RequiredLevel: { Value: 'None' } } },
  ]
  for (const c of cols) {
    if (await exists(`EntityDefinitions(LogicalName='${ENTITY}')/Attributes(LogicalName='${c.logical}')?$select=LogicalName`)) { console.log(`  · ${c.logical} exists`); continue }
    await req(`EntityDefinitions(LogicalName='${ENTITY}')/Attributes`, 'POST', c.p)
    console.log(`  ✓ ${c.logical}`)
  }
  await req('PublishAllXml', 'POST', {})
  await new Promise((r) => setTimeout(r, 10000))
}

async function entitySetName() {
  const d = await req(`EntityDefinitions(LogicalName='${ENTITY}')?$select=EntitySetName`, 'GET')
  return d.EntitySetName
}

async function run() {
  const data = JSON.parse(readFileSync(resolve(import.meta.dirname, '../../data/_task_templates.json'), 'utf8'))
  if (!seedOnly) { console.log('Schema:'); await ensureTable(); await ensureColumns() }
  const SET = await entitySetName()
  console.log('entity set:', SET)

  // clear existing rows (idempotent re-seed)
  const existing = await req(`${SET}?$select=pth_tasktemplateid&$top=5000`, 'GET')
  for (const r of existing.value) await fetch(`${API}/${SET}(${r.pth_tasktemplateid})`, { method: 'DELETE', headers })
  if (existing.value.length) console.log(`  cleared ${existing.value.length} existing rows`)

  let n = 0
  for (const t of data.types) {
    for (const task of t.tasks) {
      const body = {
        pth_name: task.task.slice(0, 400),
        pth_projecttype: t.type,
        pth_sequence: task.seq,
        pth_responsible: task.responsible.join(' / ') || null,
      }
      if (task.leadtimeWeeks != null) body.pth_leadtimeweeks = task.leadtimeWeeks
      if (task.inputs) body.pth_inputs = task.inputs.slice(0, 2000)
      if (task.workloadPct != null) body.pth_workloadpct = task.workloadPct
      await req(SET, 'POST', body)
      n++
    }
    console.log(`  ${t.type}: ${t.tasks.length}`)
  }
  console.log(`\n✅ Seeded ${n} task-template rows.`)
}
run().catch((e) => { console.error(e.message); process.exit(1) })
