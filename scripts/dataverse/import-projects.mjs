/**
 * import-projects.mjs — Bulk-import projects + their tasks from the parsed
 * area/location workbooks (data/_projects.json).
 *
 *  1. Adds columns: pth_projecttype (pth_project); pth_responsible,
 *     pth_leadtimeweeks, pth_inputs, pth_workloadpct (pth_activity).
 *  2. Creates each project (PM = area manager, site from location, category
 *     from type) and its tasks (unassigned owner, NOT_STARTED).
 *
 * Usage:
 *   DATAVERSE_URL=https://org91869c7b.crm4.dynamics.com node scripts/dataverse/import-projects.mjs [--skip-schema]
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
const headers = {
  Authorization: `Bearer ${token}`,
  'Content-Type': 'application/json; charset=utf-8',
  Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0',
}
const label = (t) => ({ '@odata.type': 'Microsoft.Dynamics.CRM.Label', LocalizedLabels: [{ '@odata.type': 'Microsoft.Dynamics.CRM.LocalizedLabel', Label: t, LanguageCode: 1033 }] })

const CATEGORY = { ECR: 100000000, 'Path Forward': 100000001, CIP: 100000002, 'New Programs': 100000003 }
const SITE = { TCA: 100000000, SLP: 100000001 }
const AREA_NAME = { 100000000: 'PPS', 100000001: 'ENG', 100000002: 'QMM', 100000003: 'LOD', 100000004: 'LOP', 100000005: 'LOP4', 100000006: 'LOP5', 100000007: 'MSE', 100000008: 'CTG', 100000009: 'PUQ1', 100000010: 'PUQ2' }
const LOC_NAME = { 100000000: 'SlpP', 100000001: 'TlP' }
const ACT_STATUS_NOTSTARTED = 100000000
const RYG_GREEN = 100000002

async function req(path, method, body, extra) {
  const res = await fetch(`${API}/${path}`, { method, headers: { ...headers, ...extra }, body: body ? JSON.stringify(body) : undefined })
  if (!res.ok && res.status !== 204) throw new Error(`${method} ${path} (${res.status}): ${(await res.text()).slice(0, 300)}`)
  const txt = await res.text(); return txt ? JSON.parse(txt) : null
}
async function exists(path) { return (await fetch(`${API}/${path}`, { headers })).ok }
async function addCol(entity, logical, payload) {
  if (await exists(`EntityDefinitions(LogicalName='${entity}')/Attributes(LogicalName='${logical}')?$select=LogicalName`)) { console.log(`  · ${entity}.${logical} exists`); return }
  await req(`EntityDefinitions(LogicalName='${entity}')/Attributes`, 'POST', payload)
  console.log(`  ✓ ${entity}.${logical}`)
}

async function ensureSchema() {
  await addCol('pth_project', 'pth_projecttype', { '@odata.type': '#Microsoft.Dynamics.CRM.StringAttributeMetadata', SchemaName: 'pth_ProjectType', DisplayName: label('Project Type'), Description: label('Template project type'), MaxLength: 100, FormatName: { Value: 'Text' }, RequiredLevel: { Value: 'None' } })
  await addCol('pth_activity', 'pth_responsible', { '@odata.type': '#Microsoft.Dynamics.CRM.StringAttributeMetadata', SchemaName: 'pth_Responsible', DisplayName: label('Responsible'), Description: label('Responsible role/area code(s)'), MaxLength: 200, FormatName: { Value: 'Text' }, RequiredLevel: { Value: 'None' } })
  await addCol('pth_activity', 'pth_leadtimeweeks', { '@odata.type': '#Microsoft.Dynamics.CRM.IntegerAttributeMetadata', SchemaName: 'pth_LeadtimeWeeks', DisplayName: label('Leadtime (weeks)'), Description: label('Lead time in calendar weeks'), MinValue: 0, MaxValue: 520, RequiredLevel: { Value: 'None' } })
  await addCol('pth_activity', 'pth_inputs', { '@odata.type': '#Microsoft.Dynamics.CRM.MemoAttributeMetadata', SchemaName: 'pth_Inputs', DisplayName: label('Inputs'), Description: label('Required inputs / deliverables'), MaxLength: 2000, Format: 'Text', RequiredLevel: { Value: 'None' } })
  await addCol('pth_activity', 'pth_workloadpct', { '@odata.type': '#Microsoft.Dynamics.CRM.IntegerAttributeMetadata', SchemaName: 'pth_WorkloadPct', DisplayName: label('Workload %'), Description: label('Effort estimate'), MinValue: 0, MaxValue: 100000, RequiredLevel: { Value: 'None' } })
  console.log('  publishing…')
  await req('PublishAllXml', 'POST', {})
  await new Promise((r) => setTimeout(r, 12000))
}

async function loadManagers() {
  const json = await req(`pth_resources?$select=pth_name,pth_area,pth_location&$filter=pth_role eq 'Manager'&$top=500`, 'GET')
  return json.value.map((r) => ({
    name: r.pth_name,
    area: AREA_NAME[r.pth_area],
    locs: typeof r.pth_location === 'string' ? r.pth_location.split(',').map((v) => LOC_NAME[Number(v)]) : [],
  }))
}
function pickManager(managers, area, loc) {
  const want = loc === 'SLP' ? 'SlpP' : loc === 'TLP' ? 'TlP' : ''
  const inArea = managers.filter((m) => m.area === area)
  return (inArea.find((m) => m.locs.includes(want)) || inArea[0])?.name || ''
}
function catFor(type) {
  if (type === 'CIP') return 'CIP'
  if (type.includes('ECR')) return 'ECR'
  return 'New Programs'
}
function extCode(name, area, loc, idx) {
  const m = name.match(/3E\d{6,}/) || name.match(/\bMP\d+\b/)
  return m ? m[0] : `PRJ-${area}${loc}-${String(idx + 1).padStart(3, '0')}`
}
const todayISO = () => new Date().toISOString().slice(0, 10)
function projDates(tasks) {
  const starts = tasks.map((t) => t.start).filter(Boolean).sort()
  const finishes = tasks.map((t) => t.finish).filter(Boolean).sort()
  return { start: starts[0] || todayISO(), end: finishes[finishes.length - 1] || starts[starts.length - 1] || todayISO() }
}

async function create(set, body) {
  const res = await fetch(`${API}/${set}`, { method: 'POST', headers: { ...headers, Prefer: 'return=representation' }, body: JSON.stringify(body) })
  if (!res.ok) throw new Error(`POST ${set} (${res.status}): ${(await res.text()).slice(0, 300)}`)
  return res.json()
}

async function run() {
  const projects = JSON.parse(readFileSync(resolve(import.meta.dirname, '../../data/_projects.json'), 'utf8'))
  if (!skipSchema) { console.log('Schema:'); await ensureSchema() }
  const managers = await loadManagers()
  console.log(`Loaded ${managers.length} managers for PM assignment`)

  let np = 0, nt = 0
  const noPm = new Set()
  for (const [idx, p] of projects.entries()) {
    const pm = pickManager(managers, p.area, p.location)
    if (!pm) noPm.add(p.area)
    const d = projDates(p.tasks)
    const pBody = {
      pth_projectidexternal: extCode(p.name, p.area, p.location, idx),
      pth_projectname: p.name.slice(0, 200),
      pth_category: CATEGORY[catFor(p.projectType)],
      pth_sitelocation: SITE[p.location] ?? SITE.SLP,
      pth_projectobjective: (p.name || 'Imported project').slice(0, 2000),
      pth_projectmanagername: pm || 'Unassigned',
      pth_timestatus: RYG_GREEN,
      pth_criticalpathchangedflag: false,
      pth_projecttype: p.projectType,
      pth_plannedstartdate: d.start,
      pth_plannedenddate: d.end,
    }
    const proj = await create('pth_projects', pBody)
    np++
    for (const t of p.tasks) {
      const start = t.start || d.start
      const end = t.finish || start
      const aBody = {
        pth_activityidexternal: `ACT-${proj.pth_projectid.slice(0, 8)}-${nt}`,
        pth_name: t.task.slice(0, 200),
        pth_owner: 'Unassigned',
        pth_startdate: start,
        pth_enddate: end,
        pth_plannedenddate: end,
        pth_status: ACT_STATUS_NOTSTARTED,
        pth_ryg: RYG_GREEN,
        pth_iscriticalpath: false,
        pth_responsible: t.responsible || null,
        pth_inputs: t.inputs ? t.inputs.slice(0, 2000) : null,
        'pth_Project@odata.bind': `/pth_projects(${proj.pth_projectid})`,
      }
      if (t.leadtimeWeeks != null) aBody.pth_leadtimeweeks = t.leadtimeWeeks
      if (t.workloadPct != null) aBody.pth_workloadpct = t.workloadPct
      await create('pth_activities', aBody)
      nt++
    }
    console.log(`  [${p.area}/${p.location}] ${p.name.slice(0, 45)} — ${p.tasks.length} tasks${pm ? '' : ' (no PM)'}`)
  }
  console.log(`\n✅ Imported ${np} projects, ${nt} tasks.`)
  if (noPm.size) console.log(`⚠ No baseline manager for areas: ${[...noPm].join(', ')} → PM left "Unassigned".`)
}
run().catch((e) => { console.error(e.message); process.exit(1) })
