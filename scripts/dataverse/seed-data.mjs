/**
 * seed-data.mjs — Insert dummy rows into the Dataverse tables
 * provisioned by apply-schema.mjs.
 *
 * Usage:
 *   DATAVERSE_URL=https://orgname.crm.dynamics.com node scripts/dataverse/seed-data.mjs
 */
import { resolveToken } from './auth.mjs'
import process from 'node:process'

const dataverseUrl = process.env.DATAVERSE_URL
if (!dataverseUrl) {
  console.error('Missing DATAVERSE_URL env var')
  process.exit(1)
}

const token = await resolveToken(dataverseUrl)
if (!token) {
  console.error('Could not obtain a Dataverse access token.')
  process.exit(1)
}

// ─── Helpers ───────────────────────────────────────────────────────────────
const API = `${dataverseUrl.replace(/\/+$/, '')}/api/data/v9.2`

async function post(entity, body) {
  const res = await fetch(`${API}/${entity}`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json; charset=utf-8',
      Accept: 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      Prefer: 'return=representation',
    },
    body: JSON.stringify(body),
  })
  if (!res.ok) {
    const err = await res.text()
    throw new Error(`POST ${entity} failed (${res.status}): ${err}`)
  }
  return res.json()
}

function addDays(base, days) {
  const d = new Date(`${base}T00:00:00Z`)
  d.setUTCDate(d.getUTCDate() + days)
  return d.toISOString().slice(0, 10)
}

function todayISO() { return new Date().toISOString().slice(0, 10) }

function weekStart(iso) {
  const d = new Date(`${iso}T00:00:00Z`)
  const day = d.getUTCDay()
  const shift = day === 0 ? -6 : 1 - day
  d.setUTCDate(d.getUTCDate() + shift)
  return d.toISOString().slice(0, 10)
}

// ─── Static reference data (matches src/App.tsx INITIAL_PEOPLE) ────────
const CATEGORY_VALUE = { ECR: 100000000, 'Path Forward': 100000001, CIP: 100000002, 'New Programs': 100000003 }
const TIME_STATUS_VALUE = { RED: 100000000, YELLOW: 100000001, GREEN: 100000002 }
const MILESTONE_STATUS_VALUE = { OPEN: 100000000, IN_PROGRESS: 100000001, CLOSED: 100000002, DELAYED: 100000003 }
const RYG_VALUE = { RED: 100000000, YELLOW: 100000001, GREEN: 100000002 }
const ACTIVITY_STATUS_VALUE = { NOT_STARTED: 100000000, IN_PROGRESS: 100000001, BLOCKED: 100000002, CLOSED: 100000003 }

const CATEGORY_MILESTONES = {
  ECR: [
    { name: 'ECR Kickoff', offset: -21 },
    { name: 'Engineering Release Review', offset: 7 },
    { name: 'ECR Implementation Complete', offset: 42 },
  ],
  CIP: [
    { name: 'CIP Baseline', offset: -14 },
    { name: 'CIP Improvement Gate', offset: 10 },
    { name: 'CIP Sustainment Review', offset: 35 },
  ],
  'Path Forward': [
    { name: 'Path Forward Alignment', offset: -10 },
    { name: 'Path Forward Validation', offset: 14 },
    { name: 'Path Forward Sign-Off', offset: 45 },
  ],
  'New Programs': [
    { name: 'Program Kickoff', offset: -7 },
    { name: 'Prototype Gate', offset: 21 },
    { name: 'SOP Readiness', offset: 63 },
  ],
}

// ─── 1. Resources (matches INITIAL_PEOPLE who have role RESOURCE) ──────
const resources = [
  { extId: 'u_res_1', name: 'Nina Becker', role: 'Engineer', department: 'Assembly', capacity: 40 },
  { extId: 'u_res_2', name: 'Viktor Hahn', role: 'Engineer', department: 'Assembly', capacity: 40 },
  { extId: 'u_res_3', name: 'Iris Weber', role: 'Test Technician', department: 'Quality', capacity: 40 },
  { extId: 'u_res_4', name: 'Pablo Diaz', role: 'Process Engineer', department: 'Manufacturing', capacity: 40 },
  { extId: 'u_res_5', name: 'Elena Costa', role: 'Design Engineer', department: 'R&D', capacity: 40 },
  { extId: 'u_res_6', name: 'Marco Lenz', role: 'Supply Chain Lead', department: 'Logistics', capacity: 40 },
]

// ─── 2. Projects (matches generateInitialState in App.tsx) ─────────────
const projects = [
  {
    extId: 'p_1', name: 'Assembly Line CIP – TCA Station 12', category: 'CIP',
    siteLocation: 100000000, sponsor: 'Laura Schneider', pm: 'Rafael Gomez',
    objective: 'Reduce cycle time in final assembly by 8% without quality drift.',
  },
  {
    extId: 'p_2', name: 'ESC Sensor New Program', category: 'New Programs',
    siteLocation: 100000001, sponsor: 'Laura Schneider', pm: 'Sven Maurer',
    objective: 'Prepare ESC sensor manufacturing and validation readiness for SOP.',
  },
  {
    extId: 'p_3', name: 'Brake Caliper Path Forward', category: 'Path Forward',
    siteLocation: 100000000, sponsor: 'Laura Schneider', pm: 'Clara Rossi',
    objective: 'Stabilize backlog and phase-in process controls in brake caliper line.',
  },
  {
    extId: 'p_4', name: 'ABS Module ECR', category: 'ECR',
    siteLocation: 100000001, sponsor: 'Laura Schneider', pm: 'Rafael Gomez',
    objective: 'Implement ABS module engineering change with contained launch.',
  },
  {
    extId: 'p_5', name: 'Sparkplugs Manufacturing CIP', category: 'CIP',
    siteLocation: 100000000, sponsor: 'Laura Schneider', pm: 'Sven Maurer',
    objective: 'Improve scrap reduction and process capability in sparkplug production.',
  },
  {
    extId: 'p_6', name: 'ADAS Path Forward – Radar Cell', category: 'Path Forward',
    siteLocation: 100000001, sponsor: 'Laura Schneider', pm: 'Clara Rossi',
    objective: 'Recover ADAS radar cell throughput and delivery adherence.',
  },
]

// ─── Run ───────────────────────────────────────────────────────────────────
async function run() {
  const today = todayISO()

  // ── Resources ──────────────────────────────────────────────────────────
  console.log('Creating resources…')
  const resourceMap = {} // extId → Dataverse record id
  for (const r of resources) {
    const row = await post('pth_resources', {
      pth_personidexternal: r.extId,
      pth_name: r.name,
      pth_role: r.role,
      pth_department: r.department,
      pth_weeklycapacityhours: r.capacity,
    })
    resourceMap[r.extId] = row.pth_resourceid
    console.log(`  Resource: ${r.name}  →  ${row.pth_resourceid}`)
  }

  // ── Projects ───────────────────────────────────────────────────────────
  console.log('Creating projects…')
  const projectMap = {} // extId → Dataverse record id
  for (const p of projects) {
    const row = await post('pth_projects', {
      pth_projectidexternal: p.extId,
      pth_projectname: p.name,                // <-- primary name column
      pth_category: CATEGORY_VALUE[p.category],
      pth_sitelocation: p.siteLocation,
      pth_projectobjective: p.objective,
      pth_sponsorexecutive: p.sponsor,
      pth_projectmanagername: p.pm,
      pth_timestatus: TIME_STATUS_VALUE.GREEN,
      pth_criticalpathchangedflag: false,
    })
    projectMap[p.extId] = row.pth_projectid
    console.log(`  Project: ${p.name}  →  ${row.pth_projectid}`)
  }

  // ── Milestones ─────────────────────────────────────────────────────────
  console.log('Creating milestones…')
  const milestoneRows = []
  for (const [pIdx, p] of projects.entries()) {
    const templates = CATEGORY_MILESTONES[p.category]
    for (const [tIdx, tmpl] of templates.entries()) {
      const shifted = tmpl.offset + pIdx * 2 + tIdx * 3
      const targetDate = addDays(today, shifted)
      const isPast = new Date(targetDate) < new Date(today)
      const status = isPast ? 'CLOSED' : tIdx === 0 ? 'IN_PROGRESS' : 'OPEN'
      const ryg = isPast ? 'GREEN' : shifted <= 7 ? 'YELLOW' : 'GREEN'

      const row = await post('pth_milestones', {
        pth_milestoneidexternal: `ms_${p.extId}_${tIdx}`,
        pth_name: tmpl.name,                  // <-- primary name column
        pth_planneddate: targetDate,
        pth_status: MILESTONE_STATUS_VALUE[status],
        pth_ryg: RYG_VALUE[ryg],
        pth_iscriticalpath: tIdx === 1,
        // Lookup: bind to project
        'pth_Project@odata.bind': `/pth_projects(${projectMap[p.extId]})`,
      })
      milestoneRows.push({ ...row, _projectExtId: p.extId, _tIdx: tIdx })
      console.log(`  Milestone: ${tmpl.name} (${p.name})  →  ${row.pth_milestoneid}`)
    }
  }

  // ── Activities ─────────────────────────────────────────────────────────
  console.log('Creating activities…')
  const activityRows = []
  const activityOwnerIds = resources.map((r) => r.extId)

  for (const [pIdx, p] of projects.entries()) {
    const pMilestones = milestoneRows.filter((m) => m._projectExtId === p.extId)
    const doneThreshold = [2, 4, 3, 1, 5, 3][pIdx % 6]

    for (let i = 0; i < 6; i++) {
      const start = addDays(today, -18 + i * 8 - pIdx * 2)
      const baseEnd = addDays(start, 7 + (i % 3))
      const isDone = i < doneThreshold
      const overdue = i === doneThreshold && pIdx % 2 === 0
      const dueSoon = i === doneThreshold + 1
      const endDate = overdue
        ? addDays(today, -3 - pIdx)
        : dueSoon
          ? addDays(today, 3 + (pIdx % 2))
          : baseEnd

      const statusKey = isDone ? 'CLOSED' : i % 2 === 0 ? 'IN_PROGRESS' : 'NOT_STARTED'
      const ryg = isDone ? 'GREEN' : overdue ? 'RED' : dueSoon ? 'YELLOW' : 'GREEN'

      const ms = pMilestones[i % pMilestones.length]

      const body = {
        pth_activityidexternal: `act_${p.extId}_${i}`,
        pth_name: `${p.name.split(' ')[0]} Activity ${i + 1}`,
        pth_owner: resources[(pIdx + i) % activityOwnerIds.length].name,
        pth_startdate: start,
        pth_enddate: endDate,
        pth_plannedenddate: baseEnd,
        pth_status: ACTIVITY_STATUS_VALUE[statusKey],
        pth_ryg: RYG_VALUE[ryg],
        pth_iscriticalpath: i === 2 || i === 3,
        // Lookups
        'pth_Project@odata.bind': `/pth_projects(${projectMap[p.extId]})`,
      }

      if (isDone) body.pth_closeddate = addDays(endDate, -1)
      if (ms) body['pth_Milestone@odata.bind'] = `/pth_milestones(${ms.pth_milestoneid})`

      const row = await post('pth_activities', body)
      activityRows.push({ ...row, _projectExtId: p.extId, _idx: i, _ownerExtId: activityOwnerIds[(pIdx + i) % activityOwnerIds.length] })
      console.log(`  Activity: ${body.pth_name} (${p.name})  →  ${row.pth_activityid}`)
    }
  }

  // ── Assignments ────────────────────────────────────────────────────────
  console.log('Creating assignments…')
  const week0 = weekStart(today)
  const weekStarts = Array.from({ length: 8 }, (_, idx) => addDays(week0, idx * 7))

  for (const [wIndex, wk] of weekStarts.entries()) {
    for (const [pIndex, res] of resources.entries()) {
      const actA = activityRows[(wIndex * 2 + pIndex) % activityRows.length]
      const actB = activityRows[(wIndex * 2 + pIndex + 7) % activityRows.length]
      const base = 22 + ((pIndex + wIndex) % 3) * 8
      const loadBump = res.extId === 'u_res_2' && wIndex % 3 === 0 ? 16
        : res.extId === 'u_res_4' && wIndex % 4 === 1 ? 11
        : 0

      await post('pth_assignments', {
        pth_assignmentidexternal: `as_${res.extId}_${wIndex}_a`,
        pth_assignedhours: base,
        pth_weekstartdate: wk,
        'pth_Activity@odata.bind': `/pth_activities(${actA.pth_activityid})`,
        'pth_Resource@odata.bind': `/pth_resources(${resourceMap[res.extId]})`,
      })
      await post('pth_assignments', {
        pth_assignmentidexternal: `as_${res.extId}_${wIndex}_b`,
        pth_assignedhours: 12 + loadBump,
        pth_weekstartdate: wk,
        'pth_Activity@odata.bind': `/pth_activities(${actB.pth_activityid})`,
        'pth_Resource@odata.bind': `/pth_resources(${resourceMap[res.extId]})`,
      })
    }
    console.log(`  Week ${wk}: ${resources.length * 2} assignments`)
  }

  // ── PPM Settings ───────────────────────────────────────────────────────
  console.log('Creating PPM settings…')
  await post('pth_ppmsettings', {
    pth_name: 'Default',
    pth_warningdaysthreshold: 7,
    pth_reportweekwindow: 4,
    pth_defaultweeklycapacity: 40,
    pth_utilizationwarningthreshold: 90,
    pth_utilizationcriticalthreshold: 100,
  })
  console.log('  PPM Settings: Default')

  console.log('\n✅ Seed data complete!')
  console.log(`  ${resources.length} resources`)
  console.log(`  ${projects.length} projects`)
  console.log(`  ${milestoneRows.length} milestones`)
  console.log(`  ${activityRows.length} activities`)
  console.log(`  ${weekStarts.length * resources.length * 2} assignments`)
  console.log('  1 PPM Settings row')
}

run().catch((err) => {
  console.error(err.message)
  process.exit(1)
})
