/**
 * cleanup-dupes.mjs — Remove duplicate records created by multiple seed runs.
 *
 * Strategy:
 *   1. Projects  — keep first of each name (by createdon), delete the rest
 *   2. Resources — keep first of each name, delete the rest
 *   3. Milestones / Activities / Assignments that reference deleted projects
 *      or deleted resources are orphaned → delete them too
 *
 * Usage:
 *   DATAVERSE_URL=https://org....crm.dynamics.com node scripts/cleanup-dupes.mjs
 */

import { readFileSync } from 'node:fs'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = dirname(fileURLToPath(import.meta.url))
const envPath = resolve(__dirname, '..', '.env.local')
const envContent = readFileSync(envPath, 'utf-8')
const DV_URL = envContent.match(/VITE_DATAVERSE_URL=(.*)/)?.[1]?.trim().replace(/\/+$/, '')
const token = envContent.match(/VITE_DATAVERSE_TOKEN=(.*)/)?.[1]?.trim()

if (!DV_URL || !token) { console.error('Missing VITE_DATAVERSE_URL or VITE_DATAVERSE_TOKEN in .env.local — run: npm run dev:dv'); process.exit(1) }
const BASE = `${DV_URL}/api/data/v9.2`
const headers = {
  Authorization: `Bearer ${token}`,
  Accept: 'application/json',
  'OData-MaxVersion': '4.0',
  'OData-Version': '4.0',
}

async function fetchAll(entity) {
  const res = await fetch(`${BASE}/${entity}?$orderby=createdon`, { headers })
  if (!res.ok) throw new Error(`GET ${entity}: ${res.status}`)
  return (await res.json()).value
}

async function del(entity, id) {
  const res = await fetch(`${BASE}/${entity}(${id})`, { method: 'DELETE', headers })
  if (!res.ok && res.status !== 404) {
    console.error(`  ✗ DELETE ${entity}(${id}): ${res.status}`)
    return false
  }
  return true
}

// ── 1. Deduplicate projects ──────────────────────────────────────────────

console.log('\n── Projects ──')
const projects = await fetchAll('pth_projects')
const keepProjectIds = new Set()
const deleteProjectIds = new Set()
const seenProjectNames = new Set()

for (const p of projects) {
  const name = p.pth_projectname
  if (seenProjectNames.has(name)) {
    deleteProjectIds.add(p.pth_projectid)
  } else {
    seenProjectNames.add(name)
    keepProjectIds.add(p.pth_projectid)
  }
}
console.log(`  Total: ${projects.length}  Keep: ${keepProjectIds.size}  Delete: ${deleteProjectIds.size}`)

// ── 2. Deduplicate resources ─────────────────────────────────────────────

console.log('\n── Resources ──')
const resources = await fetchAll('pth_resources')
const keepResourceIds = new Set()
const deleteResourceIds = new Set()
const seenResourceNames = new Set()

for (const r of resources) {
  const name = r.pth_name
  if (seenResourceNames.has(name)) {
    deleteResourceIds.add(r.pth_resourceid)
  } else {
    seenResourceNames.add(name)
    keepResourceIds.add(r.pth_resourceid)
  }
}
console.log(`  Total: ${resources.length}  Keep: ${keepResourceIds.size}  Delete: ${deleteResourceIds.size}`)

// ── 3. Identify orphaned milestones (linked to deleted projects) ─────────

console.log('\n── Milestones ──')
const milestones = await fetchAll('pth_milestones')
const deleteMilestoneIds = []
for (const m of milestones) {
  const projRef = m._pth_project_value
  if (projRef && deleteProjectIds.has(projRef)) {
    deleteMilestoneIds.push(m.pth_milestoneid)
  }
}
// Also deduplicate milestones within kept projects (same name + same project)
const seenMs = new Set()
for (const m of milestones) {
  const projRef = m._pth_project_value
  if (projRef && keepProjectIds.has(projRef)) {
    const key = m.pth_milestonename + '|' + projRef
    if (seenMs.has(key)) {
      deleteMilestoneIds.push(m.pth_milestoneid)
    } else {
      seenMs.add(key)
    }
  }
}
console.log(`  Total: ${milestones.length}  Delete: ${deleteMilestoneIds.length}`)

// ── 4. Identify orphaned activities ──────────────────────────────────────

console.log('\n── Activities ──')
const activities = await fetchAll('pth_activities')
const deleteActivityIds = new Set()
for (const a of activities) {
  const projRef = a._pth_project_value
  if (projRef && deleteProjectIds.has(projRef)) {
    deleteActivityIds.add(a.pth_activityid)
  }
}
// Also deduplicate within kept projects
const seenAct = new Set()
for (const a of activities) {
  const projRef = a._pth_project_value
  if (projRef && keepProjectIds.has(projRef)) {
    const key = a.pth_activityname + '|' + projRef
    if (seenAct.has(key)) {
      deleteActivityIds.add(a.pth_activityid)
    } else {
      seenAct.add(key)
    }
  }
}
console.log(`  Total: ${activities.length}  Delete: ${deleteActivityIds.size}`)

// ── 5. Identify orphaned assignments ─────────────────────────────────────

console.log('\n── Assignments ──')
const assignments = await fetchAll('pth_assignments')
const deleteAssignmentIds = []
for (const a of assignments) {
  const actRef = a._pth_activity_value
  const resRef = a._pth_resource_value
  if ((actRef && deleteActivityIds.has(actRef)) || (resRef && deleteResourceIds.has(resRef))) {
    deleteAssignmentIds.push(a.pth_assignmentid)
  }
}
console.log(`  Total: ${assignments.length}  Delete: ${deleteAssignmentIds.length}`)

// ── Execute deletes ──────────────────────────────────────────────────────

const totalDeletes = deleteAssignmentIds.length + deleteActivityIds.size +
  deleteMilestoneIds.length + deleteResourceIds.size + deleteProjectIds.size
console.log(`\n🗑  Total records to delete: ${totalDeletes}`)

// Delete in reverse-dependency order: assignments → activities → milestones → resources → projects
let deleted = 0

console.log('\nDeleting assignments...')
for (const id of deleteAssignmentIds) {
  if (await del('pth_assignments', id)) deleted++
}

console.log('Deleting activities...')
for (const id of deleteActivityIds) {
  if (await del('pth_activities', id)) deleted++
}

console.log('Deleting milestones...')
for (const id of deleteMilestoneIds) {
  if (await del('pth_milestones', id)) deleted++
}

console.log('Deleting resources...')
for (const id of deleteResourceIds) {
  if (await del('pth_resources', id)) deleted++
}

console.log('Deleting projects...')
for (const id of deleteProjectIds) {
  if (await del('pth_projects', id)) deleted++
}

console.log(`\n✅ Done — deleted ${deleted}/${totalDeletes} records`)
