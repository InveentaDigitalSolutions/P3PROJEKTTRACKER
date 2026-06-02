/**
 * dataverse.ts — Thin Dataverse data layer for the Project Tracker Hub.
 *
 * **Primary path (Power Apps)** — uses the generated SDK services produced by
 * `pac code add-data-source`.  The services call `getClient()` from
 * `@microsoft/power-apps/data` which handles auth automatically inside
 * Power Apps at runtime.
 *
 * **Fallback (standalone dev mode)** — direct `fetch()` calls against the
 * Dataverse Web API using `VITE_DATAVERSE_URL` + `VITE_DATAVERSE_TOKEN`
 * env vars set in `.env.local`.
 *
 * If neither path succeeds the module returns `null` and the app uses its
 * built-in demo data.
 */

import { getContext } from '@microsoft/power-apps/app'
import { Pth_projectsService }   from './generated/services/Pth_projectsService'
import { Pth_milestonesService }  from './generated/services/Pth_milestonesService'
import { Pth_activitiesService }  from './generated/services/Pth_activitiesService'
import { Pth_resourcesService }   from './generated/services/Pth_resourcesService'
import { Pth_assignmentsService } from './generated/services/Pth_assignmentsService'
import { Pth_ppmsettingsService } from './generated/services/Pth_ppmsettingsService'

// ── Types mirroring App.tsx ──────────────────────────────────────────────

export type ProjectCategory = 'ECR' | 'CIP' | 'Path Forward' | 'New Programs'
export type RygStatus = 'RED' | 'YELLOW' | 'GREEN'
export type ActivityState = 'NOT_STARTED' | 'IN_PROGRESS' | 'DONE' | 'NOT_APPLICABLE'

/* Choice → label maps for Dataverse option-set values */
const CATEGORY_MAP: Record<number, ProjectCategory> = {
  100000000: 'ECR',
  100000001: 'Path Forward',
  100000002: 'CIP',
  100000003: 'New Programs',
}
const RYG_MAP: Record<number, RygStatus> = {
  100000000: 'RED',
  100000001: 'YELLOW',
  100000002: 'GREEN',
}
const ACTIVITY_STATUS_MAP: Record<number, ActivityState> = {
  100000000: 'NOT_STARTED',
  100000001: 'IN_PROGRESS',
  100000002: 'NOT_STARTED',     // BLOCKED → treat as not started in the app
  100000003: 'DONE',            // CLOSED → DONE
  100000004: 'NOT_APPLICABLE',  // N/A → not applicable (won't be done)
}

// ── Direct Web API fetch helpers (standalone / dev mode) ─────────────────

const DV_URL = (typeof import.meta !== 'undefined' && (import.meta as any).env?.VITE_DATAVERSE_URL as string) || ''
const DV_TOKEN = (typeof import.meta !== 'undefined' && (import.meta as any).env?.VITE_DATAVERSE_TOKEN as string) || ''

// In dev mode (Vite), route through the Vite proxy to avoid CORS.
// In production (Power Apps), use the full Dataverse URL.
const isDev = (typeof import.meta !== 'undefined' && (import.meta as any).env?.DEV) || false
const API_BASE = DV_URL
  ? (isDev ? '/api/data/v9.2' : `${DV_URL.replace(/\/+$/, '')}/api/data/v9.2`)
  : ''

/**
 * True when we should write straight to the Web API instead of the Power Apps
 * SDK — i.e. standalone dev mode (token present, not running inside the Power
 * Apps host). The SDK's getClient() hangs outside the host, so create/update/
 * delete must bypass it here, mirroring the read path in fetchAllFromDataverse.
 */
function preferDirectWrite(): boolean {
  const insidePowerApps = typeof window !== 'undefined' && !!(window as { __powerAppsContext__?: unknown }).__powerAppsContext__
  return !!(API_BASE && DV_TOKEN && !insidePowerApps)
}

async function dvGet<T>(entity: string, query = ''): Promise<T[]> {
  if (!API_BASE || !DV_TOKEN) return []
  const url = `${API_BASE}/${entity}${query ? '?' + query : ''}`
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${DV_TOKEN}`,
      Accept: 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
    },
  })
  if (!res.ok) {
    console.error(`Dataverse GET ${entity} failed (${res.status})`)
    return []
  }
  const json = await res.json()
  return (json.value ?? []) as T[]
}

// ── Row → App model mappers ──────────────────────────────────────────────

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Row = Record<string, any>

/* Choice column maps for pth_sitelocation */
const SITE_MAP: Record<number, string> = {
  100000000: 'site_tca',
  100000001: 'site_slp',
}
const SITE_TO_DV: Record<string, number> = {
  site_tca: 100000000,
  site_slp: 100000001,
}

/* Resource Area choice + Location multi-select maps */
const RESOURCE_AREA_MAP: Record<number, string> = {
  100000000: 'PPS', 100000001: 'ENG', 100000002: 'QMM', 100000003: 'LOD',
  100000004: 'LOP', 100000005: 'LOP4', 100000006: 'LOP5', 100000007: 'MSE',
  100000008: 'CTG', 100000009: 'PUQ1', 100000010: 'PUQ2',
}
const RESOURCE_LOC_MAP: Record<number, string> = {
  100000000: 'SlpP',
  100000001: 'TlP',
}

function mapProject(r: Row) {
  return {
    id: r.pth_projectid as string,
    name: (r.pth_projectname ?? '') as string,
    projectIdExternal: (r.pth_projectidexternal ?? '') as string,
    category: CATEGORY_MAP[r.pth_category as number] ?? ('CIP' as ProjectCategory),
    siteId: SITE_MAP[r.pth_sitelocation as number] ?? 'site_tca',
    sponsorExecutiveId: 'u_dir_1',
    projectManagerId: 'u_pm_1',
    clientIds: [] as string[],
    supplierIds: [] as string[],
    objective: (r.pth_projectobjective ?? '') as string,
    clientCode: (r.pth_boschcode ?? '') as string,
    budgetAllocated: Number(r.pth_budgetallocated ?? 0),

    plannedStartDate: ((r.pth_plannedstartdate ?? '') as string).slice(0, 10),
    plannedEndDate: ((r.pth_plannedenddate ?? '') as string).slice(0, 10),
    projectType: (r.pth_projecttype ?? '') as string,
    _sponsorName: (r.pth_sponsorexecutive ?? '') as string,
    _pmName: (r.pth_projectmanagername ?? '') as string,
    _timeStatus: RYG_MAP[r.pth_timestatus as number] ?? ('GREEN' as RygStatus),
  }
}

function mapMilestone(r: Row) {
  return {
    id: r.pth_milestoneid as string,
    projectId: (r._pth_project_value ?? '') as string,
    name: (r.pth_name ?? '') as string,
    targetDate: ((r.pth_planneddate ?? '') as string).slice(0, 10),
    doneDate: r.pth_status === 100000002
      ? ((r.pth_planneddate ?? '') as string).slice(0, 10)
      : undefined,
  }
}

function mapActivity(r: Row) {
  const state = ACTIVITY_STATUS_MAP[r.pth_status as number] ?? 'NOT_STARTED'
  return {
    id: r.pth_activityid as string,
    projectId: (r._pth_project_value ?? '') as string,
    milestoneId: (r._pth_milestone_value ?? undefined) as string | undefined,
    name: (r.pth_name ?? '') as string,
    ownerId: 'u_res_1',
    _ownerName: (r.pth_owner ?? '') as string,
    startDate: ((r.pth_startdate ?? '') as string).slice(0, 10),
    endDate: ((r.pth_enddate ?? '') as string).slice(0, 10),
    doneDate: state === 'DONE'
      ? ((r.pth_closeddate ?? r.pth_enddate ?? '') as string).slice(0, 10) || undefined
      : undefined,
    state,
    criticalPath: !!r.pth_iscriticalpath,
    responsible: (r.pth_responsible ?? '') as string,
    workloadPct: r.pth_workloadpct == null ? null : Number(r.pth_workloadpct),
  }
}

function mapResource(r: Row) {
  const name = (r.pth_name ?? '') as string
  const initials = name
    .split(' ')
    .map((w: string) => w[0])
    .join('')
    .toUpperCase()
    .slice(0, 2)
  const locVal = r.pth_location
  const location = typeof locVal === 'string' && locVal
    ? locVal.split(',').map((v: string) => RESOURCE_LOC_MAP[Number(v)]).filter(Boolean)
    : []
  return {
    id: r.pth_resourceid as string,
    extId: (r.pth_personidexternal ?? '') as string,
    name,
    initials,
    role: 'RESOURCE' as const,
    department: (r.pth_department ?? '') as string,
    weeklyCapacity: Number(r.pth_weeklycapacityhours ?? 40),
    area: (RESOURCE_AREA_MAP[r.pth_area as number] ?? '') as string,
    location: location as string[],
    managerId: (r._pth_manager_value ?? undefined) as string | undefined,
    resourceKind: (((r.pth_role ?? '') as string) || 'Associate') as string,
  }
}

function mapAssignment(r: Row) {
  return {
    id: r.pth_assignmentid as string,
    personId: (r._pth_resource_value ?? '') as string,
    activityId: (r._pth_activity_value ?? '') as string,
    weekStart: ((r.pth_weekstartdate ?? '') as string).slice(0, 10),
    assignedHours: Number(r.pth_assignedhours ?? 0),
    capacityHours: 40,
  }
}

function mapSettings(r: Row) {
  return {
    dvSettingsId: (r.pth_ppmsettingid ?? r.pth_ppmsettingsid ?? '') as string,
    warningDaysThreshold: Number(r.pth_warningdaysthreshold ?? 7),
    upcomingWeeksWindow: Number(r.pth_reportweekwindow ?? 4),
    lastClosedWeeksWindow: Number(r.pth_reportweekwindow ?? 4),
    defaultWeeklyCapacity: Number(r.pth_defaultweeklycapacity ?? 40),
    workloadWarningThreshold: Number(r.pth_utilizationwarningthreshold ?? 90),
    workloadCriticalThreshold: Number(r.pth_utilizationcriticalthreshold ?? 100),
  }
}

// ── Public API ────────────────────────────────────────────────────────────

export interface DataverseAppData {
  projects: ReturnType<typeof mapProject>[]
  milestones: ReturnType<typeof mapMilestone>[]
  activities: ReturnType<typeof mapActivity>[]
  resources: ReturnType<typeof mapResource>[]
  assignments: ReturnType<typeof mapAssignment>[]
  settings: ReturnType<typeof mapSettings>
}

// ── SDK-based fetch (Power Apps runtime) ─────────────────────────────────

/** Safely extract the data array from an IOperationResult — handles .data, .value, and array-at-root */
function extractArray<T>(result: unknown): T[] {
  if (!result) return []
  const r = result as Record<string, unknown>
  // Standard SDK shape: { data: T[] }
  if (Array.isArray(r.data)) return r.data as T[]
  // Some SDK versions return { value: T[] }
  if (Array.isArray(r.value)) return r.value as T[]
  // Result itself is an array
  if (Array.isArray(result)) return result as T[]
  return []
}

async function fetchWithSDK(): Promise<DataverseAppData | null> {
  console.log('[DV-SDK] Starting SDK fetch...')
  const [projectRes, milestoneRes, activityRes, resourceRes, assignmentRes, settingsRes] = await Promise.all([
    Pth_projectsService.getAll({ orderBy: ['pth_projectname asc'] }),
    Pth_milestonesService.getAll({ orderBy: ['pth_planneddate asc'] }),
    Pth_activitiesService.getAll({ orderBy: ['pth_startdate asc'] }),
    Pth_resourcesService.getAll({ orderBy: ['pth_name asc'] }),
    Pth_assignmentsService.getAll({ orderBy: ['pth_weekstartdate asc'] }),
    Pth_ppmsettingsService.getAll({ top: 1 }),
  ])

  const projectRows   = extractArray<Row>(projectRes)
  const milestoneRows = extractArray<Row>(milestoneRes)
  const activityRows  = extractArray<Row>(activityRes)
  const resourceRows  = extractArray<Row>(resourceRes)
  const assignmentRows = extractArray<Row>(assignmentRes)
  const settingsRows  = extractArray<Row>(settingsRes)

  console.log('[DV-SDK] Rows — projects:', projectRows.length,
    'milestones:', milestoneRows.length,
    'activities:', activityRows.length,
    'resources:', resourceRows.length,
    'assignments:', assignmentRows.length,
    'settings:', settingsRows.length)
  if (projectRows.length > 0) {
    console.log('[DV-SDK] First project raw keys:', Object.keys(projectRows[0]).join(', '))
    console.log('[DV-SDK] First project raw:', JSON.stringify(projectRows[0]).slice(0, 600))
  }

  const projects   = projectRows.map((r) => mapProject(r))
  const milestones = milestoneRows.map((r) => mapMilestone(r))
  const activities = activityRows.map((r) => mapActivity(r))
  const resources  = resourceRows.map((r) => mapResource(r))
  const assignments = assignmentRows.map((r) => mapAssignment(r))
  const settings = settingsRows.length > 0
    ? mapSettings(settingsRows[0])
    : mapSettings({})

  console.log('[DV-SDK] Mapped projects:', projects.length)
  if (projects.length > 0) console.log('[DV-SDK] First mapped project:', JSON.stringify(projects[0]))

  return { projects, milestones, activities, resources, assignments, settings }
}

// ── Direct-fetch fallback (standalone dev mode) ──────────────────────────

async function fetchWithDirectAPI(): Promise<DataverseAppData | null> {
  if (!API_BASE || !DV_TOKEN) return null

  const [projectRows, milestoneRows, activityRows, resourceRows, assignmentRows, settingsRows] = await Promise.all([
    dvGet<Row>('pth_projects', '$orderby=pth_projectname'),
    dvGet<Row>('pth_milestones', '$orderby=pth_planneddate'),
    dvGet<Row>('pth_activities', '$orderby=pth_startdate'),
    dvGet<Row>('pth_resources', '$orderby=pth_name'),
    dvGet<Row>('pth_assignments', '$orderby=pth_weekstartdate'),
    dvGet<Row>('pth_ppmsettings', '$top=1'),
  ])

  return {
    projects: projectRows.map(mapProject),
    milestones: milestoneRows.map(mapMilestone),
    activities: activityRows.map(mapActivity),
    resources: resourceRows.map(mapResource),
    assignments: assignmentRows.map(mapAssignment),
    settings: settingsRows.length > 0 ? mapSettings(settingsRows[0]) : mapSettings({}),
  }
}

/**
 * Fetch all app data from Dataverse.
 *
 * 1. Tries the generated SDK services (works inside Power Apps).
 *    Times out after 5 s so the app doesn't hang in dev mode.
 * 2. Falls back to direct Web API fetch (standalone dev with env vars).
 * 3. Returns `null` if neither path succeeds (demo mode).
 */
export async function fetchAllFromDataverse(): Promise<DataverseAppData | null> {
  console.log('[DV] fetchAllFromDataverse called')

  // Detect if we're running inside Power Apps (iframe with __powerAppsContext__)
  const insidePowerApps = typeof window !== 'undefined' && !!(window as any).__powerAppsContext__

  // If env vars are set and we're NOT inside Power Apps, skip SDK entirely
  if (API_BASE && DV_TOKEN && !insidePowerApps) {
    console.log('[DV] Direct API env vars detected & not inside Power Apps — skipping SDK, going straight to direct API')
    try {
      const result = await fetchWithDirectAPI()
      console.log('[DV] Direct API result:', result ? `${result.projects.length} projects` : 'null')
      return result
    } catch (err) {
      console.error('[DV] Direct API fetch failed:', err)
      return null
    }
  }

  // 1. SDK path (Power Apps runtime) — with generous timeout for cold start
  try {
    const timeout = new Promise<never>((_, reject) =>
      setTimeout(() => reject(new Error('SDK timeout after 30s')), 30000),
    )
    console.log('[DV] Attempting SDK fetch (30s timeout)...')
    const data = await Promise.race([fetchWithSDK(), timeout])
    console.log('[DV] SDK fetch returned. projects:', data?.projects?.length,
      'milestones:', data?.milestones?.length,
      'activities:', data?.activities?.length,
      'resources:', data?.resources?.length,
      'assignments:', data?.assignments?.length)
    if (data) {
      console.log('[DV] ✅ SDK path succeeded')
      return data
    }
    console.warn('[DV] SDK returned null, falling through to direct API')
  } catch (err) {
    console.warn('[DV] SDK fetch failed/timed-out:', err)
  }

  // 2. Direct API path (standalone dev mode)
  try {
    console.log('[DV] Attempting direct API fetch... API_BASE:', API_BASE ? 'set' : 'empty', 'DV_TOKEN:', DV_TOKEN ? 'set' : 'empty')
    const result = await fetchWithDirectAPI()
    console.log('[DV] Direct API result:', result ? `${result.projects.length} projects` : 'null')
    return result
  } catch (err) {
    console.error('[DV] Direct API fetch failed:', err)
    return null
  }
}

/**
 * Returns `true` so the app always attempts to load from Dataverse.
 * `fetchAllFromDataverse` handles graceful fallback internally.
 */
export function isDataverseConfigured(): boolean {
  return true
}

// ── Dataverse Users (systemusers table) ──────────────────────────────────

export interface DataverseUser {
  id: string
  fullname: string
  email: string
  jobtitle: string
}

/**
 * Fetch enabled interactive users from the Dataverse `systemusers` table.
 * Filters: `isdisabled eq false` + `accessmode eq 0` (Read-Write interactive users).
 */
export async function fetchDataverseUsers(): Promise<DataverseUser[]> {
  console.log('[DV] Fetching systemusers...')
  const rows = await dvGet<Row>(
    'systemusers',
    '$select=systemuserid,fullname,internalemailaddress,jobtitle&$filter=isdisabled eq false and accessmode eq 0&$orderby=fullname',
  )
  console.log('[DV] systemusers fetched:', rows.length)
  return rows.map((r) => ({
    id: r.systemuserid as string,
    fullname: (r.fullname ?? '') as string,
    email: (r.internalemailaddress ?? '') as string,
    jobtitle: (r.jobtitle ?? '') as string,
  }))
}

// ── Reverse maps (app enum → Dataverse option-set value) ─────────────────

const CATEGORY_TO_DV: Record<ProjectCategory, number> = {
  ECR: 100000000,
  'Path Forward': 100000001,
  CIP: 100000002,
  'New Programs': 100000003,
}

const RYG_TO_DV: Record<RygStatus, number> = {
  RED: 100000000,
  YELLOW: 100000001,
  GREEN: 100000002,
}

const ACTIVITY_STATE_TO_DV: Record<ActivityState, number> = {
  NOT_STARTED: 100000000,
  IN_PROGRESS: 100000001,
  DONE: 100000003,            // maps to CLOSED in Dataverse
  NOT_APPLICABLE: 100000004,  // N/A
}

// ── Generic direct Web API write helpers ──────────────────────────────────

async function dvPost(entity: string, body: Record<string, unknown>): Promise<string | null> {
  if (!API_BASE || !DV_TOKEN) return null
  const url = `${API_BASE}/${entity}`
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${DV_TOKEN}`,
      'Content-Type': 'application/json',
      Accept: 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      Prefer: 'return=representation',
    },
    body: JSON.stringify(body),
  })
  if (!res.ok) {
    const errText = await res.text()
    console.error(`Dataverse POST ${entity} failed (${res.status}):`, errText)
    return null
  }
  const json = await res.json()
  // Return the primary key — try well-known id columns then any column ending in "id"
  return (json.pth_projectid ?? json.pth_milestoneid ?? json.pth_activityid ?? json.pth_resourceid ?? json.pth_assignmentid ?? json.pth_ppmsettingid ?? json[Object.keys(json).find(k => k.endsWith('id')) ?? ''] ?? null) as string | null
}

async function dvPatch(entity: string, id: string, body: Record<string, unknown>): Promise<boolean> {
  if (!API_BASE || !DV_TOKEN) return false
  const res = await fetch(`${API_BASE}/${entity}(${id})`, {
    method: 'PATCH',
    headers: {
      Authorization: `Bearer ${DV_TOKEN}`,
      'Content-Type': 'application/json',
      Accept: 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
    },
    body: JSON.stringify(body),
  })
  if (res.ok || res.status === 204) return true
  console.error(`Dataverse PATCH ${entity}(${id}) failed (${res.status}):`, await res.text())
  return false
}

async function dvDelete(entity: string, id: string): Promise<boolean> {
  if (!API_BASE || !DV_TOKEN) return false
  const res = await fetch(`${API_BASE}/${entity}(${id})`, {
    method: 'DELETE',
    headers: {
      Authorization: `Bearer ${DV_TOKEN}`,
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
    },
  })
  if (res.ok || res.status === 204) return true
  console.error(`Dataverse DELETE ${entity}(${id}) failed (${res.status}):`, await res.text())
  return false
}

export interface CreateProjectPayload {
  name: string
  clientCode: string
  category: ProjectCategory
  budgetAllocated: number
  site: string
  objective: string
  sponsorExecutive: string
  projectManagerName: string
  plannedStartDate: string
  plannedEndDate: string
  projectIdExternal?: string
  projectType?: string
}

/** Build the Dataverse column payload from the app-level create payload. */
function buildProjectBody(p: CreateProjectPayload): Record<string, unknown> {
  const body: Record<string, unknown> = {
    pth_projectname: p.name,
    pth_boschcode: p.clientCode,
    pth_category: CATEGORY_TO_DV[p.category],
    pth_budgetallocated: p.budgetAllocated,
    pth_sitelocation: SITE_TO_DV[p.site] ?? 100000000,
    pth_projectobjective: p.objective,
    pth_sponsorexecutive: p.sponsorExecutive,
    pth_projectmanagername: p.projectManagerName,
    pth_monitoringcriteria: 'On Time, On Cost, On Quality',
    pth_plannedstartdate: p.plannedStartDate,
    pth_plannedenddate: p.plannedEndDate,
    pth_projectidexternal: p.projectIdExternal ?? `PRJ-${Date.now().toString(36).toUpperCase()}`,
    pth_timestatus: RYG_TO_DV.GREEN,
    pth_criticalpathchangedflag: 0,
  }
  if (p.projectType) body.pth_projecttype = p.projectType
  return body
}

/**
 * Deletes a project row from Dataverse by its GUID.
 * Tries the SDK service first, falls back to direct DELETE.
 */
export async function deleteProjectInDataverse(id: string): Promise<boolean> {
  if (preferDirectWrite()) return dvDelete('pth_projects', id)
  try { await Pth_projectsService.delete(id); return true } catch (err) { console.warn('SDK delete project unavailable:', err) }
  return dvDelete('pth_projects', id)
}

/**
 * Creates a new project row in Dataverse.
 * Tries the SDK service first, falls back to direct POST.
 * Returns the Dataverse-generated GUID (pth_projectid) on success, or `null`.
 */
export async function createProjectInDataverse(p: CreateProjectPayload): Promise<string | null> {
  const body = buildProjectBody(p)
  if (preferDirectWrite()) return dvPost('pth_projects', body)

  // 1. SDK path
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const result = await Pth_projectsService.create(body as any)
    console.log('[DV] SDK create result:', JSON.stringify(result))
    // The returned record may live under result.data or directly on result
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const record = (result as any).data ?? result
    const id = record?.pth_projectid as string | undefined
    if (id) {
      console.log('[DV] Project created via SDK, id =', id)
      return id
    }
    console.warn('[DV] SDK create succeeded but no pth_projectid in result')
  } catch (err) {
    console.error('[DV] SDK create failed:', err)
  }

  // 2. Direct API path
  return dvPost('pth_projects', body)
}

/**
 * Updates an existing project row in Dataverse.
 * Tries the SDK service first, falls back to direct PATCH.
 */
export async function updateProjectInDataverse(id: string, p: CreateProjectPayload): Promise<boolean> {
  const body = buildProjectBody(p)
  if (preferDirectWrite()) return dvPatch('pth_projects', id, body)
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await Pth_projectsService.update(id, body as any)
    console.log('[DV] Project updated via SDK, id =', id)
    return true
  } catch (err) { console.error('[DV] SDK update project failed:', err) }
  return dvPatch('pth_projects', id, body)
}

// ═══════════════════════════════════════════════════════════════════════════
// ── MILESTONE CRUD ────────────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════

const MILESTONE_STATUS_TO_DV: Record<string, number> = {
  OPEN: 100000000,
  IN_PROGRESS: 100000001,
  CLOSED: 100000002,
  DELAYED: 100000003,
}

export interface MilestonePayload {
  name: string
  targetDate: string
  projectId: string
  doneDate?: string
}

function buildMilestoneBody(m: MilestonePayload): Record<string, unknown> {
  const isDone = !!m.doneDate
  return {
    pth_name: m.name,
    pth_planneddate: m.targetDate,
    pth_status: isDone ? MILESTONE_STATUS_TO_DV.CLOSED : MILESTONE_STATUS_TO_DV.OPEN,
    pth_iscriticalpath: false,
    pth_ryg: RYG_TO_DV.GREEN,
    'pth_project@odata.bind': `/pth_projects(${m.projectId})`,
    pth_milestoneidexternal: `MS-${Date.now().toString(36).toUpperCase()}`,
  }
}

export async function createMilestoneInDataverse(m: MilestonePayload): Promise<string | null> {
  const body = buildMilestoneBody(m)
  if (preferDirectWrite()) return dvPost('pth_milestones', body)
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const result = await Pth_milestonesService.create(body as any)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const record = (result as any).data ?? result
    const id = record?.pth_milestoneid as string | undefined
    if (id) { console.log('[DV] Milestone created via SDK, id =', id); return id }
  } catch (err) { console.error('[DV] SDK create milestone failed:', err) }
  return dvPost('pth_milestones', body)
}

export async function updateMilestoneInDataverse(id: string, m: MilestonePayload): Promise<boolean> {
  const body: Record<string, unknown> = {
    pth_name: m.name,
    pth_planneddate: m.targetDate,
    pth_status: m.doneDate ? MILESTONE_STATUS_TO_DV.CLOSED : MILESTONE_STATUS_TO_DV.OPEN,
    'pth_project@odata.bind': `/pth_projects(${m.projectId})`,
  }
  if (preferDirectWrite()) return dvPatch('pth_milestones', id, body)
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await Pth_milestonesService.update(id, body as any)
    console.log('[DV] Milestone updated via SDK, id =', id); return true
  } catch (err) { console.error('[DV] SDK update milestone failed:', err) }
  return dvPatch('pth_milestones', id, body)
}

export async function deleteMilestoneInDataverse(id: string): Promise<boolean> {
  if (preferDirectWrite()) return dvDelete('pth_milestones', id)
  try { await Pth_milestonesService.delete(id); return true } catch (err) { console.warn('SDK delete milestone unavailable:', err) }
  return dvDelete('pth_milestones', id)
}

// ═══════════════════════════════════════════════════════════════════════════
// ── ACTIVITY CRUD ─────────────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════

export interface ActivityPayload {
  name: string
  ownerName: string
  startDate: string
  endDate: string
  doneDate?: string
  state: ActivityState
  criticalPath: boolean
  projectId: string
  milestoneId?: string
  responsible?: string
  leadtimeWeeks?: number | null
  workloadPct?: number | null
  inputs?: string
}

function buildActivityBody(a: ActivityPayload): Record<string, unknown> {
  const body: Record<string, unknown> = {
    pth_name: a.name,
    pth_owner: a.ownerName,
    pth_startdate: a.startDate,
    pth_enddate: a.endDate,
    pth_plannedenddate: a.endDate,
    pth_status: ACTIVITY_STATE_TO_DV[a.state],
    pth_iscriticalpath: a.criticalPath,
    pth_ryg: RYG_TO_DV.GREEN,
    'pth_project@odata.bind': `/pth_projects(${a.projectId})`,
    pth_activityidexternal: `ACT-${Date.now().toString(36).toUpperCase()}-${Math.floor(performance.now() % 100000)}`,
  }
  if (a.doneDate) body.pth_closeddate = a.doneDate
  if (a.milestoneId) body['pth_milestone@odata.bind'] = `/pth_milestones(${a.milestoneId})`
  if (a.responsible) body.pth_responsible = a.responsible
  if (a.leadtimeWeeks != null) body.pth_leadtimeweeks = a.leadtimeWeeks
  if (a.workloadPct != null) body.pth_workloadpct = a.workloadPct
  if (a.inputs) body.pth_inputs = a.inputs.slice(0, 2000)
  return body
}

export async function createActivityInDataverse(a: ActivityPayload): Promise<string | null> {
  const body = buildActivityBody(a)
  if (preferDirectWrite()) return dvPost('pth_activities', body)
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const result = await Pth_activitiesService.create(body as any)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const record = (result as any).data ?? result
    const id = record?.pth_activityid as string | undefined
    if (id) { console.log('[DV] Activity created via SDK, id =', id); return id }
  } catch (err) { console.error('[DV] SDK create activity failed:', err) }
  return dvPost('pth_activities', body)
}

export async function updateActivityInDataverse(id: string, a: ActivityPayload): Promise<boolean> {
  const body: Record<string, unknown> = {
    pth_name: a.name,
    pth_owner: a.ownerName,
    pth_startdate: a.startDate,
    pth_enddate: a.endDate,
    pth_plannedenddate: a.endDate,
    pth_status: ACTIVITY_STATE_TO_DV[a.state],
    pth_iscriticalpath: a.criticalPath,
    'pth_project@odata.bind': `/pth_projects(${a.projectId})`,
  }
  if (a.doneDate) body.pth_closeddate = a.doneDate
  if (a.milestoneId) body['pth_milestone@odata.bind'] = `/pth_milestones(${a.milestoneId})`
  if (preferDirectWrite()) return dvPatch('pth_activities', id, body)
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await Pth_activitiesService.update(id, body as any)
    console.log('[DV] Activity updated via SDK, id =', id); return true
  } catch (err) { console.error('[DV] SDK update activity failed:', err) }
  return dvPatch('pth_activities', id, body)
}

export async function deleteActivityInDataverse(id: string): Promise<boolean> {
  if (preferDirectWrite()) return dvDelete('pth_activities', id)
  try { await Pth_activitiesService.delete(id); return true } catch (err) { console.warn('SDK delete activity unavailable:', err) }
  return dvDelete('pth_activities', id)
}

// ═══════════════════════════════════════════════════════════════════════════
// ── RESOURCE (PERSON) CRUD ────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════

export interface ResourcePayload {
  name: string
  role: string
  department?: string
  weeklyCapacity?: number
  area?: string          // e.g. 'PPS', 'ENG' …
  location?: string[]    // ['SlpP', 'TlP']
  managerId?: string     // Dataverse resource GUID of the manager
  kind?: string          // 'Manager' | 'Associate' (stored in pth_role)
}

/* Reverse maps: label → Dataverse option-set value */
const AREA_TO_DV: Record<string, number> = Object.fromEntries(
  Object.entries(RESOURCE_AREA_MAP).map(([v, k]) => [k, Number(v)]),
)
const LOC_TO_DV: Record<string, number> = Object.fromEntries(
  Object.entries(RESOURCE_LOC_MAP).map(([v, k]) => [k, Number(v)]),
)

function applyResourceAttrs(body: Record<string, unknown>, r: ResourcePayload): void {
  if (r.area && AREA_TO_DV[r.area] !== undefined) body.pth_area = AREA_TO_DV[r.area]
  if (r.location && r.location.length) {
    body.pth_location = r.location.map((l) => LOC_TO_DV[l]).filter((v) => v !== undefined).join(',')
  }
  if (r.managerId) body['pth_Manager@odata.bind'] = `/pth_resources(${r.managerId})`
}

function buildResourceBody(r: ResourcePayload): Record<string, unknown> {
  const body: Record<string, unknown> = {
    pth_name: r.name,
    pth_role: r.kind ?? r.role,
    pth_department: r.department ?? r.area ?? '',
    pth_weeklycapacityhours: r.weeklyCapacity ?? 40,
    pth_personidexternal: `RES-${Date.now().toString(36).toUpperCase()}`,
  }
  applyResourceAttrs(body, r)
  return body
}

export async function createResourceInDataverse(r: ResourcePayload): Promise<string | null> {
  const body = buildResourceBody(r)
  if (preferDirectWrite()) return dvPost('pth_resources', body)
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const result = await Pth_resourcesService.create(body as any)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const record = (result as any).data ?? result
    const id = record?.pth_resourceid as string | undefined
    if (id) { console.log('[DV] Resource created via SDK, id =', id); return id }
  } catch (err) { console.error('[DV] SDK create resource failed:', err) }
  return dvPost('pth_resources', body)
}

export async function updateResourceInDataverse(id: string, r: ResourcePayload): Promise<boolean> {
  const body: Record<string, unknown> = {
    pth_name: r.name,
    pth_role: r.kind ?? r.role,
    pth_department: r.department ?? r.area ?? '',
    pth_weeklycapacityhours: r.weeklyCapacity ?? 40,
  }
  applyResourceAttrs(body, r)
  if (preferDirectWrite()) return dvPatch('pth_resources', id, body)
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await Pth_resourcesService.update(id, body as any)
    console.log('[DV] Resource updated via SDK, id =', id); return true
  } catch (err) { console.error('[DV] SDK update resource failed:', err) }
  return dvPatch('pth_resources', id, body)
}

export async function deleteResourceInDataverse(id: string): Promise<boolean> {
  if (preferDirectWrite()) return dvDelete('pth_resources', id)
  try { await Pth_resourcesService.delete(id); return true } catch (err) { console.warn('SDK delete resource unavailable:', err) }
  return dvDelete('pth_resources', id)
}

// ═══════════════════════════════════════════════════════════════════════════
// ── ASSIGNMENT CRUD ───────────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════

export interface AssignmentPayload {
  activityId: string
  resourceId: string
  assignedHours: number
  weekStartDate: string
}

function buildAssignmentBody(a: AssignmentPayload): Record<string, unknown> {
  return {
    pth_assignedhours: a.assignedHours,
    pth_weekstartdate: a.weekStartDate,
    'pth_activity@odata.bind': `/pth_activities(${a.activityId})`,
    'pth_resource@odata.bind': `/pth_resources(${a.resourceId})`,
    pth_assignmentidexternal: `ASN-${Date.now().toString(36).toUpperCase()}`,
  }
}

export async function createAssignmentInDataverse(a: AssignmentPayload): Promise<string | null> {
  const body = buildAssignmentBody(a)
  if (preferDirectWrite()) return dvPost('pth_assignments', body)
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const result = await Pth_assignmentsService.create(body as any)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const record = (result as any).data ?? result
    const id = record?.pth_assignmentid as string | undefined
    if (id) { console.log('[DV] Assignment created via SDK, id =', id); return id }
  } catch (err) { console.error('[DV] SDK create assignment failed:', err) }
  return dvPost('pth_assignments', body)
}

export async function updateAssignmentInDataverse(id: string, a: Partial<AssignmentPayload>): Promise<boolean> {
  const body: Record<string, unknown> = {}
  if (a.assignedHours !== undefined) body.pth_assignedhours = a.assignedHours
  if (a.weekStartDate) body.pth_weekstartdate = a.weekStartDate
  if (a.activityId) body['pth_activity@odata.bind'] = `/pth_activities(${a.activityId})`
  if (a.resourceId) body['pth_resource@odata.bind'] = `/pth_resources(${a.resourceId})`
  if (preferDirectWrite()) return dvPatch('pth_assignments', id, body)
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await Pth_assignmentsService.update(id, body as any)
    console.log('[DV] Assignment updated via SDK, id =', id); return true
  } catch (err) { console.error('[DV] SDK update assignment failed:', err) }
  return dvPatch('pth_assignments', id, body)
}

export async function deleteAssignmentInDataverse(id: string): Promise<boolean> {
  if (preferDirectWrite()) return dvDelete('pth_assignments', id)
  try { await Pth_assignmentsService.delete(id); return true } catch (err) { console.warn('SDK delete assignment unavailable:', err) }
  return dvDelete('pth_assignments', id)
}

// ═══════════════════════════════════════════════════════════════════════════
// ── SETTINGS CRUD ─────────────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════

export interface SettingsPayload {
  warningDaysThreshold: number
  upcomingWeeksWindow: number
  defaultWeeklyCapacity: number
  workloadWarningThreshold: number
  workloadCriticalThreshold: number
}

export async function updateSettingsInDataverse(id: string, s: SettingsPayload): Promise<boolean> {
  const body: Record<string, unknown> = {
    pth_warningdaysthreshold: s.warningDaysThreshold,
    pth_reportweekwindow: s.upcomingWeeksWindow,
    pth_defaultweeklycapacity: s.defaultWeeklyCapacity,
    pth_utilizationwarningthreshold: s.workloadWarningThreshold,
    pth_utilizationcriticalthreshold: s.workloadCriticalThreshold,
  }
  if (preferDirectWrite()) return dvPatch('pth_ppmsettings', id, body)
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await Pth_ppmsettingsService.update(id, body as any)
    console.log('[DV] Settings updated via SDK, id =', id); return true
  } catch (err) { console.error('[DV] SDK update settings failed:', err) }
  return dvPatch('pth_ppmsettings', id, body)
}

// ═══════════════════════════════════════════════════════════════════════════
// ── TASK TEMPLATES (editable standard per project type) ───────────────────
// ═══════════════════════════════════════════════════════════════════════════

export interface TaskTemplateRow {
  id: string
  projectType: string
  sequence: number
  task: string
  responsible: string
  leadtimeWeeks: number | null
  inputs: string
  workloadPct: number | null
}

export interface TaskTemplatePayload {
  projectType: string
  sequence: number
  task: string
  responsible: string
  leadtimeWeeks: number | null
  inputs: string
  workloadPct: number | null
}

function mapTaskTemplate(r: Row): TaskTemplateRow {
  return {
    id: r.pth_tasktemplateid as string,
    projectType: (r.pth_projecttype ?? '') as string,
    sequence: Number(r.pth_sequence ?? 0),
    task: (r.pth_name ?? '') as string,
    responsible: (r.pth_responsible ?? '') as string,
    leadtimeWeeks: r.pth_leadtimeweeks == null ? null : Number(r.pth_leadtimeweeks),
    inputs: (r.pth_inputs ?? '') as string,
    workloadPct: r.pth_workloadpct == null ? null : Number(r.pth_workloadpct),
  }
}

function buildTaskTemplateBody(p: TaskTemplatePayload): Record<string, unknown> {
  return {
    pth_name: p.task,
    pth_projecttype: p.projectType,
    pth_sequence: p.sequence,
    pth_responsible: p.responsible || null,
    pth_inputs: p.inputs || null,
    pth_leadtimeweeks: p.leadtimeWeeks,
    pth_workloadpct: p.workloadPct,
  }
}

/** Fetch all task-template rows (no generated SDK service — direct Web API only). */
export async function fetchTaskTemplates(): Promise<TaskTemplateRow[]> {
  const rows = await dvGet<Row>('pth_tasktemplates', '$orderby=pth_projecttype,pth_sequence&$top=1000')
  return rows.map(mapTaskTemplate)
}

export async function createTaskTemplate(p: TaskTemplatePayload): Promise<string | null> {
  return dvPost('pth_tasktemplates', buildTaskTemplateBody(p))
}

export async function updateTaskTemplate(id: string, p: TaskTemplatePayload): Promise<boolean> {
  return dvPatch('pth_tasktemplates', id, buildTaskTemplateBody(p))
}

export async function deleteTaskTemplate(id: string): Promise<boolean> {
  return dvDelete('pth_tasktemplates', id)
}

// ═══════════════════════════════════════════════════════════════════════════
// ── AAD / Microsoft Graph People Lookup ───────────────────────────────────
// ═══════════════════════════════════════════════════════════════════════════

export interface AadUser {
  id: string
  displayName: string
  mail: string | null
  jobTitle: string | null
  department: string | null
}

/**
 * Acquire a Microsoft Graph access token.
 * 1. Power Apps context helpers (utils.getGraphToken / webAPI.getGraphToken)
 * 2. Fallback: VITE_GRAPH_TOKEN env var for standalone dev mode
 */
async function getGraphToken(): Promise<string | null> {
  // 1. Try Power Apps context
  try {
    const ctx = getContext()
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctxAny = ctx as any
    if (typeof ctxAny?.utils?.getGraphToken === 'function') {
      return await ctxAny.utils.getGraphToken()
    }
    if (typeof ctxAny?.webAPI?.getGraphToken === 'function') {
      return await ctxAny.webAPI.getGraphToken()
    }
  } catch { /* not in Power Apps */ }

  // 2. Fallback to env var
  return (typeof import.meta !== 'undefined' && (import.meta as any).env?.VITE_GRAPH_TOKEN as string) || null
}

// ── Current-user profile (photo + job title) from Graph /me ──────────────

export interface UserProfile {
  photoUrl: string | null
  jobTitle: string | null
  department: string | null
  mail: string | null
}

/**
 * Fetches the signed-in user's profile photo and job title from
 * Microsoft Graph `/me`.  Returns `null` fields when unavailable.
 */
export async function fetchCurrentUserProfile(): Promise<UserProfile> {
  const result: UserProfile = { photoUrl: null, jobTitle: null, department: null, mail: null }

  const token = await getGraphToken()
  if (!token) {
    console.warn('[AAD] No Graph token — cannot fetch user profile')
    return result
  }

  // Fetch profile metadata (job title, department, mail)
  try {
    const res = await fetch(
      'https://graph.microsoft.com/v1.0/me?$select=jobTitle,department,mail',
      { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' } },
    )
    if (res.ok) {
      const json = await res.json()
      result.jobTitle = json.jobTitle ?? null
      result.department = json.department ?? null
      result.mail = json.mail ?? null
    }
  } catch (err) {
    console.error('[AAD] Graph /me profile error:', err)
  }

  // Fetch profile photo as blob URL
  try {
    const photoRes = await fetch(
      'https://graph.microsoft.com/v1.0/me/photo/$value',
      { headers: { Authorization: `Bearer ${token}` } },
    )
    if (photoRes.ok) {
      const blob = await photoRes.blob()
      result.photoUrl = URL.createObjectURL(blob)
    }
  } catch (err) {
    console.error('[AAD] Graph /me/photo error:', err)
  }

  return result
}

/**
 * Searches Azure AD (Entra ID) users matching a display name prefix.
 *
 * Inside Power Apps the access token is obtained via the platform context.
 * In standalone dev mode, falls back to the VITE_GRAPH_TOKEN env var.
 */
export async function searchAadUsers(query: string): Promise<AadUser[]> {
  if (!query || query.length < 2) return []

  const token = await getGraphToken()
  if (!token) {
    console.warn('[AAD] No Graph token available – skipping AAD lookup')
    return []
  }

  const filter = encodeURIComponent(`startswith(displayName,'${query.replace(/'/g, "''")}')`)
  const url = `https://graph.microsoft.com/v1.0/users?$filter=${filter}&$top=10&$select=id,displayName,mail,jobTitle,department`

  try {
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
    })
    if (!res.ok) {
      console.error(`[AAD] Graph search failed (${res.status}):`, await res.text())
      return []
    }
    const json = await res.json()
    return ((json.value ?? []) as AadUser[]).map((u) => ({
      id: u.id,
      displayName: u.displayName,
      mail: u.mail,
      jobTitle: u.jobTitle,
      department: u.department,
    }))
  } catch (err) {
    console.error('[AAD] Graph search error:', err)
    return []
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// ── Email Notifications via Power Automate Webhook ────────────────────────
// ═══════════════════════════════════════════════════════════════════════════

export interface ProjectCreationEmailPayload {
  /** Power Automate HTTP trigger URL */
  webhookUrl: string
  /** Recipient email (creator) */
  recipientEmail: string
  recipientName: string
  /** Project details */
  projectName: string
  projectCode: string
  category: string
  objective: string
  budgetAllocated: number
  siteName: string
  sponsorExecutive: string
  projectManager: string
  plannedStartDate: string
  plannedEndDate: string
  clientCode: string
  milestones: Array<{ name: string; targetDate: string }>
}

/**
 * Sends a project-creation notification email via a Power Automate HTTP
 * webhook.  The webhook flow is expected to receive a JSON body and use the
 * Office 365 / Outlook connector to send a formatted HTML email.
 *
 * Returns `true` if the call succeeded (2xx), `false` otherwise.
 */
export async function sendProjectCreationEmail(payload: ProjectCreationEmailPayload): Promise<boolean> {
  const { webhookUrl } = payload
  if (!webhookUrl || !webhookUrl.includes('logic.azure.com')) {
    console.warn('[Email] No valid Power Automate webhook URL — email not sent')
    return false
  }

  const milestoneRows = payload.milestones
    .map((m) => `<tr><td style="padding:6px 12px;border:1px solid #e5e7eb">${m.name}</td><td style="padding:6px 12px;border:1px solid #e5e7eb">${m.targetDate}</td></tr>`)
    .join('')

  const htmlBody = `
<div style="font-family:'Segoe UI',Roboto,Arial,sans-serif;max-width:640px;margin:0 auto">
  <div style="background:#007bc0;padding:20px 24px;border-radius:8px 8px 0 0">
    <h1 style="color:#fff;font-size:20px;margin:0">New Project Created</h1>
  </div>
  <div style="background:#ffffff;padding:24px;border:1px solid #e5e7eb;border-top:none;border-radius:0 0 8px 8px">
    <p style="margin:0 0 16px;color:#333">Hi <strong>${payload.recipientName}</strong>,</p>
    <p style="margin:0 0 20px;color:#555">Your project has been successfully created in the <strong>Project Tracker Hub</strong>. Here are the details:</p>
    <table style="width:100%;border-collapse:collapse;margin-bottom:20px">
      <tr><td style="padding:8px 12px;background:#f3f4f6;font-weight:600;width:40%;border:1px solid #e5e7eb">Project Name</td><td style="padding:8px 12px;border:1px solid #e5e7eb">${payload.projectName}</td></tr>
      <tr><td style="padding:8px 12px;background:#f3f4f6;font-weight:600;border:1px solid #e5e7eb">Project Code</td><td style="padding:8px 12px;border:1px solid #e5e7eb">${payload.projectCode}</td></tr>
      <tr><td style="padding:8px 12px;background:#f3f4f6;font-weight:600;border:1px solid #e5e7eb">Client Code</td><td style="padding:8px 12px;border:1px solid #e5e7eb">${payload.clientCode || '—'}</td></tr>
      <tr><td style="padding:8px 12px;background:#f3f4f6;font-weight:600;border:1px solid #e5e7eb">Category</td><td style="padding:8px 12px;border:1px solid #e5e7eb">${payload.category}</td></tr>
      <tr><td style="padding:8px 12px;background:#f3f4f6;font-weight:600;border:1px solid #e5e7eb">Budget Allocated</td><td style="padding:8px 12px;border:1px solid #e5e7eb">$${payload.budgetAllocated.toLocaleString()}</td></tr>
      <tr><td style="padding:8px 12px;background:#f3f4f6;font-weight:600;border:1px solid #e5e7eb">Site</td><td style="padding:8px 12px;border:1px solid #e5e7eb">${payload.siteName || '—'}</td></tr>
      <tr><td style="padding:8px 12px;background:#f3f4f6;font-weight:600;border:1px solid #e5e7eb">Sponsor / Executive</td><td style="padding:8px 12px;border:1px solid #e5e7eb">${payload.sponsorExecutive || '—'}</td></tr>
      <tr><td style="padding:8px 12px;background:#f3f4f6;font-weight:600;border:1px solid #e5e7eb">Project Manager</td><td style="padding:8px 12px;border:1px solid #e5e7eb">${payload.projectManager || '—'}</td></tr>
      <tr><td style="padding:8px 12px;background:#f3f4f6;font-weight:600;border:1px solid #e5e7eb">Planned Start</td><td style="padding:8px 12px;border:1px solid #e5e7eb">${payload.plannedStartDate}</td></tr>
      <tr><td style="padding:8px 12px;background:#f3f4f6;font-weight:600;border:1px solid #e5e7eb">Planned End</td><td style="padding:8px 12px;border:1px solid #e5e7eb">${payload.plannedEndDate}</td></tr>

      <tr><td style="padding:8px 12px;background:#f3f4f6;font-weight:600;border:1px solid #e5e7eb">Objective</td><td style="padding:8px 12px;border:1px solid #e5e7eb">${payload.objective || '—'}</td></tr>
    </table>
    ${payload.milestones.length > 0 ? `
    <h3 style="margin:0 0 8px;font-size:14px;color:#333">Auto-Generated Milestones</h3>
    <table style="width:100%;border-collapse:collapse;margin-bottom:20px">
      <tr><th style="padding:8px 12px;background:#007bc0;color:#fff;text-align:left;border:1px solid #e5e7eb">Milestone</th><th style="padding:8px 12px;background:#007bc0;color:#fff;text-align:left;border:1px solid #e5e7eb">Target Date</th></tr>
      ${milestoneRows}
    </table>` : ''}
    <p style="margin:16px 0 0;font-size:12px;color:#9ca3af">This is an automated notification from Project Tracker Hub.</p>
  </div>
</div>`

  const body = {
    trigger: 'project_created',
    recipientEmail: payload.recipientEmail,
    recipientName: payload.recipientName,
    subject: `Project Created: ${payload.projectName} (${payload.projectCode})`,
    htmlBody,
    projectName: payload.projectName,
    projectCode: payload.projectCode,
    category: payload.category,
    budgetAllocated: payload.budgetAllocated,
    plannedStartDate: payload.plannedStartDate,
    plannedEndDate: payload.plannedEndDate,
  }

  try {
    const res = await fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    })
    if (res.ok || res.status === 202) {
      console.log('[Email] Project creation notification sent successfully')
      return true
    }
    console.error(`[Email] Webhook responded with ${res.status}:`, await res.text())
    return false
  } catch (err) {
    console.error('[Email] Failed to trigger project creation email:', err)
    return false
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// ── Email Notifications via Azure Function + Graph API ────────────────────
// ═══════════════════════════════════════════════════════════════════════════

const FUNC_NOTIFY_URL = (typeof import.meta !== 'undefined' && (import.meta as any).env?.VITE_FUNC_NOTIFY_URL as string) || ''

export interface FuncProjectCreatedPayload {
  projectName: string
  projectCode: string
  category: string
  objective: string
  siteName: string
  sponsorExecutive: string
  projectManager: string
  budgetAllocated: number
  plannedStartDate: string
  plannedEndDate: string
  clientCode: string
  milestones: Array<{ name: string; targetDate: string }>
  creatorName: string
  creatorEmail: string
}

/**
 * Sends a project-creation notification via the `onProjectCreated` Azure
 * Function (HTTP trigger).  The function uses Microsoft Graph to send a
 * rich HTML email.
 *
 * Configure `VITE_FUNC_NOTIFY_URL` in `.env.local` (or as a Vite env)
 * e.g. `https://pth-notifications.azurewebsites.net/api/onProjectCreated?code=<func-key>`
 *
 * Returns `true` on success, `false` if skipped or failed.
 */
export async function sendProjectCreationViaFunction(payload: FuncProjectCreatedPayload): Promise<boolean> {
  if (!FUNC_NOTIFY_URL) {
    console.log('[FuncEmail] No VITE_FUNC_NOTIFY_URL configured — skipping Azure Function notification')
    return false
  }

  try {
    const res = await fetch(FUNC_NOTIFY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    })
    if (res.ok) {
      const data = await res.json()
      console.log('[FuncEmail] Project creation notification sent via Azure Function:', data)
      return true
    }
    console.error(`[FuncEmail] Azure Function responded with ${res.status}:`, await res.text())
    return false
  } catch (err) {
    console.error('[FuncEmail] Failed to call Azure Function:', err)
    return false
  }
}
