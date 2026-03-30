/**
 * dataverse.ts — Dataverse Web API client for the notification functions.
 *
 * Uses `@azure/identity` ClientSecretCredential to obtain a token scoped
 * to the Dataverse environment, then performs OData queries.
 */

import { ClientSecretCredential } from '@azure/identity'
import { config } from './config.js'

// ── Types ────────────────────────────────────────────────────────────────

export type RygStatus = 'RED' | 'YELLOW' | 'GREEN'
export type ActivityState = 'NOT_STARTED' | 'IN_PROGRESS' | 'BLOCKED' | 'DONE'
export type MilestoneStatus = 'OPEN' | 'IN_PROGRESS' | 'CLOSED' | 'DELAYED'

export interface DvProject {
  id: string
  name: string
  projectCode: string
  category: string
  site: string
  sponsorExecutive: string
  projectManagerName: string
  timeStatus: RygStatus
  plannedStartDate: string | null
  plannedEndDate: string | null
}

export interface DvActivity {
  id: string
  projectId: string
  name: string
  owner: string
  startDate: string
  endDate: string
  plannedEndDate: string
  closedDate: string | null
  status: ActivityState
  ryg: RygStatus
  isCriticalPath: boolean
}

export interface DvMilestone {
  id: string
  projectId: string
  name: string
  plannedDate: string
  status: MilestoneStatus
  ryg: RygStatus
  isCriticalPath: boolean
}

export interface DvUser {
  id: string
  fullname: string
  email: string
}

// ── Choice maps ──────────────────────────────────────────────────────────

const RYG_MAP: Record<number, RygStatus> = {
  100000000: 'RED',
  100000001: 'YELLOW',
  100000002: 'GREEN',
}

const ACTIVITY_STATUS_MAP: Record<number, ActivityState> = {
  100000000: 'NOT_STARTED',
  100000001: 'IN_PROGRESS',
  100000002: 'BLOCKED',
  100000003: 'DONE',
}

const MILESTONE_STATUS_MAP: Record<number, MilestoneStatus> = {
  100000000: 'OPEN',
  100000001: 'IN_PROGRESS',
  100000002: 'CLOSED',
  100000003: 'DELAYED',
}

const CATEGORY_MAP: Record<number, string> = {
  100000000: 'ECR',
  100000001: 'Path Forward',
  100000002: 'CIP',
  100000003: 'New Programs',
}

// ── Credential / fetch helpers ───────────────────────────────────────────

let _credential: ClientSecretCredential | null = null

function getCredential(): ClientSecretCredential {
  if (!_credential) {
    _credential = new ClientSecretCredential(
      config.tenantId(),
      config.clientId(),
      config.clientSecret(),
    )
  }
  return _credential
}

async function getToken(): Promise<string> {
  const cred = getCredential()
  const dvUrl = config.dataverseUrl().replace(/\/+$/, '')
  const res = await cred.getToken(`${dvUrl}/.default`)
  return res.token
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Row = Record<string, any>

async function dvGet<T>(entity: string, query = ''): Promise<T[]> {
  const dvUrl = config.dataverseUrl().replace(/\/+$/, '')
  const token = await getToken()
  const url = `${dvUrl}/api/data/v9.2/${entity}${query ? '?' + query : ''}`
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      Prefer: 'odata.maxpagesize=5000',
    },
  })
  if (!res.ok) {
    const text = await res.text()
    throw new Error(`Dataverse GET ${entity} failed (${res.status}): ${text}`)
  }
  const json = (await res.json()) as { value?: T[] }
  return (json.value ?? []) as T[]
}

// ── Public query functions ───────────────────────────────────────────────

export async function fetchProjects(): Promise<DvProject[]> {
  const rows = await dvGet<Row>(
    'pth_projectses',
    '$select=pth_projectid,pth_projectname,pth_projectidexternal,pth_category,pth_site,pth_sponsorexecutive,pth_projectmanagername,pth_timestatus,pth_plannedstartdate,pth_plannedenddate',
  )
  return rows.map((r) => ({
    id: r.pth_projectid,
    name: r.pth_projectname ?? '',
    projectCode: r.pth_projectidexternal ?? '',
    category: CATEGORY_MAP[r.pth_category as number] ?? 'CIP',
    site: r.pth_site ?? '',
    sponsorExecutive: r.pth_sponsorexecutive ?? '',
    projectManagerName: r.pth_projectmanagername ?? '',
    timeStatus: RYG_MAP[r.pth_timestatus as number] ?? 'GREEN',
    plannedStartDate: r.pth_plannedstartdate ?? null,
    plannedEndDate: r.pth_plannedenddate ?? null,
  }))
}

export async function fetchActivities(): Promise<DvActivity[]> {
  const rows = await dvGet<Row>(
    'pth_activitieses',
    '$select=pth_activityid,_pth_project_value,pth_name,pth_owner,pth_startdate,pth_enddate,pth_plannedenddate,pth_closeddate,pth_status,pth_ryg,pth_iscriticalpath',
  )
  return rows.map((r) => ({
    id: r.pth_activityid,
    projectId: r._pth_project_value ?? '',
    name: r.pth_name ?? '',
    owner: r.pth_owner ?? '',
    startDate: (r.pth_startdate ?? '').slice(0, 10),
    endDate: (r.pth_enddate ?? '').slice(0, 10),
    plannedEndDate: (r.pth_plannedenddate ?? '').slice(0, 10),
    closedDate: r.pth_closeddate ? (r.pth_closeddate as string).slice(0, 10) : null,
    status: ACTIVITY_STATUS_MAP[r.pth_status as number] ?? 'NOT_STARTED',
    ryg: RYG_MAP[r.pth_ryg as number] ?? 'GREEN',
    isCriticalPath: !!r.pth_iscriticalpath,
  }))
}

export async function fetchMilestones(): Promise<DvMilestone[]> {
  const rows = await dvGet<Row>(
    'pth_milestoneses',
    '$select=pth_milestoneid,_pth_project_value,pth_name,pth_planneddate,pth_status,pth_ryg,pth_iscriticalpath',
  )
  return rows.map((r) => ({
    id: r.pth_milestoneid,
    projectId: r._pth_project_value ?? '',
    name: r.pth_name ?? '',
    plannedDate: (r.pth_planneddate ?? '').slice(0, 10),
    status: MILESTONE_STATUS_MAP[r.pth_status as number] ?? 'OPEN',
    ryg: RYG_MAP[r.pth_ryg as number] ?? 'GREEN',
    isCriticalPath: !!r.pth_iscriticalpath,
  }))
}

/** Fetch interactive (non-disabled, non-service) users from systemusers */
export async function fetchUsers(): Promise<DvUser[]> {
  const rows = await dvGet<Row>(
    'systemusers',
    '$select=systemuserid,fullname,internalemailaddress&$filter=isdisabled eq false and accessmode eq 0',
  )
  return rows.map((r) => ({
    id: r.systemuserid,
    fullname: r.fullname ?? '',
    email: r.internalemailaddress ?? '',
  }))
}
