import { useCallback, useEffect, useMemo, useState, type ReactElement } from 'react'
import { AnimatePresence, motion } from 'framer-motion'
import { getContext, type IContext } from '@microsoft/power-apps/app'
import logoImg from '/logo.png?inline'
import { fetchAllFromDataverse, fetchDataverseUsers, isDataverseConfigured, createProjectInDataverse, deleteProjectInDataverse, updateProjectInDataverse, createMilestoneInDataverse, updateMilestoneInDataverse, deleteMilestoneInDataverse, createActivityInDataverse, updateActivityInDataverse, deleteActivityInDataverse, createResourceInDataverse, updateResourceInDataverse, deleteResourceInDataverse, updateSettingsInDataverse, searchAadUsers, fetchCurrentUserProfile, sendProjectCreationEmail, sendProjectCreationViaFunction, fetchTaskTemplates, createTaskTemplate, updateTaskTemplate, deleteTaskTemplate, type TaskTemplateRow, type AadUser, type DataverseUser } from './dataverse'
import {
  Bell,
  BellRing,
  Briefcase,
  CalendarClock,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Clock3,
  Database,
  FolderKanban,
  LayoutDashboard,
  LayoutGrid,
  List,
  Menu,
  Moon,
  Pencil,
  GripVertical,
  Plus,
  Save,
  Search,
  ShieldAlert,
  SlidersHorizontal,
  Sun,
  Trash2,
  User,
  Users,
  X,
  Zap,
  Send,
  ExternalLink,
  CheckCircle,
  XCircle,
  ToggleLeft,
  ToggleRight,
  PanelLeftClose,
  PanelLeftOpen,
} from 'lucide-react'
import { useForm } from 'react-hook-form'
import './App.css'

/** Apply dynamic CSS properties via callback ref (avoids inline style attributes) */
function dynRef(props: Record<string, string>) {
  return (el: HTMLElement | null) => {
    if (!el) return
    for (const [key, value] of Object.entries(props)) {
      el.style.setProperty(key, value)
    }
  }
}

type FlowTriggerType = 'overdue_milestones' | 'delayed_activities' | 'workload_threshold' | 'project_created'

type PowerAutomateFlow = {
  id: string
  name: string
  description: string
  triggerType: FlowTriggerType
  enabled: boolean
}

type FlowLogEntry = {
  id: string
  flowName: string
  triggeredAt: string
  status: 'success' | 'failed'
  payload: string
}

type Role = 'PROJECT_MANAGER' | 'MANAGER' | 'DIRECTOR'
type NavPage = 'overview' | 'workload' | 'reports' | 'alerts' | 'admin' | 'masterData' | 'createProject'
type MasterDataTab = 'people' | 'customers' | 'suppliers' | 'templates'
type ReportActivityTab = 'upcoming' | 'closed' | 'delayed'
type ProjectCategory = 'ECR' | 'CIP' | 'Path Forward' | 'New Programs'
type RygStatus = 'RED' | 'YELLOW' | 'GREEN'
type ReportTab = 'projectsStatus' | 'taskTracking' | 'prioritization' | 'workloadOverview'
type ActivityState = 'NOT_STARTED' | 'IN_PROGRESS' | 'DONE' | 'NOT_APPLICABLE'

type Site = { id: string; name: string }
type Partner = { id: string; name: string }

type Person = {
  id: string
  name: string
  initials: string
  role: Role | 'RESOURCE'
  managerId?: string
  area?: string
  location?: string[]
  resourceKind?: string
  weeklyCapacity?: number
}



type Project = {
  id: string
  projectCode: string
  name: string
  clientCode: string
  category: ProjectCategory
  budgetAllocated: number
  siteIds: string[]
  sponsorExecutiveId: string
  projectManagerId: string
  clientIds: string[]
  supplierIds: string[]
  objective: string
  plannedStartDate: string
  plannedEndDate: string
  projectType?: string
}

type Milestone = {
  id: string
  projectId: string
  name: string
  targetDate: string
  doneDate?: string
}

type Activity = {
  id: string
  projectId: string
  milestoneId?: string
  name: string
  ownerId: string
  startDate: string
  endDate: string
  doneDate?: string
  state: ActivityState
  criticalPath: boolean
  responsible?: string
  workloadPct?: number | null
}

type Dependency = {
  id: string
  projectId: string
  activityId: string
  predecessorActivityId: string
}

type Assignment = {
  id: string
  personId: string
  activityId: string
  weekStart: string
  assignedHours: number
  capacityHours: number
}

type NotificationItem = {
  id: string
  createdAt: string
  message: string
  severity: RygStatus
  recipients: Role[]
}

type AppSettings = {
  dvSettingsId?: string
  warningDaysThreshold: number
  upcomingWeeksWindow: number
  lastClosedWeeksWindow: number
  defaultWeeklyCapacity: number
  workloadWarningThreshold: number
  workloadCriticalThreshold: number
}

type AppState = {
  projects: Project[]
  milestones: Milestone[]
  activities: Activity[]
  dependencies: Dependency[]
  assignments: Assignment[]
  notifications: NotificationItem[]
  settings: AppSettings
  people: Person[]
  clients: Partner[]
  suppliers: Partner[]
}

type NewProjectForm = {
  name: string
  clientCode: string
  category: ProjectCategory
  projectType: string
  budgetAllocated: number
  siteIds: string[]
  sponsorExecutiveId: string
  projectManagerId: string
  clientIds: string[]
  supplierIds: string[]
  objective: string
  plannedStartDate: string
  plannedEndDate: string
}

type ProjectOverviewForm = {
  name: string
  clientCode: string
  category: ProjectCategory
  budgetAllocated: number
  siteIds: string[]
  sponsorExecutiveId: string
  projectManagerId: string
  clientIds: string[]
  supplierIds: string[]
  objective: string
  plannedStartDate: string
  plannedEndDate: string
}

type MilestoneForm = {
  name: string
  targetDate: string
}

type ActivityForm = {
  name: string
  ownerId: string
  startDate: string
  endDate: string
  workloadPct: number
  responsible: string
  state: ActivityState
}

type ActivityEditMode = 'create' | 'edit'

/** Inline task draft used during project creation */
type InlineTask = {
  _key: string
  name: string
  ownerId: string
  startDate: string
  endDate: string
}

/** Inline milestone draft used during project creation */
type InlineMilestone = {
  _key: string
  name: string
  targetDate: string
  tasks: InlineTask[]
}

const ROLE_LABEL: Record<Role, string> = {
  PROJECT_MANAGER: 'Project Manager',
  MANAGER: 'Manager',
  DIRECTOR: 'Director',
}

/**
 * ── PROJECT TASK STANDARD ──────────────────────────────────────────────────
 * Per-category standard breakdown: each category has phase milestones, and
 * each milestone has standard tasks. Offsets are in DAYS relative to the
 * project's planned start date. Generated tasks start UNASSIGNED — the manager
 * assigns each to a specific resource afterward.
 *
 * ⚠️ PLACEHOLDER CONTENT — replace with the real company standard
 *    (source: data/task-template.csv).
 */
type TaskTemplate = { name: string; area: string; startOffset: number; duration: number; criticalPath: boolean }
type MilestoneTemplate = { name: string; offset: number; tasks: TaskTemplate[] }

const CATEGORY_TEMPLATE: Record<ProjectCategory, MilestoneTemplate[]> = {
  ECR: [
    { name: 'ECR Kickoff', offset: 0, tasks: [
      { name: 'Requirements Gathering', area: 'PPS', startOffset: 0, duration: 5, criticalPath: true },
      { name: 'Technical Design Review', area: 'ENG', startOffset: 5, duration: 7, criticalPath: true },
    ] },
    { name: 'Engineering Release Review', offset: 17, tasks: [
      { name: 'Process FMEA Update', area: 'QMM', startOffset: 12, duration: 5, criticalPath: false },
    ] },
    { name: 'ECR Implementation Complete', offset: 42, tasks: [
      { name: 'Validation Testing', area: 'QMM', startOffset: 17, duration: 10, criticalPath: true },
    ] },
  ],
  CIP: [
    { name: 'CIP Baseline', offset: 0, tasks: [
      { name: 'Baseline Data Collection', area: 'QMM', startOffset: 0, duration: 5, criticalPath: false },
    ] },
    { name: 'CIP Improvement Gate', offset: 10, tasks: [
      { name: 'Improvement Workshop', area: 'MSE', startOffset: 5, duration: 3, criticalPath: true },
    ] },
    { name: 'CIP Sustainment Review', offset: 35, tasks: [
      { name: 'Sustainment Audit', area: 'QMM', startOffset: 30, duration: 4, criticalPath: false },
    ] },
  ],
  'Path Forward': [
    { name: 'Path Forward Alignment', offset: 0, tasks: [
      { name: 'Stakeholder Alignment', area: 'PPS', startOffset: 0, duration: 3, criticalPath: true },
    ] },
    { name: 'Path Forward Validation', offset: 14, tasks: [
      { name: 'Process Controls Phase-in', area: 'MSE', startOffset: 10, duration: 14, criticalPath: true },
    ] },
    { name: 'Path Forward Sign-Off', offset: 45, tasks: [
      { name: 'Final Sign-Off', area: 'PPS', startOffset: 40, duration: 2, criticalPath: false },
    ] },
  ],
  'New Programs': [
    { name: 'Program Kickoff', offset: 0, tasks: [
      { name: 'Program Kickoff', area: 'PPS', startOffset: 0, duration: 3, criticalPath: true },
    ] },
    { name: 'Prototype Gate', offset: 21, tasks: [
      { name: 'Prototype Build', area: 'ENG', startOffset: 15, duration: 21, criticalPath: true },
    ] },
    { name: 'SOP Readiness', offset: 63, tasks: [
      { name: 'SOP Readiness Review', area: 'QMM', startOffset: 60, duration: 5, criticalPath: true },
    ] },
  ],
}

// Derived: milestone names + offsets only (keeps existing references working)
const CATEGORY_MILESTONES: Record<ProjectCategory, Array<{ name: string; offset: number }>> =
  Object.fromEntries(
    (Object.entries(CATEGORY_TEMPLATE) as Array<[ProjectCategory, MilestoneTemplate[]]>)
      .map(([cat, ms]) => [cat, ms.map((m) => ({ name: m.name, offset: m.offset }))]),
  ) as Record<ProjectCategory, Array<{ name: string; offset: number }>>

const sites: Site[] = [
  { id: 'site_tca', name: 'SlpP' },
  { id: 'site_slp', name: 'TlP' },
]

/** Auto-increment counter for generating project codes */
let _projectCodeSeq = 0
function nextProjectCode(siteIds: string[]): string {
  _projectCodeSeq++
  const siteAbbr = siteIds.length > 0
    ? (sites.find((s) => s.id === siteIds[0])?.name ?? 'XX').toUpperCase().replace(/[^A-Z]/g, '')
    : 'XX'
  return `PTH-${siteAbbr}-${String(_projectCodeSeq).padStart(4, '0')}`
}



const INITIAL_CLIENTS: Partner[] = [
  { id: 'client_ford', name: 'Ford' },
  { id: 'client_gm', name: 'GM' },
  { id: 'client_vw', name: 'VW' },
  { id: 'client_toyota', name: 'Toyota' },
  { id: 'client_stellantis', name: 'Stellantis' },
  { id: 'client_bmw', name: 'BMW' },
  { id: 'client_mercedes', name: 'Mercedes-Benz' },
  { id: 'client_na', name: 'Not Applicable' },
]

const INITIAL_SUPPLIERS: Partner[] = [
  { id: 'supplier_rbcb', name: 'RBCB' },
  { id: 'supplier_dmp', name: 'DMP' },
  { id: 'supplier_a', name: 'Supplier A' },
  { id: 'supplier_conti', name: 'Continental' },
  { id: 'supplier_bosch', name: 'Bosch' },
  { id: 'supplier_zf', name: 'ZF Friedrichshafen' },
  { id: 'supplier_na', name: 'Not Applicable' },
]

const INITIAL_PEOPLE: Person[] = [
  { id: 'u_dir_1', name: 'Laura Schneider', initials: 'LS', role: 'DIRECTOR' },
  { id: 'u_mgr_1', name: 'Tobias Klein', initials: 'TK', role: 'MANAGER', managerId: 'u_dir_1' },
  { id: 'u_mgr_2', name: 'Marta Alonso', initials: 'MA', role: 'MANAGER', managerId: 'u_dir_1' },
  { id: 'u_pm_1', name: 'Rafael Gomez', initials: 'RG', role: 'PROJECT_MANAGER', managerId: 'u_mgr_1' },
  { id: 'u_pm_2', name: 'Sven Maurer', initials: 'SM', role: 'PROJECT_MANAGER', managerId: 'u_mgr_1' },
  { id: 'u_pm_3', name: 'Clara Rossi', initials: 'CR', role: 'PROJECT_MANAGER', managerId: 'u_mgr_2' },
  { id: 'u_pm_4', name: 'Jan Krüger', initials: 'JK', role: 'PROJECT_MANAGER', managerId: 'u_mgr_2' },
  { id: 'u_res_1', name: 'Nina Becker', initials: 'NB', role: 'RESOURCE', managerId: 'u_mgr_1' },
  { id: 'u_res_2', name: 'Viktor Hahn', initials: 'VH', role: 'RESOURCE', managerId: 'u_mgr_1' },
  { id: 'u_res_3', name: 'Iris Weber', initials: 'IW', role: 'RESOURCE', managerId: 'u_mgr_1' },
  { id: 'u_res_4', name: 'Pablo Diaz', initials: 'PD', role: 'RESOURCE', managerId: 'u_mgr_2' },
  { id: 'u_res_5', name: 'Elena Costa', initials: 'EC', role: 'RESOURCE', managerId: 'u_mgr_2' },
  { id: 'u_res_6', name: 'Marco Lenz', initials: 'ML', role: 'RESOURCE', managerId: 'u_mgr_2' },
  { id: 'u_res_7', name: 'Annika Braun', initials: 'AB', role: 'RESOURCE', managerId: 'u_mgr_1' },
  { id: 'u_res_8', name: 'Felix Hartmann', initials: 'FH', role: 'RESOURCE', managerId: 'u_mgr_1' },
  { id: 'u_res_9', name: 'Sofia Martín', initials: 'SM', role: 'RESOURCE', managerId: 'u_mgr_2' },
  { id: 'u_res_10', name: 'Lukas Richter', initials: 'LR', role: 'RESOURCE', managerId: 'u_mgr_2' },
]

function uid(prefix: string): string {
  return `${prefix}_${Math.random().toString(36).slice(2, 9)}`
}

function addDays(date: string, days: number): string {
  const base = new Date(`${date}T00:00:00`)
  base.setDate(base.getDate() + days)
  return base.toISOString().slice(0, 10)
}

function todayISO(): string {
  return new Date().toISOString().slice(0, 10)
}

function toDate(iso: string): Date {
  return new Date(`${iso}T00:00:00`)
}

function dayDiff(fromISO: string, toISO: string): number {
  const a = toDate(toISO).getTime(); const b = toDate(fromISO).getTime()
  if (isNaN(a) || isNaN(b)) return 0   // missing date → no delta (e.g. N/A tasks)
  return Math.ceil((a - b) / (1000 * 60 * 60 * 24))
}

function weekStartISO(iso: string): string {
  const date = toDate(iso)
  const day = date.getDay()
  const shift = day === 0 ? -6 : 1 - day
  date.setDate(date.getDate() + shift)
  return date.toISOString().slice(0, 10)
}

/** ISO-8601 calendar week number (Monday-based, week 1 contains Jan 4) — fully UTC-safe */
function isoCalendarWeek(iso: string): number {
  const [y, m, d] = iso.split('-').map(Number)
  const date = new Date(Date.UTC(y, m - 1, d))
  // Move to Thursday of current week (ISO weeks are Thursday-anchored)
  date.setUTCDate(date.getUTCDate() + 4 - (date.getUTCDay() || 7))
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1))
  return Math.ceil(((date.getTime() - yearStart.getTime()) / 86400000 + 1) / 7)
}

/** Monday (ISO start) of a given ISO year + week — UTC-safe */
function mondayOfCW(year: number, cw: number): string {
  const jan4 = new Date(Date.UTC(year, 0, 4))
  const dow = jan4.getUTCDay() || 7 // Mon=1..Sun=7
  const d = new Date(Date.UTC(year, 0, 4 - dow + 1 + (cw - 1) * 7))
  return d.toISOString().slice(0, 10)
}

/** Get all CW numbers that overlap a given YYYY-MM — builds ISO strings directly, no timezone pitfalls */
function cwsInMonth(ym: string): number[] {
  const [y, m] = ym.split('-').map(Number)
  const daysInMonth = new Date(Date.UTC(y, m, 0)).getUTCDate()
  const cws = new Set<number>()
  for (let day = 1; day <= daysInMonth; day++) {
    const iso = `${ym}-${String(day).padStart(2, '0')}`
    cws.add(isoCalendarWeek(iso))
  }
  return [...cws].sort((a, b) => a - b)
}

function formatDate(iso: string): string {
  const d = toDate(iso)
  if (isNaN(d.getTime())) return '—'
  return new Intl.DateTimeFormat('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }).format(d)
}

function formatShortDate(iso: string): string {
  const d = toDate(iso)
  if (isNaN(d.getTime())) return '—'   // blank/invalid (e.g. N/A tasks have no dates)
  return new Intl.DateTimeFormat('en-US', { month: 'short', day: 'numeric' }).format(d)
}

function calendarWeek(iso: string): string {
  const d = new Date(toDate(iso).getTime())
  d.setDate(d.getDate() + 3 - ((d.getDay() + 6) % 7))
  const week1 = new Date(d.getFullYear(), 0, 4)
  const weekNum = Math.round(((d.getTime() - week1.getTime()) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7) + 1
  return `CW ${String(weekNum).padStart(2, '0')}`
}

function rygFromDueDate(dueISO: string, warningDays: number): RygStatus {
  const days = dayDiff(todayISO(), dueISO)
  if (days < 0) return 'RED'
  if (days <= warningDays) return 'YELLOW'
  return 'GREEN'
}

/** Buffer (in days) a task should ideally finish ahead of its due date. */
const STATUS_BUFFER_DAYS = 14 // 2 weeks

/**
 * Completion-vs-buffer status for a single task.
 *  - Done:   GREEN if finished on/before (due − buffer); YELLOW if finished by
 *            the due date (inside the buffer window); RED if finished late.
 *  - Open:   RED if already past due; YELLOW if due within the buffer window
 *            (at risk of missing the buffer); GREEN otherwise.
 */
function rygForActivity(act: { endDate: string; doneDate?: string }, bufferDays = STATUS_BUFFER_DAYS): RygStatus {
  const due = act.endDate
  if (!due) return 'GREEN'
  if (act.doneDate) {
    const slack = dayDiff(act.doneDate, due) // due − done; positive = finished early
    if (slack >= bufferDays) return 'GREEN'
    if (slack >= 0) return 'YELLOW'
    return 'RED'
  }
  const daysToDue = dayDiff(todayISO(), due) // due − today
  if (daysToDue < 0) return 'RED'
  if (daysToDue <= bufferDays) return 'YELLOW'
  return 'GREEN'
}

function rankStatus(status: RygStatus): number {
  if (status === 'RED') return 0
  if (status === 'YELLOW') return 1
  return 2
}

function generateInitialState(): AppState {
  const projects: Project[] = [
    {
      id: 'p_1',
      projectCode: nextProjectCode(['site_tca']),
      name: 'Assembly Line CIP – TCA Station 12',
      clientCode: 'BC-001',
      category: 'CIP',
      budgetAllocated: 120000,
      siteIds: ['site_tca'],
      sponsorExecutiveId: 'u_dir_1',
      projectManagerId: 'u_pm_1',
      clientIds: ['client_ford'],
      supplierIds: ['supplier_rbcb'],
      objective: 'Reduce cycle time in final assembly by 8% without quality drift.',
      plannedStartDate: addDays(todayISO(), -30),
      plannedEndDate: addDays(todayISO(), 60),
    },
    {
      id: 'p_2',
      projectCode: nextProjectCode(['site_slp']),
      name: 'ESC Sensor New Program',
      clientCode: 'BC-002',
      category: 'New Programs',
      budgetAllocated: 250000,
      siteIds: ['site_slp'],
      sponsorExecutiveId: 'u_dir_1',
      projectManagerId: 'u_pm_2',
      clientIds: ['client_gm'],
      supplierIds: ['supplier_dmp'],
      objective: 'Prepare ESC sensor manufacturing and validation readiness for SOP.',
      plannedStartDate: addDays(todayISO(), -14),
      plannedEndDate: addDays(todayISO(), 90),
    },
    {
      id: 'p_3',
      projectCode: nextProjectCode(['site_tca', 'site_slp']),
      name: 'Brake Caliper Path Forward',
      clientCode: 'BC-003',
      category: 'Path Forward',
      budgetAllocated: 80000,
      siteIds: ['site_tca', 'site_slp'],
      sponsorExecutiveId: 'u_dir_1',
      projectManagerId: 'u_pm_3',
      clientIds: ['client_vw'],
      supplierIds: ['supplier_a'],
      objective: 'Stabilize backlog and phase-in process controls in brake caliper line.',
      plannedStartDate: addDays(todayISO(), -21),
      plannedEndDate: addDays(todayISO(), 45),
    },
    {
      id: 'p_4',
      projectCode: nextProjectCode(['site_slp']),
      name: 'ABS Module ECR',
      clientCode: 'BC-004',
      category: 'ECR',
      budgetAllocated: 65000,
      siteIds: ['site_slp'],
      sponsorExecutiveId: 'u_dir_1',
      projectManagerId: 'u_pm_1',
      clientIds: ['client_toyota'],
      supplierIds: ['supplier_rbcb'],
      objective: 'Implement ABS module engineering change with contained launch.',
      plannedStartDate: addDays(todayISO(), -7),
      plannedEndDate: addDays(todayISO(), 42),
    },
    {
      id: 'p_5',
      projectCode: nextProjectCode(['site_tca']),
      name: 'Sparkplugs Manufacturing CIP',
      clientCode: 'BC-005',
      category: 'CIP',
      budgetAllocated: 95000,
      siteIds: ['site_tca'],
      sponsorExecutiveId: 'u_dir_1',
      projectManagerId: 'u_pm_2',
      clientIds: ['client_ford', 'client_gm'],
      supplierIds: ['supplier_dmp'],
      objective: 'Improve scrap reduction and process capability in sparkplug production.',
      plannedStartDate: addDays(todayISO(), -10),
      plannedEndDate: addDays(todayISO(), 50),
    },
    {
      id: 'p_6',
      projectCode: nextProjectCode(['site_tca', 'site_slp']),
      name: 'ADAS Path Forward – Radar Cell',
      clientCode: 'BC-006',
      category: 'Path Forward',
      budgetAllocated: 180000,
      siteIds: ['site_tca', 'site_slp'],
      sponsorExecutiveId: 'u_dir_1',
      projectManagerId: 'u_pm_3',
      clientIds: ['client_vw', 'client_toyota'],
      supplierIds: ['supplier_a', 'supplier_rbcb'],
      objective: 'Recover ADAS radar cell throughput and delivery adherence.',
      plannedStartDate: addDays(todayISO(), -28),
      plannedEndDate: addDays(todayISO(), 70),
    },
    {
      id: 'p_7',
      projectCode: nextProjectCode(['site_tca']),
      name: 'Steering Column ECR – Phase 2',
      clientCode: 'BC-007',
      category: 'ECR',
      budgetAllocated: 72000,
      siteIds: ['site_tca'],
      sponsorExecutiveId: 'u_dir_1',
      projectManagerId: 'u_pm_4',
      clientIds: ['client_bmw'],
      supplierIds: ['supplier_zf'],
      objective: 'Implement revised steering column per updated BMW specification.',
      plannedStartDate: addDays(todayISO(), -5),
      plannedEndDate: addDays(todayISO(), 55),
    },
    {
      id: 'p_8',
      projectCode: nextProjectCode(['site_slp']),
      name: 'Battery Pack New Program – Gen2',
      clientCode: 'BC-008',
      category: 'New Programs',
      budgetAllocated: 420000,
      siteIds: ['site_slp'],
      sponsorExecutiveId: 'u_dir_1',
      projectManagerId: 'u_pm_1',
      clientIds: ['client_stellantis'],
      supplierIds: ['supplier_conti', 'supplier_bosch'],
      objective: 'Launch Gen2 battery pack assembly for Stellantis EV platform.',
      plannedStartDate: addDays(todayISO(), -35),
      plannedEndDate: addDays(todayISO(), 120),
    },
    {
      id: 'p_9',
      projectCode: nextProjectCode(['site_tca']),
      name: 'Wiring Harness CIP – Defect Reduction',
      clientCode: 'BC-009',
      category: 'CIP',
      budgetAllocated: 55000,
      siteIds: ['site_tca'],
      sponsorExecutiveId: 'u_dir_1',
      projectManagerId: 'u_pm_2',
      clientIds: ['client_mercedes'],
      supplierIds: ['supplier_dmp'],
      objective: 'Reduce wiring harness defect rate from 3.2% to <1% via poke-yoke.',
      plannedStartDate: addDays(todayISO(), -18),
      plannedEndDate: addDays(todayISO(), 40),
    },
    {
      id: 'p_10',
      projectCode: nextProjectCode(['site_tca', 'site_slp']),
      name: 'EV Drivetrain Path Forward',
      clientCode: 'BC-010',
      category: 'Path Forward',
      budgetAllocated: 310000,
      siteIds: ['site_tca', 'site_slp'],
      sponsorExecutiveId: 'u_dir_1',
      projectManagerId: 'u_pm_3',
      clientIds: ['client_vw', 'client_bmw'],
      supplierIds: ['supplier_zf', 'supplier_bosch'],
      objective: 'Stabilize EV drivetrain test throughput and reduce rework loops.',
      plannedStartDate: addDays(todayISO(), -42),
      plannedEndDate: addDays(todayISO(), 85),
    },
    {
      id: 'p_11',
      projectCode: nextProjectCode(['site_slp']),
      name: 'Fuel Injector ECR – Emission Compliance',
      clientCode: 'BC-011',
      category: 'ECR',
      budgetAllocated: 88000,
      siteIds: ['site_slp'],
      sponsorExecutiveId: 'u_dir_1',
      projectManagerId: 'u_pm_4',
      clientIds: ['client_ford', 'client_gm'],
      supplierIds: ['supplier_rbcb'],
      objective: 'Re-engineer fuel injector nozzle to meet Euro 7 emission limits.',
      plannedStartDate: addDays(todayISO(), -12),
      plannedEndDate: addDays(todayISO(), 65),
    },
    {
      id: 'p_12',
      projectCode: nextProjectCode(['site_tca']),
      name: 'Paint Shop CIP – VOC Reduction',
      clientCode: 'BC-012',
      category: 'CIP',
      budgetAllocated: 145000,
      siteIds: ['site_tca'],
      sponsorExecutiveId: 'u_dir_1',
      projectManagerId: 'u_pm_1',
      clientIds: ['client_toyota', 'client_stellantis'],
      supplierIds: ['supplier_a', 'supplier_conti'],
      objective: 'Cut VOC emissions in paint shop by 25% while maintaining finish quality.',
      plannedStartDate: addDays(todayISO(), -20),
      plannedEndDate: addDays(todayISO(), 75),
    },
  ]

  const milestones: Milestone[] = projects.flatMap((project, index) =>
    CATEGORY_MILESTONES[project.category].map((template, templateIndex) => {
      const shifted = template.offset + index * 2 + templateIndex * 3
      return {
        id: uid('ms'),
        projectId: project.id,
        name: template.name,
        targetDate: addDays(todayISO(), shifted),
      }
    }),
  )

  const activityOwners = ['u_res_1', 'u_res_2', 'u_res_3', 'u_res_4', 'u_res_5', 'u_res_6', 'u_res_7', 'u_res_8', 'u_res_9', 'u_res_10']
  const activityTemplates = [
    'Requirements Gathering', 'Technical Design Review', 'Prototype Build',
    'Validation Testing', 'Process FMEA Update', 'Tooling Procurement',
    'Pilot Run Coordination', 'Supplier Qualification', 'Documentation Package',
    'Customer Sample Submission',
  ]
  const activities: Activity[] = []

  projects.forEach((project, projectIdx) => {
    const pMilestones = milestones.filter((ms) => ms.projectId === project.id)
    const doneThreshold = [2, 4, 3, 1, 5, 3, 4, 2, 3, 5, 1, 4][projectIdx % 12]
    const actCount = 6 + (projectIdx % 4) // 6-9 activities per project for variety
    for (let i = 0; i < actCount; i += 1) {
      const start = addDays(todayISO(), -18 + i * 8 - projectIdx * 2)
      const end = addDays(start, 7 + (i % 3))
      const isDone = i < doneThreshold
      const overdue = i === (doneThreshold) && projectIdx % 2 === 0
      const dueSoon = i === (doneThreshold + 1)
      const endDate = overdue ? addDays(todayISO(), -3 - projectIdx) : dueSoon ? addDays(todayISO(), 3 + (projectIdx % 2)) : end

      activities.push({
        id: uid('act'),
        projectId: project.id,
        milestoneId: pMilestones[i % pMilestones.length]?.id,
        name: activityTemplates[(projectIdx + i) % activityTemplates.length],
        ownerId: activityOwners[(projectIdx + i) % activityOwners.length],
        startDate: start,
        endDate,
        doneDate: isDone ? addDays(endDate, -1) : undefined,
        state: isDone ? 'DONE' : i % 2 === 0 ? 'IN_PROGRESS' : 'NOT_STARTED',
        criticalPath: i === 2 || i === 3,
      })
    }
  })

  const dependencies: Dependency[] = projects.flatMap((project) => {
    const pActivities = activities.filter((activity) => activity.projectId === project.id)
    return pActivities.slice(1).map((activity, idx) => ({
      id: uid('dep'),
      projectId: project.id,
      activityId: activity.id,
      predecessorActivityId: pActivities[idx].id,
    }))
  })

  const week0 = weekStartISO(todayISO())
  const weekStarts = Array.from({ length: 8 }, (_, idx) => addDays(week0, idx * 7))

  const assignments: Assignment[] = weekStarts.flatMap((week, wIndex) =>
    INITIAL_PEOPLE
      .filter((p) => p.role === 'RESOURCE')
      .flatMap((person, pIndex) => {
        const activityA = activities[(wIndex * 2 + pIndex) % activities.length]
        const activityB = activities[(wIndex * 2 + pIndex + 7) % activities.length]
        const base = 22 + ((pIndex + wIndex) % 3) * 8
        const loadBump = person.id === 'u_res_2' && wIndex % 3 === 0 ? 16 : person.id === 'u_res_4' && wIndex % 4 === 1 ? 11 : 0

        return [
          {
            id: uid('as'),
            personId: person.id,
            activityId: activityA.id,
            weekStart: week,
            assignedHours: base,
            capacityHours: 40,
          },
          {
            id: uid('as'),
            personId: person.id,
            activityId: activityB.id,
            weekStart: week,
            assignedHours: 12 + loadBump,
            capacityHours: 40,
          },
        ]
      }),
  )

  const notifications: NotificationItem[] = [
    {
      id: uid('ntf'),
      createdAt: todayISO(),
      severity: 'RED',
      message: '2 milestones are overdue this week.',
      recipients: ['PROJECT_MANAGER', 'MANAGER', 'DIRECTOR'],
    },
    {
      id: uid('ntf'),
      createdAt: todayISO(),
      severity: 'RED',
      message: 'Battery Pack New Program – Gen2: Prototype Build is 5 days overdue.',
      recipients: ['PROJECT_MANAGER', 'MANAGER', 'DIRECTOR'],
    },
    {
      id: uid('ntf'),
      createdAt: addDays(todayISO(), -1),
      severity: 'YELLOW',
      message: '7 activities are due within the warning threshold.',
      recipients: ['PROJECT_MANAGER', 'MANAGER', 'DIRECTOR'],
    },
    {
      id: uid('ntf'),
      createdAt: addDays(todayISO(), -1),
      severity: 'YELLOW',
      message: 'Viktor Hahn is at 105% utilization this week (CW 16).',
      recipients: ['MANAGER', 'DIRECTOR'],
    },
    {
      id: uid('ntf'),
      createdAt: addDays(todayISO(), -2),
      severity: 'GREEN',
      message: 'Assembly Line CIP – TCA Station 12: CIP Improvement Gate completed on time.',
      recipients: ['PROJECT_MANAGER', 'MANAGER'],
    },
    {
      id: uid('ntf'),
      createdAt: addDays(todayISO(), -3),
      severity: 'RED',
      message: 'EV Drivetrain Path Forward: critical path changed — Tooling Procurement now blocking.',
      recipients: ['PROJECT_MANAGER', 'MANAGER', 'DIRECTOR'],
    },
    {
      id: uid('ntf'),
      createdAt: addDays(todayISO(), -4),
      severity: 'YELLOW',
      message: 'Annika Braun approaching 90% capacity next week (CW 17).',
      recipients: ['MANAGER'],
    },
    {
      id: uid('ntf'),
      createdAt: addDays(todayISO(), -5),
      severity: 'GREEN',
      message: 'Sparkplugs Manufacturing CIP: Supplier Qualification completed ahead of schedule.',
      recipients: ['PROJECT_MANAGER'],
    },
    {
      id: uid('ntf'),
      createdAt: addDays(todayISO(), -6),
      severity: 'YELLOW',
      message: 'Steering Column ECR – Phase 2: Technical Design Review due in 3 days.',
      recipients: ['PROJECT_MANAGER', 'MANAGER'],
    },
    {
      id: uid('ntf'),
      createdAt: addDays(todayISO(), -7),
      severity: 'RED',
      message: 'Fuel Injector ECR: Validation Testing blocked — awaiting supplier nozzle samples.',
      recipients: ['PROJECT_MANAGER', 'MANAGER', 'DIRECTOR'],
    },
  ]

  return {
    projects,
    milestones,
    activities,
    dependencies,
    assignments,
    notifications,
    people: INITIAL_PEOPLE,
    clients: INITIAL_CLIENTS,
    suppliers: INITIAL_SUPPLIERS,
    settings: {
      warningDaysThreshold: 7,
      upcomingWeeksWindow: 4,
      lastClosedWeeksWindow: 4,
      defaultWeeklyCapacity: 40,
      workloadWarningThreshold: 90,
      workloadCriticalThreshold: 100,
    },
  }
}

/**
 * Project status = worst task status under the completion-vs-buffer rule.
 * A task is rated by whether it's (projected to be) done a buffer ahead of due.
 */
function seededProjectStatus(state: AppState, projectId: string): RygStatus {
  // Exclude N/A tasks — they won't be done, so they don't affect project health.
  const projectActivities = state.activities.filter((a) => a.projectId === projectId && a.state !== 'NOT_APPLICABLE')
  if (projectActivities.length === 0) return 'GREEN'

  const statuses = projectActivities.map((activity) => rygForActivity(activity))
  if (statuses.includes('RED')) return 'RED'
  if (statuses.includes('YELLOW')) return 'YELLOW'
  return 'GREEN'
}

function projectNextMilestone(state: AppState, projectId: string): Milestone | undefined {
  return state.milestones
    .filter((ms) => ms.projectId === projectId && !ms.doneDate)
    .sort((a, b) => (a.targetDate < b.targetDate ? -1 : 1))[0]
}

function getScopeProjectIds(role: Role, userId: string, projects: Project[], peopleList: Person[]): Set<string> {
  if (role === 'DIRECTOR') return new Set(projects.map((p) => p.id))
  if (role === 'PROJECT_MANAGER') {
    return new Set(projects.filter((p) => p.projectManagerId === userId).map((p) => p.id))
  }
  const teamPms = peopleList.filter((person) => person.managerId === userId && person.role === 'PROJECT_MANAGER').map((person) => person.id)
  return new Set(projects.filter((project) => teamPms.includes(project.projectManagerId)).map((p) => p.id))
}

function getVisibleResourceIds(role: Role, userId: string, peopleList: Person[]): Set<string> {
  if (role === 'DIRECTOR') return new Set(peopleList.filter((person) => person.role === 'RESOURCE').map((p) => p.id))
  if (role === 'MANAGER') {
    return new Set(peopleList.filter((person) => person.managerId === userId && person.role === 'RESOURCE').map((p) => p.id))
  }
  return new Set(peopleList.filter((person) => person.role === 'RESOURCE').map((p) => p.id))
}

function makeAutoMilestones(projectId: string, category: ProjectCategory): Milestone[] {
  return CATEGORY_MILESTONES[category].map((template) => ({
    id: uid('ms'),
    projectId,
    name: template.name,
    targetDate: addDays(todayISO(), template.offset),
  }))
}

function statusBadge(status: RygStatus): ReactElement {
  if (status === 'RED') {
    return (
      <span className="inline-flex items-center gap-1.5 rounded-lg bg-red-500/8 px-2.5 py-1 text-[11px] font-medium text-red-600 dark:bg-red-500/15 dark:text-red-400">
        <span className="h-1.5 w-1.5 rounded-full bg-red-500" />
        Critical
      </span>
    )
  }
  if (status === 'YELLOW') {
    return (
      <span className="inline-flex items-center gap-1.5 rounded-lg bg-amber-500/8 px-2.5 py-1 text-[11px] font-medium text-amber-600 dark:bg-amber-500/15 dark:text-amber-400">
        <span className="h-1.5 w-1.5 rounded-full bg-amber-500" />
        At Risk
      </span>
    )
  }
  return (
    <span className="inline-flex items-center gap-1.5 rounded-lg bg-emerald-500/8 px-2.5 py-1 text-[11px] font-medium text-emerald-600 dark:bg-emerald-500/15 dark:text-emerald-400">
      <span className="h-1.5 w-1.5 rounded-full bg-emerald-500" />
      On Track
    </span>
  )
}

function activityStateBadge(actState: ActivityState): ReactElement {
  if (actState === 'DONE')
    return <span className="inline-flex items-center gap-1.5 rounded-lg bg-emerald-500/8 px-2.5 py-1 text-[11px] font-medium text-emerald-600 dark:bg-emerald-500/15 dark:text-emerald-400">Complete</span>
  if (actState === 'IN_PROGRESS')
    return <span className="inline-flex items-center gap-1.5 rounded-lg bg-pth-blue/8 px-2.5 py-1 text-[11px] font-medium text-pth-blue dark:bg-pth-blue/15">In Progress</span>
  if (actState === 'NOT_APPLICABLE')
    return <span className="inline-flex items-center gap-1.5 rounded-lg bg-pth-muted/8 px-2.5 py-1 text-[11px] font-medium text-pth-muted/70 dark:bg-pth-muted/10">N/A</span>
  return <span className="inline-flex items-center gap-1.5 rounded-lg bg-pth-muted/8 px-2.5 py-1 text-[11px] font-medium text-pth-muted dark:bg-pth-muted/15">Open</span>
}

function categoryBadge(cat: ProjectCategory): ReactElement {
  const colours: Record<ProjectCategory, string> = {
    ECR: 'bg-amber-500/8 text-amber-600 dark:bg-amber-500/15 dark:text-amber-400',
    CIP: 'bg-red-500/8 text-red-600 dark:bg-red-500/15 dark:text-red-400',
    'Path Forward': 'bg-emerald-500/8 text-emerald-600 dark:bg-emerald-500/15 dark:text-emerald-400',
    'New Programs': 'bg-pth-blue/8 text-pth-blue dark:bg-pth-blue/15',
  }
  return <span className={`inline-flex rounded-lg px-2.5 py-1 text-[11px] font-semibold uppercase tracking-wider ${colours[cat]}`}>{cat}</span>
}

export default function App(): ReactElement {
  const [state, setState] = useState<AppState>(() => generateInitialState())
  const [currentUserId, setCurrentUserId] = useState('u_dir_1')
  const [paContext, setPaContext] = useState<IContext | null>(null)
  const [userPhotoUrl, setUserPhotoUrl] = useState<string | null>(null)
  const [userJobTitle, setUserJobTitle] = useState<string | null>(null)
  const [dvLoading, setDvLoading] = useState(false)
  const [dvConnected, setDvConnected] = useState(false)
  const [dvUsers, setDvUsers] = useState<DataverseUser[]>([])
  const [navPage, setNavPage] = useState<NavPage>('overview')
  const [reportTab, setReportTab] = useState<ReportTab>('projectsStatus')
  const [selectedProjectId, setSelectedProjectId] = useState<string | null>(state.projects[0]?.id ?? null)
  const [sidebarOpen, setSidebarOpen] = useState(false)
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false)
  const [filterCategory, setFilterCategory] = useState<string>('all')
  const [filterStatus, setFilterStatus] = useState<string>('all')
  const [filterManager, setFilterManager] = useState<string>('all')
  const [detailProjectId, setDetailProjectId] = useState<string | null>(null)
  const [notificationsOpen, setNotificationsOpen] = useState(false)
  // Month + calendar week horizon filter
  const currentYM = todayISO().slice(0, 7) // e.g. '2026-02'
  const [filterMonth, setFilterMonth] = useState(currentYM)
  const [filterCW, setFilterCW] = useState<number | 'all'>('all')

  // Derive horizon start & end from month + CW selection
  const horizonRange = useMemo(() => {
    const [y, m] = filterMonth.split('-').map(Number)
    if (filterCW === 'all') {
      const start = `${filterMonth}-01`
      const last = new Date(y, m, 0) // last day of month
      const end = last.toISOString().slice(0, 10)
      return { start, end }
    }
    const start = mondayOfCW(y, filterCW)
    const end = addDays(start, 6) // Sunday
    return { start, end }
  }, [filterMonth, filterCW])

  // Month options: 1 month back + 5 months forward
  const monthOptions = useMemo(() => {
    const opts: Array<{ value: string; label: string }> = []
    const base = new Date(todayISO() + 'T00:00:00')
    for (let i = -1; i <= 5; i++) {
      const d = new Date(base.getFullYear(), base.getMonth() + i, 1)
      const val = d.toISOString().slice(0, 7)
      const lbl = new Intl.DateTimeFormat('en-US', { month: 'long', year: 'numeric' }).format(d)
      opts.push({ value: val, label: lbl })
    }
    return opts
  }, [])

  // Calendar weeks for the selected month
  const cwOptions = useMemo(() => cwsInMonth(filterMonth), [filterMonth])
  const [projectSearch, setProjectSearch] = useState('')
  const [workloadWeekOffset, setWorkloadWeekOffset] = useState(0)
  const [workloadSearch, setWorkloadSearch] = useState('')
  const [workloadView, setWorkloadView] = useState<'cards' | 'list'>('cards')
  const [workloadSiteFilter, setWorkloadSiteFilter] = useState<'all' | 'SlpP' | 'TlP'>('all')
  const [workloadAreaFilter, setWorkloadAreaFilter] = useState<string>('all')
  const [expandedRiskProjectId, setExpandedRiskProjectId] = useState<string | null>(null)
  const [milestoneDialogOpen, setMilestoneDialogOpen] = useState(false)
  const [milestoneEditId, setMilestoneEditId] = useState<string | null>(null)
  const [editingProject, setEditingProject] = useState(false)
  const [creatingSaving, setCreatingSaving] = useState(false)
  const [inlineMilestones, setInlineMilestones] = useState<InlineMilestone[]>([])
  const [viewMode, setViewMode] = useState<'gallery' | 'list'>('gallery')
  const [activityDialogOpen, setActivityDialogOpen] = useState(false)
  const [activityMilestoneId, setActivityMilestoneId] = useState<string | null>(null)
  const [activityEditMode, setActivityEditMode] = useState<ActivityEditMode>('create')
  const [editingActivityId, setEditingActivityId] = useState<string | null>(null)
  const [expandedWorkloadPersonId, setExpandedWorkloadPersonId] = useState<string | null>(null)
  const [masterDataTab, setMasterDataTab] = useState<MasterDataTab>('people')
  const [teamSiteFilter, setTeamSiteFilter] = useState<'all' | 'SlpP' | 'TlP'>('all')
  const [teamAreaFilter, setTeamAreaFilter] = useState<string>('all')
  // ── Task template editor ──
  const [taskTemplates, setTaskTemplates] = useState<TaskTemplateRow[]>([])
  const [ttType, setTtType] = useState<string>('')
  const [ttDialogOpen, setTtDialogOpen] = useState(false)
  const [ttEditId, setTtEditId] = useState<string | null>(null)
  const [ttForm, setTtForm] = useState<{ task: string; responsible: string; sequence: number; leadtimeWeeks: string; workloadPct: string; inputs: string }>({ task: '', responsible: '', sequence: 0, leadtimeWeeks: '', workloadPct: '', inputs: '' })
  const [reportProjectId, setReportProjectId] = useState<string | null>(null)
  const [reportActivityTab, setReportActivityTab] = useState<ReportActivityTab>('upcoming')
  const [rptSearch, setRptSearch] = useState('')
  const [rptCategory, setRptCategory] = useState<string>('all')
  const [rptStatus, setRptStatus] = useState<string>('all')
  const [rptManager, setRptManager] = useState<string>('all')
  const [rptSite, setRptSite] = useState<string>('all')
  // ── Task Tracking (preventive) filters ──
  const [ttArea, setTtArea] = useState<string>('all')
  const [ttLoc, setTtLoc] = useState<string>('all')
  const [ttProjType, setTtProjType] = useState<string>('all')
  const [ttRisk, setTtRisk] = useState<string>('atrisk')
  const [ttView, setTtView] = useState<'tasks' | 'area' | 'project' | 'owner'>('tasks')
  const [ttOwner, setTtOwner] = useState<string>('all')
  const [ganttLabelWidth, setGanttLabelWidth] = useState(176) // px, user-resizable
  const [taskNaFilter, setTaskNaFilter] = useState<'active' | 'na' | 'all'>('active')
  const [resourcePaneSearch, setResourcePaneSearch] = useState('')
  const [resourcePaneArea, setResourcePaneArea] = useState<string>('all')
  const [ttExpanded, setTtExpanded] = useState<Set<string>>(new Set())
  const [dragPersonId, setDragPersonId] = useState<string | null>(null)
  const [dropTargetId, setDropTargetId] = useState<string | null>(null)
  const [mdDialogOpen, setMdDialogOpen] = useState(false)
  const [mdEditId, setMdEditId] = useState<string | null>(null)
  const [mdForm, setMdForm] = useState<{ name: string; initials: string; role: Person['role']; managerId: string; area: string; location: string[]; kind: string; capacity: number }>({ name: '', initials: '', role: 'RESOURCE', managerId: '', area: '', location: [], kind: 'Associate', capacity: 40 })
  const [aadResults, setAadResults] = useState<AadUser[]>([])
  const [aadLoading, setAadLoading] = useState(false)
  const [aadQuery, setAadQuery] = useState('')
  const [deleteProjectDialogOpen, setDeleteProjectDialogOpen] = useState(false)
  const [deletingProject, setDeletingProject] = useState(false)

  /* ── Power Automate state ── */
  const [paFlows, setPaFlows] = useState<PowerAutomateFlow[]>([
    { id: 'flow-1', name: 'Overdue Milestones Alert', description: 'Sends a Teams notification when milestones are past due', triggerType: 'overdue_milestones', enabled: true },
    { id: 'flow-2', name: 'Delayed Activities Digest', description: 'Emails project managers a daily summary of delayed activities', triggerType: 'delayed_activities', enabled: true },
    { id: 'flow-3', name: 'Workload Threshold Breach', description: 'Posts to Teams channel when a resource exceeds capacity', triggerType: 'workload_threshold', enabled: false },
    { id: 'flow-4', name: 'Project Created Notification', description: 'Sends an email to the creator with all project details when a new project is created', triggerType: 'project_created', enabled: true },
  ])
  const [paWebhookUrl, setPaWebhookUrl] = useState('https://prod-00.westus.logic.azure.com:443/workflows/94283d50-c20e...')
  const [paFlowLog, setPaFlowLog] = useState<FlowLogEntry[]>([
    { id: 'flog-1', flowName: 'Overdue Milestones Alert', triggeredAt: addDays(todayISO(), -1) + 'T08:00:00Z', status: 'success', payload: '{"milestones":2,"recipients":["u_pm_1","u_pm_3"],"channel":"Teams"}' },
    { id: 'flog-2', flowName: 'Delayed Activities Digest', triggeredAt: addDays(todayISO(), -1) + 'T07:30:00Z', status: 'success', payload: '{"activities":5,"digest":"daily","sent_to":"pm-group@p3-group.com"}' },
    { id: 'flog-3', flowName: 'Workload Threshold Breach', triggeredAt: addDays(todayISO(), -2) + 'T14:15:00Z', status: 'failed', payload: '{"error":"Webhook URL unreachable — 503 Service Unavailable"}' },
    { id: 'flog-4', flowName: 'Overdue Milestones Alert', triggeredAt: addDays(todayISO(), -3) + 'T08:00:00Z', status: 'success', payload: '{"milestones":1,"recipients":["u_pm_2"],"channel":"Teams"}' },
    { id: 'flog-5', flowName: 'Project Created Notification', triggeredAt: addDays(todayISO(), -5) + 'T11:22:00Z', status: 'success', payload: '{"project":"Paint Shop CIP – VOC Reduction","creator":"u_pm_1","email":"rafael.gomez@p3-group.com"}' },
    { id: 'flog-6', flowName: 'Delayed Activities Digest', triggeredAt: addDays(todayISO(), -6) + 'T07:30:00Z', status: 'success', payload: '{"activities":3,"digest":"daily","sent_to":"pm-group@p3-group.com"}' },
    { id: 'flog-7', flowName: 'Workload Threshold Breach', triggeredAt: addDays(todayISO(), -7) + 'T10:05:00Z', status: 'success', payload: '{"resource":"Viktor Hahn","utilization":"108%","week":"CW 15","channel":"Teams"}' },
  ])
  const [paTestingFlowId, setPaTestingFlowId] = useState<string | null>(null)

  const [darkMode, setDarkMode] = useState(() => {
    if (typeof window !== 'undefined') {
      return window.matchMedia('(prefers-color-scheme: dark)').matches
    }
    return false
  })

  useEffect(() => {
    document.documentElement.classList.toggle('dark', darkMode)
  }, [darkMode])

  /* ── Dataverse: load real data on mount when configured ── */
  const loadFromDataverse = useCallback(async () => {
    console.log('[APP] loadFromDataverse called, isDataverseConfigured:', isDataverseConfigured())
    if (!isDataverseConfigured()) return
    setDvLoading(true)
    try {
      // Fetch Dataverse users in parallel with main data.
      // Isolate each call so a failure (e.g. systemusers 403) can't reject the
      // whole load and drop the projects/activities.
      const [dvRes, usersRes] = await Promise.allSettled([
        fetchAllFromDataverse(),
        fetchDataverseUsers(),
      ])
      const dv = dvRes.status === 'fulfilled' ? dvRes.value : null
      const users = usersRes.status === 'fulfilled' ? usersRes.value : []
      console.log('[APP] fetchAllFromDataverse returned:', dv
        ? `projects=${dv.projects.length}, milestones=${dv.milestones.length}, activities=${dv.activities.length}, resources=${dv.resources.length}, assignments=${dv.assignments.length}`
        : 'NULL')
      console.log('[APP] Dataverse users fetched:', users.length)
      if (users.length > 0) setDvUsers(users)
      if (!dv) { console.warn('[APP] dv is null — keeping demo data'); return }

      // Build a name→user lookup for resolving sponsor/PM from stored text names
      const userByName = new Map(users.map((u) => [u.fullname.toLowerCase(), u]))

      // Even if projects is empty, still apply resources/settings from Dataverse
      if (dv.projects.length === 0) {
        console.warn('[APP] 0 projects from Dataverse — applying resources & settings only')
      }

      // Build a resource lookup by Dataverse id → person-compatible record
      const resourceById = new Map(dv.resources.map((r) => [r.id, r]))
      const resourceByName = new Map(dv.resources.map((r) => [r.name.toLowerCase(), r]))

      // Map resources into the Person[] shape the app expects
      const dvPeople: Person[] = [
        // Keep the existing non-RESOURCE people (director, managers, PMs)
        ...INITIAL_PEOPLE.filter((p) => p.role !== 'RESOURCE'),
        // Add Dataverse resources as RESOURCE persons
        ...dv.resources.map((r) => ({
          id: r.id,
          name: r.name,
          initials: r.initials,
          role: 'RESOURCE' as const,
          managerId: r.managerId,
          area: r.area,
          location: r.location,
          resourceKind: r.resourceKind,
          weeklyCapacity: r.weeklyCapacity,
        })),
      ]

      // Resolve activity ownerIds — try matching by name against resources
      const dvActivities: Activity[] = dv.activities.map((a) => {
        const ownerName = (a as any)._ownerName as string | undefined
        const matched = ownerName ? resourceByName.get(ownerName.toLowerCase()) : undefined
        return {
          id: a.id,
          projectId: a.projectId,
          milestoneId: a.milestoneId,
          name: a.name,
          ownerId: matched?.id ?? '',
          startDate: a.startDate,
          endDate: a.endDate,
          doneDate: a.doneDate,
          state: a.state,
          criticalPath: a.criticalPath,
          responsible: (a as { responsible?: string }).responsible ?? '',
          workloadPct: (a as { workloadPct?: number | null }).workloadPct ?? null,
        }
      })

      // Map assignments — resolve personId from _pth_resource_value (already a Dataverse GUID)
      const dvAssignments: Assignment[] = dv.assignments.map((a) => {
        const res = resourceById.get(a.personId)
        return {
          id: a.id,
          personId: a.personId,
          activityId: a.activityId,
          weekStart: a.weekStart,
          assignedHours: a.assignedHours,
          capacityHours: res?.weeklyCapacity ?? 40,
        }
      })

      // Build dependencies (not in Dataverse — generate from activity ordering)
      const projectIdMap = new Map<string, typeof dvActivities>()
      for (const act of dvActivities) {
        const list = projectIdMap.get(act.projectId) ?? []
        list.push(act)
        projectIdMap.set(act.projectId, list)
      }
      const dvDependencies: Dependency[] = []
      for (const [projId, acts] of projectIdMap) {
        const sorted = [...acts].sort((a, b) => a.startDate.localeCompare(b.startDate))
        for (let i = 1; i < sorted.length; i++) {
          dvDependencies.push({
            id: uid('dep'),
            projectId: projId,
            activityId: sorted[i].id,
            predecessorActivityId: sorted[i - 1].id,
          })
        }
      }

      console.log('[APP] Setting state with Dataverse data:',
        'projects:', dv.projects.length,
        'milestones:', dv.milestones.length,
        'activities:', dvActivities.length,
        'assignments:', dvAssignments.length,
        'people:', dvPeople.length,
        'deps:', dvDependencies.length)

      setState((prev) => ({
        ...prev,
        projects: dv.projects.map((p) => {
          // Resolve sponsor/PM names to Dataverse user IDs
          const sponsorMatch = p._sponsorName ? userByName.get(p._sponsorName.toLowerCase()) : undefined
          const pmMatch = p._pmName ? userByName.get(p._pmName.toLowerCase()) : undefined
          return {
          id: p.id,
          projectCode: p.projectIdExternal || nextProjectCode(p.siteId ? [p.siteId] : []),
          name: p.name,
          clientCode: p.clientCode || '',
          category: p.category,
          budgetAllocated: p.budgetAllocated || 0,
          siteIds: p.siteId ? [p.siteId] : [],
          sponsorExecutiveId: sponsorMatch?.id ?? p.sponsorExecutiveId,
          projectManagerId: pmMatch?.id ?? p.projectManagerId,
          clientIds: p.clientIds,
          supplierIds: p.supplierIds,
          objective: p.objective,
          plannedStartDate: p.plannedStartDate || todayISO(),
          plannedEndDate: p.plannedEndDate || addDays(todayISO(), 30),
          projectType: (p as { projectType?: string }).projectType ?? '',
        }}),
        milestones: dv.milestones.map((m) => ({
          id: m.id,
          projectId: m.projectId,
          name: m.name,
          targetDate: m.targetDate,
          doneDate: m.doneDate,
        })),
        activities: dvActivities,
        dependencies: dvDependencies,
        assignments: dvAssignments,
        people: dvPeople,
        settings: dv.settings,
      }))
      setDvConnected(true)
    } catch (err) {
      console.error('Dataverse load failed, using demo data:', err)
    } finally {
      setDvLoading(false)
    }
  }, [])

  useEffect(() => { loadFromDataverse() }, [loadFromDataverse])

  // Load editable task templates from Dataverse
  const reloadTaskTemplates = useCallback(async () => {
    try {
      const rows = await fetchTaskTemplates()
      setTaskTemplates(rows)
      setTtType((cur) => cur || rows[0]?.projectType || '')
    } catch (err) { console.warn('[APP] task templates load failed:', err) }
  }, [])
  useEffect(() => { reloadTaskTemplates() }, [reloadTaskTemplates])

  /* ── Power Apps context: retrieve signed-in user ── */
  useEffect(() => {
    let cancelled = false
    getContext()
      .then(async (ctx) => {
        if (cancelled) return
        setPaContext(ctx)
        const fullName = ctx.user.fullName ?? ''
        if (fullName) {
          const initials = fullName.split(' ').map((w) => w[0]).join('').toUpperCase().slice(0, 2)
          setState((prev) => {
            // Try to match the signed-in user to an existing person by name
            const matched = prev.people.find((p) => p.name.toLowerCase() === fullName.toLowerCase())
            if (matched) {
              setCurrentUserId(matched.id)
              return {
                ...prev,
                people: prev.people.map((p) =>
                  p.id === matched.id ? { ...p, name: fullName, initials } : p,
                ),
              }
            }
            // No match — update the default director entry with the real user info
            setCurrentUserId('u_dir_1')
            return {
              ...prev,
              people: prev.people.map((p) =>
                p.id === 'u_dir_1' ? { ...p, name: fullName, initials } : p,
              ),
            }
          })
        }

        // Fetch profile photo + job title from Microsoft Graph with proper auth
        if (!cancelled) {
          try {
            const profile = await fetchCurrentUserProfile()
            if (!cancelled) {
              if (profile.photoUrl) setUserPhotoUrl(profile.photoUrl)
              if (profile.jobTitle) setUserJobTitle(profile.jobTitle)
            }
          } catch {
            // Graph profile not available — keep defaults
          }
        }
      })
      .catch(() => {
        // Running outside Power Apps host — keep demo defaults
      })
    return () => { cancelled = true }
  }, [])

  /* ── AAD people search debounce ── */
  useEffect(() => {
    if (!mdDialogOpen || mdEditId) { setAadResults([]); return }
    if (aadQuery.length < 2) { setAadResults([]); return }
    let cancelled = false
    setAadLoading(true)
    const timeout = setTimeout(async () => {
      const results = await searchAadUsers(aadQuery)
      if (!cancelled) { setAadResults(results); setAadLoading(false) }
    }, 350)
    return () => { cancelled = true; clearTimeout(timeout) }
  }, [aadQuery, mdDialogOpen, mdEditId])

  const people = state.people
  const clients = state.clients
  const suppliers = state.suppliers

  const currentUser = useMemo(() => people.find((person) => person.id === currentUserId) ?? people[0], [currentUserId, people])
  const role: Role = currentUser.role === 'RESOURCE' ? 'PROJECT_MANAGER' : (currentUser.role as Role)

  // Prefer Power Apps context user identity; fall back to person record
  const displayName = paContext?.user.fullName || currentUser.name
  const displayRole = userJobTitle || ROLE_LABEL[role] || 'Team Member'
  const displayInitials = (paContext?.user.fullName ?? '').trim()
    ? paContext!.user.fullName!.split(' ').map((w) => w[0]).join('').toUpperCase().slice(0, 2)
    : currentUser.initials

  const scopedProjectIds = useMemo(() => getScopeProjectIds(role, currentUserId, state.projects, people), [role, currentUserId, state.projects, people])
  const scopedProjects = useMemo(() => state.projects.filter((project) => scopedProjectIds.has(project.id)), [state.projects, scopedProjectIds])

  useEffect(() => {
    if (!selectedProjectId && scopedProjects.length > 0) {
      setSelectedProjectId(scopedProjects[0].id)
      return
    }
    if (selectedProjectId && !scopedProjects.some((project) => project.id === selectedProjectId)) {
      setSelectedProjectId(scopedProjects[0]?.id ?? null)
    }
  }, [selectedProjectId, scopedProjects])

  const selectedProject = useMemo(() => state.projects.find((project) => project.id === selectedProjectId) ?? null, [state.projects, selectedProjectId])
  const selectedProjectMilestones = useMemo(() => {
    if (!selectedProject) return []
    return state.milestones.filter((ms) => ms.projectId === selectedProject.id).sort((a, b) => (a.targetDate < b.targetDate ? -1 : 1))
  }, [selectedProject, state.milestones])

  const selectedProjectActivities = useMemo(() => {
    if (!selectedProject) return []
    return state.activities.filter((activity) => activity.projectId === selectedProject.id).sort((a, b) => (a.startDate < b.startDate ? -1 : 1))
  }, [selectedProject, state.activities])


  const visibleNotifications = useMemo(() => {
    return state.notifications
      .filter((notification) => notification.recipients.includes(role))
      .sort((a, b) => (a.createdAt < b.createdAt ? 1 : -1))
  }, [state.notifications, role])

  const filteredProjects = useMemo(() => {
    let result = scopedProjects

    // Text search
    const q = projectSearch.trim().toLowerCase()
    if (q) {
      const activityProjectIds = new Set(
        state.activities
          .filter((activity) => activity.name.toLowerCase().includes(q))
          .map((activity) => activity.projectId),
      )
      result = result.filter((project) => {
        if (project.name.toLowerCase().includes(q)) return true
        if (project.objective.toLowerCase().includes(q)) return true
        return activityProjectIds.has(project.id)
      })
    }

    // Category filter
    if (filterCategory !== 'all') {
      result = result.filter((p) => p.category === filterCategory)
    }

    // Status filter
    if (filterStatus !== 'all') {
      result = result.filter((p) => seededProjectStatus(state, p.id) === filterStatus)
    }

    // Manager filter
    if (filterManager !== 'all') {
      result = result.filter((p) => p.projectManagerId === filterManager)
    }

    // Month / CW horizon filter — keep projects whose activity date range overlaps the selected window
    result = result.filter((p) => {
      const acts = state.activities.filter((a) => a.projectId === p.id)
      if (acts.length === 0) return true // show projects without activities so they remain visible
      return acts.some((a) => a.startDate <= horizonRange.end && a.endDate >= horizonRange.start)
    })

    return result
  }, [projectSearch, scopedProjects, state, filterCategory, filterStatus, filterManager, horizonRange])

  const projectRows = useMemo(() => {
    return filteredProjects.map((project) => {
      const pm = dvUsers.find((u) => u.id === project.projectManagerId)?.fullname ?? people.find((person) => person.id === project.projectManagerId)?.name ?? 'Unknown'
      const site = project.siteIds.map((id) => sites.find((s) => s.id === id)?.name ?? id).join(', ') || 'Unknown'
      const status = seededProjectStatus(state, project.id)
      const nextMilestone = projectNextMilestone(state, project.id)
      const delayedCount = state.activities.filter(
        (activity) => activity.projectId === project.id && !activity.doneDate && rygFromDueDate(activity.endDate, state.settings.warningDaysThreshold) === 'RED',
      ).length
      const pActs = state.activities.filter((a) => a.projectId === project.id)
      const dueDate = pActs.length > 0 ? pActs.map((a) => a.endDate).sort().slice(-1)[0] : null
      const progressPct = pActs.length > 0 ? Math.round((pActs.filter((a) => a.state === 'DONE').length / pActs.length) * 100) : 0
      return { project, pm, site, status, nextMilestone, delayedCount, dueDate, progressPct }
    })
  }, [filteredProjects, state])

  const kpiCounts = useMemo(() => {
    const total = scopedProjects.length
    const statuses = scopedProjects.map((p) => seededProjectStatus(state, p.id))
    const onTrack = statuses.filter((s) => s === 'GREEN').length
    const atRisk = statuses.filter((s) => s === 'YELLOW').length
    const critical = statuses.filter((s) => s === 'RED').length
    const completed = scopedProjects.filter((p) => state.activities.filter((a) => a.projectId === p.id).every((a) => a.state === 'DONE')).length
    return { total, onTrack, atRisk, critical, completed }
  }, [scopedProjects, state])

  /* ── Report-tab filtered projects (independent from dashboard filters) ── */
  const rptFiltered = useMemo(() => {
    let result = scopedProjects
    const q = rptSearch.toLowerCase().trim()
    if (q) {
      result = result.filter((p) => p.name.toLowerCase().includes(q) || p.objective.toLowerCase().includes(q) || (p.projectCode ?? '').toLowerCase().includes(q))
    }
    if (rptCategory !== 'all') result = result.filter((p) => p.category === rptCategory)
    if (rptStatus !== 'all') result = result.filter((p) => seededProjectStatus(state, p.id) === rptStatus)
    if (rptManager !== 'all') result = result.filter((p) => p.projectManagerId === rptManager)
    if (rptSite !== 'all') result = result.filter((p) => p.siteIds.includes(rptSite))
    return result
  }, [scopedProjects, state, rptSearch, rptCategory, rptStatus, rptManager, rptSite])

  const rptRows = useMemo(() => {
    return rptFiltered.map((project) => {
      const pm = dvUsers.find((u) => u.id === project.projectManagerId)?.fullname ?? people.find((person) => person.id === project.projectManagerId)?.name ?? 'Unknown'
      const site = project.siteIds.map((id) => sites.find((s) => s.id === id)?.name ?? id).join(', ') || 'Unknown'
      const status = seededProjectStatus(state, project.id)
      const nextMilestone = projectNextMilestone(state, project.id)
      const delayedCount = state.activities.filter((a) => a.projectId === project.id && !a.doneDate && rygFromDueDate(a.endDate, state.settings.warningDaysThreshold) === 'RED').length
      const pActs = state.activities.filter((a) => a.projectId === project.id)
      // Next task: earliest-due, not-done, dated, in-scope (skip N/A) task.
      const nextTask = pActs
        .filter((a) => a.state !== 'NOT_APPLICABLE' && a.state !== 'DONE' && a.endDate)
        .sort((a, b) => (a.endDate < b.endDate ? -1 : 1))[0]
      const progressPct = pActs.length > 0 ? Math.round((pActs.filter((a) => a.state === 'DONE').length / pActs.length) * 100) : 0
      return { project, pm, site, status, nextMilestone, nextTask, delayedCount, progressPct }
    })
  }, [rptFiltered, state])

  const rptKpi = useMemo(() => {
    const total = rptFiltered.length
    const statuses = rptFiltered.map((p) => seededProjectStatus(state, p.id))
    const onTrack = statuses.filter((s) => s === 'GREEN').length
    const atRisk = statuses.filter((s) => s === 'YELLOW').length
    const critical = statuses.filter((s) => s === 'RED').length
    return { total, onTrack, atRisk, critical }
  }, [rptFiltered, state])

  /* ── Report-specific computed data (scoped to reportProjectId) ── */
  const reportProject = useMemo(() => (reportProjectId ? state.projects.find((p) => p.id === reportProjectId) ?? null : null), [reportProjectId, state.projects])

  const reportMilestones = useMemo(() => {
    if (!reportProject) return []
    return state.milestones.filter((ms) => ms.projectId === reportProject.id).sort((a, b) => (a.targetDate < b.targetDate ? -1 : 1))
  }, [reportProject, state.milestones])

  const reportActivities = useMemo(() => {
    if (!reportProject) return []
    return state.activities.filter((a) => a.projectId === reportProject.id).sort((a, b) => (a.startDate < b.startDate ? -1 : 1))
  }, [reportProject, state.activities])

  const reportNextMilestone = useMemo(() => {
    if (!reportProject) return undefined
    return projectNextMilestone(state, reportProject.id)
  }, [reportProject, state])

  const reportUpcoming = useMemo(() => {
    if (!reportProject) return []
    return state.activities.filter((a) => a.projectId === reportProject.id && !a.doneDate && a.startDate <= horizonRange.end && a.endDate >= horizonRange.start)
  }, [reportProject, horizonRange, state.activities])

  const reportClosed = useMemo(() => {
    if (!reportProject) return []
    const since = addDays(todayISO(), -state.settings.lastClosedWeeksWindow * 7)
    return state.activities.filter((a) => a.projectId === reportProject.id && !!a.doneDate && (a.doneDate ?? '') >= since)
  }, [reportProject, state.activities, state.settings.lastClosedWeeksWindow])

  const reportDelayed = useMemo(() => {
    if (!reportProject) return []
    return state.activities
      .filter((a) => a.projectId === reportProject.id && !a.doneDate && rygFromDueDate(a.endDate, state.settings.warningDaysThreshold) === 'RED')
      .map((a) => ({ ...a, delayDays: Math.abs(dayDiff(todayISO(), a.endDate)) }))
  }, [reportProject, state.activities, state.settings.warningDaysThreshold])

  const upcomingMilestonesRisk = useMemo(() => {
    return state.milestones
      .filter((milestone) => milestone.targetDate >= horizonRange.start && milestone.targetDate <= horizonRange.end)
      .map((milestone) => {
        const project = state.projects.find((p) => p.id === milestone.projectId)
        const status = rygFromDueDate(milestone.targetDate, state.settings.warningDaysThreshold)
        return {
          projectName: project?.name ?? 'Unknown',
          milestone,
          status,
          daysRemaining: dayDiff(todayISO(), milestone.targetDate),
        }
      })
      .sort((a, b) => {
        if (rankStatus(a.status) !== rankStatus(b.status)) return rankStatus(a.status) - rankStatus(b.status)
        return a.milestone.targetDate < b.milestone.targetDate ? -1 : 1
      })
  }, [horizonRange, state.milestones, state.projects, state.settings.warningDaysThreshold])

  const delayedGrouped = useMemo(() => {
    return state.projects
      .map((project) => {
        const items = state.activities.filter(
          (activity) =>
            activity.projectId === project.id &&
            !activity.doneDate &&
            rygFromDueDate(activity.endDate, state.settings.warningDaysThreshold) === 'RED',
        )
        return {
          project,
          items,
          criticalImpact: items.some((item) => item.criticalPath),
        }
      })
      .filter((group) => group.items.length > 0)
  }, [state.activities, state.projects, state.settings.warningDaysThreshold])

  const currentWeek = useMemo(() => addDays(weekStartISO(todayISO()), workloadWeekOffset * 7), [workloadWeekOffset])

  const workloadRows = useMemo(() => {
    const visibleResources = getVisibleResourceIds(role, currentUserId, people)
    return people
      .filter((person) => person.role === 'RESOURCE' && visibleResources.has(person.id))
      .map((person) => {
        const personAssignments = state.assignments.filter((assignment) => assignment.personId === person.id && assignment.weekStart === currentWeek)
        const capacityHours = personAssignments[0]?.capacityHours ?? person.weeklyCapacity ?? state.settings.defaultWeeklyCapacity
        const assignedHoursFromAssignments = personAssignments.reduce((sum, row) => sum + row.assignedHours, 0)

        // ── Load from owned tasks active this week ──
        // A task the person owns contributes its workloadPct (as % of weekly
        // capacity) for every week between its start and end. Skip N/A, Done,
        // and dateless tasks. This is what makes drag-drop AND combo-box
        // assignment update the person's utilization immediately.
        const weekEnd = addDays(currentWeek, 6)
        const ownedTasks = state.activities.filter((a) =>
          a.ownerId === person.id &&
          a.state !== 'NOT_APPLICABLE' && a.state !== 'DONE' &&
          a.startDate && a.endDate &&
          a.startDate <= weekEnd && a.endDate >= currentWeek,
        )
        const ownedLoadPct = ownedTasks.reduce((sum, a) => sum + (a.workloadPct ?? 0), 0)
        const ownedHours = (ownedLoadPct / 100) * capacityHours

        const assignedHours = assignedHoursFromAssignments + ownedHours
        const utilization = Math.round((assignedHours / Math.max(capacityHours, 1)) * 100)
        const relatedProjects = Array.from(
          new Set(
            [
              ...personAssignments.map((assignment) => state.activities.find((activity) => activity.id === assignment.activityId)?.projectId),
              ...ownedTasks.map((a) => a.projectId),
            ]
              .filter((value): value is string => !!value)
              .map((projectId) => state.projects.find((project) => project.id === projectId)?.name)
              .filter((name): name is string => !!name),
          ),
        )

        return {
          person,
          personAssignments,
          assignedHours,
          capacityHours,
          utilization,
          relatedProjects,
        }
      })
  }, [currentUserId, currentWeek, role, people, state.activities, state.assignments, state.projects, state.settings.defaultWeeklyCapacity])

  // Peak weekly utilization per resource across ALL their owned active tasks.
  // hours/week for a task = workloadPct% × weekly capacity, applied to every
  // week the task spans. The "peak" week is the busiest → the overbooking
  // signal shown in the project-detail Available Resources sidebar (which has
  // no week selector, unlike the Workload page).
  const peakUtilByPerson = useMemo(() => {
    const out = new Map<string, { peakUtil: number; peakHours: number; capacity: number; taskCount: number }>()
    for (const person of people) {
      if (person.role !== 'RESOURCE') continue
      const capacity = person.weeklyCapacity ?? state.settings.defaultWeeklyCapacity
      const owned = state.activities.filter((a) =>
        a.ownerId === person.id && a.state !== 'NOT_APPLICABLE' && a.state !== 'DONE' &&
        a.startDate && a.endDate && a.startDate <= a.endDate,
      )
      const weekHours = new Map<string, number>()
      for (const a of owned) {
        const hours = ((a.workloadPct ?? 0) / 100) * capacity
        if (hours <= 0) continue
        let wk = weekStartISO(a.startDate)
        const lastWk = weekStartISO(a.endDate)
        let guard = 0
        while (wk <= lastWk && guard < 520) {
          weekHours.set(wk, (weekHours.get(wk) ?? 0) + hours)
          wk = addDays(wk, 7); guard++
        }
      }
      const peakHours = weekHours.size ? Math.max(...weekHours.values()) : 0
      out.set(person.id, {
        peakHours,
        capacity,
        peakUtil: Math.round((peakHours / Math.max(capacity, 1)) * 100),
        taskCount: owned.length,
      })
    }
    return out
  }, [people, state.activities, state.settings.defaultWeeklyCapacity])

  const overviewForm = useForm<ProjectOverviewForm>({
    values: {
      name: selectedProject?.name ?? '',
      clientCode: selectedProject?.clientCode ?? '',
      category: selectedProject?.category ?? 'CIP',
      budgetAllocated: selectedProject?.budgetAllocated ?? 0,
      siteIds: selectedProject?.siteIds ?? [],
      sponsorExecutiveId: selectedProject?.sponsorExecutiveId ?? 'u_dir_1',
      projectManagerId: selectedProject?.projectManagerId ?? 'u_pm_1',
      clientIds: selectedProject?.clientIds ?? [],
      supplierIds: selectedProject?.supplierIds ?? [],
      objective: selectedProject?.objective ?? '',
      plannedStartDate: selectedProject?.plannedStartDate ?? todayISO(),
      plannedEndDate: selectedProject?.plannedEndDate ?? addDays(todayISO(), 30),
    },
  })

  const milestoneForm = useForm<MilestoneForm>({
    defaultValues: {
      name: '',
      targetDate: addDays(todayISO(), 7),
    },
  })

  const activityForm = useForm<ActivityForm>({
    defaultValues: {
      name: '',
      ownerId: people.filter((p) => p.role === 'RESOURCE')[0]?.id ?? '',
      startDate: todayISO(),
      endDate: addDays(todayISO(), 7),
      workloadPct: 0,
      responsible: '',
      state: 'NOT_STARTED',
    },
  })

  useEffect(() => {
    if (!selectedProject) return
    overviewForm.reset({
      name: selectedProject.name,
      clientCode: selectedProject.clientCode,
      category: selectedProject.category,
      budgetAllocated: selectedProject.budgetAllocated,
      siteIds: selectedProject.siteIds,
      sponsorExecutiveId: selectedProject.sponsorExecutiveId,
      projectManagerId: selectedProject.projectManagerId,
      clientIds: selectedProject.clientIds,
      supplierIds: selectedProject.supplierIds,
      objective: selectedProject.objective,
      plannedStartDate: selectedProject.plannedStartDate,
      plannedEndDate: selectedProject.plannedEndDate,
    })
  }, [selectedProject, overviewForm])

  const newProjectForm = useForm<NewProjectForm>({
    defaultValues: {
      name: '',
      clientCode: '',
      category: 'CIP',
      projectType: '',
      budgetAllocated: 0,
      siteIds: [],
      sponsorExecutiveId: people.find((p) => p.role === 'DIRECTOR')?.id ?? '',
      projectManagerId: people.find((p) => p.role === 'PROJECT_MANAGER')?.id ?? '',
      clientIds: [],
      supplierIds: [],
      objective: '',
      plannedStartDate: todayISO(),
      plannedEndDate: addDays(todayISO(), 30),
    },
  })

  async function handleCreateProject(values: NewProjectForm): Promise<void> {
    setCreatingSaving(true)
    const siteNames = values.siteIds.map((id) => sites.find((s) => s.id === id)?.name ?? id).join(', ')
    const pmName = dvUsers.find((u) => u.id === values.projectManagerId)?.fullname
      ?? people.find((p) => p.id === values.projectManagerId)?.name ?? ''
    const sponsorName = dvUsers.find((u) => u.id === values.sponsorExecutiveId)?.fullname
      ?? people.find((p) => p.id === values.sponsorExecutiveId)?.name ?? ''
    const projectCode = nextProjectCode(values.siteIds)

    let dvId: string | null = null
    if (isDataverseConfigured()) {
      dvId = await createProjectInDataverse({
        name: values.name,
        clientCode: values.clientCode,
        category: values.category,
        budgetAllocated: values.budgetAllocated,
        site: values.siteIds[0] ?? 'site_tca',
        objective: values.objective,
        sponsorExecutive: sponsorName,
        projectManagerName: pmName,
        plannedStartDate: values.plannedStartDate,
        plannedEndDate: values.plannedEndDate,
        projectIdExternal: projectCode,
        projectType: values.projectType || undefined,
      })
    }

    const newProject: Project = {
      id: dvId ?? uid('proj'),
      projectCode,
      name: values.name,
      clientCode: values.clientCode,
      category: values.category,
      budgetAllocated: values.budgetAllocated,
      siteIds: values.siteIds,
      sponsorExecutiveId: values.sponsorExecutiveId,
      projectManagerId: values.projectManagerId,
      clientIds: values.clientIds,
      supplierIds: values.supplierIds,
      objective: values.objective,
      plannedStartDate: values.plannedStartDate,
      plannedEndDate: values.plannedEndDate,
    }

    const autoMs = makeAutoMilestones(newProject.id, values.category)

    // ── Generate standard tasks from the selected Project Type template ──
    // Tasks are dated by cumulative leadtime (weeks) from the planned start,
    // start UNASSIGNED, and carry responsible / workload / inputs. They are
    // created directly under the project (no milestone) per the flat-list model.
    const templateTasks = values.projectType
      ? taskTemplates.filter((t) => t.projectType === values.projectType).sort((a, b) => a.sequence - b.sequence)
      : []
    const genActivities: Activity[] = []
    {
      let cursor = values.plannedStartDate || todayISO()
      for (const t of templateTasks) {
        const weeks = t.leadtimeWeeks && t.leadtimeWeeks > 0 ? t.leadtimeWeeks : 1
        const start = cursor
        const end = addDays(start, weeks * 7)
        genActivities.push({
          id: uid('act'),
          projectId: newProject.id,
          name: t.task,
          ownerId: '',                       // UNASSIGNED — manager assigns later
          startDate: start,
          endDate: end,
          state: 'NOT_STARTED' as ActivityState,
          criticalPath: false,
          _responsible: t.responsible,
          _leadtimeWeeks: t.leadtimeWeeks,
          _workloadPct: t.workloadPct,
          _inputs: t.inputs,
        } as Activity & Record<string, unknown>)
        cursor = end
      }
    }

    // ── Build user-defined inline milestones + tasks ──
    const userMilestones: Milestone[] = inlineMilestones.map((im) => ({
      id: uid('ms'),
      projectId: newProject.id,
      name: im.name,
      targetDate: im.targetDate,
    }))
    const userActivities: Activity[] = inlineMilestones.flatMap((im, idx) =>
      im.tasks.map((t) => ({
        id: uid('act'),
        projectId: newProject.id,
        milestoneId: userMilestones[idx].id,
        name: t.name,
        ownerId: t.ownerId,
        startDate: t.startDate,
        endDate: t.endDate,
        state: 'NOT_STARTED' as ActivityState,
        criticalPath: false,
      })),
    )

    const allMilestones = [...autoMs, ...userMilestones]
    const allEmailMs = allMilestones.map((ms) => ({ name: ms.name, targetDate: ms.targetDate }))

    // Sync milestones to Dataverse (fire-and-forget, update IDs)
    if (isDataverseConfigured()) {
      // Standard template tasks — created directly under the project (no milestone)
      genActivities.forEach(async (act) => {
        const a = act as Activity & Record<string, unknown>
        const dvActId = await createActivityInDataverse({
          name: act.name, ownerName: 'Unassigned', startDate: act.startDate, endDate: act.endDate,
          state: 'NOT_STARTED', criticalPath: false, projectId: newProject.id,
          responsible: a._responsible as string | undefined,
          leadtimeWeeks: a._leadtimeWeeks as number | null | undefined,
          workloadPct: a._workloadPct as number | null | undefined,
          inputs: a._inputs as string | undefined,
        })
        if (dvActId) act.id = dvActId
      })

      Promise.all(
        allMilestones.map(async (ms) => {
          const dvMsId = await createMilestoneInDataverse({ name: ms.name, targetDate: ms.targetDate, projectId: newProject.id })
          if (dvMsId) ms.id = dvMsId
        }),
      ).then(() => {
        // After milestones are created, sync activities with correct milestone IDs
        userActivities.forEach(async (act) => {
          const ownerName = people.find((p) => p.id === act.ownerId)?.name ?? ''
          const dvActId = await createActivityInDataverse({
            name: act.name, ownerName, startDate: act.startDate, endDate: act.endDate,
            state: 'NOT_STARTED', criticalPath: false, projectId: newProject.id, milestoneId: act.milestoneId,
          })
          if (dvActId) act.id = dvActId
        })
      }).catch((err) => console.warn('Failed to sync milestones/activities to Dataverse:', err))
    }

    setState((prev) => ({
      ...prev,
      projects: [...prev.projects, newProject],
      milestones: [...prev.milestones, ...allMilestones],
      activities: [...prev.activities, ...userActivities, ...genActivities],
    }))

    // ── Send project creation email notification ──
    const projectCreatedFlow = paFlows.find((f) => f.triggerType === 'project_created' && f.enabled)
    if (projectCreatedFlow) {
      const creatorEmail = (paContext as any)?.user?.email
        ?? (paContext as any)?.user?.loginName
        ?? (paContext as any)?.user?.userName
        ?? ''
      const creatorName = paContext?.user.fullName ?? displayName

      sendProjectCreationEmail({
        webhookUrl: paWebhookUrl,
        recipientEmail: creatorEmail,
        recipientName: creatorName,
        projectName: values.name,
        projectCode,
        category: values.category,
        objective: values.objective,
        budgetAllocated: values.budgetAllocated,
        siteName: siteNames,
        sponsorExecutive: sponsorName,
        projectManager: pmName,
        plannedStartDate: values.plannedStartDate,
        plannedEndDate: values.plannedEndDate,
        clientCode: values.clientCode,
        milestones: allEmailMs,
      }).then((sent) => {
        const logEntry: FlowLogEntry = {
          id: uid('flog'),
          flowName: projectCreatedFlow.name,
          triggeredAt: new Date().toISOString(),
          status: sent ? 'success' : 'failed',
          payload: sent
            ? `Email sent to ${creatorEmail || creatorName} for project "${values.name}"`
            : 'Webhook call failed — verify Power Automate URL',
        }
        setPaFlowLog((prev) => [logEntry, ...prev].slice(0, 50))
      })

      // Also add in-app notification
      setState((prev) => ({
        ...prev,
        notifications: [
          {
            id: uid('ntf'),
            createdAt: todayISO(),
            severity: 'GREEN' as RygStatus,
            message: `Project "${values.name}" (${projectCode}) created successfully. Notification email queued.`,
            recipients: ['PROJECT_MANAGER', 'MANAGER', 'DIRECTOR'],
          },
          ...prev.notifications,
        ],
      }))
    }

    // ── Send via Azure Function + Graph API (parallel, fire-and-forget) ──
    const creatorEmail = (paContext as any)?.user?.email
      ?? (paContext as any)?.user?.loginName
      ?? (paContext as any)?.user?.userName
      ?? ''
    const creatorNameForFunc = paContext?.user.fullName ?? displayName
    sendProjectCreationViaFunction({
      projectName: values.name,
      projectCode,
      category: values.category,
      objective: values.objective,
      siteName: siteNames,
      sponsorExecutive: sponsorName,
      projectManager: pmName,
      budgetAllocated: values.budgetAllocated,
      plannedStartDate: values.plannedStartDate,
      plannedEndDate: values.plannedEndDate,
      clientCode: values.clientCode,
      milestones: allEmailMs,
      creatorName: creatorNameForFunc,
      creatorEmail,
    }).catch((err) => console.warn('[FuncEmail] notification failed:', err))

    setSelectedProjectId(newProject.id)
    setNavPage('overview')
    setCreatingSaving(false)
    newProjectForm.reset()
    setInlineMilestones([])
  }

  async function saveOverview(values: ProjectOverviewForm): Promise<void> {
    if (!selectedProject) return

    const pmName = dvUsers.find((u) => u.id === values.projectManagerId)?.fullname
      ?? people.find((p) => p.id === values.projectManagerId)?.name ?? ''
    const sponsorName = dvUsers.find((u) => u.id === values.sponsorExecutiveId)?.fullname
      ?? people.find((p) => p.id === values.sponsorExecutiveId)?.name ?? ''

    // Patch Dataverse
    if (isDataverseConfigured()) {
      try {
        await updateProjectInDataverse(selectedProject.id, {
          name: values.name,
          clientCode: values.clientCode,
          category: values.category,
          budgetAllocated: values.budgetAllocated,
          site: values.siteIds[0] ?? 'site_tca',
          objective: values.objective,
          sponsorExecutive: sponsorName,
          projectManagerName: pmName,
          plannedStartDate: values.plannedStartDate,
          plannedEndDate: values.plannedEndDate,
        })
      } catch (err) {
        console.warn('Dataverse update failed:', err)
      }
    }

    setState((previous) => {
      const categoryChanged = selectedProject.category !== values.category
      const nextProject: Project = {
        ...selectedProject,
        name: values.name,
        clientCode: values.clientCode,
        category: values.category,
        budgetAllocated: values.budgetAllocated,
        siteIds: values.siteIds,
        sponsorExecutiveId: values.sponsorExecutiveId,
        projectManagerId: values.projectManagerId,
        clientIds: values.clientIds,
        supplierIds: values.supplierIds,
        objective: values.objective,
        plannedStartDate: values.plannedStartDate,
        plannedEndDate: values.plannedEndDate,
      }

      const nextProjects = previous.projects.map((project) => (project.id === selectedProject.id ? nextProject : project))
      const nextMilestones = categoryChanged
        ? [...previous.milestones.filter((milestone) => milestone.projectId !== selectedProject.id), ...makeAutoMilestones(selectedProject.id, values.category)]
        : previous.milestones

      return {
        ...previous,
        projects: nextProjects,
        milestones: nextMilestones,
      }
    })
  }

  async function handleDeleteProject(): Promise<void> {
    if (!selectedProject) return
    setDeletingProject(true)
    const pid = selectedProject.id

    // Cascade-delete from Dataverse: children first (activities, milestones), then project
    if (isDataverseConfigured()) {
      try {
        const childActivities = state.activities.filter((a) => a.projectId === pid)
        const childMilestones = state.milestones.filter((m) => m.projectId === pid)
        // Delete activities (assignments cascade via Dataverse referential constraints)
        await Promise.all(childActivities.map((a) => deleteActivityInDataverse(a.id).catch(() => {})))
        // Delete milestones
        await Promise.all(childMilestones.map((m) => deleteMilestoneInDataverse(m.id).catch(() => {})))
        // Delete project
        await deleteProjectInDataverse(pid)
      } catch (err) { console.warn('Dataverse cascade-delete failed:', err) }
    }

    setState((prev) => ({
      ...prev,
      projects: prev.projects.filter((p) => p.id !== pid),
      milestones: prev.milestones.filter((m) => m.projectId !== pid),
      activities: prev.activities.filter((a) => a.projectId !== pid),
      assignments: prev.assignments.filter((asgn) =>
        !prev.activities.some((a) => a.projectId === pid && a.id === asgn.activityId)
      ),
    }))

    setDeleteProjectDialogOpen(false)
    setDeletingProject(false)
    setDetailProjectId(null)
    setEditingProject(false)
    setSelectedProjectId(null)
  }

  function openCreateMilestone(): void {
    setMilestoneEditId(null)
    milestoneForm.reset({ name: '', targetDate: addDays(todayISO(), 7) })
    setMilestoneDialogOpen(true)
  }

  function openEditMilestone(ms: Milestone): void {
    setMilestoneEditId(ms.id)
    milestoneForm.reset({ name: ms.name, targetDate: ms.targetDate })
    setMilestoneDialogOpen(true)
  }

  async function submitMilestone(values: MilestoneForm): Promise<void> {
    if (!selectedProject) return

    if (!milestoneEditId) {
      // Create
      let dvId: string | null = null
      if (isDataverseConfigured()) {
        dvId = await createMilestoneInDataverse({ name: values.name, targetDate: values.targetDate, projectId: selectedProject.id })
      }
      const newId = dvId ?? uid('ms')
      setState((prev) => ({
        ...prev,
        milestones: [...prev.milestones, { id: newId, projectId: selectedProject.id, name: values.name, targetDate: values.targetDate }],
      }))
    } else {
      // Update
      if (isDataverseConfigured()) {
        const existing = state.milestones.find((ms) => ms.id === milestoneEditId)
        await updateMilestoneInDataverse(milestoneEditId, { name: values.name, targetDate: values.targetDate, projectId: selectedProject.id, doneDate: existing?.doneDate })
      }
      setState((prev) => ({
        ...prev,
        milestones: prev.milestones.map((ms) => (ms.id === milestoneEditId ? { ...ms, name: values.name, targetDate: values.targetDate } : ms)),
      }))
    }

    setMilestoneDialogOpen(false)
  }

  function deleteMilestone(milestoneId: string): void {
    if (isDataverseConfigured()) { deleteMilestoneInDataverse(milestoneId) }
    setState((prev) => ({
      ...prev,
      milestones: prev.milestones.filter((ms) => ms.id !== milestoneId),
      activities: prev.activities.map((activity) => (activity.milestoneId === milestoneId ? { ...activity, milestoneId: undefined } : activity)),
    }))
  }

  function openCreateActivity(milestoneId: string | null): void {
    setActivityMilestoneId(milestoneId)
    setActivityEditMode('create')
    setEditingActivityId(null)
    activityForm.reset({
      name: '',
      ownerId: '',
      startDate: todayISO(),
      endDate: addDays(todayISO(), 7),
      workloadPct: 0,
      responsible: '',
      state: 'NOT_STARTED',
    })
    setActivityDialogOpen(true)
  }

  function openEditActivity(activity: Activity): void {
    setActivityMilestoneId(activity.milestoneId ?? null)
    setActivityEditMode('edit')
    setEditingActivityId(activity.id)
    activityForm.reset({
      name: activity.name,
      ownerId: activity.ownerId,
      startDate: activity.startDate,
      endDate: activity.endDate,
      workloadPct: activity.workloadPct ?? 0,
      responsible: activity.responsible ?? '',
      state: activity.state,
    })
    setActivityDialogOpen(true)
  }

  async function submitActivity(values: ActivityForm): Promise<void> {
    if (!selectedProject) return
    const ownerName = people.find((p) => p.id === values.ownerId)?.name ?? ''
    const workloadPct = Number(values.workloadPct) || 0
    const responsible = values.responsible.trim()
    const newState = values.state
    if (activityEditMode === 'edit' && editingActivityId) {
      const existing = state.activities.find((a) => a.id === editingActivityId)
      // When marking DONE, stamp a completion date (end date, or today); clear it otherwise.
      const doneDate = newState === 'DONE' ? (existing?.doneDate ?? values.endDate ?? todayISO()) : undefined
      if (isDataverseConfigured() && existing) {
        await updateActivityInDataverse(editingActivityId, {
          name: values.name, ownerName, startDate: values.startDate, endDate: values.endDate,
          doneDate, state: newState, criticalPath: existing.criticalPath,
          projectId: existing.projectId, milestoneId: existing.milestoneId,
          workloadPct, responsible,
        })
      }
      setState((prev) => ({
        ...prev,
        activities: prev.activities.map((a) =>
          a.id === editingActivityId
            ? { ...a, name: values.name, ownerId: values.ownerId, startDate: values.startDate, endDate: values.endDate, workloadPct, responsible, state: newState, doneDate }
            : a,
        ),
      }))
    } else {
      const doneDate = newState === 'DONE' ? (values.endDate || todayISO()) : undefined
      let dvId: string | null = null
      if (isDataverseConfigured()) {
        dvId = await createActivityInDataverse({
          name: values.name, ownerName, startDate: values.startDate, endDate: values.endDate,
          doneDate, state: newState, criticalPath: false, projectId: selectedProject.id,
          milestoneId: activityMilestoneId ?? undefined,
          workloadPct, responsible,
        })
      }
      const newId = dvId ?? uid('act')
      setState((prev) => ({
        ...prev,
        activities: [
          ...prev.activities,
          {
            id: newId,
            projectId: selectedProject.id,
            milestoneId: activityMilestoneId ?? undefined,
            name: values.name,
            ownerId: values.ownerId,
            startDate: values.startDate,
            endDate: values.endDate,
            state: newState,
            doneDate,
            criticalPath: false,
            workloadPct,
            responsible,
          },
        ],
      }))
    }
    setActivityDialogOpen(false)
    setEditingActivityId(null)
  }

  function handleDropOnMilestone(milestoneId: string): void {
    if (!dragPersonId || !selectedProject) return
    const person = people.find((p) => p.id === dragPersonId)
    if (!person) return
    setActivityMilestoneId(milestoneId)
    activityForm.reset({
      name: `${person.name} – New Task`,
      ownerId: dragPersonId,
      startDate: todayISO(),
      endDate: addDays(todayISO(), 7),
    })
    setActivityDialogOpen(true)
    setDragPersonId(null)
    setDropTargetId(null)
  }

  function handleDropOnTask(activityId: string): void {
    if (!dragPersonId) return
    const personId = dragPersonId
    const act = state.activities.find((a) => a.id === activityId)
    setDragPersonId(null)
    setDropTargetId(null)
    if (!act) return

    // If the task has no workload yet, open the edit dialog pre-filled with the
    // owner so the user must supply a workload % (prompt-on-assign).
    if (act.workloadPct == null || act.workloadPct === 0) {
      openEditActivity({ ...act, ownerId: personId })
      return
    }

    // Otherwise assign directly + persist owner to Dataverse.
    const ownerName = people.find((p) => p.id === personId)?.name ?? ''
    setState((prev) => ({
      ...prev,
      activities: prev.activities.map((a) => (a.id === activityId ? { ...a, ownerId: personId } : a)),
    }))
    if (isDataverseConfigured()) {
      updateActivityInDataverse(activityId, {
        name: act.name, ownerName, startDate: act.startDate, endDate: act.endDate,
        doneDate: act.doneDate, state: act.state, criticalPath: act.criticalPath,
        projectId: act.projectId, milestoneId: act.milestoneId,
        workloadPct: act.workloadPct, responsible: act.responsible,
      })
    }
  }

  function simulateWeeklyAlerts(): void {
    const delayedMilestones = state.milestones.filter((milestone) => !milestone.doneDate && rygFromDueDate(milestone.targetDate, state.settings.warningDaysThreshold) === 'RED').length
    const delayedActs = state.activities.filter((activity) => !activity.doneDate && rygFromDueDate(activity.endDate, state.settings.warningDaysThreshold) === 'RED').length
    const overAllocated = workloadRows.filter((row) => row.utilization > state.settings.workloadCriticalThreshold).length

    const generated: NotificationItem[] = [
      {
        id: uid('ntf'),
        createdAt: todayISO(),
        severity: delayedMilestones > 0 ? 'RED' : 'GREEN',
        message: `Weekly digest: ${delayedMilestones} delayed milestones sent to all team members.`,
        recipients: ['PROJECT_MANAGER', 'MANAGER', 'DIRECTOR'],
      },
      {
        id: uid('ntf'),
        createdAt: todayISO(),
        severity: delayedActs > 0 ? 'RED' : 'YELLOW',
        message: `Weekly digest: ${delayedActs} delayed activities sent to activity owners.`,
        recipients: ['PROJECT_MANAGER', 'MANAGER', 'DIRECTOR'],
      },
      {
        id: uid('ntf'),
        createdAt: todayISO(),
        severity: overAllocated > 0 ? 'RED' : 'GREEN',
        message: `Weekly digest: ${overAllocated} resources over 100% utilization.`,
        recipients: ['MANAGER', 'DIRECTOR'],
      },
    ]

    setState((prev) => ({
      ...prev,
      notifications: [...generated, ...prev.notifications],
    }))

    // Trigger enabled Power Automate flows
    triggerPowerAutomateFlows(generated)
  }

  /* ── Power Automate Handlers ── */
  function togglePaFlow(flowId: string): void {
    setPaFlows((prev) => prev.map((f) => (f.id === flowId ? { ...f, enabled: !f.enabled } : f)))
  }

  function triggerPowerAutomateFlows(notifications: NotificationItem[]): void {
    const enabledFlows = paFlows.filter((f) => f.enabled)
    const newLogs: FlowLogEntry[] = enabledFlows.map((flow) => {
      const relevant = notifications.filter((n) => {
        if (flow.triggerType === 'overdue_milestones') return n.message.includes('milestone')
        if (flow.triggerType === 'delayed_activities') return n.message.includes('activit')
        if (flow.triggerType === 'workload_threshold') return n.message.includes('utilization')
        if (flow.triggerType === 'project_created') return n.message.includes('created')
        return false
      })
      return {
        id: uid('flog'),
        flowName: flow.name,
        triggeredAt: new Date().toISOString(),
        status: 'success' as const,
        payload: `${relevant.length} notification(s) → ${flow.name}`,
      }
    })
    setPaFlowLog((prev) => [...newLogs, ...prev].slice(0, 50))
  }

  function testPowerAutomateFlow(flowId: string): void {
    const flow = paFlows.find((f) => f.id === flowId)
    if (!flow) return
    setPaTestingFlowId(flowId)
    // Simulate async HTTP call to Power Automate webhook
    setTimeout(() => {
      const status: FlowLogEntry['status'] = Math.random() > 0.15 ? 'success' : 'failed'
      setPaFlowLog((prev) => [
        {
          id: uid('flog'),
          flowName: flow.name,
          triggeredAt: new Date().toISOString(),
          status,
          payload: status === 'success' ? 'Test payload sent successfully' : 'Connection timed out – verify webhook URL',
        },
        ...prev,
      ].slice(0, 50))
      setPaTestingFlowId(null)
    }, 1200)
  }

  /* ── Master Data CRUD ── */
  function openAddPerson(): void {
    setMdEditId(null)
    setMdForm({ name: '', initials: '', role: 'RESOURCE', managerId: '', area: '', location: [], kind: 'Associate', capacity: 40 })
    setAadQuery('')
    setAadResults([])
    setMdDialogOpen(true)
  }
  function openEditPerson(p: Person): void {
    setMdEditId(p.id)
    setMdForm({ name: p.name, initials: p.initials, role: p.role, managerId: p.managerId ?? '', area: p.area ?? '', location: p.location ?? [], kind: p.resourceKind ?? 'Associate', capacity: p.weeklyCapacity ?? 40 })
    setMdDialogOpen(true)
  }
  async function submitPerson(): Promise<void> {
    if (!mdForm.name.trim()) return
    const initials = mdForm.initials.trim() || mdForm.name.split(' ').map((w) => w[0]).join('').toUpperCase().slice(0, 2)
    const isResource = mdForm.role === 'RESOURCE'
    const payload = {
      name: mdForm.name.trim(),
      role: mdForm.role,
      kind: isResource ? mdForm.kind : undefined,
      area: isResource ? (mdForm.area || undefined) : undefined,
      location: isResource ? mdForm.location : undefined,
      managerId: mdForm.managerId || undefined,
      department: isResource ? (mdForm.area || undefined) : undefined,
      weeklyCapacity: mdForm.capacity,
    }
    const personFields: Partial<Person> = {
      managerId: mdForm.managerId || undefined,
      area: isResource ? (mdForm.area || undefined) : undefined,
      location: isResource ? mdForm.location : undefined,
      resourceKind: isResource ? mdForm.kind : undefined,
      weeklyCapacity: mdForm.capacity,
    }
    if (mdEditId) {
      if (isDataverseConfigured()) {
        await updateResourceInDataverse(mdEditId, payload)
      }
      setState((prev) => ({ ...prev, people: prev.people.map((p) => p.id === mdEditId ? { ...p, name: mdForm.name.trim(), initials, role: mdForm.role, ...personFields } : p) }))
    } else {
      let dvId: string | null = null
      if (isDataverseConfigured()) {
        dvId = await createResourceInDataverse(payload)
      }
      const newPerson: Person = { id: dvId ?? uid('u'), name: mdForm.name.trim(), initials, role: mdForm.role, ...personFields }
      setState((prev) => ({ ...prev, people: [...prev.people, newPerson] }))
    }
    setMdDialogOpen(false)
  }
  function deletePerson(id: string): void {
    if (isDataverseConfigured()) { deleteResourceInDataverse(id) }
    setState((prev) => ({
      ...prev,
      people: prev.people.filter((p) => p.id !== id),
      activities: prev.activities.map((a) => a.ownerId === id ? { ...a, ownerId: prev.people.find((pp) => pp.role === 'RESOURCE' && pp.id !== id)?.id ?? a.ownerId } : a),
    }))
  }

  // ── Task template editor handlers ──
  function openAddTemplate(): void {
    setTtEditId(null)
    const maxSeq = taskTemplates.filter((t) => t.projectType === ttType).reduce((m, t) => Math.max(m, t.sequence), 0)
    setTtForm({ task: '', responsible: '', sequence: maxSeq + 1, leadtimeWeeks: '', workloadPct: '', inputs: '' })
    setTtDialogOpen(true)
  }
  function openEditTemplate(t: TaskTemplateRow): void {
    setTtEditId(t.id)
    setTtForm({ task: t.task, responsible: t.responsible, sequence: t.sequence, leadtimeWeeks: t.leadtimeWeeks == null ? '' : String(t.leadtimeWeeks), workloadPct: t.workloadPct == null ? '' : String(t.workloadPct), inputs: t.inputs })
    setTtDialogOpen(true)
  }
  async function submitTemplate(): Promise<void> {
    if (!ttForm.task.trim() || !ttType) return
    const payload = {
      projectType: ttType,
      sequence: Number(ttForm.sequence) || 0,
      task: ttForm.task.trim(),
      responsible: ttForm.responsible.trim(),
      leadtimeWeeks: ttForm.leadtimeWeeks === '' ? null : Number(ttForm.leadtimeWeeks),
      inputs: ttForm.inputs.trim(),
      workloadPct: ttForm.workloadPct === '' ? null : Number(ttForm.workloadPct),
    }
    if (ttEditId) await updateTaskTemplate(ttEditId, payload)
    else await createTaskTemplate(payload)
    await reloadTaskTemplates()
    setTtDialogOpen(false)
  }
  async function deleteTemplate(id: string): Promise<void> {
    await deleteTaskTemplate(id)
    setTaskTemplates((prev) => prev.filter((t) => t.id !== id))
  }

  function addPartner(type: 'clients' | 'suppliers', name: string): void {
    if (!name.trim()) return
    setState((prev) => ({
      ...prev,
      [type]: [...prev[type], { id: uid(type === 'clients' ? 'client' : 'supplier'), name: name.trim() }],
    }))
  }
  function editPartner(type: 'clients' | 'suppliers', id: string, name: string): void {
    if (!name.trim()) return
    setState((prev) => ({
      ...prev,
      [type]: prev[type].map((item) => item.id === id ? { ...item, name: name.trim() } : item),
    }))
  }
  function deletePartner(type: 'clients' | 'suppliers', id: string): void {
    setState((prev) => ({
      ...prev,
      [type]: prev[type].filter((item) => item.id !== id),
    }))
  }

  const navItems: Array<{ key: NavPage; label: string; icon: ReactElement }> = [
    { key: 'overview', label: 'Dashboard', icon: <LayoutDashboard size={16} /> },
    { key: 'workload', label: 'Workload', icon: <Users size={16} /> },
    { key: 'reports', label: 'Reports', icon: <FolderKanban size={16} /> },
  ]

  const setupItems: Array<{ key: NavPage; label: string; icon: ReactElement }> = [
    { key: 'alerts', label: 'Alerts', icon: <BellRing size={16} /> },
    { key: 'admin', label: 'Admin', icon: <SlidersHorizontal size={16} /> },
    { key: 'masterData', label: 'Master Data', icon: <Database size={16} /> },
  ]

  const timelineItems = useMemo(() => {
    if (!selectedProject) return { minDate: todayISO(), maxDate: addDays(todayISO(), 1), activities: [], milestones: [] as Milestone[] }
    const activities = selectedProjectActivities
    const allDates = [
      ...activities.flatMap((activity) => [activity.startDate, activity.endDate]),
      ...selectedProjectMilestones.map((milestone) => milestone.targetDate),
    ].filter((d): d is string => !!d).sort()   // drop blanks (N/A tasks carry no dates)
    const minDate = allDates[0] ?? todayISO()
    const maxDate = allDates[allDates.length - 1] ?? addDays(todayISO(), 1)
    return { minDate, maxDate, activities, milestones: selectedProjectMilestones }
  }, [selectedProject, selectedProjectActivities, selectedProjectMilestones])



  return (
    <div className="relative flex min-h-screen bg-pth-bg font-sans text-pth-text md:bg-pth-outer md:pr-3 md:py-3">
      {/* Subtle blue-to-dark-gray gradient behind everything */}
      <div className="pointer-events-none absolute inset-0 hidden md:block" style={{ background: 'linear-gradient(135deg, #242a32 0%, #31343A 35%, #31343A 100%)' }} />
      {/* ─── DATAVERSE LOADING OVERLAY ─── */}
      {dvLoading && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-pth-bg/80 backdrop-blur-sm">
          <div className="flex flex-col items-center gap-4 rounded-2xl bg-pth-card p-8 shadow-elevated ring-1 ring-pth-border/20">
            <div className="h-10 w-10 animate-spin rounded-full border-4 border-pth-blue/20 border-t-pth-blue" />
            <div className="text-sm font-medium text-pth-muted">Loading…</div>
          </div>
        </div>
      )}
      {/* ─── SIDEBAR (desktop: sticky column) ─── */}
      <aside
        className={`fixed bottom-0 left-0 top-14 z-40 flex flex-col bg-transparent text-pth-sidebar-text transition-all duration-200 ease-in-out ${sidebarCollapsed ? 'w-16' : 'w-56'} ${sidebarOpen ? 'translate-x-0' : '-translate-x-full'} md:sticky md:top-0 md:z-auto md:h-screen md:translate-x-0 md:shrink-0`}
      >
        {/* Brand area — single source of logo */}
        <button
          type="button"
          className={`hidden w-full items-center border-b border-pth-sidebar-border transition-colors hover:bg-pth-sidebar-hover md:flex ${sidebarCollapsed ? 'justify-center px-2 py-4' : 'gap-3 px-4 py-4'}`}
          onClick={() => setSidebarCollapsed((c) => !c)}
          aria-label="Toggle sidebar"
        >
          <div className="flex h-9 w-9 shrink-0 items-center justify-center">
            <img src={logoImg} alt="Project Tracker Hub" className="h-9 w-9 object-contain" />
          </div>
          {!sidebarCollapsed && (
            <div className="min-w-0 flex-1 text-left">
              <div className="text-sm font-bold tracking-tight truncate text-pth-sidebar-text">Project Tracker Hub</div>
              <div className="text-[11px] text-pth-sidebar-muted truncate">{displayName}</div>
              <div className="text-[10px] text-pth-sidebar-faint truncate">{displayRole}</div>
            </div>
          )}
        </button>

        <div className={`flex-1 overflow-y-auto ${sidebarCollapsed ? 'p-2' : 'p-4'}`}>
          {!sidebarCollapsed && (
            <div className="mb-2 px-3 text-[10px] font-semibold uppercase tracking-wider text-pth-sidebar-faint">Main Menu</div>
          )}

          <nav className="space-y-1">
            {navItems.map((item) => (
              <button
                key={item.key}
                type="button"
                className={`relative flex w-full items-center gap-3 rounded-lg ${sidebarCollapsed ? 'justify-center px-2 py-2.5' : 'px-3 py-2.5'} text-sm font-medium transition-all duration-200 ${
                  navPage === item.key
                    ? 'text-pth-sidebar-text'
                    : 'text-pth-sidebar-muted hover:bg-pth-sidebar-hover hover:text-pth-sidebar-text'
                }`}
                onClick={() => {
                  setNavPage(item.key)
                  setSidebarOpen(false)
                }}
                title={sidebarCollapsed ? item.label : undefined}
              >
                {navPage === item.key && (
                  <motion.div layoutId="nav-active" className="absolute inset-0 rounded-lg bg-pth-sidebar-active" transition={{ type: 'spring', duration: 0.4, bounce: 0.15 }} />
                )}
                <span className={`relative z-10 flex items-center ${sidebarCollapsed ? '' : 'gap-3'} ${navPage === item.key ? 'text-pth-sidebar-accent' : ''}`}>
                  {item.icon}
                  {!sidebarCollapsed && <span className={navPage === item.key ? 'text-pth-sidebar-text' : ''}>{item.label}</span>}
                </span>
                {navPage === item.key && !sidebarCollapsed && <span className="relative z-10 ml-auto h-1.5 w-1.5 rounded-full bg-pth-sidebar-dot" />}
              </button>
            ))}

          {!sidebarCollapsed && (
            <div className="mb-2 mt-6 px-3 text-[10px] font-semibold uppercase tracking-wider text-pth-sidebar-faint">Setup</div>
          )}

            {setupItems.map((item) => (
              <button
                key={item.key}
                type="button"
                className={`relative flex w-full items-center gap-3 rounded-lg ${sidebarCollapsed ? 'justify-center px-2 py-2.5' : 'px-3 py-2.5'} text-sm font-medium transition-all duration-200 ${
                  navPage === item.key
                    ? 'text-pth-sidebar-text'
                    : 'text-pth-sidebar-muted hover:bg-pth-sidebar-hover hover:text-pth-sidebar-text'
                }`}
                onClick={() => {
                  setNavPage(item.key)
                  setSidebarOpen(false)
                }}
                title={sidebarCollapsed ? item.label : undefined}
              >
                {navPage === item.key && (
                  <motion.div layoutId="nav-active" className="absolute inset-0 rounded-lg bg-pth-sidebar-active" transition={{ type: 'spring', duration: 0.4, bounce: 0.15 }} />
                )}
                <span className={`relative z-10 flex items-center ${sidebarCollapsed ? '' : 'gap-3'} ${navPage === item.key ? 'text-pth-sidebar-accent' : ''}`}>
                  {item.icon}
                  {!sidebarCollapsed && <span className={navPage === item.key ? 'text-pth-sidebar-text' : ''}>{item.label}</span>}
                </span>
                {navPage === item.key && !sidebarCollapsed && <span className="relative z-10 ml-auto h-1.5 w-1.5 rounded-full bg-pth-sidebar-dot" />}
              </button>
            ))}
          </nav>

          {/* Mobile-only selectors */}
          {!sidebarCollapsed && (
            <div className="mt-4 space-y-3 border-t border-pth-sidebar-border pt-4 md:hidden">
              <label className="block text-xs font-medium text-pth-sidebar-muted">
                Month
                <select
                  className="mt-1 h-9 w-full rounded-lg border border-pth-sidebar-border bg-pth-sidebar-hover px-3 text-sm text-pth-sidebar-text"
                  value={filterMonth}
                  onChange={(e) => { setFilterMonth(e.target.value); setFilterCW('all') }}
                >
                  {monthOptions.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
                </select>
              </label>
              <label className="block text-xs font-medium text-pth-sidebar-muted">
                Week
                <select
                  className="mt-1 h-9 w-full rounded-lg border border-pth-sidebar-border bg-pth-sidebar-hover px-3 text-sm text-pth-sidebar-text"
                  value={filterCW}
                  onChange={(e) => setFilterCW(e.target.value === 'all' ? 'all' : Number(e.target.value))}
                >
                  <option value="all">All Weeks</option>
                  {cwOptions.map((cw) => <option key={cw} value={cw}>{`CW ${cw}`}</option>)}
                </select>
              </label>
            </div>
          )}
        </div>

        {/* ── Sidebar footer: version ── */}
        <div className={`shrink-0 border-t border-pth-sidebar-border ${sidebarCollapsed ? 'px-2 py-3' : 'px-4 py-4'}`}>
          <div className={`flex flex-col ${sidebarCollapsed ? 'items-center' : 'items-start'} gap-1`}>
            <span className={`font-bold text-pth-sidebar-faint ${sidebarCollapsed ? 'text-[8px]' : 'text-[10px]'}`}>
              {sidebarCollapsed ? 'v1.0' : 'Project Tracker Hub v1.0'}
            </span>
          </div>
        </div>
      </aside>

      {/* ─── NOTIFICATIONS DROPDOWN ─── */}
      <AnimatePresence>
        {notificationsOpen && (
          <>
            <div className="fixed inset-0 z-[65]" onClick={() => setNotificationsOpen(false)} />
            <motion.div
              initial={{ opacity: 0, y: -8, scale: 0.96 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: -8, scale: 0.96 }}
              transition={{ duration: 0.15 }}
              className="glass-card fixed right-4 top-[58px] z-[70] w-[380px] rounded-xl border border-pth-border/20 shadow-elevated md:right-6"
            >
              <div className="flex items-center justify-between border-b border-pth-border/20 px-4 py-3">
                <h3 className="text-sm font-semibold">Notifications</h3>
                <button type="button" title="Close notifications" className="rounded-lg p-1 text-pth-muted transition-colors hover:bg-pth-hover" onClick={() => setNotificationsOpen(false)}>
                  <X size={14} />
                </button>
              </div>
              <div className="max-h-[400px] overflow-y-auto p-2">
                {visibleNotifications.map((notification) => (
                  <div
                    key={notification.id}
                    className="rounded-lg p-3 transition-colors hover:bg-pth-subtle"
                  >
                    <div className="mb-1.5 flex items-center justify-between">
                      {statusBadge(notification.severity)}
                      <span className="text-[11px] text-pth-muted">{formatDate(notification.createdAt)}</span>
                    </div>
                    <p className="text-xs leading-relaxed text-pth-text/80">{notification.message}</p>
                  </div>
                ))}
              </div>
            </motion.div>
          </>
        )}
      </AnimatePresence>

      {/* Sidebar overlay (mobile) */}
      <AnimatePresence>
        {sidebarOpen && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-30 bg-black/30 backdrop-blur-sm md:hidden"
            onClick={() => setSidebarOpen(false)}
          />
        )}
      </AnimatePresence>

      {/* ─── RIGHT COLUMN: header + content ─── */}
      <div className="relative flex min-w-0 flex-1 flex-col bg-pth-bg text-pth-text md:ml-3 md:rounded-2xl md:shadow-lg md:overflow-hidden md:ring-1 md:ring-pth-sidebar-border">
        {/* Decorative gradient fade at top */}
        <div className="pointer-events-none absolute inset-x-0 top-0 z-0 hidden h-48 md:block md:rounded-t-2xl" style={{ background: 'linear-gradient(180deg, rgba(0,86,145,0.07) 0%, rgba(0,86,145,0.02) 40%, transparent 100%)' }} />
        {/* ─── HEADER ─── */}
        <header className="glass-card sticky top-0 z-50 border-b border-pth-border/30 shadow-header md:rounded-t-2xl" style={{ position: 'relative', zIndex: 50 }}>
          <div className="flex h-14 items-center gap-3 px-4 md:px-6">
            {/* Mobile hamburger */}
            <button
              type="button"
              className="inline-flex h-9 w-9 items-center justify-center rounded-lg text-pth-muted transition-colors hover:bg-pth-hover md:hidden"
              onClick={() => setSidebarOpen((open) => !open)}
              aria-label="Toggle sidebar"
            >
              {sidebarOpen ? <X size={18} /> : <Menu size={18} />}
            </button>

            {/* Logo — visible only on small screens */}
            <div className="flex items-center gap-2 md:hidden">
              <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-pth-blue">
                <FolderKanban size={16} className="text-white" />
              </div>
              <span className="text-sm font-bold tracking-tight sm:inline">PROJECT HUB</span>
            </div>

            {/* Desktop sidebar collapse toggle */}
            <button
              type="button"
              className="hidden h-9 w-9 items-center justify-center rounded-lg text-pth-muted transition-colors hover:bg-pth-hover md:inline-flex"
              onClick={() => setSidebarCollapsed((c) => !c)}
              aria-label={sidebarCollapsed ? 'Expand sidebar' : 'Collapse sidebar'}
            >
              {sidebarCollapsed ? <PanelLeftOpen size={18} /> : <PanelLeftClose size={18} />}
            </button>

            {/* Spacer */}
            <div className="flex-1" />

            {/* Dataverse connection badge */}
            {dvConnected && (
              <div className="hidden items-center gap-1.5 rounded-lg bg-emerald-500/8 px-2.5 py-1 text-[11px] font-medium text-emerald-600 dark:bg-emerald-500/15 dark:text-emerald-400 sm:inline-flex">
                <Database size={12} />
                Dataverse
              </div>
            )}

            {/* Theme toggle */}
            <button
              type="button"
              className="inline-flex h-9 w-9 items-center justify-center rounded-lg text-pth-muted transition-colors hover:bg-pth-hover"
              onClick={() => setDarkMode((prev) => !prev)}
              aria-label="Toggle dark mode"
            >
              {darkMode ? <Sun size={18} /> : <Moon size={18} />}
            </button>

            {/* Notifications */}
            <button
              type="button"
              className="relative inline-flex h-9 w-9 items-center justify-center rounded-lg text-pth-muted transition-colors hover:bg-pth-hover"
              onClick={() => setNotificationsOpen((open) => !open)}
              aria-label="Notifications"
            >
              <Bell size={18} />
              {visibleNotifications.length > 0 && (
                <span className="absolute -right-0.5 -top-0.5 inline-flex h-4 min-w-4 items-center justify-center rounded-full bg-pth-red px-1 text-[10px] font-bold text-white">
                  {visibleNotifications.length}
                </span>
              )}
            </button>

            {/* User avatar */}
            <div
              className="relative inline-flex h-9 w-9 shrink-0 items-center justify-center rounded-full bg-pth-btn text-xs font-bold text-white overflow-hidden"
              title={`${displayName} — ${displayRole}`}
            >
              {userPhotoUrl
                ? <img src={userPhotoUrl} alt={displayName} className="h-full w-full object-cover" />
                : displayInitials
              }
            </div>
          </div>
        </header>

        {/* ─── MAIN CONTENT ─── */}
        <main className="relative z-10 min-w-0 flex-1 p-4 md:p-6">
          <AnimatePresence mode="wait">
            <motion.section
              key={navPage}
              initial={{ opacity: 0, y: 8 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -8 }}
              transition={{ duration: 0.2, ease: 'easeOut' }}
              className="space-y-4"
            >
              {navPage === 'overview' && !detailProjectId && (
                <>
                  {/* ── Page header ── */}
                  <div className="flex flex-col gap-2 md:flex-row md:items-end md:justify-between">
                    <div>
                      <p className="text-xs font-medium uppercase tracking-wider text-pth-muted">PROJECT TRACKER HUB</p>
                      <h1 className="text-xl font-bold tracking-tight md:text-2xl">Project Dashboard</h1>
                    </div>
                    <div className="flex items-center gap-2">
                      <select title="Month" className="h-9 rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-xs transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" value={filterMonth} onChange={(e) => { setFilterMonth(e.target.value); setFilterCW('all') }}>
                        {monthOptions.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
                      </select>
                      <select title="Calendar Week" className="h-9 rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-xs transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" value={filterCW} onChange={(e) => setFilterCW(e.target.value === 'all' ? 'all' : Number(e.target.value))}>
                        <option value="all">All Weeks</option>
                        {cwOptions.map((cw) => <option key={cw} value={cw}>{`CW ${cw}`}</option>)}
                      </select>
                      <button type="button" onClick={() => { newProjectForm.reset(); setNavPage('createProject') }} className="inline-flex h-9 items-center gap-1.5 rounded-lg bg-pth-btn px-3 text-xs font-semibold text-white shadow-sm transition-all hover:bg-pth-btn-hover active:scale-[0.98]">
                        <Plus size={14} /> New Project
                      </button>
                    </div>
                  </div>

                  {/* ── KPI Cards ── */}
                  <div className="grid grid-cols-2 gap-3 md:grid-cols-5">
                    {[
                      { label: 'TOTAL', value: kpiCounts.total, icon: <Briefcase size={16} className="text-pth-muted" />, color: '' },
                      { label: 'ON TRACK', value: kpiCounts.onTrack, icon: <span className="h-3 w-3 rounded-full bg-emerald-500" />, color: 'text-emerald-600 dark:text-emerald-400' },
                      { label: 'AT RISK', value: kpiCounts.atRisk, icon: <span className="h-3 w-3 rounded-full bg-amber-500" />, color: 'text-amber-600 dark:text-amber-400' },
                      { label: 'CRITICAL', value: kpiCounts.critical, icon: <span className="h-3 w-3 rounded-full bg-pth-red" />, color: 'text-pth-red' },
                      { label: 'COMPLETED', value: kpiCounts.completed, icon: <span className="h-3 w-3 rounded-full border-2 border-pth-muted" />, color: '' },
                    ].map((kpi) => (
                      <div key={kpi.label} className="glass-card rounded-xl p-4 shadow-card">
                        <div className="flex items-center justify-between">
                          <span className="text-[11px] font-semibold uppercase tracking-wider text-pth-muted">{kpi.label}</span>
                          {kpi.icon}
                        </div>
                        <div className={`mt-2 text-2xl font-bold ${kpi.color}`}>{kpi.value}</div>
                      </div>
                    ))}
                  </div>


                  {/* ── Filters ── */}
                  <div className="flex flex-col gap-3 md:flex-row md:items-center">
                    <div className="relative flex-1">
                      <input
                        className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle pl-10 pr-4 text-sm placeholder:text-pth-muted transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20"
                        placeholder="Search projects..."
                        value={projectSearch}
                        onChange={(e) => setProjectSearch(e.target.value)}
                      />
                      <svg className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-pth-muted" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2"><circle cx="11" cy="11" r="8" /><path d="m21 21-4.35-4.35" /></svg>
                    </div>
                    <select title="Category filter" className="h-10 rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" value={filterCategory} onChange={(e) => setFilterCategory(e.target.value)}>
                      <option value="all">All Categories</option>
                      {(['ECR', 'CIP', 'Path Forward', 'New Programs'] as ProjectCategory[]).map((c) => <option key={c} value={c}>{c}</option>)}
                    </select>
                    <select title="Status filter" className="h-10 rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)}>
                      <option value="all">All Status</option>
                      <option value="GREEN">On Track</option>
                      <option value="YELLOW">At Risk</option>
                      <option value="RED">Critical</option>
                    </select>
                    <select title="Manager filter" className="h-10 rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" value={filterManager} onChange={(e) => setFilterManager(e.target.value)}>
                      <option value="all">All Managers</option>
                      {dvConnected && dvUsers.length > 0
                        ? dvUsers.map((u) => <option key={u.id} value={u.id}>{u.fullname}</option>)
                        : people.filter((p) => p.role === 'PROJECT_MANAGER').map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                    </select>
                  </div>

                  {/* ── Active Projects header with counter + view toggle ── */}
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-3">
                      <h2 className="text-xs font-semibold uppercase tracking-wider text-pth-muted">Active Projects</h2>
                      <span className="rounded-md bg-pth-subtle px-2 py-0.5 text-[11px] font-medium text-pth-muted">
                        Showing {filteredProjects.length} of {scopedProjects.length}
                      </span>
                    </div>
                    <div className="flex items-center gap-2">
                    <div className="flex items-center gap-1 rounded-lg bg-pth-subtle p-0.5">
                      <button type="button" title="Gallery view" className={`inline-flex h-7 w-7 items-center justify-center rounded-md transition-colors ${viewMode === 'gallery' ? 'bg-pth-card shadow-sm text-pth-text' : 'text-pth-muted hover:text-pth-text'}`} onClick={() => setViewMode('gallery')}>
                        <LayoutGrid size={14} />
                      </button>
                      <button type="button" title="List view" className={`inline-flex h-7 w-7 items-center justify-center rounded-md transition-colors ${viewMode === 'list' ? 'bg-pth-card shadow-sm text-pth-text' : 'text-pth-muted hover:text-pth-text'}`} onClick={() => setViewMode('list')}>
                        <List size={14} />
                      </button>
                    </div>
                    </div>
                  </div>

                  {/* ── Projects: Gallery View ── */}
                  {viewMode === 'gallery' && (
                  <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-3">
                    {projectRows.map((row) => {
                      const projectMilestones = state.milestones.filter((ms) => ms.projectId === row.project.id)
                      const projectActivities = state.activities.filter((a) => a.projectId === row.project.id)
                      const msStatus = projectMilestones.length > 0 ? projectMilestones.map((ms) => rygFromDueDate(ms.targetDate, state.settings.warningDaysThreshold)).sort((a, b) => rankStatus(a) - rankStatus(b))[0] : 'GREEN'
                      const actStatus = projectActivities.length > 0 ? projectActivities.map((a) => rygFromDueDate(a.endDate, state.settings.warningDaysThreshold)).sort((a, b) => rankStatus(a) - rankStatus(b))[0] : 'GREEN'
                      const qualityStatus: RygStatus = 'GREEN'
                      const pm = people.find((p) => p.id === row.project.projectManagerId)
                      const sponsor = people.find((p) => p.id === row.project.sponsorExecutiveId)
                      const endDate = projectActivities.length > 0 ? projectActivities.map((a) => a.endDate).sort().slice(-1)[0] : null

                      return (
                        <button
                          key={row.project.id}
                          type="button"
                          onClick={() => {
                            setDetailProjectId(row.project.id)
                            setSelectedProjectId(row.project.id)
                          }}
                          className="glass-card group cursor-pointer rounded-xl p-5 text-left shadow-card transition-all duration-200 hover:shadow-elevated"
                        >
                          {/* Top: category + ID + status */}
                          <div className="flex items-start justify-between gap-2">
                            <div className="flex items-center gap-2">
                              {categoryBadge(row.project.category)}
                              <span className="text-[11px] font-medium text-pth-muted">{row.project.projectCode}</span>
                            </div>
                            {statusBadge(row.status)}
                          </div>

                          {/* Title */}
                          <h3 className="mt-3 text-sm font-semibold leading-snug group-hover:text-pth-blue transition-colors">{row.project.name}</h3>

                          {/* Description */}
                          <p className="mt-1.5 line-clamp-2 text-xs leading-relaxed text-pth-muted">{row.project.objective}</p>

                          {/* Meta: Due date + Site */}
                          <div className="mt-4 flex items-center gap-6 text-[11px] text-pth-muted">
                            {endDate && (
                              <div>
                                <span className="block text-[10px] font-semibold uppercase tracking-wider text-pth-muted/60">DUE</span>
                                <span className="font-medium text-pth-text">{formatDate(endDate)}</span>
                              </div>
                            )}
                            <div>
                              <span className="block text-[10px] font-semibold uppercase tracking-wider text-pth-muted/60">SITE</span>
                              <span className="font-medium text-pth-text">{row.site}</span>
                            </div>
                          </div>

                          {/* Progress bar */}
                          <div className="mt-3">
                            <div className="mb-1 flex items-center justify-between text-[10px]">
                              <span className="font-semibold uppercase tracking-wider text-pth-muted/60">Progress</span>
                              <span className="font-medium text-pth-muted">{row.progressPct}%</span>
                            </div>
                            <div className="h-1.5 w-full overflow-hidden rounded-full bg-pth-border/30">
                              <div
                                ref={dynRef({ width: `${row.progressPct}%` })}
                                className="h-full rounded-full transition-all bg-pth-blue"
                              />
                            </div>
                          </div>

                          {/* RYG triple indicators */}
                          <div className="mt-4 flex items-center gap-3">
                            <div className="flex items-center gap-1.5 text-[11px]">
                              <span className="font-semibold text-pth-muted">T</span>
                              {statusBadge(msStatus)}
                            </div>
                            <div className="flex items-center gap-1.5 text-[11px]">
                              <span className="font-semibold text-pth-muted">C</span>
                              {statusBadge(actStatus)}
                            </div>
                            <div className="flex items-center gap-1.5 text-[11px]">
                              <span className="font-semibold text-pth-muted">Q</span>
                              {statusBadge(qualityStatus)}
                            </div>
                          </div>

                          {/* PM + Sponsor */}
                          <div className="mt-4 flex items-center gap-2 border-t border-pth-border/15 pt-3">
                            <div className="flex h-7 w-7 items-center justify-center rounded-full bg-pth-subtle text-[10px] font-bold text-pth-text">{pm?.initials ?? '??'}</div>
                            <div className="min-w-0 text-xs leading-tight">
                              <div className="truncate font-medium">{pm?.name ?? 'Unknown'}</div>
                              <div className="truncate text-pth-muted">{sponsor?.name ?? 'Unknown'}</div>
                            </div>
                          </div>
                        </button>
                      )
                    })}
                  </div>
                  )}

                  {/* ── Projects: List View ── */}
                  {viewMode === 'list' && (
                  <div className="glass-card overflow-hidden rounded-xl shadow-card">
                    <table className="w-full text-sm">
                      <thead>
                        <tr className="bg-pth-subtle text-[11px] font-semibold uppercase tracking-wider text-pth-muted">
                          <th className="px-4 py-3 text-left">ID</th>
                          <th className="px-4 py-3 text-left">Project</th>
                          <th className="hidden px-4 py-3 text-left md:table-cell">Category</th>
                          <th className="px-4 py-3 text-left">Status</th>
                          <th className="hidden px-4 py-3 text-left md:table-cell">Manager</th>
                          <th className="hidden px-4 py-3 text-left lg:table-cell">Site</th>
                          <th className="px-4 py-3 text-left">Due</th>
                          <th className="px-4 py-3 text-left">Progress</th>
                          <th className="hidden px-4 py-3 text-left lg:table-cell">Delayed</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-pth-border/10">
                        {projectRows.map((row) => (
                          <tr
                            key={row.project.id}
                            className="cursor-pointer transition-colors hover:bg-pth-hover"
                            onClick={() => {
                              setDetailProjectId(row.project.id)
                              setSelectedProjectId(row.project.id)
                            }}
                          >
                            <td className="px-4 py-3">
                              <span className="text-xs font-mono font-medium text-pth-blue">{row.project.projectCode}</span>
                            </td>
                            <td className="px-4 py-3">
                              <div className="font-medium">{row.project.name}</div>
                            </td>
                            <td className="hidden px-4 py-3 md:table-cell">{categoryBadge(row.project.category)}</td>
                            <td className="px-4 py-3">{statusBadge(row.status)}</td>
                            <td className="hidden px-4 py-3 text-pth-muted md:table-cell">{row.pm}</td>
                            <td className="hidden px-4 py-3 text-pth-muted lg:table-cell">{row.site}</td>
                            <td className="px-4 py-3 text-xs text-pth-muted whitespace-nowrap">{row.dueDate ? formatDate(row.dueDate) : '—'}</td>
                            <td className="px-4 py-3">
                              <div className="flex items-center gap-2">
                                <div className="h-1.5 w-20 overflow-hidden rounded-full bg-pth-border/30">
                                  <div
                                    ref={dynRef({ width: `${row.progressPct}%` })}
                                    className="h-full rounded-full transition-all bg-pth-blue"
                                  />
                                </div>
                                <span className="text-[11px] font-medium text-pth-muted">{row.progressPct}%</span>
                              </div>
                            </td>
                            <td className="hidden px-4 py-3 lg:table-cell">{row.delayedCount > 0 ? <span className="text-xs font-semibold text-pth-red">{row.delayedCount}</span> : <span className="text-xs text-pth-muted">0</span>}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  )}
                </>
              )}

              {/* ── Project Detail (drill-in from card) ── */}
              {navPage === 'overview' && detailProjectId && selectedProject && (
                <>
                  <button type="button" className="inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-sm font-medium text-pth-muted transition-colors hover:bg-pth-hover hover:text-pth-text" onClick={() => { setDetailProjectId(null); setEditingProject(false) }}>
                    <ChevronLeft size={16} /> Back to Dashboard
                  </button>

                  {(() => {
                    const projStatus = seededProjectStatus(state, selectedProject.id)
                    const dvPm = dvUsers.find((u) => u.id === selectedProject.projectManagerId)
                    const pm = dvPm ? { id: dvPm.id, name: dvPm.fullname } : people.find((p) => p.id === selectedProject.projectManagerId)
                    const dvSponsor = dvUsers.find((u) => u.id === selectedProject.sponsorExecutiveId)
                    const sponsor = dvSponsor ? { id: dvSponsor.id, name: dvSponsor.fullname } : people.find((p) => p.id === selectedProject.sponsorExecutiveId)
                    const pActs = selectedProjectActivities
                    const pMss = selectedProjectMilestones
                    const projClients = selectedProject.clientIds.map((cid) => clients.find((c) => c.id === cid)?.name).filter(Boolean).join(', ') || 'Not Applicable'
                    const projSuppliers = selectedProject.supplierIds.map((sid) => suppliers.find((s) => s.id === sid)?.name).filter(Boolean).join(', ') || '—'
                    const plannedStart = pActs.length > 0 ? [...pActs.map((a) => a.startDate)].sort()[0] : null
                    const plannedEnd = pActs.length > 0 ? [...pActs.map((a) => a.endDate)].sort().slice(-1)[0] : null
                    const realStart = pActs.filter((a) => a.state !== 'NOT_STARTED').length > 0 ? [...pActs.filter((a) => a.state !== 'NOT_STARTED').map((a) => a.startDate)].sort()[0] : null
                    const msWorst: RygStatus = pMss.length > 0 ? pMss.map((ms) => rygFromDueDate(ms.targetDate, state.settings.warningDaysThreshold)).sort((a, b) => rankStatus(a) - rankStatus(b))[0] : 'GREEN'
                    const actWorst: RygStatus = pActs.filter((a) => !a.doneDate).length > 0 ? pActs.filter((a) => !a.doneDate).map((a) => rygFromDueDate(a.endDate, state.settings.warningDaysThreshold)).sort((a, b) => rankStatus(a) - rankStatus(b))[0] : 'GREEN'
                    const tlMin = timelineItems.minDate
                    const tlMax = timelineItems.maxDate
                    const tlRange = Math.max(dayDiff(tlMin, tlMax), 1)
                    const todayPct = Math.min(Math.max((dayDiff(tlMin, todayISO()) / tlRange) * 100, 0), 100)

                    return (
                      <motion.div
                        layout
                        initial={{ opacity: 0, y: 12 }}
                        animate={{ opacity: 1, y: 0 }}
                        transition={{ duration: 0.25, ease: 'easeOut' }}
                        className="space-y-6"
                      >
                        {/* ── Project Header ── */}
                        <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                          <div>
                            <div className="flex items-center gap-3 flex-wrap">
                              <h1 className="text-xl font-bold md:text-2xl">{selectedProject.name}</h1>
                              {statusBadge(projStatus)}
                            </div>
                            <p className="mt-1 text-sm text-pth-muted">Code: {selectedProject.projectCode}</p>
                            <div className="mt-2 flex items-center gap-2">
                              {categoryBadge(selectedProject.category)}
                            </div>
                          </div>
                          <div className="flex items-center gap-2">
                            <button type="button" className="inline-flex items-center gap-2 rounded-lg border border-pth-border/40 bg-pth-card px-4 py-2 text-sm font-medium text-pth-text shadow-sm transition-colors hover:bg-pth-hover" onClick={() => setEditingProject(true)}>
                              <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2"><path strokeLinecap="round" strokeLinejoin="round" d="M16.862 4.487l1.687-1.688a1.875 1.875 0 112.652 2.652L10.582 16.07a4.5 4.5 0 01-1.897 1.13L6 18l.8-2.685a4.5 4.5 0 011.13-1.897l8.932-8.931zm0 0L19.5 7.125M18 14v4.75A2.25 2.25 0 0115.75 21H5.25A2.25 2.25 0 013 18.75V8.25A2.25 2.25 0 015.25 6H10" /></svg>
                              Edit Project
                            </button>
                            <button type="button" className="inline-flex items-center gap-2 rounded-lg border border-red-200 bg-pth-card px-4 py-2 text-sm font-medium text-pth-red shadow-sm transition-colors hover:bg-red-50 dark:border-red-900/40 dark:hover:bg-red-900/20" onClick={() => setDeleteProjectDialogOpen(true)}>
                              <Trash2 size={16} />
                              Delete
                            </button>
                          </div>
                        </div>

                        {/* ── T / C / Q / Risk Indicators ── */}
                        <div className="grid grid-cols-2 gap-3 md:grid-cols-4">
                          {[
                            { label: 'Time', icon: <Clock3 size={16} className="text-pth-blue" />, status: msWorst },
                            { label: 'Cost', icon: <span className="text-sm font-bold text-pth-blue">$</span>, status: actWorst },
                            { label: 'Quality', icon: <span className="text-sm text-pth-blue">◉</span>, status: 'GREEN' as RygStatus },
                            { label: 'Risk', icon: <ShieldAlert size={16} className="text-pth-blue" />, status: projStatus },
                          ].map((ind) => (
                            <div key={ind.label} className="glass-card rounded-xl p-4 shadow-card">
                              <div className="flex items-center gap-2">
                                {ind.icon}
                                <span className="text-sm font-semibold">{ind.label}</span>
                              </div>
                              <div className="mt-2">{statusBadge(ind.status)}</div>
                            </div>
                          ))}
                        </div>

                        {/* ── Project Information + Stakeholders ── */}
                        <div className="grid gap-4 md:grid-cols-3">
                          <div className="md:col-span-2 glass-card rounded-xl p-5 shadow-card">
                            <h2 className="text-base font-semibold tracking-tight mb-4">Project Information</h2>
                            <div className="space-y-4">
                              <div>
                                <span className="text-xs font-medium text-pth-blue">Objective</span>
                                <p className="mt-0.5 text-sm leading-relaxed">{selectedProject.objective}</p>
                              </div>
                              <div className="grid grid-cols-2 gap-x-6 gap-y-3">
                                <div>
                                  <span className="text-xs font-medium text-pth-blue">Executive Sponsor</span>
                                  <p className="mt-0.5 text-sm font-medium">{sponsor?.name ?? '—'}</p>
                                </div>
                                <div>
                                  <span className="text-xs font-medium text-pth-blue">Project Responsible</span>
                                  <p className="mt-0.5 text-sm font-medium">{pm?.name ?? '—'}</p>
                                </div>
                                <div>
                                  <span className="text-xs font-medium text-pth-blue">Planned Start</span>
                                  <p className="mt-0.5 text-sm font-medium">{plannedStart ? formatDate(plannedStart) : '—'}</p>
                                </div>
                                <div>
                                  <span className="text-xs font-medium text-pth-blue">Planned End</span>
                                  <p className="mt-0.5 text-sm font-medium">{plannedEnd ? formatDate(plannedEnd) : '—'}</p>
                                </div>
                                <div>
                                  <span className="text-xs font-medium text-pth-blue">Real Start</span>
                                  <p className="mt-0.5 text-sm font-medium">{realStart ? formatDate(realStart) : '—'}</p>
                                </div>
                              </div>
                            </div>
                          </div>
                          <div className="glass-card rounded-xl p-5 shadow-card">
                            <h2 className="text-base font-semibold tracking-tight mb-4">Stakeholders</h2>
                            <div className="space-y-4">
                              <div>
                                <div className="flex items-center gap-1.5 text-xs font-medium text-pth-muted"><Users size={12} /> Customers</div>
                                <p className="mt-0.5 text-sm">{projClients}</p>
                              </div>
                              <div>
                                <div className="flex items-center gap-1.5 text-xs font-medium text-pth-muted"><Briefcase size={12} /> Suppliers</div>
                                <p className="mt-0.5 text-sm">{projSuppliers}</p>
                              </div>
                              <div>
                                <div className="flex items-center gap-1.5 text-xs font-medium text-pth-muted"><span className="text-xs font-bold">$</span> Budget</div>
                                <p className="mt-0.5 text-sm font-medium">—</p>
                              </div>
                            </div>
                          </div>
                        </div>

                        {/* ── Project Timeline ── */}
                        <div className="glass-card rounded-xl p-6 shadow-card">
                          <div className="flex items-center gap-3 mb-8">
                            <h2 className="text-base font-semibold tracking-tight">Project Timeline</h2>
                            <span className="rounded-md bg-pth-blue px-2.5 py-0.5 text-xs font-bold text-white shadow-sm">Today</span>
                          </div>
                          <div className="relative mx-4">
                            <div className="h-2 w-full rounded-full bg-pth-border/20" />
                            <div ref={dynRef({ width: `${todayPct}%` })} className="absolute left-0 top-0 h-2 rounded-full bg-gradient-to-r from-pth-blue to-pth-blue/60" />
                            <div ref={dynRef({ left: `${todayPct}%` })} className="absolute top-1/2 -translate-x-1/2 -translate-y-1/2 h-3.5 w-3.5 rounded-full border-2 border-white bg-pth-blue shadow-md dark:border-pth-bg" />
                            <div className="relative mt-5">
                              {pMss.map((ms) => {
                                const pos = Math.min(Math.max((dayDiff(tlMin, ms.targetDate) / tlRange) * 100, 0), 100)
                                const isDone = !!ms.doneDate
                                return (
                                  <div key={ms.id} ref={dynRef({ left: `${pos}%` })} className="absolute top-0 -translate-x-1/2">
                                    <div className="flex flex-col items-center">
                                      <div className="h-4 w-px bg-pth-border/30" />
                                      <div className={`flex h-6 w-6 items-center justify-center rounded-full border-2 shadow-sm ${isDone ? 'border-emerald-500 bg-emerald-50 dark:bg-emerald-900/30' : 'border-pth-border/50 bg-pth-card'}`}>
                                        {isDone && <svg className="h-3 w-3 text-emerald-600" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="3"><path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" /></svg>}
                                      </div>
                                      <span className="mt-1.5 max-w-[80px] truncate text-center text-[11px] font-medium leading-tight">{ms.name}</span>
                                      <span className="text-[10px] text-pth-blue">{formatShortDate(ms.targetDate)}</span>
                                    </div>
                                  </div>
                                )
                              })}
                            </div>
                          </div>
                          <div className="mt-24 flex items-center justify-between text-xs text-pth-muted">
                            <span>{formatDate(tlMin)}</span>
                            <span>{formatDate(tlMax)}</span>
                          </div>

                          {/* ── Task bars (each dated task across the timeline) ── */}
                          {(() => {
                            const dated = pActs.filter((a) => a.startDate && a.endDate).sort((a, b) => (a.startDate < b.startDate ? -1 : 1))
                            if (dated.length === 0) return null
                            const colorFor = (a: Activity) =>
                              a.state === 'DONE' ? 'bg-emerald-500'
                                : a.state === 'IN_PROGRESS' ? 'bg-pth-blue'
                                : rygForActivity(a) === 'RED' ? 'bg-pth-red'
                                : rygForActivity(a) === 'YELLOW' ? 'bg-amber-500'
                                : 'bg-pth-border'
                            return (
                              <div className="mt-6 border-t border-pth-border/15 pt-4">
                                <div className="mb-2 flex items-center justify-between">
                                  <h3 className="text-xs font-semibold uppercase tracking-wider text-pth-muted">Tasks ({dated.length})</h3>
                                  <div className="flex items-center gap-3 text-[10px] text-pth-muted">
                                    <span className="flex items-center gap-1"><span className="h-2 w-2 rounded-full bg-emerald-500" />Done</span>
                                    <span className="flex items-center gap-1"><span className="h-2 w-2 rounded-full bg-pth-blue" />In progress</span>
                                    <span className="flex items-center gap-1"><span className="h-2 w-2 rounded-full bg-amber-500" />At risk</span>
                                    <span className="flex items-center gap-1"><span className="h-2 w-2 rounded-full bg-pth-red" />Late</span>
                                  </div>
                                </div>
                                <div className="max-h-[340px] space-y-1 overflow-y-auto pr-1">
                                  {dated.map((a) => {
                                    const left = Math.min(Math.max((dayDiff(tlMin, a.startDate) / tlRange) * 100, 0), 100)
                                    const rawW = (dayDiff(a.startDate, a.endDate) / tlRange) * 100
                                    const width = Math.min(Math.max(rawW, 1.5), 100 - left)
                                    return (
                                      <div key={a.id} className="group flex items-center gap-2">
                                        <span className="w-40 shrink-0 truncate text-[11px] text-pth-text" title={a.name}>{a.name}</span>
                                        <div className="relative h-4 flex-1 rounded bg-pth-border/10">
                                          <div ref={dynRef({ left: `${left}%`, width: `${width}%` })} className={`absolute top-0 h-4 rounded ${colorFor(a)} opacity-80`} title={`${formatShortDate(a.startDate)} – ${formatShortDate(a.endDate)}`} />
                                        </div>
                                        <span className="w-28 shrink-0 text-right text-[10px] text-pth-muted">{formatShortDate(a.startDate)}–{formatShortDate(a.endDate)}</span>
                                      </div>
                                    )
                                  })}
                                </div>
                              </div>
                            )
                          })()}
                        </div>

                        {/* ── Milestones + Available Resources ── */}
                        <div className="grid gap-6 lg:grid-cols-[1fr_280px]">
                          {/* Left: Milestones */}
                          <div>
                            <div className="flex items-center justify-between mb-5">
                              <h2 className="text-base font-semibold tracking-tight">Milestones</h2>
                              <button type="button" className="inline-flex items-center gap-1.5 rounded-lg bg-pth-btn px-4 py-2 text-sm font-semibold text-white shadow-sm transition-all hover:bg-pth-btn-hover active:scale-[0.98]" onClick={openCreateMilestone}>
                                <Plus size={14} /> Add Milestone
                              </button>
                            </div>
                            <div className="space-y-0">
                              {pMss.map((ms, msIdx) => {
                                const msRyg = rygFromDueDate(ms.targetDate, state.settings.warningDaysThreshold)
                                const isDone = !!ms.doneDate
                                const msActs = pActs.filter((a) => a.milestoneId === ms.id)
                                const isLast = msIdx === pMss.length - 1
                                const ownerCounts: Record<string, number> = {}
                                msActs.forEach((a) => { ownerCounts[a.ownerId] = (ownerCounts[a.ownerId] || 0) + 1 })
                                const topOwnerId = Object.entries(ownerCounts).sort((a, b) => b[1] - a[1])[0]?.[0]
                                const msOwner = topOwnerId ? people.find((p) => p.id === topOwnerId) : pm
                                const isMsDropTarget = dropTargetId === `ms_${ms.id}`
                                return (
                                  <div
                                    key={ms.id}
                                    className={`relative pl-10 rounded-lg transition-all ${isMsDropTarget ? 'bg-pth-blue/5 ring-2 ring-pth-blue/30 ring-inset' : ''}`}
                                    onDragOver={(e) => { e.preventDefault(); setDropTargetId(`ms_${ms.id}`) }}
                                    onDragLeave={() => setDropTargetId(null)}
                                    onDrop={(e) => { e.preventDefault(); handleDropOnMilestone(ms.id) }}
                                  >
                                    {!isLast && <div className="absolute left-[11px] top-6 bottom-0 w-0.5 bg-pth-border/20" />}
                                    <div className={`absolute left-0 top-0.5 flex h-6 w-6 items-center justify-center rounded-full border-2 shadow-sm ${isDone ? 'border-emerald-500 bg-emerald-50 dark:bg-emerald-900/30' : msRyg === 'RED' ? 'border-pth-red bg-pth-red/5' : 'border-pth-blue/40 bg-pth-blue/5 dark:border-pth-blue/50 dark:bg-pth-blue/10'}`}>
                                      {isDone && <svg className="h-3 w-3 text-emerald-600" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="3"><path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" /></svg>}
                                    </div>
                                    <div className="pb-8">
                                      {/* Milestone header with name, badges, and actions */}
                                      <div className="flex items-start justify-between gap-2">
                                        <div className="flex items-center gap-2 flex-wrap">
                                          <h3 className="text-sm font-bold">{ms.name}</h3>
                                          {isDone ? <span className="rounded-md bg-pth-subtle px-2 py-0.5 text-[11px] font-semibold uppercase tracking-wider text-pth-muted">Completed</span> : statusBadge(msRyg)}
                                          {msActs.some((a) => a.criticalPath && !a.doneDate && rygFromDueDate(a.endDate, state.settings.warningDaysThreshold) === 'RED') && (
                                            <span className="rounded-md bg-pth-red/10 px-2 py-0.5 text-[11px] font-semibold text-pth-red">Critical</span>
                                          )}
                                          {isMsDropTarget && <span className="rounded-md bg-pth-blue/10 px-2 py-0.5 text-[11px] font-semibold text-pth-blue animate-pulse">Drop to add task</span>}
                                        </div>
                                        <div className="flex items-center gap-1 shrink-0">
                                          <button type="button" className="rounded-lg p-1.5 text-pth-muted transition-colors hover:bg-pth-hover hover:text-pth-text" onClick={() => openEditMilestone(ms)} title="Edit milestone">
                                            <Pencil size={13} />
                                          </button>
                                          <button type="button" className="rounded-lg p-1.5 text-pth-muted transition-colors hover:bg-red-50 hover:text-pth-red dark:hover:bg-pth-red/10" onClick={() => deleteMilestone(ms.id)} title="Delete milestone">
                                            <Trash2 size={13} />
                                          </button>
                                        </div>
                                      </div>
                                      <div className="mt-1 flex items-center gap-3 text-xs text-pth-muted">
                                        <span className="flex items-center gap-1"><CalendarClock size={12} /> {formatDate(ms.targetDate)}</span>
                                        {msOwner && <span>{msOwner.name}</span>}
                                      </div>
                                      {msActs.length > 0 && (
                                        <div className="mt-4 ml-2">
                                          <div className="flex items-center justify-between mb-2">
                                            <span className="text-xs font-semibold text-pth-muted">Tasks</span>
                                            <span className="text-[11px] text-pth-muted">{msActs.filter((a) => a.state === 'DONE').length}/{msActs.length} complete</span>
                                          </div>
                                          <div className="space-y-0.5">
                                            {msActs.map((act) => {
                                              const actOwner = people.find((p) => p.id === act.ownerId)
                                              const hours = Math.max(dayDiff(act.startDate, act.endDate), 1) * 8
                                              const isTaskDropTarget = dropTargetId === `act_${act.id}`
                                              return (
                                                <div
                                                  key={act.id}
                                                  className={`group/task flex items-center justify-between rounded-lg px-3 py-2 transition-all ${isTaskDropTarget ? 'bg-pth-blue/10 ring-2 ring-pth-blue/30 ring-inset' : 'hover:bg-pth-subtle'}`}
                                                  onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); setDropTargetId(`act_${act.id}`) }}
                                                  onDragLeave={(e) => { e.stopPropagation(); setDropTargetId(null) }}
                                                  onDrop={(e) => { e.preventDefault(); e.stopPropagation(); handleDropOnTask(act.id) }}
                                                >
                                                  <div className="flex items-center gap-2.5 min-w-0">
                                                    <span className={`h-2 w-2 shrink-0 rounded-full ${act.state === 'DONE' ? 'bg-emerald-500' : act.state === 'IN_PROGRESS' ? 'bg-pth-blue' : 'bg-pth-border'}`} />
                                                    <div className="min-w-0">
                                                      <div className="text-sm font-medium truncate">{act.name}</div>
                                                      <div className="text-[11px] text-pth-muted">
                                                        {isTaskDropTarget ? <span className="text-pth-blue font-medium animate-pulse">Reassign to dropped person</span> : <>{actOwner?.name ?? 'Unknown'} &middot; {formatShortDate(act.startDate)} &middot; {hours}h</>}
                                                      </div>
                                                    </div>
                                                  </div>
                                                  <div className="flex items-center gap-1.5">
                                                    <button type="button" className="rounded-lg p-1 text-pth-muted opacity-0 transition-all group-hover/task:opacity-100 hover:bg-pth-hover hover:text-pth-text" onClick={() => openEditActivity(act)} title="Edit task">
                                                      <Pencil size={12} />
                                                    </button>
                                                    {activityStateBadge(act.state)}
                                                  </div>
                                                </div>
                                              )
                                            })}
                                          </div>
                                        </div>
                                      )}
                                      <div className="mt-3">
                                        <button type="button" className="inline-flex items-center gap-1 rounded-lg bg-pth-blue/10 px-2.5 py-1 text-xs font-medium text-pth-blue transition-colors hover:bg-pth-blue/20" onClick={() => openCreateActivity(ms.id)}>
                                          <Plus size={12} /> Add Task
                                        </button>
                                      </div>
                                    </div>
                                  </div>
                                )
                              })}
                            </div>
                            {(() => {
                              const allTasks = pActs.filter((a) => !a.milestoneId)
                              const naCount = allTasks.filter((a) => a.state === 'NOT_APPLICABLE').length
                              const activeCount = allTasks.length - naCount
                              const shownTasks = allTasks.filter((a) =>
                                taskNaFilter === 'all' ? true
                                  : taskNaFilter === 'na' ? a.state === 'NOT_APPLICABLE'
                                  : a.state !== 'NOT_APPLICABLE',
                              )
                              return (
                              <div className="mt-4 rounded-xl border border-pth-border/20 bg-pth-subtle/30 p-4">
                                <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
                                  <h3 className="text-sm font-semibold text-pth-muted">Tasks ({shownTasks.length})</h3>
                                  <div className="flex items-center gap-2">
                                    <div className="flex items-center rounded-lg border border-pth-border/40 bg-pth-card p-0.5 text-xs">
                                      {([['active', `To assign (${activeCount})`], ['na', `N/A (${naCount})`], ['all', 'All']] as const).map(([k, lbl]) => (
                                        <button key={k} type="button" onClick={() => setTaskNaFilter(k)} className={`rounded-md px-2.5 py-1 font-medium transition-colors ${taskNaFilter === k ? 'bg-pth-blue/10 text-pth-blue' : 'text-pth-muted hover:text-pth-text'}`}>{lbl}</button>
                                      ))}
                                    </div>
                                    <button type="button" onClick={() => openCreateActivity(null)} className="inline-flex items-center gap-1 rounded-lg bg-pth-btn px-2.5 py-1.5 text-xs font-medium text-white shadow-sm transition-colors hover:bg-pth-btn-hover">
                                      <Plus size={13} /> Add Task
                                    </button>
                                  </div>
                                </div>
                                <div className="space-y-0.5">
                                  {shownTasks.map((act) => {
                                    const actOwner = people.find((p) => p.id === act.ownerId)
                                    const isTaskDropTarget = dropTargetId === `act_${act.id}`
                                    return (
                                      <div
                                        key={act.id}
                                        className={`group/task flex items-center justify-between rounded-lg px-3 py-2 transition-all ${act.state === 'NOT_APPLICABLE' ? 'opacity-45' : ''} ${isTaskDropTarget ? 'bg-pth-blue/10 ring-2 ring-pth-blue/30 ring-inset' : 'hover:bg-pth-subtle'}`}
                                        onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); setDropTargetId(`act_${act.id}`) }}
                                        onDragLeave={(e) => { e.stopPropagation(); setDropTargetId(null) }}
                                        onDrop={(e) => { e.preventDefault(); e.stopPropagation(); handleDropOnTask(act.id) }}
                                      >
                                        <div className="flex items-center gap-2.5 min-w-0">
                                          <span className={`h-2 w-2 shrink-0 rounded-full ${act.state === 'DONE' ? 'bg-emerald-500' : act.state === 'IN_PROGRESS' ? 'bg-pth-blue' : 'bg-pth-border'}`} />
                                          <div className="min-w-0">
                                            <div className={`text-sm font-medium truncate ${act.state === 'NOT_APPLICABLE' ? 'line-through' : ''}`}>{act.name}</div>
                                            <div className="text-[11px] text-pth-muted">
                                              {act.state === 'NOT_APPLICABLE' ? 'Not applicable for this project' : isTaskDropTarget ? <span className="text-pth-blue font-medium animate-pulse">Reassign to dropped person</span> : <>{actOwner?.name ?? 'Unknown'}{act.startDate ? <> &middot; {formatShortDate(act.startDate)}</> : ''}</>}
                                            </div>
                                          </div>
                                        </div>
                                        <div className="flex items-center gap-1.5">
                                          <button type="button" className="rounded-lg p-1 text-pth-muted opacity-0 transition-all group-hover/task:opacity-100 hover:bg-pth-hover hover:text-pth-text" onClick={() => openEditActivity(act)} title="Edit task">
                                            <Pencil size={12} />
                                          </button>
                                          {activityStateBadge(act.state)}
                                        </div>
                                      </div>
                                    )
                                  })}
                                </div>
                              </div>
                              )
                            })()}
                          </div>

                          {/* Right: Available Resources (sticky sidebar) */}
                          <div className="hidden lg:block">
                            <div className="sticky top-4">
                              <div className="glass-card rounded-xl border border-pth-border/20 shadow-card">
                                <div className="border-b border-pth-border/15 px-4 py-3">
                                  <h3 className="text-sm font-bold">Available Resources</h3>
                                  <p className="mt-0.5 text-[11px] text-pth-muted">Drag a person onto a milestone or task</p>
                                  <div className="relative mt-2">
                                    <Search size={13} className="pointer-events-none absolute left-2.5 top-1/2 -translate-y-1/2 text-pth-muted" />
                                    <input
                                      type="text"
                                      placeholder="Search name or area…"
                                      value={resourcePaneSearch}
                                      onChange={(e) => setResourcePaneSearch(e.target.value)}
                                      className="h-8 w-full rounded-lg border border-pth-border/40 bg-pth-subtle py-1 pl-8 pr-7 text-xs outline-none placeholder:text-pth-muted/60 focus:border-pth-blue focus:bg-pth-card"
                                    />
                                    {resourcePaneSearch && (
                                      <button type="button" title="Clear" onClick={() => setResourcePaneSearch('')} className="absolute right-2 top-1/2 -translate-y-1/2 text-pth-muted hover:text-pth-text"><X size={12} /></button>
                                    )}
                                  </div>
                                  {/* Area filter chips */}
                                  {(() => {
                                    const areas = Array.from(new Set(workloadRows.map((r) => r.person.area).filter(Boolean) as string[])).sort()
                                    if (areas.length === 0) return null
                                    return (
                                      <div className="mt-2 flex flex-wrap gap-1">
                                        {['all', ...areas].map((ar) => (
                                          <button key={ar} type="button" onClick={() => setResourcePaneArea(ar)}
                                            className={`rounded-full px-2 py-0.5 text-[10px] font-medium transition-colors ${resourcePaneArea === ar ? 'bg-pth-blue text-white' : 'bg-pth-subtle text-pth-muted hover:text-pth-text'}`}>
                                            {ar === 'all' ? 'All' : ar}
                                          </button>
                                        ))}
                                      </div>
                                    )
                                  })()}
                                </div>
                                <div className="max-h-[60vh] overflow-y-auto p-2 space-y-1">
                                  {(() => {
                                    const rq = resourcePaneSearch.trim().toLowerCase()
                                    const sorted = [...workloadRows]
                                      .filter((r) => resourcePaneArea === 'all' || (r.person.area ?? '') === resourcePaneArea)
                                      .filter((r) => !rq || r.person.name.toLowerCase().includes(rq) || (r.person.area ?? '').toLowerCase().includes(rq) || r.person.initials.toLowerCase().includes(rq))
                                      .sort((a, b) => (peakUtilByPerson.get(a.person.id)?.peakUtil ?? 0) - (peakUtilByPerson.get(b.person.id)?.peakUtil ?? 0))
                                    if (sorted.length === 0) return <div className="px-3 py-6 text-center text-[11px] text-pth-muted">No resources match{resourcePaneArea !== 'all' ? ` ${resourcePaneArea}` : ''}{resourcePaneSearch ? ` “${resourcePaneSearch}”` : ''}.</div>
                                    return sorted.map((row) => {
                                      const peak = peakUtilByPerson.get(row.person.id) ?? { peakUtil: 0, peakHours: 0, capacity: row.capacityHours, taskCount: 0 }
                                      const utilization = peak.peakUtil
                                      const available = Math.max(peak.capacity - peak.peakHours, 0)
                                      const critical = utilization > state.settings.workloadCriticalThreshold
                                      const warning = utilization >= state.settings.workloadWarningThreshold && utilization <= state.settings.workloadCriticalThreshold
                                      return (
                                        <div
                                          key={row.person.id}
                                          draggable
                                          onDragStart={(e) => { e.dataTransfer.effectAllowed = 'copy'; setDragPersonId(row.person.id) }}
                                          onDragEnd={() => { setDragPersonId(null); setDropTargetId(null) }}
                                          className={`flex items-center gap-2.5 rounded-lg px-3 py-2.5 cursor-grab active:cursor-grabbing select-none transition-all hover:bg-pth-hover ${dragPersonId === row.person.id ? 'opacity-50 ring-2 ring-pth-blue/30' : ''}`}
                                        >
                                          <GripVertical size={12} className="shrink-0 text-pth-border" />
                                          <div className="flex h-7 w-7 shrink-0 items-center justify-center rounded-full bg-pth-subtle text-[10px] font-bold">{row.person.initials}</div>
                                          <div className="min-w-0 flex-1">
                                            <div className="text-xs font-medium truncate">{row.person.name}</div>
                                            <div className="flex items-center gap-1.5 mt-0.5">
                                              <div className="h-1 flex-1 rounded-full bg-pth-border/20 overflow-hidden">
                                                <div ref={dynRef({ width: `${Math.min(utilization, 100)}%` })} className={`h-full rounded-full ${critical ? 'bg-pth-red' : warning ? 'bg-amber-500' : 'bg-emerald-500'}`} />
                                              </div>
                                              <span className={`text-[10px] font-semibold ${critical ? 'text-pth-red' : warning ? 'text-amber-600' : 'text-emerald-600'}`}>{utilization}%</span>
                                            </div>
                                            <div className="text-[10px] text-pth-muted mt-0.5">{peak.peakHours.toFixed(0)}h peak · {available}h free{peak.taskCount ? ` · ${peak.taskCount} tasks` : ''}</div>
                                          </div>
                                        </div>
                                      )
                                    })
                                  })()}
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                      </motion.div>
                    )
                  })()}
                </>
              )}

              {navPage === 'workload' && (() => {
                const overloaded = workloadRows.filter((r) => r.utilization > 100).length
                const atRisk = workloadRows.filter((r) => r.utilization >= 90 && r.utilization <= 100).length
                const moderate = workloadRows.filter((r) => r.utilization >= 50 && r.utilization < 90).length
                const low = workloadRows.filter((r) => r.utilization < 50).length

                // compute week end date for display
                const weekEnd = addDays(currentWeek, 6)

                // filter by search query + site + area
                const wq = workloadSearch.trim().toLowerCase()
                const wlAreaList = Array.from(new Set(workloadRows.map((r) => r.person.area).filter(Boolean) as string[])).sort()
                const filtered = workloadRows.filter((r) => {
                  if (workloadSiteFilter !== 'all' && !(r.person.location ?? []).includes(workloadSiteFilter)) return false
                  if (workloadAreaFilter !== 'all' && (r.person.area ?? '') !== workloadAreaFilter) return false
                  if (wq && !(
                    r.person.name.toLowerCase().includes(wq) ||
                    r.person.initials.toLowerCase().includes(wq) ||
                    (r.person.area ?? '').toLowerCase().includes(wq) ||
                    (ROLE_LABEL[r.person.role as Role] ?? '').toLowerCase().includes(wq)
                  )) return false
                  return true
                })

                // sort: overloaded first, then at-risk, then by descending util
                const sorted = [...filtered].sort((a, b) => b.utilization - a.utilization)

                return (
                <div className="space-y-4">
                  {/* ── Page header ── */}
                  <div className="flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
                    <div>
                      <h1 className="text-xl font-bold tracking-tight md:text-2xl">Workload Overview</h1>
                      <p className="mt-0.5 text-xs text-pth-muted">Team allocation and capacity analysis</p>
                    </div>
                    <div className="flex items-center gap-2">
                      <button type="button" title="Previous week" className="inline-flex h-9 w-9 items-center justify-center rounded-lg text-pth-muted transition-colors hover:bg-pth-hover hover:text-pth-text" onClick={() => setWorkloadWeekOffset((v) => v - 1)}>
                        <ChevronLeft size={16} />
                      </button>
                      <div className="rounded-lg bg-pth-subtle px-4 py-2 text-xs font-medium">{formatShortDate(currentWeek)} – {formatShortDate(weekEnd)}</div>
                      <button type="button" title="Next week" className="inline-flex h-9 w-9 items-center justify-center rounded-lg text-pth-muted transition-colors hover:bg-pth-hover hover:text-pth-text" onClick={() => setWorkloadWeekOffset((v) => v + 1)}>
                        <ChevronRight size={16} />
                      </button>
                    </div>
                  </div>

                  {/* ── Search + filters + view toggle ── */}
                  <div className="flex flex-wrap items-center gap-2">
                    <div className="relative min-w-[200px] flex-1">
                      <Search size={15} className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-pth-muted" />
                      <input
                        type="text"
                        placeholder="Search team members..."
                        value={workloadSearch}
                        onChange={(e) => setWorkloadSearch(e.target.value)}
                        className="w-full rounded-lg border border-pth-border/40 bg-pth-card py-2 pl-9 pr-3 text-sm outline-none placeholder:text-pth-muted/60 focus:border-pth-blue"
                      />
                    </div>
                    <select title="Site" value={workloadSiteFilter} onChange={(e) => setWorkloadSiteFilter(e.target.value as 'all' | 'SlpP' | 'TlP')} className="h-9 rounded-lg border border-pth-border/40 bg-pth-card px-3 text-sm outline-none focus:border-pth-blue">
                      <option value="all">All sites</option>
                      <option value="SlpP">SlpP</option>
                      <option value="TlP">TlP</option>
                    </select>
                    <select title="Area" value={workloadAreaFilter} onChange={(e) => setWorkloadAreaFilter(e.target.value)} className="h-9 rounded-lg border border-pth-border/40 bg-pth-card px-3 text-sm outline-none focus:border-pth-blue">
                      <option value="all">All areas</option>
                      {wlAreaList.map((a) => <option key={a} value={a}>{a}</option>)}
                    </select>
                    <div className="flex items-center rounded-lg border border-pth-border/40 bg-pth-card p-0.5">
                      <button type="button" title="Card view" onClick={() => setWorkloadView('cards')} className={`inline-flex h-8 w-8 items-center justify-center rounded-md transition-colors ${workloadView === 'cards' ? 'bg-pth-blue/10 text-pth-blue' : 'text-pth-muted hover:text-pth-text'}`}>
                        <LayoutGrid size={15} />
                      </button>
                      <button type="button" title="List view" onClick={() => setWorkloadView('list')} className={`inline-flex h-8 w-8 items-center justify-center rounded-md transition-colors ${workloadView === 'list' ? 'bg-pth-blue/10 text-pth-blue' : 'text-pth-muted hover:text-pth-text'}`}>
                        <List size={15} />
                      </button>
                    </div>
                  </div>

                  {/* ── KPI Cards ── */}
                  <div className="grid grid-cols-2 gap-3 md:grid-cols-4">
                    {[
                      { label: 'Overloaded (>100%)', value: overloaded, icon: <ShieldAlert size={18} className="text-pth-red" />, ring: '' },
                      { label: 'At Risk (90–100%)', value: atRisk, icon: <Users size={18} className="text-amber-500" />, ring: '' },
                      { label: 'Moderate (50–90%)', value: moderate, icon: <Users size={18} className="text-blue-500" />, ring: '' },
                      { label: 'Low (<50%)', value: low, icon: <Users size={18} className="text-emerald-500" />, ring: '' },
                    ].map((kpi) => (
                      <div key={kpi.label} className={`glass-card rounded-xl p-4 shadow-card ${kpi.ring}`}>
                        <div className="flex items-center justify-between">
                          <span className="text-[11px] font-semibold uppercase tracking-wider text-pth-muted">{kpi.label}</span>
                          {kpi.icon}
                        </div>
                        <div className="mt-2 text-2xl font-bold">{kpi.value}</div>
                      </div>
                    ))}
                  </div>

                  {/* ── List view ── */}
                  {workloadView === 'list' && (
                    <div className="glass-card overflow-hidden rounded-xl border border-pth-border/20 shadow-card">
                      <div className="overflow-x-auto">
                        <table className="w-full text-sm">
                          <thead>
                            <tr className="border-b border-pth-border/15 text-left text-[11px] uppercase tracking-wide text-pth-muted">
                              <th className="px-4 py-2.5 font-semibold">Member</th>
                              <th className="px-4 py-2.5 font-semibold">Area</th>
                              <th className="px-4 py-2.5 font-semibold">Location</th>
                              <th className="px-4 py-2.5 font-semibold">Manager</th>
                              <th className="px-4 py-2.5 font-semibold whitespace-nowrap">Hours</th>
                              <th className="px-4 py-2.5 font-semibold">Utilization</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-pth-border/10">
                            {sorted.map((row) => {
                              const lcritical = row.utilization > 100
                              const lwarning = row.utilization >= 90 && row.utilization <= 100
                              const lbar = lcritical ? 'bg-pth-red' : lwarning ? 'bg-amber-400' : row.utilization >= 50 ? 'bg-blue-500' : 'bg-emerald-500'
                              const ltext = lcritical ? 'text-pth-red' : lwarning ? 'text-amber-600 dark:text-amber-400' : row.utilization >= 50 ? 'text-blue-600 dark:text-blue-400' : 'text-emerald-600 dark:text-emerald-400'
                              const mgrName = row.person.managerId ? (people.find((m) => m.id === row.person.managerId)?.name ?? '—') : '—'
                              return (
                                <tr key={row.person.id} className="transition-colors hover:bg-pth-hover/50">
                                  <td className="px-4 py-2.5">
                                    <div className="flex items-center gap-2.5">
                                      <div className="flex h-7 w-7 shrink-0 items-center justify-center rounded-full bg-pth-subtle text-[10px] font-bold">{row.person.initials}</div>
                                      <div className="min-w-0">
                                        <div className="truncate font-medium">{row.person.name}</div>
                                        <div className="text-[10px] text-pth-muted">{row.person.resourceKind ?? (ROLE_LABEL[row.person.role as Role] ?? 'Resource')}</div>
                                      </div>
                                    </div>
                                  </td>
                                  <td className="px-4 py-2.5 text-pth-muted">{row.person.area || '—'}</td>
                                  <td className="px-4 py-2.5">
                                    <div className="flex gap-1">
                                      {(row.person.location ?? []).map((loc) => <span key={loc} className="rounded-full bg-pth-blue/10 px-1.5 py-0.5 text-[10px] font-medium text-pth-blue">{loc}</span>)}
                                      {(row.person.location?.length ?? 0) === 0 && <span className="text-pth-muted">—</span>}
                                    </div>
                                  </td>
                                  <td className="max-w-[160px] truncate px-4 py-2.5 text-pth-muted">{mgrName}</td>
                                  <td className="whitespace-nowrap px-4 py-2.5 text-pth-muted">{row.assignedHours}.0 / {row.capacityHours}</td>
                                  <td className="px-4 py-2.5">
                                    <div className="flex items-center gap-2">
                                      <div className="h-1.5 w-20 overflow-hidden rounded-full bg-pth-border/20">
                                        <div ref={dynRef({ width: `${Math.min(row.utilization, 100)}%` })} className={`h-full rounded-full ${lbar}`} />
                                      </div>
                                      <span className={`min-w-[2.5rem] text-right text-xs font-semibold ${ltext}`}>{row.utilization}%</span>
                                    </div>
                                  </td>
                                </tr>
                              )
                            })}
                            {sorted.length === 0 && (
                              <tr><td colSpan={6} className="px-4 py-8 text-center text-sm text-pth-muted">No team members match the filters.</td></tr>
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  {/* ── Person Cards Grid ── */}
                  {workloadView === 'cards' && (
                  <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-3">
                    {sorted.map((row) => {
                      const critical = row.utilization > 100
                      const warning = row.utilization >= 90 && row.utilization <= 100
                      const barColor = critical ? 'bg-pth-red' : warning ? 'bg-amber-400' : row.utilization >= 50 ? 'bg-blue-500' : 'bg-emerald-500'
                      const textColor = critical ? 'text-pth-red' : warning ? 'text-amber-600 dark:text-amber-400' : row.utilization >= 50 ? 'text-blue-600 dark:text-blue-400' : 'text-emerald-600 dark:text-emerald-400'
                      const ringClass = ''
                      const roleLabel = ROLE_LABEL[row.person.role as Role] ?? 'Team Member'

                      // Group assignments by project for breakdown
                      const projectBreakdown: Array<{name: string; hours: number}> = []
                      const projectHoursMap = new Map<string, number>()
                      row.personAssignments.forEach((asg) => {
                        const act = state.activities.find((a) => a.id === asg.activityId)
                        if (!act) return
                        const proj = state.projects.find((p) => p.id === act.projectId)
                        if (!proj) return
                        projectHoursMap.set(proj.name, (projectHoursMap.get(proj.name) ?? 0) + asg.assignedHours)
                      })
                      projectHoursMap.forEach((hours, name) => projectBreakdown.push({ name, hours }))
                      projectBreakdown.sort((a, b) => b.hours - a.hours)
                      const visibleProjects = projectBreakdown.slice(0, 3)
                      const remaining = projectBreakdown.length - 3

                      return (
                        <div key={row.person.id} className={`glass-card rounded-xl p-5 shadow-card ${ringClass}`}>
                          {/* Header: avatar + name + utilization */}
                          <div className="flex items-center justify-between">
                            <div className="flex items-center gap-3">
                              <div className={`flex h-10 w-10 items-center justify-center rounded-full text-sm font-bold text-white ${critical ? 'bg-pth-red' : warning ? 'bg-amber-500' : row.utilization >= 50 ? 'bg-blue-500' : 'bg-emerald-500'}`}>{row.person.initials}</div>
                              <div>
                                <div className="text-sm font-semibold">{row.person.name}</div>
                                <div className="text-[11px] text-pth-muted">
                                  {row.person.resourceKind ?? roleLabel}{row.person.area ? ` · ${row.person.area}` : ''}
                                </div>
                                {(row.person.location?.length ?? 0) > 0 && (
                                  <div className="mt-1 flex gap-1">
                                    {row.person.location!.map((loc) => (
                                      <span key={loc} className="rounded-full bg-pth-blue/10 px-1.5 py-0.5 text-[9px] font-medium text-pth-blue">{loc}</span>
                                    ))}
                                  </div>
                                )}
                              </div>
                            </div>
                            <span className={`text-xl font-bold ${textColor}`}>{row.utilization}%</span>
                          </div>

                          {/* Workload bar */}
                          <div className="mt-4">
                            <div className="flex items-center justify-between text-[11px] text-pth-muted mb-1.5">
                              <span>Workload</span>
                              <span>{row.assignedHours}.0 / {row.capacityHours} hrs</span>
                            </div>
                            <div className="h-2 w-full overflow-hidden rounded-full bg-pth-border/15">
                              <div ref={dynRef({ width: `${Math.min(row.utilization, 100)}%` })} className={`h-full rounded-full ${barColor}`} />
                            </div>
                          </div>

                          {/* Project breakdown */}
                          {projectBreakdown.length > 0 && (
                            <div className="mt-4">
                              <div className="text-[11px] font-semibold text-pth-muted mb-2">Projects</div>
                              <div className="space-y-1">
                                {visibleProjects.map((p) => (
                                  <div key={p.name} className="flex items-center justify-between text-xs">
                                    <span className="truncate pr-3 text-pth-text">{p.name}</span>
                                    <span className="shrink-0 text-pth-muted">{p.hours}.0h</span>
                                  </div>
                                ))}
                                {remaining > 0 && (
                                  <div className="text-[11px] text-pth-muted">+{remaining} more</div>
                                )}
                              </div>
                            </div>
                          )}
                        </div>
                      )
                    })}
                  </div>
                  )}
                </div>
                )
              })()}

              {navPage === 'reports' && (
              <div className="glass-card rounded-xl border border-pth-border/20 shadow-card">
                  {/* Report tab bar */}
                  <div className="flex items-center gap-1 border-b border-pth-border/15 px-4 pt-3">
                    {([
                      { key: 'projectsStatus', label: 'Projects Status' },
                      { key: 'taskTracking', label: 'Task Tracking' },
                      { key: 'prioritization', label: 'Prioritization & Risks' },
                      { key: 'workloadOverview', label: 'Workload Overview' },
                    ] as Array<{ key: ReportTab; label: string }>).map((tab) => (
                      <button
                        key={tab.key}
                        type="button"
                        className={`relative px-4 py-2.5 text-sm font-medium transition-colors ${
                          reportTab === tab.key
                            ? 'text-pth-blue'
                            : 'text-pth-muted hover:text-pth-text'
                        }`}
                        onClick={() => setReportTab(tab.key)}
                      >
                        {tab.label}
                        {reportTab === tab.key && (
                          <motion.div layoutId="report-tab-underline" className="absolute inset-x-0 -bottom-px h-0.5 rounded-full bg-pth-blue" transition={{ type: 'spring', duration: 0.35, bounce: 0.15 }} />
                        )}
                      </button>
                    ))}
                  </div>

                  <div className="p-5">
                    {/* ── PROJECTS STATUS TAB (merged overview + project deep-dive) ── */}
                    {reportTab === 'projectsStatus' && (
                      <div className="space-y-4">
                        {/* Filter bar */}
                        <div className="flex flex-wrap items-center gap-2">
                          <div className="relative flex-1 min-w-[180px]">
                            <input type="text" placeholder="Search projects…" className="w-full rounded-lg border border-pth-border/40 bg-pth-subtle py-2 pl-9 pr-3 text-sm outline-none placeholder:text-pth-muted/60 focus:border-pth-blue focus:ring-2 focus:ring-pth-blue/20" value={rptSearch} onChange={(e) => setRptSearch(e.target.value)} />
                            <svg className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-pth-muted" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2"><circle cx="11" cy="11" r="8" /><path d="m21 21-4.35-4.35" /></svg>
                          </div>
                          <select title="Category" className="h-9 rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm outline-none focus:border-pth-blue" value={rptCategory} onChange={(e) => setRptCategory(e.target.value)}>
                            <option value="all">All Categories</option>
                            {(['ECR', 'CIP', 'Path Forward', 'New Programs'] as ProjectCategory[]).map((c) => <option key={c} value={c}>{c}</option>)}
                          </select>
                          <select title="Status" className="h-9 rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm outline-none focus:border-pth-blue" value={rptStatus} onChange={(e) => setRptStatus(e.target.value)}>
                            <option value="all">All Status</option>
                            <option value="GREEN">On Track</option>
                            <option value="YELLOW">At Risk</option>
                            <option value="RED">Critical</option>
                          </select>
                          <select title="Manager" className="h-9 rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm outline-none focus:border-pth-blue" value={rptManager} onChange={(e) => setRptManager(e.target.value)}>
                            <option value="all">All Managers</option>
                            {dvConnected && dvUsers.length > 0
                              ? dvUsers.map((u) => <option key={u.id} value={u.id}>{u.fullname}</option>)
                              : people.filter((p) => p.role === 'PROJECT_MANAGER').map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                          </select>
                          <select title="Site" className="h-9 rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm outline-none focus:border-pth-blue" value={rptSite} onChange={(e) => setRptSite(e.target.value)}>
                            <option value="all">All Sites</option>
                            {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                          </select>
                          {(rptSearch || rptCategory !== 'all' || rptStatus !== 'all' || rptManager !== 'all' || rptSite !== 'all') && (
                            <button type="button" className="text-xs text-pth-muted hover:text-pth-text transition-colors" onClick={() => { setRptSearch(''); setRptCategory('all'); setRptStatus('all'); setRptManager('all'); setRptSite('all') }}>✕ Clear filters</button>
                          )}
                        </div>

                        {/* KPI summary cards */}
                        <div className="grid gap-4 sm:grid-cols-2 md:grid-cols-4">
                          {[
                            { label: 'Total Projects', value: rptKpi.total, color: 'text-pth-blue' },
                            { label: 'On Track', value: rptKpi.onTrack, color: 'text-emerald-600' },
                            { label: 'At Risk', value: rptKpi.atRisk, color: 'text-amber-600' },
                            { label: 'Critical', value: rptKpi.critical, color: 'text-pth-red' },
                          ].map((kpi) => (
                            <div key={kpi.label} className="rounded-xl border border-pth-border/15 bg-pth-card p-4 text-center">
                              <div className="text-[11px] font-semibold uppercase tracking-wider text-pth-muted">{kpi.label}</div>
                              <div className={`mt-1 text-2xl font-bold ${kpi.color}`}>{kpi.value}</div>
                            </div>
                          ))}
                        </div>

                        {/* ── Projects Timeline — Gantt Chart ── */}
                        {rptFiltered.length > 0 && (() => {
                          const bars = rptFiltered.map((project) => {
                            const pActs = state.activities.filter((a) => a.projectId === project.id)
                            if (pActs.length === 0) return null
                            // Only dated tasks define the bar span (N/A tasks have no dates).
                            const starts = pActs.map((a) => a.startDate).filter(Boolean).sort()
                            const ends = pActs.map((a) => a.endDate).filter(Boolean).sort()
                            if (starts.length === 0 || ends.length === 0) return null
                            const minStart = starts[0]
                            const maxEnd = ends[ends.length - 1]
                            const doneCount = pActs.filter((a) => a.state === 'DONE').length
                            const doneRatio = doneCount / pActs.length
                            const status = seededProjectStatus(state, project.id)
                            const milestones = state.milestones.filter((m) => m.projectId === project.id)
                            return { project, minStart, maxEnd, doneRatio, status, activities: pActs, milestones }
                          }).filter(Boolean) as Array<{project: Project; minStart: string; maxEnd: string; doneRatio: number; status: RygStatus; activities: Activity[]; milestones: Milestone[]}>

                          if (bars.length === 0) return null

                          // Full data range: union of all bar dates, milestones, and today
                          const allStarts = bars.map((b) => b.minStart)
                          const allEnds = bars.map((b) => b.maxEnd)
                          const allMsDates = bars.flatMap((b) => b.milestones.map((m) => m.targetDate))
                          const allDates = [...allStarts, ...allEnds, ...allMsDates, todayISO()].filter(Boolean).sort()
                          const dataMin = allDates[0] ?? todayISO()
                          const dataMax = allDates[allDates.length - 1] ?? addDays(todayISO(), 30)

                          // Pad by 7 days on each side
                          const globalMin = addDays(dataMin, -7)
                          const globalMax = addDays(dataMax, 7)
                          const globalRange = Math.max(dayDiff(globalMin, globalMax), 1)

                          // Pixel density: scale so the full range is readable
                          const pxPerDay = 10
                          const innerWidth = globalRange * pxPerDay

                          const todayPct = ((dayDiff(globalMin, todayISO()) / globalRange) * 100)
                          const pct = (d: string) => ((dayDiff(globalMin, d) / globalRange) * 100)

                          // Auto-scroll to show the selected horizon window
                          const scrollRef = (el: HTMLDivElement | null) => {
                            if (!el) return
                            const horizonPos = (dayDiff(globalMin, horizonRange.start) / globalRange) * innerWidth
                            const scrollTarget = Math.max(horizonPos - 40, 0)
                            if (Math.abs(el.scrollLeft - scrollTarget) > 50) {
                              el.scrollLeft = scrollTarget
                            }
                          }

                          // Month labels
                          const monthLabels: Array<{label: string; pct: number; isQuarter: boolean}> = []
                          const cursor = new Date(toDate(globalMin))
                          cursor.setDate(1)
                          if (cursor < toDate(globalMin)) cursor.setMonth(cursor.getMonth() + 1)
                          while (cursor <= toDate(globalMax)) {
                            const iso = cursor.toISOString().slice(0, 10)
                            const p = (dayDiff(globalMin, iso) / globalRange) * 100
                            const isQuarter = cursor.getMonth() % 3 === 0
                            monthLabels.push({ label: new Intl.DateTimeFormat('en-US', { month: 'short', year: 'numeric' }).format(cursor), pct: p, isQuarter })
                            cursor.setMonth(cursor.getMonth() + 1)
                          }

                          // Week tick marks for finer granularity
                          const weekTicks: Array<{pct: number}> = []
                          const wCursor = new Date(toDate(globalMin))
                          const dayOfWeek = wCursor.getDay()
                          wCursor.setDate(wCursor.getDate() + ((8 - dayOfWeek) % 7)) // next Monday
                          while (wCursor <= toDate(globalMax)) {
                            const iso = wCursor.toISOString().slice(0, 10)
                            weekTicks.push({ pct: (dayDiff(globalMin, iso) / globalRange) * 100 })
                            wCursor.setDate(wCursor.getDate() + 7)
                          }

                          return (
                            <div className="glass-card rounded-xl p-5 shadow-card">
                              <div className="flex items-center justify-between mb-4">
                                <h2 className="text-xs font-semibold uppercase tracking-wider text-pth-muted">Projects Timeline</h2>
                                <div className="flex items-center gap-4 text-[10px] text-pth-muted">
                                  <span className="flex items-center gap-1"><span className="h-2 w-2 rounded-full bg-emerald-500" /> On Track</span>
                                  <span className="flex items-center gap-1"><span className="h-2 w-2 rounded-full bg-amber-400" /> At Risk</span>
                                  <span className="flex items-center gap-1"><span className="h-2 w-2 rounded-full bg-pth-red" /> Critical</span>
                                  <span className="flex items-center gap-1"><span className="h-2.5 w-2.5 rotate-45 bg-pth-blue" /> Milestone</span>
                                </div>
                              </div>

                              {/* Scrollable Gantt area */}
                              <div className="flex">
                                {/* Resizable left column: project names */}
                                <div ref={dynRef({ width: `${ganttLabelWidth}px` })} className="relative shrink-0 z-10">
                                  <div className="h-6 border-b border-pth-border/15" /> {/* spacer for month header */}
                                  {bars.map((bar) => {
                                    const dotColor = bar.status === 'RED' ? 'bg-pth-red' : bar.status === 'YELLOW' ? 'bg-amber-400' : 'bg-emerald-500'
                                    return (
                                      <div
                                        key={`label-${bar.project.id}`}
                                        className="flex items-center h-7 pr-3 gap-2 cursor-pointer group"
                                        onClick={() => { setDetailProjectId(bar.project.id); setSelectedProjectId(bar.project.id) }}
                                        title={bar.project.name}
                                      >
                                        <div className={`h-2 w-2 shrink-0 rounded-full ${dotColor}`} />
                                        <span className="truncate text-xs font-medium group-hover:text-pth-blue transition-colors">{bar.project.name}</span>
                                        <span className="ml-auto shrink-0 text-[10px] font-semibold tabular-nums text-pth-muted">{Math.round(bar.doneRatio * 100)}%</span>
                                      </div>
                                    )
                                  })}
                                  <div className="h-7 mt-1 pt-2 border-t border-pth-border/10 text-[10px] text-pth-muted">{bars.length} project{bars.length !== 1 ? 's' : ''}</div>
                                  {/* Drag handle to resize the name column */}
                                  <div
                                    role="separator"
                                    title="Drag to resize"
                                    onMouseDown={(e) => {
                                      e.preventDefault()
                                      const startX = e.clientX; const startW = ganttLabelWidth
                                      const onMove = (ev: MouseEvent) => setGanttLabelWidth(Math.min(Math.max(startW + (ev.clientX - startX), 120), 520))
                                      const onUp = () => { window.removeEventListener('mousemove', onMove); window.removeEventListener('mouseup', onUp) }
                                      window.addEventListener('mousemove', onMove); window.addEventListener('mouseup', onUp)
                                    }}
                                    className="absolute right-0 top-0 h-full w-1.5 cursor-col-resize bg-transparent hover:bg-pth-blue/30"
                                  />
                                </div>

                                {/* Scrollable right area */}
                                <div className="flex-1 overflow-x-auto min-w-0" ref={scrollRef}>
                                  <div ref={dynRef({ width: `${innerWidth}px` })} className="relative">
                                    {/* Month header */}
                                    <div className="relative h-6 border-b border-pth-border/15">
                                      {monthLabels.map((m, i) => (
                                        <span key={`${m.label}-${i}`} ref={dynRef({left: `${m.pct}%`})} className={`absolute -translate-x-1/2 text-[10px] whitespace-nowrap ${m.isQuarter ? 'font-bold text-pth-text' : 'font-medium text-pth-muted/60'}`}>{m.label}</span>
                                      ))}
                                    </div>

                                    {/* Chart rows */}
                                    <div className="relative">
                                      {/* Vertical grid — week ticks (light) */}
                                      {weekTicks.map((w, i) => (
                                        <div key={`wk-${i}`} ref={dynRef({left: `${w.pct}%`})} className="absolute top-0 bottom-0 w-px bg-pth-border/5" />
                                      ))}
                                      {/* Vertical grid — month lines */}
                                      {monthLabels.map((m, i) => (
                                        <div key={`grid-${i}`} ref={dynRef({left: `${m.pct}%`})} className={`absolute top-0 bottom-0 w-px ${m.isQuarter ? 'bg-pth-border/25' : 'bg-pth-border/10'}`} />
                                      ))}

                                      {/* Today line */}
                                      <div ref={dynRef({left: `${todayPct}%`})} className="absolute top-0 bottom-0 z-20 -translate-x-1/2">
                                        <div className="relative h-full">
                                          <div className="h-full w-px bg-pth-blue/50 border-l border-dashed border-pth-blue/30" />
                                          <span className="absolute -top-0.5 left-1/2 -translate-x-1/2 rounded bg-pth-blue px-1.5 py-0.5 text-[9px] font-bold text-white shadow-sm whitespace-nowrap">Today</span>
                                        </div>
                                      </div>

                                      {/* Project bars */}
                                      {bars.map((bar) => {
                                        const left = pct(bar.minStart)
                                        const right = pct(bar.maxEnd)
                                        const width = Math.max(right - left, 0.5)
                                        const barFill = bar.status === 'RED' ? 'bg-pth-red' : bar.status === 'YELLOW' ? 'bg-amber-400' : 'bg-emerald-500'
                                        const barBg = bar.status === 'RED' ? 'bg-pth-red/20' : bar.status === 'YELLOW' ? 'bg-amber-400/20' : 'bg-emerald-500/20'
                                        const progressPctVal = Math.round(bar.doneRatio * 100)
                                        return (
                                          <div
                                            key={bar.project.id}
                                            className="relative h-7 cursor-pointer group"
                                            onClick={() => { setDetailProjectId(bar.project.id); setSelectedProjectId(bar.project.id) }}
                                          >
                                            {/* Background bar */}
                                            <div ref={dynRef({left: `${left}%`, width: `${width}%`})} className={`absolute top-1 bottom-1 overflow-hidden rounded ${barBg} group-hover:ring-1 group-hover:ring-pth-blue/30 transition-shadow`}>
                                              <div ref={dynRef({width: `${progressPctVal}%`})} className={`h-full rounded ${barFill}`} />
                                            </div>
                                            {/* Progress label */}
                                            {(width / 100) * innerWidth > 40 && (
                                              <span ref={dynRef({left: `${left + width / 2}%`})} className="absolute top-1/2 -translate-x-1/2 -translate-y-1/2 text-[9px] font-bold text-pth-text/70 z-10">{progressPctVal}%</span>
                                            )}
                                            {/* Milestone diamonds */}
                                            {bar.milestones.map((ms) => {
                                              const msPct = pct(ms.targetDate)
                                              const isDone = !!ms.doneDate
                                              const isLate = !isDone && rygFromDueDate(ms.targetDate, state.settings.warningDaysThreshold) === 'RED'
                                              return (
                                                <div key={ms.id} ref={dynRef({left: `${msPct}%`})} className="absolute top-1/2 -translate-x-1/2 -translate-y-1/2 z-10" title={`${ms.name} — ${formatShortDate(ms.targetDate)}`}>
                                                  <div className={`h-2.5 w-2.5 rotate-45 border ${isDone ? 'bg-emerald-500 border-emerald-600' : isLate ? 'bg-pth-red border-pth-red' : 'bg-pth-blue border-pth-blue'} shadow-sm`} />
                                                </div>
                                              )
                                            })}
                                          </div>
                                        )
                                      })}
                                    </div>

                                    {/* Date axis footer with regular tick labels */}
                                    <div className="relative mt-1 h-5 border-t border-pth-border/10">
                                      {/* Bi-weekly date ticks for fine granularity */}
                                      {(() => {
                                        const ticks: Array<{label: string; pct: number}> = []
                                        const tc = new Date(toDate(globalMin))
                                        const dayOff = tc.getDay()
                                        tc.setDate(tc.getDate() + ((8 - dayOff) % 7)) // next Monday
                                        while (tc <= toDate(globalMax)) {
                                          const iso = tc.toISOString().slice(0, 10)
                                          const p = (dayDiff(globalMin, iso) / globalRange) * 100
                                          if (p >= 0 && p <= 100) {
                                            ticks.push({ label: `${tc.getDate()} ${new Intl.DateTimeFormat('en-US', { month: 'short' }).format(tc)}`, pct: p })
                                          }
                                          tc.setDate(tc.getDate() + 14) // every 2 weeks
                                        }
                                        return ticks.map((t, i) => (
                                          <span key={`dtick-${i}`} ref={dynRef({ left: `${t.pct}%` })} className="absolute top-1 -translate-x-1/2 text-[9px] text-pth-muted whitespace-nowrap">{t.label}</span>
                                        ))
                                      })()}
                                    </div>
                                  </div>
                                </div>
                              </div>
                            </div>
                          )
                        })()}

                        {/* All projects status table */}
                        <div className="overflow-auto rounded-xl border border-pth-border/20">
                          <table className="w-full text-sm">
                            <thead>
                              <tr className="border-b border-pth-border/15 bg-pth-subtle text-xs font-medium uppercase tracking-wider text-pth-muted">
                                {['Project', 'Category', 'PM', 'Status', 'Progress', 'Next Task'].map((h) => (
                                  <th key={h} className="px-4 py-3 text-left font-medium">{h}</th>
                                ))}
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-pth-border/10">
                              {rptRows.map((row) => (
                                <tr key={row.project.id} className="cursor-pointer transition-colors hover:bg-pth-hover" onClick={() => setReportProjectId(row.project.id)}>
                                  <td className="px-4 py-3 font-medium">{row.project.name}</td>
                                  <td className="px-4 py-3 text-pth-muted">{row.project.category}</td>
                                  <td className="px-4 py-3 text-pth-muted">{row.pm}</td>
                                  <td className="px-4 py-3">{statusBadge(row.status)}</td>
                                  <td className="px-4 py-3">
                                    <div className="flex items-center gap-2">
                                      <div className="h-1.5 w-16 overflow-hidden rounded-full bg-pth-border/20">
                                        <div ref={dynRef({ width: `${row.progressPct}%` })} className="h-full rounded-full bg-pth-blue" />
                                      </div>
                                      <span className="text-xs text-pth-muted">{row.progressPct}%</span>
                                    </div>
                                  </td>
                                  <td className="px-4 py-3 text-xs text-pth-muted">{row.nextTask ? `${row.nextTask.name} (${dayDiff(todayISO(), row.nextTask.endDate)}d)` : '—'}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>

                        {/* ── Project deep-dive selector ── */}
                        <div className="flex items-center gap-3">
                          <label className="text-xs font-semibold uppercase tracking-wider text-pth-muted">Deep Dive</label>
                          <select
                            title="Select a project to deep-dive"
                            className="rounded-lg border border-pth-border/40 bg-pth-card px-3 py-1.5 text-sm outline-none focus:border-pth-blue"
                            value={reportProjectId ?? ''}
                            onChange={(e) => setReportProjectId(e.target.value || null)}
                          >
                            <option value="">Select a project…</option>
                            {rptFiltered.map((p) => (
                              <option key={p.id} value={p.id}>{p.name}</option>
                            ))}
                          </select>
                          {reportProject && (
                            <button type="button" className="text-xs text-pth-muted hover:text-pth-text transition-colors" onClick={() => setReportProjectId(null)}>✕ Clear</button>
                          )}
                        </div>

                        {/* When project selected → detail view */}
                        {reportProject && (
                          <div className="space-y-4">
                            {/* Header card */}
                            <div className="rounded-xl border border-pth-border/15 bg-pth-subtle p-4">
                              <div className="mb-3 text-xs font-semibold uppercase tracking-wider text-pth-muted">Project Overview</div>
                              <div className="grid gap-3 text-sm sm:grid-cols-3 md:grid-cols-5">
                                <div><div className="text-[11px] text-pth-muted">Name</div><div className="font-medium">{reportProject.name}</div></div>
                                <div><div className="text-[11px] text-pth-muted">Category</div><div className="font-medium">{reportProject.category}</div></div>
                                <div><div className="text-[11px] text-pth-muted">PM</div><div className="font-medium">{people.find((p) => p.id === reportProject.projectManagerId)?.name ?? '—'}</div></div>
                                <div><div className="text-[11px] text-pth-muted">Site</div><div className="font-medium">{reportProject.siteIds.map((id) => sites.find((s) => s.id === id)?.name ?? id).join(', ') || '—'}</div></div>
                                <div><div className="text-[11px] text-pth-muted">Status</div><div className="mt-0.5">{statusBadge(seededProjectStatus(state, reportProject.id))}</div></div>
                              </div>
                            </div>

                            {/* Next milestone */}
                            <div className="rounded-xl border border-pth-border/15 bg-pth-card p-4">
                              <div className="mb-1 text-xs font-semibold uppercase tracking-wider text-pth-muted">Next Milestone</div>
                              {reportNextMilestone ? (
                                <div className="flex items-center gap-3 text-sm">
                                  <span className="font-medium">{reportNextMilestone.name}</span>
                                  <span className="text-pth-muted">—</span>
                                  <span className="text-pth-muted">{formatDate(reportNextMilestone.targetDate)}</span>
                                  <span className={`rounded-md px-2 py-0.5 text-xs font-medium ${dayDiff(todayISO(), reportNextMilestone.targetDate) < 0 ? 'bg-pth-red/10 text-pth-red' : 'bg-pth-subtle'}`}>{dayDiff(todayISO(), reportNextMilestone.targetDate)}d</span>
                                </div>
                              ) : (
                                <div className="text-sm text-pth-muted">No upcoming milestone</div>
                              )}
                            </div>

                            {/* Gantt chart */}
                            {reportActivities.length > 0 && (() => {
                              const allDates = [...reportActivities.map((a) => a.startDate), ...reportActivities.map((a) => a.endDate), ...reportMilestones.map((m) => m.targetDate)]
                              const minDate = allDates.sort()[0]
                              const maxDate = allDates.sort().slice(-1)[0]
                              const totalDays = Math.max(dayDiff(minDate, maxDate), 1)
                              const pct = (d: string) => `${Math.max(0, Math.min(100, (dayDiff(minDate, d) / totalDays) * 100))}%`
                              const todayPct = `${Math.max(0, Math.min(100, (dayDiff(minDate, todayISO()) / totalDays) * 100))}%`
                              return (
                                <div className="rounded-xl border border-pth-border/15 bg-pth-card p-4">
                                  <div className="mb-3 text-xs font-semibold uppercase tracking-wider text-pth-muted">Gantt Chart</div>
                                  <div className="relative min-h-[120px]">
                                    {/* Today marker */}
                                    <div ref={dynRef({ left: todayPct })} className="absolute top-0 bottom-0 w-px bg-pth-blue/40 z-10">
                                      <span className="absolute -top-4 left-1/2 -translate-x-1/2 text-[9px] font-medium text-pth-blue">Today</span>
                                    </div>
                                    {/* Activity bars */}
                                    <div className="space-y-1.5">
                                      {reportActivities.map((activity) => {
                                        const left = pct(activity.startDate)
                                        const right = `${100 - parseFloat(pct(activity.endDate))}%`
                                        const isDone = activity.state === 'DONE'
                                        const isDelayed = !isDone && rygFromDueDate(activity.endDate, state.settings.warningDaysThreshold) === 'RED'
                                        return (
                                          <div key={activity.id} className="flex items-center gap-2">
                                            <div className="w-28 shrink-0 truncate text-[11px] text-pth-muted">{activity.name}</div>
                                            <div className="relative h-5 flex-1 rounded-full bg-pth-subtle">
                                              <div ref={dynRef({ left, right })} className={`absolute top-0 bottom-0 rounded-full ${isDone ? 'bg-emerald-500/70' : isDelayed ? 'bg-pth-red/70' : 'bg-pth-blue/70'}`} />
                                            </div>
                                          </div>
                                        )
                                      })}
                                    </div>
                                    {/* Milestone diamonds */}
                                    {reportMilestones.map((ms) => (
                                      <div key={ms.id} ref={dynRef({ left: pct(ms.targetDate) })} className="absolute -translate-x-1/2" title={`${ms.name} — ${formatDate(ms.targetDate)}`}>
                                        <div className={`h-3 w-3 rotate-45 ${ms.doneDate ? 'bg-emerald-500' : 'bg-pth-blue'} border border-white dark:border-pth-card`} />
                                      </div>
                                    ))}
                                  </div>
                                  <div className="mt-2 flex justify-between text-[10px] text-pth-muted">
                                    <span>{formatShortDate(minDate)}</span>
                                    <span>{formatShortDate(maxDate)}</span>
                                  </div>
                                </div>
                              )
                            })()}

                            {/* Activity tabs: upcoming / closed / delayed */}
                            <div className="rounded-xl border border-pth-border/15 bg-pth-card overflow-hidden">
                              <div className="flex items-center gap-1 border-b border-pth-border/15 px-4 pt-2">
                                {([
                                  { key: 'upcoming' as ReportActivityTab, label: 'Upcoming', count: reportUpcoming.length },
                                  { key: 'closed' as ReportActivityTab, label: 'Recently Closed', count: reportClosed.length },
                                  { key: 'delayed' as ReportActivityTab, label: 'Delayed', count: reportDelayed.length },
                                ]).map((t) => (
                                  <button
                                    key={t.key}
                                    type="button"
                                    className={`relative px-3 py-2 text-xs font-medium transition-colors ${reportActivityTab === t.key ? 'text-pth-blue' : 'text-pth-muted hover:text-pth-text'}`}
                                    onClick={() => setReportActivityTab(t.key)}
                                  >
                                    {t.label} <span className="ml-1 text-[10px] opacity-60">({t.count})</span>
                                    {reportActivityTab === t.key && (
                                      <motion.div layoutId="rpt-activity-underline" className="absolute inset-x-0 -bottom-px h-0.5 rounded-full bg-pth-blue" transition={{ type: 'spring', duration: 0.35, bounce: 0.15 }} />
                                    )}
                                  </button>
                                ))}
                              </div>
                              <div className="p-4 space-y-2">
                                {reportActivityTab === 'upcoming' && (
                                  reportUpcoming.length === 0 ? <div className="text-xs text-pth-muted">No upcoming activities</div> : reportUpcoming.map((a) => (
                                    <div key={a.id} className="flex items-center justify-between rounded-lg bg-pth-subtle px-3 py-2 text-xs">
                                      <span className="font-medium">{a.name}</span>
                                      <span className="text-pth-muted">{people.find((p) => p.id === a.ownerId)?.name ?? 'Unknown'}</span>
                                    </div>
                                  ))
                                )}
                                {reportActivityTab === 'closed' && (
                                  reportClosed.length === 0 ? <div className="text-xs text-pth-muted">No recently closed activities</div> : reportClosed.map((a) => (
                                    <div key={a.id} className="flex items-center justify-between rounded-lg bg-pth-subtle px-3 py-2 text-xs">
                                      <span className="font-medium">{a.name}</span>
                                      <span className="text-pth-muted">{formatDate(a.doneDate ?? a.endDate)}</span>
                                    </div>
                                  ))
                                )}
                                {reportActivityTab === 'delayed' && (
                                  reportDelayed.length === 0 ? <div className="text-xs text-pth-muted">No delayed activities</div> : reportDelayed.map((a) => (
                                    <div key={a.id} className="flex items-center justify-between rounded-lg bg-pth-card px-3 py-2 text-xs">
                                      <span className="font-medium">{a.name}</span>
                                      <span className="font-semibold text-pth-red">{a.delayDays}d late</span>
                                    </div>
                                  ))
                                )}
                              </div>
                            </div>
                          </div>
                        )}
                      </div>
                    )}

                    {/* ── TASK TRACKING (preventive) TAB ── */}
                    {reportTab === 'taskTracking' && (() => {
                      // N/A tasks won't be done — exclude from this preventive tracking view entirely.
                      const rows = state.activities.filter((a) => a.state !== 'NOT_APPLICABLE').map((a) => {
                        const proj = state.projects.find((p) => p.id === a.projectId)
                        const status = rygForActivity(a)
                        const daysToDue = a.doneDate ? null : dayDiff(todayISO(), a.endDate)
                        const owner = people.find((p) => p.id === a.ownerId)
                        const locNames = proj ? proj.siteIds.map((id) => sites.find((s) => s.id === id)?.name ?? id) : []
                        return { a, proj, status, daysToDue, ownerName: owner?.name ?? 'Unassigned', locNames }
                      })
                      const areaTokens = Array.from(new Set(
                        state.activities.flatMap((a) => (a.responsible ?? '').split(/[+/,]/).map((t) => t.trim()).filter(Boolean)),
                      )).sort()
                      const projTypes = Array.from(new Set(state.projects.map((p) => p.projectType).filter(Boolean) as string[])).sort()
                      const ownerNames = Array.from(new Set(rows.map((r) => r.ownerName))).sort((a, b) => (a === 'Unassigned' ? 1 : b === 'Unassigned' ? -1 : a.localeCompare(b)))
                      const filtered = rows.filter((r) => {
                        if (ttArea !== 'all' && !(r.a.responsible ?? '').toLowerCase().includes(ttArea.toLowerCase())) return false
                        if (ttLoc !== 'all' && !r.locNames.includes(ttLoc)) return false
                        if (ttProjType !== 'all' && (r.proj?.projectType ?? '') !== ttProjType) return false
                        if (ttOwner !== 'all' && r.ownerName !== ttOwner) return false
                        if (ttRisk !== 'all' && r.a.doneDate) return false
                        if (ttRisk === 'atrisk' && r.status === 'GREEN') return false
                        if (ttRisk === 'red' && r.status !== 'RED') return false
                        if (ttRisk === 'green' && r.status !== 'GREEN') return false
                        return true
                      }).sort((x, y) => {
                        const rank = (s: RygStatus) => (s === 'RED' ? 0 : s === 'YELLOW' ? 1 : 2)
                        if (rank(x.status) !== rank(y.status)) return rank(x.status) - rank(y.status)
                        return (x.daysToDue ?? 99999) - (y.daysToDue ?? 99999)
                      })
                      const redCount = rows.filter((r) => !r.a.doneDate && r.status === 'RED').length
                      const yellowCount = rows.filter((r) => !r.a.doneDate && r.status === 'YELLOW').length
                      const dueSoon = rows.filter((r) => !r.a.doneDate && r.daysToDue !== null && r.daysToDue >= 0 && r.daysToDue <= 14).length
                      const sel = 'h-9 rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm outline-none focus:border-pth-blue'

                      // ── Build groups for By-Area / By-Project drill-down ──
                      type TtRow = typeof filtered[number]
                      const rollup = (items: TtRow[]): RygStatus => {
                        const open = items.filter((r) => !r.a.doneDate)
                        if (open.some((r) => r.status === 'RED')) return 'RED'
                        if (open.some((r) => r.status === 'YELLOW')) return 'YELLOW'
                        return 'GREEN'
                      }
                      const groupKeys = (r: TtRow): string[] => ttView === 'area'
                        ? ((r.a.responsible ?? '').split(/[+/,]/).map((t) => t.trim()).filter(Boolean) || ['—'])
                        : ttView === 'owner'
                        ? [r.ownerName || 'Unassigned']
                        : [r.proj?.name ?? '—']
                      const groupMap = new Map<string, TtRow[]>()
                      if (ttView !== 'tasks') {
                        for (const r of filtered) {
                          for (const k of (groupKeys(r).length ? groupKeys(r) : ['—'])) {
                            if (!groupMap.has(k)) groupMap.set(k, [])
                            groupMap.get(k)!.push(r)
                          }
                        }
                      }
                      const groups = Array.from(groupMap.entries()).map(([key, items]) => ({
                        key, items,
                        roll: rollup(items),
                        red: items.filter((r) => !r.a.doneDate && r.status === 'RED').length,
                        yellow: items.filter((r) => !r.a.doneDate && r.status === 'YELLOW').length,
                        green: items.filter((r) => !r.a.doneDate && r.status === 'GREEN').length,
                        done: items.filter((r) => r.a.doneDate).length,
                      })).sort((a, b) => {
                        const rank = (s: RygStatus) => (s === 'RED' ? 0 : s === 'YELLOW' ? 1 : 2)
                        if (rank(a.roll) !== rank(b.roll)) return rank(a.roll) - rank(b.roll)
                        return b.red - a.red
                      })
                      const toggleGroup = (k: string) => setTtExpanded((prev) => {
                        const n = new Set(prev); n.has(k) ? n.delete(k) : n.add(k); return n
                      })
                      const dotClass = (s: RygStatus) => s === 'RED' ? 'bg-pth-red' : s === 'YELLOW' ? 'bg-amber-500' : 'bg-emerald-500'
                      return (
                        <div className="space-y-4">
                          <p className="text-xs text-pth-muted">Preventive view — surfaces tasks at risk of missing their due date (target: finish ≥2 weeks early). Filter by area, location and project type to see what won't be met in time.</p>
                          <div className="flex flex-wrap items-center gap-2">
                            <select title="Area" className={sel} value={ttArea} onChange={(e) => setTtArea(e.target.value)}>
                              <option value="all">All Areas</option>
                              {areaTokens.map((t) => <option key={t} value={t}>{t}</option>)}
                            </select>
                            <select title="Location" className={sel} value={ttLoc} onChange={(e) => setTtLoc(e.target.value)}>
                              <option value="all">All Locations</option>
                              {sites.map((s) => <option key={s.id} value={s.name}>{s.name}</option>)}
                            </select>
                            <select title="Project Type" className={sel} value={ttProjType} onChange={(e) => setTtProjType(e.target.value)}>
                              <option value="all">All Project Types</option>
                              {projTypes.map((t) => <option key={t} value={t}>{t}</option>)}
                            </select>
                            <select title="Owner" className={sel} value={ttOwner} onChange={(e) => setTtOwner(e.target.value)}>
                              <option value="all">All Owners</option>
                              {ownerNames.map((o) => <option key={o} value={o}>{o}</option>)}
                            </select>
                            <select title="Risk" className={sel} value={ttRisk} onChange={(e) => setTtRisk(e.target.value)}>
                              <option value="atrisk">At risk (RED + YELLOW)</option>
                              <option value="red">Will miss (RED)</option>
                              <option value="green">On track (GREEN)</option>
                              <option value="all">All tasks</option>
                            </select>
                            {(ttArea !== 'all' || ttLoc !== 'all' || ttProjType !== 'all' || ttOwner !== 'all' || ttRisk !== 'atrisk') && (
                              <button type="button" className="text-xs text-pth-muted hover:text-pth-text" onClick={() => { setTtArea('all'); setTtLoc('all'); setTtProjType('all'); setTtOwner('all'); setTtRisk('atrisk') }}>✕ Clear</button>
                            )}
                            <div className="ml-auto flex items-center rounded-lg border border-pth-border/40 bg-pth-subtle p-0.5 text-xs">
                              {([['tasks', 'Tasks'], ['area', 'By Area'], ['project', 'By Project'], ['owner', 'By Owner']] as const).map(([k, lbl]) => (
                                <button key={k} type="button" onClick={() => setTtView(k)} className={`rounded-md px-2.5 py-1 font-medium transition-colors ${ttView === k ? 'bg-pth-card text-pth-blue shadow-sm' : 'text-pth-muted hover:text-pth-text'}`}>{lbl}</button>
                              ))}
                            </div>
                            <span className="text-xs text-pth-muted">{filtered.length} tasks</span>
                          </div>
                          <div className="grid grid-cols-3 gap-3">
                            <div className="rounded-xl border border-pth-border/15 bg-pth-card p-4 text-center">
                              <div className="text-[11px] font-semibold uppercase tracking-wider text-pth-muted">Will miss due date</div>
                              <div className="mt-1 text-2xl font-bold text-pth-red">{redCount}</div>
                            </div>
                            <div className="rounded-xl border border-pth-border/15 bg-pth-card p-4 text-center">
                              <div className="text-[11px] font-semibold uppercase tracking-wider text-pth-muted">At risk (no buffer)</div>
                              <div className="mt-1 text-2xl font-bold text-amber-600">{yellowCount}</div>
                            </div>
                            <div className="rounded-xl border border-pth-border/15 bg-pth-card p-4 text-center">
                              <div className="text-[11px] font-semibold uppercase tracking-wider text-pth-muted">Due within 2 weeks</div>
                              <div className="mt-1 text-2xl font-bold text-pth-blue">{dueSoon}</div>
                            </div>
                          </div>
                          {/* Flat task list */}
                          {ttView === 'tasks' && (
                          <div className="overflow-x-auto rounded-xl border border-pth-border/15">
                            <table className="w-full text-sm">
                              <thead>
                                <tr className="border-b border-pth-border/15 text-left text-[11px] uppercase tracking-wide text-pth-muted">
                                  <th className="px-3 py-2 font-semibold">Status</th>
                                  <th className="px-3 py-2 font-semibold">Task</th>
                                  <th className="px-3 py-2 font-semibold">Project</th>
                                  <th className="px-3 py-2 font-semibold">Type</th>
                                  <th className="px-3 py-2 font-semibold">Location</th>
                                  <th className="px-3 py-2 font-semibold">Area</th>
                                  <th className="px-3 py-2 font-semibold">Owner</th>
                                  <th className="whitespace-nowrap px-3 py-2 font-semibold">Due</th>
                                  <th className="whitespace-nowrap px-3 py-2 font-semibold">Days left</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-pth-border/10">
                                {filtered.map((r) => {
                                  const dot = dotClass(r.status)
                                  const dleft = r.daysToDue
                                  return (
                                    <tr key={r.a.id} className="transition-colors hover:bg-pth-hover/40">
                                      <td className="px-3 py-2"><span className={`inline-block h-2.5 w-2.5 rounded-full ${dot}`} /></td>
                                      <td className="px-3 py-2 font-medium">{r.a.name}{r.a.criticalPath && <span className="ml-1.5 rounded bg-pth-red/10 px-1 py-0.5 text-[9px] font-semibold text-pth-red">CP</span>}</td>
                                      <td className="max-w-[180px] truncate px-3 py-2 text-pth-muted" title={r.proj?.name}>{r.proj?.name ?? '—'}</td>
                                      <td className="px-3 py-2 text-pth-muted">{r.proj?.projectType || '—'}</td>
                                      <td className="px-3 py-2 text-pth-muted">{r.locNames.join(', ') || '—'}</td>
                                      <td className="px-3 py-2 text-pth-muted">{r.a.responsible || '—'}</td>
                                      <td className="px-3 py-2 text-pth-muted">{r.ownerName}</td>
                                      <td className="whitespace-nowrap px-3 py-2 text-pth-muted">{r.a.doneDate ? 'Done' : formatShortDate(r.a.endDate)}</td>
                                      <td className={`whitespace-nowrap px-3 py-2 font-medium ${dleft != null && dleft < 0 ? 'text-pth-red' : dleft != null && dleft <= 14 ? 'text-amber-600' : 'text-pth-muted'}`}>
                                        {r.a.doneDate ? '—' : dleft != null ? (dleft < 0 ? `${-dleft}d overdue` : `${dleft}d`) : '—'}
                                      </td>
                                    </tr>
                                  )
                                })}
                                {filtered.length === 0 && <tr><td colSpan={9} className="px-3 py-8 text-center text-pth-muted">No tasks match these filters.</td></tr>}
                              </tbody>
                            </table>
                          </div>
                          )}

                          {/* Grouped by Area / Project — expandable */}
                          {ttView !== 'tasks' && (
                          <div className="space-y-2">
                            {groups.map((g) => {
                              const open = ttExpanded.has(g.key)
                              return (
                                <div key={g.key} className="overflow-hidden rounded-xl border border-pth-border/15">
                                  <button type="button" onClick={() => toggleGroup(g.key)} className="flex w-full items-center gap-3 bg-pth-subtle/50 px-4 py-2.5 text-left transition-colors hover:bg-pth-subtle">
                                    <ChevronDown size={15} className={`shrink-0 text-pth-muted transition-transform ${open ? '' : '-rotate-90'}`} />
                                    <span className={`h-2.5 w-2.5 shrink-0 rounded-full ${dotClass(g.roll)}`} />
                                    <span className="font-semibold">{g.key}</span>
                                    <span className="text-xs text-pth-muted">{g.items.length} task{g.items.length === 1 ? '' : 's'}</span>
                                    <span className="ml-auto flex items-center gap-2 text-[11px]">
                                      {g.red > 0 && <span className="rounded-full bg-pth-red/10 px-2 py-0.5 font-semibold text-pth-red">{g.red} will miss</span>}
                                      {g.yellow > 0 && <span className="rounded-full bg-amber-500/10 px-2 py-0.5 font-semibold text-amber-600">{g.yellow} at risk</span>}
                                      {g.green > 0 && <span className="rounded-full bg-emerald-500/10 px-2 py-0.5 font-semibold text-emerald-600">{g.green} on track</span>}
                                      {g.done > 0 && <span className="rounded-full bg-pth-border/20 px-2 py-0.5 font-medium text-pth-muted">{g.done} done</span>}
                                    </span>
                                  </button>
                                  {open && (
                                    <table className="w-full text-sm">
                                      <thead>
                                        <tr className="border-y border-pth-border/15 text-left text-[10px] uppercase tracking-wide text-pth-muted">
                                          <th className="px-3 py-1.5 pl-10 font-semibold">Task</th>
                                          <th className="px-3 py-1.5 font-semibold">{ttView === 'area' ? 'Project' : 'Area'}</th>
                                          <th className="px-3 py-1.5 font-semibold">Owner</th>
                                          <th className="whitespace-nowrap px-3 py-1.5 font-semibold">Due</th>
                                          <th className="whitespace-nowrap px-3 py-1.5 font-semibold">Days left</th>
                                        </tr>
                                      </thead>
                                      <tbody className="divide-y divide-pth-border/10">
                                        {g.items.sort((x, y) => {
                                          const rank = (s: RygStatus) => (s === 'RED' ? 0 : s === 'YELLOW' ? 1 : 2)
                                          if (rank(x.status) !== rank(y.status)) return rank(x.status) - rank(y.status)
                                          return (x.daysToDue ?? 99999) - (y.daysToDue ?? 99999)
                                        }).map((r) => {
                                          const dleft = r.daysToDue
                                          return (
                                            <tr key={r.a.id} className="transition-colors hover:bg-pth-hover/40">
                                              <td className="px-3 py-1.5 pl-10"><span className={`mr-2 inline-block h-2 w-2 rounded-full ${dotClass(r.status)}`} />{r.a.name}{r.a.criticalPath && <span className="ml-1.5 rounded bg-pth-red/10 px-1 py-0.5 text-[9px] font-semibold text-pth-red">CP</span>}</td>
                                              <td className="max-w-[200px] truncate px-3 py-1.5 text-pth-muted" title={ttView === 'area' ? r.proj?.name : r.a.responsible}>{ttView === 'area' ? (r.proj?.name ?? '—') : (r.a.responsible || '—')}</td>
                                              <td className="px-3 py-1.5 text-pth-muted">{r.ownerName}</td>
                                              <td className="whitespace-nowrap px-3 py-1.5 text-pth-muted">{r.a.doneDate ? 'Done' : formatShortDate(r.a.endDate)}</td>
                                              <td className={`whitespace-nowrap px-3 py-1.5 font-medium ${dleft != null && dleft < 0 ? 'text-pth-red' : dleft != null && dleft <= 14 ? 'text-amber-600' : 'text-pth-muted'}`}>{r.a.doneDate ? '—' : dleft != null ? (dleft < 0 ? `${-dleft}d overdue` : `${dleft}d`) : '—'}</td>
                                            </tr>
                                          )
                                        })}
                                      </tbody>
                                    </table>
                                  )}
                                </div>
                              )
                            })}
                            {groups.length === 0 && <div className="rounded-xl border border-pth-border/15 px-3 py-8 text-center text-pth-muted">No tasks match these filters.</div>}
                          </div>
                          )}
                        </div>
                      )
                    })()}

                    {reportTab === 'prioritization' && (
                      <div className="space-y-4">
                        {role !== 'DIRECTOR' && (
                          <div className="rounded-lg border border-amber-200/60 bg-amber-50/80 px-4 py-3 text-xs font-medium text-amber-800 dark:border-amber-500/20 dark:bg-amber-900/20 dark:text-amber-300">
                            Prioritization & Risks is a Director view. Current role is {ROLE_LABEL[role]}.
                          </div>
                        )}

                        {/* Milestones risk table */}
                        <div className="overflow-auto rounded-xl border border-pth-border/20">
                          <table className="w-full text-sm">
                            <thead>
                              <tr className="border-b border-pth-border/15 bg-pth-subtle text-xs font-medium uppercase tracking-wider text-pth-muted">
                                {['Project', 'Milestone', 'Due Date', 'Days', 'RYG'].map((h) => (
                                  <th key={h} className="px-4 py-3 text-left font-medium">{h}</th>
                                ))}
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-pth-border/10">
                              {upcomingMilestonesRisk.map((row) => (
                                <tr key={row.milestone.id} className="transition-colors hover:bg-pth-hover">
                                  <td className="px-4 py-3 font-medium">{row.projectName}</td>
                                  <td className="px-4 py-3 text-pth-muted">{row.milestone.name}</td>
                                  <td className="px-4 py-3 text-pth-muted">{formatDate(row.milestone.targetDate)}</td>
                                  <td className="px-4 py-3">
                                    <span className={`text-xs font-semibold ${row.daysRemaining < 0 ? 'text-pth-red' : row.daysRemaining <= state.settings.warningDaysThreshold ? 'text-amber-600' : 'text-pth-muted'}`}>{row.daysRemaining}d</span>
                                  </td>
                                  <td className="px-4 py-3">{statusBadge(row.status)}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>

                        {/* Delayed activities grouped */}
                        <div className="space-y-3">
                          <h3 className="text-xs font-semibold uppercase tracking-wider text-pth-muted">Delayed Activities per Project</h3>
                          {delayedGrouped.map((group) => (
                            <div key={group.project.id} className="rounded-xl border border-pth-border/15 bg-pth-card transition-shadow hover:shadow-sm">
                              <button
                                type="button"
                                className="flex w-full items-center justify-between px-4 py-3"
                                onClick={() => setExpandedRiskProjectId((current) => (current === group.project.id ? null : group.project.id))}
                              >
                                <span className="text-sm font-medium">{group.project.name}</span>
                                <div className="flex items-center gap-2">
                                  <span className="inline-flex h-5 min-w-5 items-center justify-center rounded-full bg-pth-red/10 px-1.5 text-xs font-semibold text-pth-red">{group.items.length}</span>
                                  {group.criticalImpact && <ShieldAlert size={14} className="text-pth-red" />}
                                  <ChevronDown size={14} className={`text-pth-muted transition-transform ${expandedRiskProjectId === group.project.id ? 'rotate-180' : ''}`} />
                                </div>
                              </button>
                              <AnimatePresence>
                                {expandedRiskProjectId === group.project.id && (
                                  <motion.div initial={{ height: 0, opacity: 0 }} animate={{ height: 'auto', opacity: 1 }} exit={{ height: 0, opacity: 0 }} className="overflow-hidden">
                                    <div className="space-y-1.5 border-t border-pth-border/15 px-4 py-3">
                                      {group.items.map((activity) => (
                                        <div key={activity.id} className="flex items-center gap-2 rounded-lg bg-pth-subtle px-3 py-2 text-xs">
                                          <span className="h-1.5 w-1.5 rounded-full bg-pth-red" />
                                          <span className="font-medium">{activity.name}</span>
                                          <span className="text-pth-muted">•</span>
                                          <span className="text-pth-muted">{people.find((p) => p.id === activity.ownerId)?.name ?? 'Unknown'}</span>
                                          {activity.criticalPath && <span className="ml-auto rounded-md bg-pth-red/10 px-1.5 py-0.5 text-[10px] font-semibold text-pth-red">Critical</span>}
                                        </div>
                                      ))}
                                    </div>
                                  </motion.div>
                                )}
                              </AnimatePresence>
                            </div>
                          ))}
                        </div>
                      </div>
                    )}

                    {reportTab === 'workloadOverview' && (
                      <div className="space-y-3">
                        <div className="flex items-center justify-between">
                          <span className="text-xs font-semibold uppercase tracking-wider text-pth-muted">
                            {calendarWeek(currentWeek)} &middot; {formatShortDate(currentWeek)} – {formatShortDate(addDays(currentWeek, 6))}
                          </span>
                          <span className="text-xs text-pth-muted">{workloadRows.length} resources</span>
                        </div>
                        {workloadRows.map((row) => {
                          const warning = row.utilization >= state.settings.workloadWarningThreshold && row.utilization <= state.settings.workloadCriticalThreshold
                          const critical = row.utilization > state.settings.workloadCriticalThreshold
                          const isExpanded = expandedWorkloadPersonId === row.person.id

                          // group assignments by project for the accordion content
                          const projectBreakdown = row.personAssignments.reduce<Record<string, { name: string; hours: number }>>((acc, asgn) => {
                            const act = state.activities.find((a) => a.id === asgn.activityId)
                            const projId = act?.projectId
                            if (!projId) return acc
                            const projName = state.projects.find((p) => p.id === projId)?.name ?? 'Unknown'
                            if (!acc[projId]) acc[projId] = { name: projName, hours: 0 }
                            acc[projId].hours += asgn.assignedHours
                            return acc
                          }, {})
                          const totalHours = Object.values(projectBreakdown).reduce((s, p) => s + p.hours, 0)

                          return (
                            <div key={row.person.id} className="rounded-xl border border-pth-border/15 bg-pth-card transition-shadow hover:shadow-sm">
                              <button
                                type="button"
                                className="flex w-full items-center gap-4 px-4 py-3"
                                onClick={() => setExpandedWorkloadPersonId((cur) => (cur === row.person.id ? null : row.person.id))}
                              >
                                <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-pth-subtle text-xs font-bold">{row.person.initials}</div>
                                <div className="min-w-0 flex-1 text-left">
                                  <div className="text-sm font-medium">{row.person.name}</div>
                                  <div className="text-[11px] text-pth-muted">{row.utilization}% allocated</div>
                                </div>
                                <div className="flex items-center gap-3">
                                  <div className="w-24 h-1.5 rounded-full bg-pth-border/20 overflow-hidden">
                                    <div ref={dynRef({ width: `${Math.min(row.utilization, 150)}%` })} className={`h-full rounded-full transition-all ${critical ? 'bg-pth-red' : warning ? 'bg-amber-500' : 'bg-emerald-500'}`} />
                                  </div>
                                  <span className={`min-w-[3rem] text-right text-xs font-semibold ${critical ? 'text-pth-red' : warning ? 'text-amber-600' : 'text-emerald-600'}`}>{row.utilization}%</span>
                                  <ChevronDown size={14} className={`text-pth-muted transition-transform ${isExpanded ? 'rotate-180' : ''}`} />
                                </div>
                              </button>
                              <AnimatePresence>
                                {isExpanded && (
                                  <motion.div initial={{ height: 0, opacity: 0 }} animate={{ height: 'auto', opacity: 1 }} exit={{ height: 0, opacity: 0 }} className="overflow-hidden">
                                    <div className="border-t border-pth-border/15 px-4 py-3 space-y-1.5">
                                      {Object.entries(projectBreakdown).length > 0 ? Object.entries(projectBreakdown).map(([projId, info]) => {
                                        const pct = totalHours > 0 ? Math.round((info.hours / Math.max(row.capacityHours, 1)) * 100) : 0
                                        return (
                                        <div key={projId} className="flex items-center justify-between rounded-lg bg-pth-subtle px-3 py-2 text-xs">
                                          <div className="flex items-center gap-2 min-w-0">
                                            <span className="h-1.5 w-1.5 rounded-full bg-pth-blue" />
                                            <span className="font-medium truncate">{info.name}</span>
                                          </div>
                                          <span className="shrink-0 font-semibold">{pct}%</span>
                                        </div>)
                                      }) : (
                                        <div className="text-xs text-pth-muted px-3 py-2">No project assignments for this week</div>
                                      )}
                                    </div>
                                  </motion.div>
                                )}
                              </AnimatePresence>
                            </div>
                          )
                        })}
                      </div>
                    )}
                  </div>
                </div>
              )}

              {/* ─── ALERTS PAGE ─── */}
              {navPage === 'alerts' && (
                <div className="space-y-4">
                  {/* ─── Alert Simulation ─── */}
                  <div className="glass-card rounded-xl border border-pth-border/20 shadow-card">
                    <div className="border-b border-pth-border/15 px-5 py-4">
                      <h2 className="text-base font-semibold tracking-tight">Alert Simulation</h2>
                      <p className="mt-0.5 text-xs text-pth-muted">Generate weekly digest notifications based on current data</p>
                    </div>
                    <div className="p-5">
                      <button
                        type="button"
                        className="inline-flex items-center gap-2 rounded-lg bg-pth-btn px-5 py-2.5 text-sm font-semibold text-white shadow-sm transition-all hover:bg-pth-btn-hover active:scale-[0.98]"
                        onClick={simulateWeeklyAlerts}
                      >
                        <CalendarClock size={14} /> Simulate Weekly Alerts
                      </button>
                    </div>
                  </div>

                  {/* ─── Power Automate Integration ─── */}
                  <div className="glass-card rounded-xl border border-pth-border/20 shadow-card">
                    <div className="border-b border-pth-border/15 px-5 py-4">
                      <div className="flex items-center gap-2.5">
                        <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-gradient-to-br from-blue-600 to-violet-600 shadow-sm">
                          <Zap size={16} className="text-white" />
                        </div>
                        <div>
                          <h2 className="text-base font-semibold tracking-tight">Power Automate Integration</h2>
                          <p className="mt-0.5 text-xs text-pth-muted">Configure automated flows to notify your team via Microsoft Teams, email, and more</p>
                        </div>
                      </div>
                    </div>

                    {/* Webhook configuration */}
                    <div className="border-b border-pth-border/10 px-5 py-4">
                      <label className="space-y-1.5">
                        <span className="text-xs font-medium text-pth-muted">Webhook URL (HTTP trigger endpoint)</span>
                        <div className="flex gap-2">
                          <input
                            type="text"
                            className="h-9 flex-1 rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-xs font-mono transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20"
                            value={paWebhookUrl}
                            onChange={(e) => setPaWebhookUrl(e.target.value)}
                            placeholder="https://prod-XX.westus.logic.azure.com:443/workflows/..."
                          />
                          <a
                            href="https://make.powerautomate.com"
                            target="_blank"
                            rel="noopener noreferrer"
                            className="inline-flex h-9 items-center gap-1.5 rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-xs font-medium text-pth-muted transition-colors hover:bg-pth-card hover:text-pth-text"
                          >
                            <ExternalLink size={12} /> Open Portal
                          </a>
                        </div>
                      </label>
                      <div className="mt-2 flex items-center gap-1.5 text-[11px] text-pth-muted">
                        <span className={`h-1.5 w-1.5 rounded-full ${paWebhookUrl.includes('logic.azure.com') ? 'bg-emerald-500' : 'bg-amber-400'}`} />
                        {paWebhookUrl.includes('logic.azure.com') ? 'Endpoint configured' : 'Provide a valid Power Automate HTTP trigger URL'}
                      </div>
                    </div>

                    {/* Flow list */}
                    <div className="divide-y divide-pth-border/10">
                      {paFlows.map((flow) => (
                        <div key={flow.id} className="flex items-center gap-4 px-5 py-3.5">
                          <button
                            type="button"
                            className="shrink-0"
                            onClick={() => togglePaFlow(flow.id)}
                            title={flow.enabled ? 'Disable flow' : 'Enable flow'}
                          >
                            {flow.enabled ? (
                              <ToggleRight size={28} className="text-pth-blue" />
                            ) : (
                              <ToggleLeft size={28} className="text-pth-muted/50" />
                            )}
                          </button>
                          <div className="min-w-0 flex-1">
                            <p className={`text-sm font-medium ${flow.enabled ? 'text-pth-text' : 'text-pth-muted'}`}>{flow.name}</p>
                            <p className="mt-0.5 text-[11px] text-pth-muted">{flow.description}</p>
                          </div>
                          <span className={`rounded-full px-2 py-0.5 text-[10px] font-semibold uppercase tracking-wider ${
                            flow.triggerType === 'overdue_milestones' ? 'bg-red-500/10 text-red-600 dark:text-red-400' :
                            flow.triggerType === 'delayed_activities' ? 'bg-amber-500/10 text-amber-600 dark:text-amber-400' :
                            'bg-violet-500/10 text-violet-600 dark:text-violet-400'
                          }`}>
                            {flow.triggerType.replace(/_/g, ' ')}
                          </span>
                          <button
                            type="button"
                            className="inline-flex items-center gap-1.5 rounded-lg border border-pth-border/40 bg-pth-subtle px-3 py-1.5 text-xs font-medium text-pth-muted transition-colors hover:bg-pth-card hover:text-pth-text disabled:opacity-40"
                            onClick={() => testPowerAutomateFlow(flow.id)}
                            disabled={paTestingFlowId !== null}
                          >
                            {paTestingFlowId === flow.id ? (
                              <span className="inline-block h-3 w-3 animate-spin rounded-full border-2 border-pth-blue border-t-transparent" />
                            ) : (
                              <Send size={12} />
                            )}
                            {paTestingFlowId === flow.id ? 'Testing…' : 'Test'}
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* ─── Flow Trigger Log ─── */}
                  {paFlowLog.length > 0 && (
                    <div className="glass-card rounded-xl border border-pth-border/20 shadow-card">
                      <div className="border-b border-pth-border/15 px-5 py-4">
                        <h3 className="text-sm font-semibold">Flow Trigger Log</h3>
                        <p className="mt-0.5 text-[11px] text-pth-muted">Recent Power Automate flow executions</p>
                      </div>
                      <div className="divide-y divide-pth-border/10">
                        {paFlowLog.slice(0, 15).map((entry) => (
                          <div key={entry.id} className="flex items-center gap-3 px-5 py-2.5">
                            {entry.status === 'success' ? (
                              <CheckCircle size={14} className="shrink-0 text-emerald-500" />
                            ) : (
                              <XCircle size={14} className="shrink-0 text-red-500" />
                            )}
                            <div className="min-w-0 flex-1">
                              <p className="text-xs text-pth-text">{entry.payload}</p>
                              <p className="mt-0.5 text-[10px] text-pth-muted">{entry.flowName} · {new Date(entry.triggeredAt).toLocaleString()}</p>
                            </div>
                            <span className={`rounded px-1.5 py-0.5 text-[10px] font-semibold ${entry.status === 'success' ? 'bg-emerald-500/10 text-emerald-600 dark:text-emerald-400' : 'bg-red-500/10 text-red-600 dark:text-red-400'}`}>
                              {entry.status}
                            </span>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* ─── Recent Alerts list ─── */}
                  <div className="glass-card rounded-xl border border-pth-border/20 shadow-card">
                    <div className="border-b border-pth-border/15 px-5 py-4">
                      <h3 className="text-sm font-semibold">Recent Notifications</h3>
                    </div>
                    <div className="divide-y divide-pth-border/10">
                      {state.notifications.length === 0 ? (
                        <div className="px-5 py-10 text-center text-sm text-pth-muted">No notifications yet. Run a simulation to generate alerts.</div>
                      ) : (
                        state.notifications.slice(0, 20).map((n) => (
                          <div key={n.id} className="flex items-start gap-3 px-5 py-3">
                            <span className={`mt-0.5 h-2 w-2 shrink-0 rounded-full ${n.severity === 'RED' ? 'bg-red-500' : n.severity === 'YELLOW' ? 'bg-amber-500' : 'bg-emerald-500'}`} />
                            <div>
                              <p className="text-sm text-pth-text">{n.message}</p>
                              <p className="mt-0.5 text-[11px] text-pth-muted">{formatDate(n.createdAt)}</p>
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  </div>
                </div>
              )}

              {/* ─── ADMIN PAGE ─── */}
              {navPage === 'admin' && (
                <div className="space-y-4">
                  <div className="glass-card rounded-xl border border-pth-border/20 shadow-card">
                    <div className="border-b border-pth-border/15 px-5 py-4">
                      <h2 className="text-base font-semibold tracking-tight">Configuration</h2>
                      <p className="mt-0.5 text-xs text-pth-muted">Adjust thresholds and alert parameters</p>
                    </div>
                    <div className="grid gap-4 p-5 md:grid-cols-2 lg:grid-cols-3">
                      <label className="space-y-1.5">
                        <span className="text-xs font-medium text-pth-muted">Warning Days Threshold</span>
                        <input
                          type="number"
                          className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20"
                          value={state.settings.warningDaysThreshold}
                          onChange={(event) => setState((prev) => ({ ...prev, settings: { ...prev.settings, warningDaysThreshold: Number(event.target.value) || 0 } }))}
                        />
                      </label>

                      <label className="space-y-1.5">
                        <span className="text-xs font-medium text-pth-muted">Upcoming Weeks Window</span>
                        <input
                          type="number"
                          className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20"
                          value={state.settings.upcomingWeeksWindow}
                          onChange={(event) => setState((prev) => ({ ...prev, settings: { ...prev.settings, upcomingWeeksWindow: Number(event.target.value) || 1 } }))}
                        />
                      </label>

                      <label className="space-y-1.5">
                        <span className="text-xs font-medium text-pth-muted">Default Weekly Capacity</span>
                        <input
                          type="number"
                          className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20"
                          value={state.settings.defaultWeeklyCapacity}
                          onChange={(event) => setState((prev) => ({ ...prev, settings: { ...prev.settings, defaultWeeklyCapacity: Number(event.target.value) || 1 } }))}
                        />
                      </label>

                      <label className="space-y-1.5">
                        <span className="text-xs font-medium text-pth-muted">Workload Warning %</span>
                        <input
                          type="number"
                          className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20"
                          value={state.settings.workloadWarningThreshold}
                          onChange={(event) => setState((prev) => ({ ...prev, settings: { ...prev.settings, workloadWarningThreshold: Number(event.target.value) || 1 } }))}
                        />
                      </label>

                      <label className="space-y-1.5">
                        <span className="text-xs font-medium text-pth-muted">Workload Critical %</span>
                        <input
                          type="number"
                          className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20"
                          value={state.settings.workloadCriticalThreshold}
                          onChange={(event) => setState((prev) => ({ ...prev, settings: { ...prev.settings, workloadCriticalThreshold: Number(event.target.value) || 1 } }))}
                        />
                      </label>
                    </div>
                    {/* Save Settings button */}
                    <div className="flex justify-end border-t border-pth-border/15 px-5 py-4">
                      <button
                        type="button"
                        className="inline-flex items-center gap-2 rounded-lg bg-pth-btn px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition-all hover:bg-pth-btn-hover active:scale-[0.98]"
                        onClick={async () => {
                          if (isDataverseConfigured() && state.settings.dvSettingsId) {
                            const ok = await updateSettingsInDataverse(state.settings.dvSettingsId, {
                              warningDaysThreshold: state.settings.warningDaysThreshold,
                              upcomingWeeksWindow: state.settings.upcomingWeeksWindow,
                              defaultWeeklyCapacity: state.settings.defaultWeeklyCapacity,
                              workloadWarningThreshold: state.settings.workloadWarningThreshold,
                              workloadCriticalThreshold: state.settings.workloadCriticalThreshold,
                            })
                            if (ok) {
                              setState((prev) => ({
                                ...prev,
                                notifications: [
                                  { id: uid('ntf'), createdAt: todayISO(), severity: 'GREEN' as RygStatus, message: 'Settings saved to Dataverse.', recipients: ['PROJECT_MANAGER', 'MANAGER', 'DIRECTOR'] },
                                  ...prev.notifications,
                                ],
                              }))
                            }
                          }
                        }}
                      >
                        <CheckCircle size={14} />
                        Save Settings
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {/* ─── MASTER DATA PAGE ─── */}
              {navPage === 'masterData' && (
                <div className="space-y-4">
                  {/* Tab bar */}
                  <div className="flex gap-1 rounded-lg bg-pth-subtle p-1">
                    {([['people', 'Team'], ['templates', 'Task Templates'], ['customers', 'Customers'], ['suppliers', 'Suppliers']] as const).map(([key, label]) => (
                      <button
                        key={key}
                        type="button"
                        className={`relative flex-1 rounded-md px-4 py-2 text-sm font-medium transition-all ${
                          masterDataTab === key ? 'text-pth-blue' : 'text-pth-muted hover:text-pth-text'
                        }`}
                        onClick={() => setMasterDataTab(key)}
                      >
                        {masterDataTab === key && (
                          <motion.div layoutId="masterdata-tab" className="absolute inset-0 rounded-md bg-pth-card shadow-sm" transition={{ type: 'spring', duration: 0.35, bounce: 0.15 }} />
                        )}
                        <span className="relative z-10">{label}</span>
                      </button>
                    ))}
                  </div>

                  {/* People / Team Tab */}
                  {masterDataTab === 'people' && (() => {
                    const areaList = Array.from(new Set(people.map((p) => p.area).filter(Boolean) as string[])).sort()
                    const filtered = people.filter((p) => {
                      if (teamSiteFilter !== 'all' && !(p.location ?? []).includes(teamSiteFilter)) return false
                      if (teamAreaFilter !== 'all' && (p.area ?? '') !== teamAreaFilter) return false
                      return true
                    })
                    const grouped = new Map<string, Person[]>()
                    for (const p of filtered) {
                      const key = p.area || '—'
                      if (!grouped.has(key)) grouped.set(key, [])
                      grouped.get(key)!.push(p)
                    }
                    const groups = Array.from(grouped).sort((a, b) =>
                      a[0] === '—' ? 1 : b[0] === '—' ? -1 : a[0].localeCompare(b[0]))
                    for (const [, list] of groups) list.sort((a, b) => a.name.localeCompare(b.name))
                    return (
                    <div className="glass-card rounded-xl border border-pth-border/20 shadow-card">
                      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-pth-border/15 px-5 py-4">
                        <div>
                          <h2 className="text-base font-semibold tracking-tight">Team Members</h2>
                          <p className="mt-0.5 text-xs text-pth-muted">{filtered.length} of {people.length} members</p>
                        </div>
                        <div className="flex items-center gap-2">
                          <select value={teamSiteFilter} onChange={(e) => setTeamSiteFilter(e.target.value as 'all' | 'SlpP' | 'TlP')} className="rounded-lg border border-pth-border/30 bg-pth-card px-2.5 py-1.5 text-xs font-medium text-pth-text">
                            <option value="all">All sites</option>
                            <option value="SlpP">SlpP</option>
                            <option value="TlP">TlP</option>
                          </select>
                          <select value={teamAreaFilter} onChange={(e) => setTeamAreaFilter(e.target.value)} className="rounded-lg border border-pth-border/30 bg-pth-card px-2.5 py-1.5 text-xs font-medium text-pth-text">
                            <option value="all">All areas</option>
                            {areaList.map((a) => <option key={a} value={a}>{a}</option>)}
                          </select>
                          <button type="button" onClick={openAddPerson} className="inline-flex items-center gap-1.5 rounded-lg bg-pth-btn px-3 py-2 text-xs font-medium text-white shadow-sm transition-colors hover:bg-pth-btn-hover">
                            <Plus size={14} /> Add Member
                          </button>
                        </div>
                      </div>
                      {groups.map(([area, list]) => (
                        <div key={area}>
                          <div className="flex items-center justify-between bg-pth-subtle/60 px-5 py-1.5">
                            <span className="text-[11px] font-semibold uppercase tracking-wide text-pth-muted">{area === '—' ? 'No area' : area}</span>
                            <span className="text-[11px] text-pth-muted">{list.length}</span>
                          </div>
                          <div className="divide-y divide-pth-border/10">
                            {list.map((p) => (
                          <div key={p.id} className="group flex items-center gap-4 px-5 py-3 transition-colors hover:bg-pth-hover/50">
                            <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-full bg-pth-btn/15 text-xs font-bold text-pth-text">
                              {p.initials}
                            </div>
                            <div className="min-w-0 flex-1">
                              <p className="text-sm font-medium text-pth-text truncate">{p.name}</p>
                              <p className="text-[11px] text-pth-muted">
                                {p.resourceKind ?? ROLE_LABEL[p.role as Role] ?? 'Resource'}
                                {p.managerId && ` · Reports to ${people.find((m) => m.id === p.managerId)?.name ?? '—'}`}
                              </p>
                            </div>
                            <div className="hidden shrink-0 items-center gap-1 sm:flex">
                              {(p.location ?? []).map((loc) => (
                                <span key={loc} className="rounded-full bg-pth-blue/10 px-2 py-0.5 text-[10px] font-medium text-pth-blue">{loc}</span>
                              ))}
                            </div>
                            <span className={`shrink-0 rounded-full px-2.5 py-0.5 text-[11px] font-medium ${
                              p.resourceKind === 'Manager' ? 'bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-300'
                                : p.role === 'DIRECTOR' ? 'bg-purple-100 text-purple-700 dark:bg-purple-900/30 dark:text-purple-300'
                                : p.role === 'MANAGER' ? 'bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-300'
                                : p.role === 'PROJECT_MANAGER' ? 'bg-teal-100 text-teal-700 dark:bg-teal-900/30 dark:text-teal-300'
                                : 'bg-gray-100 text-gray-600 dark:bg-gray-800 dark:text-gray-300'
                            }`}>
                              {p.resourceKind === 'Manager' ? 'Manager' : p.role === 'DIRECTOR' ? 'Director' : p.role === 'MANAGER' ? 'Manager' : p.role === 'PROJECT_MANAGER' ? 'PM' : 'Associate'}
                            </span>
                            <div className="flex items-center gap-1 opacity-0 transition-opacity group-hover:opacity-100">
                              <button type="button" title="Edit member" onClick={() => openEditPerson(p)} className="inline-flex h-8 w-8 items-center justify-center rounded-lg text-pth-muted transition-colors hover:bg-pth-hover hover:text-pth-blue">
                                <Pencil size={14} />
                              </button>
                              <button type="button" title="Delete member" onClick={() => deletePerson(p.id)} className="inline-flex h-8 w-8 items-center justify-center rounded-lg text-pth-muted transition-colors hover:bg-pth-hover hover:text-pth-red">
                                <Trash2 size={14} />
                              </button>
                            </div>
                          </div>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                    )
                  })()}

                  {/* Task Templates Tab */}
                  {masterDataTab === 'templates' && (() => {
                    const types = Array.from(new Set(taskTemplates.map((t) => t.projectType))).sort()
                    const activeType = ttType || types[0] || ''
                    const rows = taskTemplates.filter((t) => t.projectType === activeType).sort((a, b) => a.sequence - b.sequence)
                    return (
                    <div className="space-y-4">
                      <div className="flex flex-wrap items-center gap-2">
                        <span className="text-xs font-medium text-pth-muted">Project type:</span>
                        <select title="Project type" value={activeType} onChange={(e) => setTtType(e.target.value)} className="h-9 rounded-lg border border-pth-border/40 bg-pth-card px-3 text-sm outline-none focus:border-pth-blue">
                          {types.map((t) => <option key={t} value={t}>{t}</option>)}
                        </select>
                        <button type="button" onClick={openAddTemplate} className="ml-auto inline-flex items-center gap-1.5 rounded-lg bg-pth-btn px-3 py-2 text-xs font-medium text-white shadow-sm transition-colors hover:bg-pth-btn-hover">
                          <Plus size={14} /> Add Task
                        </button>
                      </div>
                      <div className="glass-card overflow-hidden rounded-xl border border-pth-border/20 shadow-card">
                        <div className="flex items-center justify-between border-b border-pth-border/15 px-5 py-3">
                          <h2 className="text-base font-semibold tracking-tight">{activeType || 'Task Templates'}</h2>
                          <span className="text-xs text-pth-muted">{rows.length} tasks</span>
                        </div>
                        <div className="overflow-x-auto">
                          <table className="w-full text-sm">
                            <thead>
                              <tr className="border-b border-pth-border/15 text-left text-[11px] uppercase tracking-wide text-pth-muted">
                                <th className="w-10 px-3 py-2 font-semibold">#</th>
                                <th className="px-3 py-2 font-semibold">Task</th>
                                <th className="px-3 py-2 font-semibold">Responsible</th>
                                <th className="whitespace-nowrap px-3 py-2 font-semibold">Leadtime (w)</th>
                                <th className="whitespace-nowrap px-3 py-2 font-semibold">Workload %</th>
                                <th className="px-3 py-2 font-semibold">Inputs</th>
                                <th className="w-16 px-3 py-2 font-semibold"></th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-pth-border/10">
                              {rows.map((t) => (
                                <tr key={t.id} className="group transition-colors hover:bg-pth-hover/50">
                                  <td className="px-3 py-2 text-pth-muted">{t.sequence}</td>
                                  <td className="px-3 py-2 font-medium">{t.task}</td>
                                  <td className="px-3 py-2 text-pth-muted">{t.responsible || '—'}</td>
                                  <td className="px-3 py-2 text-pth-muted">{t.leadtimeWeeks ?? '—'}</td>
                                  <td className="px-3 py-2 text-pth-muted">{t.workloadPct ?? '—'}</td>
                                  <td className="max-w-[280px] truncate px-3 py-2 text-pth-muted" title={t.inputs}>{t.inputs || '—'}</td>
                                  <td className="px-3 py-2">
                                    <div className="flex items-center gap-1 opacity-0 transition-opacity group-hover:opacity-100">
                                      <button type="button" title="Edit task" onClick={() => openEditTemplate(t)} className="inline-flex h-7 w-7 items-center justify-center rounded-lg text-pth-muted hover:bg-pth-hover hover:text-pth-blue"><Pencil size={13} /></button>
                                      <button type="button" title="Delete task" onClick={() => deleteTemplate(t.id)} className="inline-flex h-7 w-7 items-center justify-center rounded-lg text-pth-muted hover:bg-pth-hover hover:text-pth-red"><Trash2 size={13} /></button>
                                    </div>
                                  </td>
                                </tr>
                              ))}
                              {rows.length === 0 && <tr><td colSpan={7} className="px-3 py-8 text-center text-pth-muted">No tasks for this type yet.</td></tr>}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>
                    )
                  })()}

                  {/* Customers Tab */}
                  {masterDataTab === 'customers' && (
                    <div className="glass-card rounded-xl border border-pth-border/20 shadow-card">
                      <div className="flex items-center justify-between border-b border-pth-border/15 px-5 py-4">
                        <div>
                          <h2 className="text-base font-semibold tracking-tight">Customers</h2>
                          <p className="mt-0.5 text-xs text-pth-muted">{clients.length} registered customers</p>
                        </div>
                        <button type="button" onClick={() => {
                          const name = window.prompt('Customer name:')
                          if (name) addPartner('clients', name)
                        }} className="inline-flex items-center gap-1.5 rounded-lg bg-pth-btn px-3 py-2 text-xs font-medium text-white shadow-sm transition-colors hover:bg-pth-btn-hover">
                          <Plus size={14} /> Add Customer
                        </button>
                      </div>
                      <div className="divide-y divide-pth-border/10">
                        {clients.map((c) => (
                          <div key={c.id} className="group flex items-center gap-4 px-5 py-3 transition-colors hover:bg-pth-hover/50">
                            <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-full bg-emerald-100 text-xs font-bold text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-300">
                              {c.name.split(' ').map((w) => w[0]).join('').slice(0, 2)}
                            </div>
                            <div className="min-w-0 flex-1">
                              <p className="text-sm font-medium text-pth-text">{c.name}</p>
                              <p className="text-[11px] text-pth-muted">Customer</p>
                            </div>
                            <span className="shrink-0 rounded-full bg-emerald-100 px-2.5 py-0.5 text-[11px] font-medium text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-300">
                              Customer
                            </span>
                            <div className="flex items-center gap-1 opacity-0 transition-opacity group-hover:opacity-100">
                              <button type="button" title="Edit customer" onClick={() => { const name = window.prompt('Customer name:', c.name); if (name) editPartner('clients', c.id, name) }} className="inline-flex h-8 w-8 items-center justify-center rounded-lg text-pth-muted transition-colors hover:bg-pth-hover hover:text-pth-blue">
                                <Pencil size={14} />
                              </button>
                              <button type="button" title="Delete customer" onClick={() => deletePartner('clients', c.id)} className="inline-flex h-8 w-8 items-center justify-center rounded-lg text-pth-muted transition-colors hover:bg-pth-hover hover:text-pth-red">
                                <Trash2 size={14} />
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Suppliers Tab */}
                  {masterDataTab === 'suppliers' && (
                    <div className="glass-card rounded-xl border border-pth-border/20 shadow-card">
                      <div className="flex items-center justify-between border-b border-pth-border/15 px-5 py-4">
                        <div>
                          <h2 className="text-base font-semibold tracking-tight">Suppliers</h2>
                          <p className="mt-0.5 text-xs text-pth-muted">{suppliers.length} registered suppliers</p>
                        </div>
                        <button type="button" onClick={() => {
                          const name = window.prompt('Supplier name:')
                          if (name) addPartner('suppliers', name)
                        }} className="inline-flex items-center gap-1.5 rounded-lg bg-pth-btn px-3 py-2 text-xs font-medium text-white shadow-sm transition-colors hover:bg-pth-btn-hover">
                          <Plus size={14} /> Add Supplier
                        </button>
                      </div>
                      <div className="divide-y divide-pth-border/10">
                        {suppliers.map((s) => (
                          <div key={s.id} className="group flex items-center gap-4 px-5 py-3 transition-colors hover:bg-pth-hover/50">
                            <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-full bg-amber-100 text-xs font-bold text-amber-700 dark:bg-amber-900/30 dark:text-amber-300">
                              {s.name.split(' ').map((w) => w[0]).join('').slice(0, 2)}
                            </div>
                            <div className="min-w-0 flex-1">
                              <p className="text-sm font-medium text-pth-text">{s.name}</p>
                              <p className="text-[11px] text-pth-muted">Supplier</p>
                            </div>
                            <span className="shrink-0 rounded-full bg-amber-100 px-2.5 py-0.5 text-[11px] font-medium text-amber-700 dark:bg-amber-900/30 dark:text-amber-300">
                              Supplier
                            </span>
                            <div className="flex items-center gap-1 opacity-0 transition-opacity group-hover:opacity-100">
                              <button type="button" title="Edit supplier" onClick={() => { const name = window.prompt('Supplier name:', s.name); if (name) editPartner('suppliers', s.id, name) }} className="inline-flex h-8 w-8 items-center justify-center rounded-lg text-pth-muted transition-colors hover:bg-pth-hover hover:text-pth-blue">
                                <Pencil size={14} />
                              </button>
                              <button type="button" title="Delete supplier" onClick={() => deletePartner('suppliers', s.id)} className="inline-flex h-8 w-8 items-center justify-center rounded-lg text-pth-muted transition-colors hover:bg-pth-hover hover:text-pth-red">
                                <Trash2 size={14} />
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              )}

              {/* ── Create New Project (full-page form) ── */}
              {navPage === 'createProject' && (
                <div className="space-y-6">
                  {/* Header */}
                  <div>
                    <button type="button" className="inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-sm font-medium text-pth-muted transition-colors hover:bg-pth-hover hover:text-pth-text" onClick={() => setNavPage('overview')}>
                      <ChevronLeft size={16} /> Back to Dashboard
                    </button>
                    <h1 className="mt-3 text-xl font-bold md:text-2xl">Create New Project</h1>
                    <p className="mt-1 text-sm text-pth-muted">Standard milestones will be generated automatically based on category.</p>
                  </div>

                  {/* Form card */}
                  <div className="glass-card rounded-xl border border-pth-border/20 p-6 shadow-card">
                    <h2 className="mb-5 text-sm font-semibold uppercase tracking-wider text-pth-muted">Project Information</h2>
                    <form className="space-y-5" onSubmit={newProjectForm.handleSubmit(handleCreateProject)}>
                      {/* Row 1: Name + Client Code */}
                      <div className="grid gap-4 md:grid-cols-2">
                        <label className="space-y-1.5">
                          <span className="text-xs font-medium text-pth-muted">Project Name <span className="text-pth-red">*</span></span>
                          <input className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" placeholder="e.g. Brake Module Redesign" {...newProjectForm.register('name', { required: true })} />
                        </label>
                        <label className="space-y-1.5">
                          <span className="text-xs font-medium text-pth-muted">Client Code</span>
                          <input className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" placeholder="e.g. BC-007" {...newProjectForm.register('clientCode')} />
                        </label>
                      </div>

                      {/* Row 2: Category + Budget */}
                      <div className="grid gap-4 md:grid-cols-2">
                        <label className="space-y-1.5">
                          <span className="text-xs font-medium text-pth-muted">Category <span className="text-pth-red">*</span></span>
                          <select className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...newProjectForm.register('category')}>
                            {(['ECR', 'CIP', 'Path Forward', 'New Programs'] as ProjectCategory[]).map((c) => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </label>
                        <label className="space-y-1.5">
                          <span className="text-xs font-medium text-pth-muted">Budget Allocated</span>
                          <input type="number" min={0} className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" placeholder="0" {...newProjectForm.register('budgetAllocated', { valueAsNumber: true })} />
                        </label>
                      </div>

                      {/* Project Type — drives standard task generation */}
                      <label className="block space-y-1.5">
                        <span className="text-xs font-medium text-pth-muted">Project Type <span className="text-[10px] font-normal text-pth-muted">(auto-creates standard tasks)</span></span>
                        <select className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...newProjectForm.register('projectType')}>
                          <option value="">— None (no standard tasks) —</option>
                          {Array.from(new Set(taskTemplates.map((t) => t.projectType))).sort().map((t) => {
                            const n = taskTemplates.filter((x) => x.projectType === t).length
                            return <option key={t} value={t}>{t} ({n} tasks)</option>
                          })}
                        </select>
                      </label>

                      {/* Project Objective */}
                      <label className="block space-y-1.5">
                        <span className="text-xs font-medium text-pth-muted">Project Objective <span className="text-pth-red">*</span></span>
                        <textarea className="min-h-[80px] w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 py-2.5 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" placeholder="Describe the project objective…" {...newProjectForm.register('objective', { required: true })} />
                      </label>

                      {/* Site Locations — checkboxes */}
                      <div className="space-y-2">
                        <span className="text-xs font-medium text-pth-muted">Site Locations</span>
                        <div className="flex flex-wrap gap-4">
                          {sites.map((s) => {
                            const checked = newProjectForm.watch('siteIds').includes(s.id)
                            return (
                              <label key={s.id} className="flex items-center gap-2 text-sm">
                                <input type="checkbox" checked={checked} onChange={(e) => { const cur = newProjectForm.getValues('siteIds'); newProjectForm.setValue('siteIds', e.target.checked ? [...cur, s.id] : cur.filter((v) => v !== s.id)) }} className="h-4 w-4 rounded border-pth-border/40 text-pth-blue focus:ring-pth-blue/20" />
                                {s.name}
                              </label>
                            )
                          })}
                        </div>
                      </div>

                      {/* Executive Sponsor + Project Manager */}
                      <div className="grid gap-4 md:grid-cols-2">
                        <label className="space-y-1.5">
                          <span className="text-xs font-medium text-pth-muted">Executive Sponsor</span>
                          <select className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...newProjectForm.register('sponsorExecutiveId')}>
                            <option value="">Select…</option>
                            {dvConnected && dvUsers.length > 0
                              ? dvUsers.map((u) => <option key={u.id} value={u.id}>{u.fullname}</option>)
                              : people.filter((p) => p.role === 'DIRECTOR').map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                          </select>
                        </label>
                        <label className="space-y-1.5">
                          <span className="text-xs font-medium text-pth-muted">Project Manager</span>
                          <select className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...newProjectForm.register('projectManagerId')}>
                            <option value="">Select…</option>
                            {dvConnected && dvUsers.length > 0
                              ? dvUsers.map((u) => <option key={u.id} value={u.id}>{u.fullname}</option>)
                              : people.filter((p) => p.role === 'PROJECT_MANAGER').map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                          </select>
                        </label>
                      </div>

                      {/* Planned Start + End dates */}
                      <div className="grid gap-4 md:grid-cols-2">
                        <label className="space-y-1.5">
                          <span className="text-xs font-medium text-pth-muted">Planned Start Date</span>
                          <input type="date" className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...newProjectForm.register('plannedStartDate')} />
                        </label>
                        <label className="space-y-1.5">
                          <span className="text-xs font-medium text-pth-muted">Planned End Date</span>
                          <input type="date" className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...newProjectForm.register('plannedEndDate')} />
                        </label>
                      </div>

                      {/* Customers — checkboxes */}
                      <div className="space-y-2">
                        <span className="text-xs font-medium text-pth-muted">Customers</span>
                        <div className="flex flex-wrap gap-4">
                          {clients.map((c) => {
                            const checked = newProjectForm.watch('clientIds').includes(c.id)
                            return (
                              <label key={c.id} className="flex items-center gap-2 text-sm">
                                <input type="checkbox" checked={checked} onChange={(e) => { const cur = newProjectForm.getValues('clientIds'); newProjectForm.setValue('clientIds', e.target.checked ? [...cur, c.id] : cur.filter((v) => v !== c.id)) }} className="h-4 w-4 rounded border-pth-border/40 text-pth-blue focus:ring-pth-blue/20" />
                                {c.name}
                              </label>
                            )
                          })}
                        </div>
                      </div>

                      {/* Suppliers — checkboxes */}
                      <div className="space-y-2">
                        <span className="text-xs font-medium text-pth-muted">Suppliers</span>
                        <div className="flex flex-wrap gap-4">
                          {suppliers.map((s) => {
                            const checked = newProjectForm.watch('supplierIds').includes(s.id)
                            return (
                              <label key={s.id} className="flex items-center gap-2 text-sm">
                                <input type="checkbox" checked={checked} onChange={(e) => { const cur = newProjectForm.getValues('supplierIds'); newProjectForm.setValue('supplierIds', e.target.checked ? [...cur, s.id] : cur.filter((v) => v !== s.id)) }} className="h-4 w-4 rounded border-pth-border/40 text-pth-blue focus:ring-pth-blue/20" />
                                {s.name}
                              </label>
                            )
                          })}
                        </div>
                      </div>

                      {/* ── Milestones & Tasks (optional, inline) ── */}
                      <div className="space-y-3 rounded-lg border border-pth-border/20 bg-pth-subtle/40 p-4">
                        <div className="flex items-center justify-between">
                          <div>
                            <span className="text-xs font-semibold uppercase tracking-wider text-pth-muted">Milestones & Tasks</span>
                            <p className="text-[11px] text-pth-muted">Add milestones and assign tasks now, or do it later after creation.</p>
                          </div>
                          <button
                            type="button"
                            className="inline-flex items-center gap-1.5 rounded-lg border border-pth-border/40 bg-pth-card px-3 py-1.5 text-xs font-medium text-pth-text shadow-sm transition-colors hover:bg-pth-hover"
                            onClick={() =>
                              setInlineMilestones((prev) => [
                                ...prev,
                                { _key: uid('ims'), name: '', targetDate: addDays(todayISO(), 14), tasks: [] },
                              ])
                            }
                          >
                            <Plus size={12} /> Add Milestone
                          </button>
                        </div>

                        {inlineMilestones.length === 0 && (
                          <p className="text-center text-xs text-pth-muted py-2">No milestones added yet — standard milestones for the selected category will still be generated automatically.</p>
                        )}

                        {inlineMilestones.map((im, mIdx) => (
                          <div key={im._key} className="space-y-2 rounded-lg border border-pth-border/30 bg-pth-card p-3">
                            {/* Milestone header row */}
                            <div className="flex items-start gap-2">
                              <div className="flex-1 grid gap-2 md:grid-cols-2">
                                <input
                                  type="text"
                                  placeholder="Milestone name"
                                  className="h-9 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20"
                                  value={im.name}
                                  onChange={(e) =>
                                    setInlineMilestones((prev) =>
                                      prev.map((m, i) => (i === mIdx ? { ...m, name: e.target.value } : m)),
                                    )
                                  }
                                />
                                <input
                                  type="date"
                                  title="Target date"
                                  className="h-9 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20"
                                  value={im.targetDate}
                                  onChange={(e) =>
                                    setInlineMilestones((prev) =>
                                      prev.map((m, i) => (i === mIdx ? { ...m, targetDate: e.target.value } : m)),
                                    )
                                  }
                                />
                              </div>
                              <button
                                type="button"
                                title="Remove milestone"
                                className="mt-1 rounded-lg p-1 text-pth-muted transition-colors hover:bg-red-50 hover:text-pth-red dark:hover:bg-red-900/20"
                                onClick={() => setInlineMilestones((prev) => prev.filter((_, i) => i !== mIdx))}
                              >
                                <Trash2 size={14} />
                              </button>
                            </div>

                            {/* Tasks under this milestone */}
                            {im.tasks.length > 0 && (
                              <div className="ml-4 space-y-1.5 border-l-2 border-pth-blue/20 pl-3">
                                {im.tasks.map((t, tIdx) => (
                                  <div key={t._key} className="flex items-center gap-2">
                                    <input
                                      type="text"
                                      placeholder="Task name"
                                      className="h-8 min-w-0 flex-1 rounded border border-pth-border/40 bg-pth-subtle px-2 text-xs transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20"
                                      value={t.name}
                                      onChange={(e) =>
                                        setInlineMilestones((prev) =>
                                          prev.map((m, mi) =>
                                            mi === mIdx ? { ...m, tasks: m.tasks.map((tk, ti) => (ti === tIdx ? { ...tk, name: e.target.value } : tk)) } : m,
                                          ),
                                        )
                                      }
                                    />
                                    <select
                                      title="Owner"
                                      className="h-8 w-36 rounded border border-pth-border/40 bg-pth-subtle px-2 text-xs transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20"
                                      value={t.ownerId}
                                      onChange={(e) =>
                                        setInlineMilestones((prev) =>
                                          prev.map((m, mi) =>
                                            mi === mIdx ? { ...m, tasks: m.tasks.map((tk, ti) => (ti === tIdx ? { ...tk, ownerId: e.target.value } : tk)) } : m,
                                          ),
                                        )
                                      }
                                    >
                                      <option value="">Owner…</option>
                                      {people.filter((p) => p.role === 'RESOURCE').map((p) => (
                                        <option key={p.id} value={p.id}>{p.name}</option>
                                      ))}
                                    </select>
                                    <input
                                      type="date"
                                      title="Start"
                                      className="h-8 w-32 rounded border border-pth-border/40 bg-pth-subtle px-2 text-xs transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20"
                                      value={t.startDate}
                                      onChange={(e) =>
                                        setInlineMilestones((prev) =>
                                          prev.map((m, mi) =>
                                            mi === mIdx ? { ...m, tasks: m.tasks.map((tk, ti) => (ti === tIdx ? { ...tk, startDate: e.target.value } : tk)) } : m,
                                          ),
                                        )
                                      }
                                    />
                                    <input
                                      type="date"
                                      title="End"
                                      className="h-8 w-32 rounded border border-pth-border/40 bg-pth-subtle px-2 text-xs transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20"
                                      value={t.endDate}
                                      onChange={(e) =>
                                        setInlineMilestones((prev) =>
                                          prev.map((m, mi) =>
                                            mi === mIdx ? { ...m, tasks: m.tasks.map((tk, ti) => (ti === tIdx ? { ...tk, endDate: e.target.value } : tk)) } : m,
                                          ),
                                        )
                                      }
                                    />
                                    <button
                                      type="button"
                                      title="Remove task"
                                      className="rounded p-0.5 text-pth-muted transition-colors hover:text-pth-red"
                                      onClick={() =>
                                        setInlineMilestones((prev) =>
                                          prev.map((m, mi) =>
                                            mi === mIdx ? { ...m, tasks: m.tasks.filter((_, ti) => ti !== tIdx) } : m,
                                          ),
                                        )
                                      }
                                    >
                                      <X size={12} />
                                    </button>
                                  </div>
                                ))}
                              </div>
                            )}

                            {/* Add task button */}
                            <button
                              type="button"
                              className="ml-4 inline-flex items-center gap-1 rounded px-2 py-1 text-[11px] font-medium text-pth-blue transition-colors hover:bg-pth-blue/5"
                              onClick={() =>
                                setInlineMilestones((prev) =>
                                  prev.map((m, i) =>
                                    i === mIdx
                                      ? {
                                          ...m,
                                          tasks: [
                                            ...m.tasks,
                                            {
                                              _key: uid('itk'),
                                              name: '',
                                              ownerId: people.filter((p) => p.role === 'RESOURCE')[0]?.id ?? '',
                                              startDate: todayISO(),
                                              endDate: addDays(todayISO(), 7),
                                            },
                                          ],
                                        }
                                      : m,
                                  ),
                                )
                              }
                            >
                              <Plus size={10} /> Add Task
                            </button>
                          </div>
                        ))}
                      </div>

                      {isDataverseConfigured() && (
                        <p className="flex items-center gap-1.5 text-[11px] text-emerald-600 dark:text-emerald-400">
                          <Database size={12} /> This project will be saved to Dataverse
                        </p>
                      )}

                      {/* Action buttons */}
                      <div className="flex items-center justify-end gap-3 border-t border-pth-border/15 pt-5">
                        <button type="button" onClick={() => { newProjectForm.reset(); setInlineMilestones([]); setNavPage('overview') }} className="rounded-lg border border-pth-border/40 bg-pth-card px-5 py-2.5 text-sm font-medium text-pth-text shadow-sm transition-colors hover:bg-pth-hover">
                          Cancel
                        </button>
                        <button type="submit" disabled={creatingSaving} className="rounded-lg bg-pth-btn px-5 py-2.5 text-sm font-semibold text-white shadow-sm transition-all hover:bg-pth-btn-hover active:scale-[0.98] inline-flex items-center justify-center gap-2 disabled:opacity-60 disabled:cursor-not-allowed">
                          {creatingSaving ? (
                            <><span className="h-4 w-4 animate-spin rounded-full border-2 border-white/30 border-t-white" /> Creating…</>
                          ) : (
                            <><Plus size={14} /> Create Project</>
                          )}
                        </button>
                      </div>
                    </form>
                  </div>
                </div>
              )}

            </motion.section>
          </AnimatePresence>
        </main>
      </div>

      {/* ─── EDIT PROJECT DIALOG ─── */}
      <AnimatePresence>
        {editingProject && selectedProject && (
          <>
            <motion.div className="fixed inset-0 z-50 bg-black/30 backdrop-blur-sm" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} onClick={() => setEditingProject(false)} />
            <motion.div
              className="glass-card fixed left-1/2 top-1/2 z-[60] w-[92vw] max-w-lg -translate-x-1/2 -translate-y-1/2 rounded-2xl border border-pth-border/20 shadow-elevated"
              initial={{ opacity: 0, y: 16, scale: 0.96 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: 16, scale: 0.96 }}
              transition={{ duration: 0.2 }}
            >
              <div className="flex items-center justify-between border-b border-pth-border/15 px-5 py-4">
                <h3 className="text-sm font-semibold">Edit Project</h3>
                <button type="button" title="Close dialog" className="rounded-lg p-1.5 text-pth-muted transition-colors hover:bg-pth-hover" onClick={() => setEditingProject(false)}>
                  <X size={16} />
                </button>
              </div>
              <form className="max-h-[65vh] overflow-y-auto p-5 space-y-4" onSubmit={overviewForm.handleSubmit((v) => { saveOverview(v); setEditingProject(false) })}>
                <div className="grid gap-4 md:grid-cols-2">
                  <label className="space-y-1.5">
                    <span className="text-xs font-medium text-pth-muted">Project Name</span>
                    <input className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...overviewForm.register('name', { required: true })} />
                  </label>
                  <label className="space-y-1.5">
                    <span className="text-xs font-medium text-pth-muted">Client Code</span>
                    <input className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...overviewForm.register('clientCode')} />
                  </label>
                  <label className="space-y-1.5">
                    <span className="text-xs font-medium text-pth-muted">Category</span>
                    <select className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...overviewForm.register('category')}>
                      {(['ECR', 'CIP', 'Path Forward', 'New Programs'] as ProjectCategory[]).map((c) => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </label>
                  <label className="space-y-1.5">
                    <span className="text-xs font-medium text-pth-muted">Budget Allocated</span>
                    <input type="number" className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...overviewForm.register('budgetAllocated', { valueAsNumber: true })} />
                  </label>
                  <div className="space-y-1.5 md:col-span-2">
                    <span className="text-xs font-medium text-pth-muted">Site Locations</span>
                    <div className="flex flex-wrap gap-3">
                      {sites.map((s) => {
                        const checked = overviewForm.watch('siteIds').includes(s.id)
                        return (
                          <label key={s.id} className="flex items-center gap-2 text-sm">
                            <input type="checkbox" checked={checked} onChange={(e) => { const cur = overviewForm.getValues('siteIds'); overviewForm.setValue('siteIds', e.target.checked ? [...cur, s.id] : cur.filter((v) => v !== s.id)) }} className="h-4 w-4 rounded border-pth-border/40 text-pth-blue focus:ring-pth-blue/20" />
                            {s.name}
                          </label>
                        )
                      })}
                    </div>
                  </div>
                  <label className="space-y-1.5">
                    <span className="text-xs font-medium text-pth-muted">Project Manager</span>
                    <select className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...overviewForm.register('projectManagerId')}>
                      {dvConnected && dvUsers.length > 0
                        ? dvUsers.map((u) => <option key={u.id} value={u.id}>{u.fullname}</option>)
                        : people.filter((p) => p.role === 'PROJECT_MANAGER').map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                    </select>
                  </label>
                  <label className="space-y-1.5">
                    <span className="text-xs font-medium text-pth-muted">Sponsor Executive</span>
                    <select className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...overviewForm.register('sponsorExecutiveId')}>
                      {dvConnected && dvUsers.length > 0
                        ? dvUsers.map((u) => <option key={u.id} value={u.id}>{u.fullname}</option>)
                        : people.filter((p) => p.role === 'DIRECTOR').map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                    </select>
                  </label>
                  <div className="space-y-1.5 md:col-span-2">
                    <span className="text-xs font-medium text-pth-muted">Clients</span>
                    <div className="flex flex-wrap gap-3">
                      {clients.map((c) => {
                        const checked = overviewForm.watch('clientIds').includes(c.id)
                        return (
                          <label key={c.id} className="flex items-center gap-2 text-sm">
                            <input type="checkbox" checked={checked} onChange={(e) => { const cur = overviewForm.getValues('clientIds'); overviewForm.setValue('clientIds', e.target.checked ? [...cur, c.id] : cur.filter((v) => v !== c.id)) }} className="h-4 w-4 rounded border-pth-border/40 text-pth-blue focus:ring-pth-blue/20" />
                            {c.name}
                          </label>
                        )
                      })}
                    </div>
                  </div>
                  <div className="space-y-1.5 md:col-span-2">
                    <span className="text-xs font-medium text-pth-muted">Suppliers</span>
                    <div className="flex flex-wrap gap-3">
                      {suppliers.map((s) => {
                        const checked = overviewForm.watch('supplierIds').includes(s.id)
                        return (
                          <label key={s.id} className="flex items-center gap-2 text-sm">
                            <input type="checkbox" checked={checked} onChange={(e) => { const cur = overviewForm.getValues('supplierIds'); overviewForm.setValue('supplierIds', e.target.checked ? [...cur, s.id] : cur.filter((v) => v !== s.id)) }} className="h-4 w-4 rounded border-pth-border/40 text-pth-blue focus:ring-pth-blue/20" />
                            {s.name}
                          </label>
                        )
                      })}
                    </div>
                  </div>
                </div>
                <label className="block space-y-1.5">
                  <span className="text-xs font-medium text-pth-muted">Objective</span>
                  <textarea className="min-h-[80px] w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 py-2.5 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...overviewForm.register('objective', { required: true })} />
                </label>
                <button type="submit" className="w-full rounded-lg bg-pth-btn py-2.5 text-sm font-semibold text-white shadow-sm transition-all hover:bg-pth-btn-hover active:scale-[0.98] inline-flex items-center justify-center gap-2">
                  <Save size={14} /> Save Changes
                </button>
              </form>
            </motion.div>
          </>
        )}
      </AnimatePresence>

      {/* ─── MILESTONE DIALOG ─── */}
      <AnimatePresence>
        {milestoneDialogOpen && (
          <>
            <motion.div className="fixed inset-0 z-50 bg-black/30 backdrop-blur-sm" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} onClick={() => setMilestoneDialogOpen(false)} />
            <motion.div
              className="glass-card fixed left-1/2 top-1/2 z-[60] w-[92vw] max-w-md -translate-x-1/2 -translate-y-1/2 rounded-2xl border border-pth-border/20 shadow-elevated"
              initial={{ opacity: 0, y: 16, scale: 0.96 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: 16, scale: 0.96 }}
              transition={{ duration: 0.2 }}
            >
              <div className="flex items-center justify-between border-b border-pth-border/15 px-5 py-4">
                <h3 className="text-sm font-semibold">{milestoneEditId ? 'Edit Milestone' : 'Create Milestone'}</h3>
                <button type="button" title="Close dialog" className="rounded-lg p-1.5 text-pth-muted transition-colors hover:bg-pth-hover" onClick={() => setMilestoneDialogOpen(false)}>
                  <X size={16} />
                </button>
              </div>
              <form className="space-y-4 p-5" onSubmit={milestoneForm.handleSubmit(submitMilestone)}>
                <label className="block space-y-1.5">
                  <span className="text-xs font-medium text-pth-muted">Name</span>
                  <input className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...milestoneForm.register('name', { required: true })} />
                </label>
                <label className="block space-y-1.5">
                  <span className="text-xs font-medium text-pth-muted">Target Date</span>
                  <input type="date" className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...milestoneForm.register('targetDate', { required: true })} />
                </label>
                <button type="submit" className="w-full rounded-lg bg-pth-btn py-2.5 text-sm font-semibold text-white shadow-sm transition-all hover:bg-pth-btn-hover active:scale-[0.98]">
                  {milestoneEditId ? 'Update' : 'Create'} Milestone
                </button>
              </form>
            </motion.div>
          </>
        )}
      </AnimatePresence>

      {/* ─── ADD TASK DIALOG ─── */}
      <AnimatePresence>
        {activityDialogOpen && (
          <>
            <motion.div className="fixed inset-0 z-50 bg-black/30 backdrop-blur-sm" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} onClick={() => setActivityDialogOpen(false)} />
            <motion.div
              className="glass-card fixed left-1/2 top-1/2 z-[60] w-[92vw] max-w-md -translate-x-1/2 -translate-y-1/2 rounded-2xl border border-pth-border/20 shadow-elevated"
              initial={{ opacity: 0, y: 16, scale: 0.96 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: 16, scale: 0.96 }}
              transition={{ duration: 0.2 }}
            >
              <div className="flex items-center justify-between border-b border-pth-border/15 px-5 py-4">
                <h3 className="text-sm font-semibold">{activityEditMode === 'edit' ? 'Edit Task' : 'Add Task'}</h3>
                <button type="button" title="Close dialog" className="rounded-lg p-1.5 text-pth-muted transition-colors hover:bg-pth-hover" onClick={() => setActivityDialogOpen(false)}>
                  <X size={16} />
                </button>
              </div>
              <form className="space-y-4 p-5" onSubmit={activityForm.handleSubmit(submitActivity)}>
                <label className="block space-y-1.5">
                  <span className="text-xs font-medium text-pth-muted">Task Name</span>
                  <input className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...activityForm.register('name', { required: true })} />
                </label>
                <div className="grid grid-cols-2 gap-3">
                  <label className="block space-y-1.5">
                    <span className="text-xs font-medium text-pth-muted">Owner</span>
                    <select className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...activityForm.register('ownerId')}>
                      <option value="">Unassigned</option>
                      {people.filter((p) => p.role === 'RESOURCE').map((p) => <option key={p.id} value={p.id}>{p.name}{p.area ? ` (${p.area})` : ''}</option>)}
                    </select>
                  </label>
                  <label className="block space-y-1.5">
                    <span className="text-xs font-medium text-pth-muted">Status</span>
                    <select className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...activityForm.register('state')}>
                      <option value="NOT_STARTED">Not started</option>
                      <option value="IN_PROGRESS">In progress</option>
                      <option value="DONE">Done</option>
                      <option value="NOT_APPLICABLE">N/A (not applicable)</option>
                    </select>
                  </label>
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <label className="block space-y-1.5">
                    <span className="text-xs font-medium text-pth-muted">Responsible (Area)</span>
                    <input className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" placeholder="e.g. ENG / QMM" {...activityForm.register('responsible')} />
                  </label>
                  <label className="block space-y-1.5">
                    <span className="text-xs font-medium text-pth-muted">Workload %</span>
                    <input type="number" min={0} max={100000} className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" placeholder="0" {...activityForm.register('workloadPct', { valueAsNumber: true })} />
                  </label>
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <label className="block space-y-1.5">
                    <span className="text-xs font-medium text-pth-muted">Start Date</span>
                    <input type="date" className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...activityForm.register('startDate', { required: true })} />
                  </label>
                  <label className="block space-y-1.5">
                    <span className="text-xs font-medium text-pth-muted">End Date</span>
                    <input type="date" className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" {...activityForm.register('endDate', { required: true })} />
                  </label>
                </div>
                <button type="submit" className="w-full rounded-lg bg-pth-btn py-2.5 text-sm font-semibold text-white shadow-sm transition-all hover:bg-pth-btn-hover active:scale-[0.98]">
                  {activityEditMode === 'edit' ? 'Save Changes' : 'Add Task'}
                </button>
              </form>
            </motion.div>
          </>
        )}
      </AnimatePresence>

      {/* ─── MASTER DATA PERSON DIALOG ─── */}
      <AnimatePresence>
        {mdDialogOpen && (
          <>
            <motion.div className="fixed inset-0 z-50 bg-black/30 backdrop-blur-sm" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} onClick={() => setMdDialogOpen(false)} />
            <motion.div
              className="glass-card fixed left-1/2 top-1/2 z-[60] w-[92vw] max-w-md -translate-x-1/2 -translate-y-1/2 rounded-2xl border border-pth-border/20 shadow-elevated"
              initial={{ opacity: 0, y: 16, scale: 0.96 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: 16, scale: 0.96 }}
              transition={{ duration: 0.2 }}
            >
              <div className="flex items-center justify-between border-b border-pth-border/15 px-5 py-4">
                <h3 className="text-sm font-semibold">{mdEditId ? 'Edit Team Member' : 'Add Team Member'}</h3>
                <button type="button" title="Close" onClick={() => setMdDialogOpen(false)} className="inline-flex h-8 w-8 items-center justify-center rounded-lg text-pth-muted hover:bg-pth-hover hover:text-pth-text"><X size={16} /></button>
              </div>
              <div className="space-y-4 p-5">
                {/* ── Full Name: AAD autocomplete when creating, plain input when editing ── */}
                <div className="relative block space-y-1.5">
                  <span className="text-xs font-medium text-pth-muted">Full Name</span>
                  <input
                    type="text"
                    className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20"
                    value={mdEditId ? mdForm.name : aadQuery}
                    onChange={(e) => {
                      if (mdEditId) { setMdForm((f) => ({ ...f, name: e.target.value })) }
                      else { setAadQuery(e.target.value); setMdForm((f) => ({ ...f, name: e.target.value })) }
                    }}
                    placeholder={mdEditId ? 'e.g. John Smith' : 'Search Azure AD / type name...'}
                  />
                  {/* AAD suggestions dropdown (only for create mode) */}
                  {!mdEditId && aadQuery.length >= 2 && (aadLoading || aadResults.length > 0) && (
                    <div className="absolute left-0 right-0 top-full z-10 mt-1 max-h-48 overflow-y-auto rounded-lg border border-pth-border/30 bg-pth-card shadow-lg">
                      {aadLoading && <div className="px-3 py-2 text-xs text-pth-muted">Searching Azure AD...</div>}
                      {!aadLoading && aadResults.map((user) => (
                        <button
                          key={user.id}
                          type="button"
                          className="flex w-full items-start gap-2 px-3 py-2 text-left text-sm transition-colors hover:bg-pth-hover"
                          onClick={() => {
                            setMdForm((f) => ({
                              ...f,
                              name: user.displayName,
                              initials: user.displayName.split(' ').map((w) => w[0]).join('').toUpperCase().slice(0, 2),
                            }))
                            setAadQuery(user.displayName)
                            setAadResults([])
                          }}
                        >
                          <div className="flex h-7 w-7 shrink-0 items-center justify-center rounded-full bg-pth-btn/15 text-[10px] font-bold text-pth-text">
                            {user.displayName.split(' ').map((w) => w[0]).join('').toUpperCase().slice(0, 2)}
                          </div>
                          <div className="min-w-0">
                            <div className="truncate text-sm font-medium text-pth-text">{user.displayName}</div>
                            <div className="truncate text-[11px] text-pth-muted">{[user.jobTitle, user.department, user.mail].filter(Boolean).join(' · ')}</div>
                          </div>
                        </button>
                      ))}
                      {!aadLoading && aadResults.length === 0 && aadQuery.length >= 2 && (
                        <div className="px-3 py-2 text-xs text-pth-muted">No users found — name will be used as-is</div>
                      )}
                    </div>
                  )}
                </div>
                <label className="block space-y-1.5">
                  <span className="text-xs font-medium text-pth-muted">Initials</span>
                  <input type="text" maxLength={3} className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm uppercase transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" value={mdForm.initials} onChange={(e) => setMdForm((f) => ({ ...f, initials: e.target.value.toUpperCase() }))} placeholder="Auto-generated if blank" />
                </label>
                <label className="block space-y-1.5">
                  <span className="text-xs font-medium text-pth-muted">Role</span>
                  <select title="Role" className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" value={mdForm.role} onChange={(e) => setMdForm((f) => ({ ...f, role: e.target.value as Person['role'] }))}>
                    <option value="RESOURCE">Resource</option>
                    <option value="PROJECT_MANAGER">Project Manager</option>
                    <option value="MANAGER">Manager</option>
                    <option value="DIRECTOR">Director</option>
                  </select>
                </label>
                {mdForm.role === 'RESOURCE' && (
                  <>
                    <div className="grid grid-cols-2 gap-3">
                      <label className="block space-y-1.5">
                        <span className="text-xs font-medium text-pth-muted">Kind</span>
                        <select title="Kind" className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" value={mdForm.kind} onChange={(e) => setMdForm((f) => ({ ...f, kind: e.target.value }))}>
                          <option value="Associate">Associate</option>
                          <option value="Manager">Manager</option>
                        </select>
                      </label>
                      <label className="block space-y-1.5">
                        <span className="text-xs font-medium text-pth-muted">Area</span>
                        <select title="Area" className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" value={mdForm.area} onChange={(e) => setMdForm((f) => ({ ...f, area: e.target.value }))}>
                          <option value="">—</option>
                          {['PPS', 'ENG', 'QMM', 'LOD', 'LOP', 'LOP4', 'LOP5', 'MSE', 'CTG', 'PUQ1', 'PUQ2'].map((a) => <option key={a} value={a}>{a}</option>)}
                        </select>
                      </label>
                    </div>
                    <div className="space-y-1.5">
                      <span className="text-xs font-medium text-pth-muted">Location</span>
                      <div className="flex gap-2">
                        {['SlpP', 'TlP'].map((loc) => {
                          const on = mdForm.location.includes(loc)
                          return (
                            <button key={loc} type="button" onClick={() => setMdForm((f) => ({ ...f, location: on ? f.location.filter((l) => l !== loc) : [...f.location, loc] }))}
                              className={`flex-1 rounded-lg border px-3 py-2 text-sm font-medium transition-colors ${on ? 'border-pth-blue bg-pth-blue/10 text-pth-blue' : 'border-pth-border/40 bg-pth-subtle text-pth-muted hover:text-pth-text'}`}>
                              {loc}
                            </button>
                          )
                        })}
                      </div>
                    </div>
                    <label className="block space-y-1.5">
                      <span className="text-xs font-medium text-pth-muted">Weekly Capacity (hours)</span>
                      <input type="number" min={0} max={168} className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" value={mdForm.capacity} onChange={(e) => setMdForm((f) => ({ ...f, capacity: Number(e.target.value) }))} />
                    </label>
                  </>
                )}
                <label className="block space-y-1.5">
                  <span className="text-xs font-medium text-pth-muted">{mdForm.role === 'RESOURCE' ? 'Manager' : 'Reports To'}</span>
                  <select title="Reports To" className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:outline-none focus:ring-2 focus:ring-pth-blue/20" value={mdForm.managerId} onChange={(e) => setMdForm((f) => ({ ...f, managerId: e.target.value }))}>
                    <option value="">None</option>
                    {people.filter((p) => p.role === 'DIRECTOR' || p.role === 'MANAGER' || p.resourceKind === 'Manager').map((p) => (
                      <option key={p.id} value={p.id}>{p.name}{p.area ? ` (${p.area})` : ` (${ROLE_LABEL[p.role as Role] ?? p.role})`}</option>
                    ))}
                  </select>
                </label>
                <button type="button" onClick={submitPerson} className="w-full rounded-lg bg-pth-btn py-2.5 text-sm font-semibold text-white shadow-sm transition-all hover:bg-pth-btn-hover active:scale-[0.98]">
                  {mdEditId ? 'Save Changes' : 'Add Member'}
                </button>
              </div>
            </motion.div>
          </>
        )}
      </AnimatePresence>

      {/* ─── TASK TEMPLATE ADD/EDIT DIALOG ─── */}
      <AnimatePresence>
        {ttDialogOpen && (
          <>
            <motion.div className="fixed inset-0 z-50 bg-black/30 backdrop-blur-sm" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} onClick={() => setTtDialogOpen(false)} />
            <motion.div className="glass-card fixed left-1/2 top-1/2 z-[60] w-[92vw] max-w-md -translate-x-1/2 -translate-y-1/2 rounded-2xl border border-pth-border/20 shadow-elevated" initial={{ opacity: 0, y: 16, scale: 0.96 }} animate={{ opacity: 1, y: 0, scale: 1 }} exit={{ opacity: 0, y: 16, scale: 0.96 }} transition={{ duration: 0.2 }}>
              <div className="flex items-center justify-between border-b border-pth-border/15 px-5 py-4">
                <h3 className="text-sm font-semibold">{ttEditId ? 'Edit Task' : 'Add Task'} — {ttType}</h3>
                <button type="button" title="Close" onClick={() => setTtDialogOpen(false)} className="inline-flex h-8 w-8 items-center justify-center rounded-lg text-pth-muted hover:bg-pth-hover hover:text-pth-text"><X size={16} /></button>
              </div>
              <div className="space-y-4 p-5">
                <label className="block space-y-1.5">
                  <span className="text-xs font-medium text-pth-muted">Task</span>
                  <input type="text" className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" value={ttForm.task} onChange={(e) => setTtForm((f) => ({ ...f, task: e.target.value }))} placeholder="e.g. KO meeting" />
                </label>
                <div className="grid grid-cols-2 gap-3">
                  <label className="block space-y-1.5">
                    <span className="text-xs font-medium text-pth-muted">Sequence</span>
                    <input type="number" className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" value={ttForm.sequence} onChange={(e) => setTtForm((f) => ({ ...f, sequence: Number(e.target.value) }))} />
                  </label>
                  <label className="block space-y-1.5">
                    <span className="text-xs font-medium text-pth-muted">Responsible</span>
                    <input type="text" className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" value={ttForm.responsible} onChange={(e) => setTtForm((f) => ({ ...f, responsible: e.target.value }))} placeholder="e.g. ENG / QMM" />
                  </label>
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <label className="block space-y-1.5">
                    <span className="text-xs font-medium text-pth-muted">Leadtime (weeks)</span>
                    <input type="number" className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" value={ttForm.leadtimeWeeks} onChange={(e) => setTtForm((f) => ({ ...f, leadtimeWeeks: e.target.value }))} placeholder="—" />
                  </label>
                  <label className="block space-y-1.5">
                    <span className="text-xs font-medium text-pth-muted">Workload %</span>
                    <input type="number" className="h-10 w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" value={ttForm.workloadPct} onChange={(e) => setTtForm((f) => ({ ...f, workloadPct: e.target.value }))} placeholder="—" />
                  </label>
                </div>
                <label className="block space-y-1.5">
                  <span className="text-xs font-medium text-pth-muted">Inputs</span>
                  <textarea rows={3} className="w-full rounded-lg border border-pth-border/40 bg-pth-subtle px-3 py-2 text-sm transition-colors focus:border-pth-blue focus:bg-pth-card focus:outline-none focus:ring-2 focus:ring-pth-blue/20" value={ttForm.inputs} onChange={(e) => setTtForm((f) => ({ ...f, inputs: e.target.value }))} placeholder="Required inputs / deliverables" />
                </label>
                <button type="button" onClick={submitTemplate} className="w-full rounded-lg bg-pth-btn py-2.5 text-sm font-semibold text-white shadow-sm transition-all hover:bg-pth-btn-hover active:scale-[0.98]">
                  {ttEditId ? 'Save Changes' : 'Add Task'}
                </button>
              </div>
            </motion.div>
          </>
        )}
      </AnimatePresence>

      {/* ─── DELETE PROJECT CONFIRMATION DIALOG ─── */}
      <AnimatePresence>
        {deleteProjectDialogOpen && selectedProject && (
          <>
            <motion.div className="fixed inset-0 z-50 bg-black/30 backdrop-blur-sm" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} onClick={() => { if (!deletingProject) setDeleteProjectDialogOpen(false) }} />
            <motion.div
              className="glass-card fixed left-1/2 top-1/2 z-[60] w-[92vw] max-w-sm -translate-x-1/2 -translate-y-1/2 rounded-2xl border border-pth-border/20 shadow-elevated"
              initial={{ opacity: 0, y: 16, scale: 0.96 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: 16, scale: 0.96 }}
              transition={{ duration: 0.2 }}
            >
              <div className="flex items-center gap-3 border-b border-pth-border/15 px-5 py-4">
                <div className="flex h-9 w-9 items-center justify-center rounded-full bg-red-100 dark:bg-red-900/30">
                  <Trash2 size={16} className="text-pth-red" />
                </div>
                <h3 className="text-sm font-semibold">Delete Project</h3>
              </div>
              <div className="space-y-3 p-5">
                <p className="text-sm text-pth-muted">
                  Are you sure you want to delete <span className="font-semibold text-pth-text">{selectedProject.name}</span>?
                </p>
                <p className="text-xs text-pth-muted/80">
                  This will permanently remove the project and all its milestones, activities, and assignments. This action cannot be undone.
                </p>
                <div className="flex items-center gap-3 pt-2">
                  <button type="button" disabled={deletingProject} className="flex-1 rounded-lg border border-pth-border/40 bg-pth-card py-2.5 text-sm font-medium text-pth-text shadow-sm transition-colors hover:bg-pth-hover disabled:opacity-50" onClick={() => setDeleteProjectDialogOpen(false)}>
                    Cancel
                  </button>
                  <button type="button" disabled={deletingProject} className="flex-1 rounded-lg bg-pth-red py-2.5 text-sm font-semibold text-white shadow-sm transition-all hover:bg-red-700 active:scale-[0.98] disabled:opacity-50" onClick={handleDeleteProject}>
                    {deletingProject ? 'Deleting…' : 'Delete'}
                  </button>
                </div>
              </div>
            </motion.div>
          </>
        )}
      </AnimatePresence>

      {/* Hidden icons to prevent tree-shaking */}
      <div className="hidden">
        <Briefcase />
        <Clock3 />
        <User />
        <ChevronDown />
      </div>
    </div>
  )
}
