export type Role = 'PROJECT_MANAGER' | 'MANAGER' | 'DIRECTOR'
export type RygStatus = 'RED' | 'YELLOW' | 'GREEN'

export type ViewKey = 'pmStatus' | 'risks' | 'workload' | 'config'

export interface UserProfile {
  id: string
  name: string
  role: Role
  managerUserId?: string
  department: string
  site: string
}

export type ProjectCategory = 'ECR' | 'Path Forward' | 'CIP' | 'New Programs'

export interface ProjectEntity {
  id: string
  name: string
  category: ProjectCategory
  site: string
  objective: string
  sponsorExecutive: string
  projectManagerId: string
  managerId: string
  timeStatus: RygStatus
  criticalPathChangedFlag: boolean
  dependencyVersion: number
  criticalPathVersion: number
}

export interface MilestoneEntity {
  id: string
  projectId: string
  name: string
  plannedDate: string
  status: 'OPEN' | 'IN_PROGRESS' | 'CLOSED' | 'DELAYED'
  ryg: RygStatus
  isCriticalPath: boolean
  dependsOnMilestoneIds: string[]
}

export interface ActivityEntity {
  id: string
  projectId: string
  milestoneId?: string
  name: string
  ownerUserId: string
  startDate: string
  endDate: string
  plannedEndDate: string
  closedDate?: string
  status: 'NOT_STARTED' | 'IN_PROGRESS' | 'BLOCKED' | 'CLOSED'
  ryg: RygStatus
  isCriticalPath: boolean
  dependsOnActivityIds: string[]
}

export interface ResourceEntity {
  personId: string
  name: string
  role: string
  department: string
  weeklyCapacityHours: number
  managerUserId: string
}

export interface AssignmentEntity {
  id: string
  activityId: string
  personId: string
  assignedHours: number
  weekStart: string
}

export interface NotificationEntity {
  id: string
  createdAt: string
  severity: RygStatus
  type:
    | 'EARLY_WARNING_ACTIVITY'
    | 'EARLY_WARNING_MILESTONE'
    | 'WORKLOAD_90'
    | 'WORKLOAD_100'
    | 'CRITICAL_PATH_CHANGED'
    | 'WEEKLY_DELAYED_MILESTONES'
    | 'WEEKLY_DELAYED_ACTIVITIES'
  message: string
  recipientUserIds: string[]
}

export interface AppSettings {
  warningDaysThreshold: number
  reportWeekWindow: number
  defaultWeeklyCapacity: number
  utilizationWarningThreshold: number
  utilizationCriticalThreshold: number
}

export interface AppData {
  users: UserProfile[]
  projects: ProjectEntity[]
  milestones: MilestoneEntity[]
  activities: ActivityEntity[]
  resources: ResourceEntity[]
  assignments: AssignmentEntity[]
  notifications: NotificationEntity[]
  settings: AppSettings
}

interface ProjectCreateInput {
  name: string
  category: ProjectCategory
  site: string
  objective: string
  sponsorExecutive: string
  projectManagerId: string
  managerId: string
}

interface TimelineRow {
  id: string
  name: string
  ownerName: string
  startDate: string
  endDate: string
  isCriticalPath: boolean
  ryg: RygStatus
}

interface ProjectSummaryRow {
  project: ProjectEntity
  managerName: string
  pmName: string
  overallRyg: RygStatus
  timeRyg: RygStatus
  costRyg: RygStatus
  qualityRyg: RygStatus
  criticalPathChanged: boolean
  endDate: string
}

interface MilestonePrioritizationRow {
  id: string
  projectName: string
  name: string
  plannedDate: string
  daysRemaining: number
  ryg: RygStatus
}

interface DelayedActivityRow {
  id: string
  name: string
  ownerName: string
  plannedEndDate: string
  delayDays: number
  status: ActivityEntity['status']
  ryg: RygStatus
  isCriticalPath: boolean
}

interface WorkloadRow {
  weekStart: string
  personId: string
  personName: string
  capacityHours: number
  assignedHours: number
  utilization: number
  utilizationRyg: RygStatus
  relatedProjectNames: string[]
}

const MS_PER_DAY = 24 * 60 * 60 * 1000

const CATEGORY_MILESTONE_TEMPLATES: Record<ProjectCategory, string[]> = {
  ECR: ['Intake Approved', 'Engineering Validation', 'Implementation Complete', 'Benefit Realized'],
  'Path Forward': ['Current-State Review', 'Roadmap Alignment', 'Pilot Execution', 'Rollout Completion'],
  CIP: ['Kaizen Identification', 'Root Cause Complete', 'Countermeasure Implemented', 'Sustainment Audit'],
  'New Programs': ['Program Kickoff', 'Design Freeze', 'Pre-Launch Gate', 'Production Readiness'],
}

function parseDate(value: string): Date {
  return new Date(`${value}T00:00:00`)
}

function isoDate(date: Date): string {
  return date.toISOString().slice(0, 10)
}

function addDays(value: string, days: number): string {
  const date = parseDate(value)
  date.setDate(date.getDate() + days)
  return isoDate(date)
}

function addWeeks(value: string, weeks: number): string {
  return addDays(value, weeks * 7)
}

function daysBetween(start: string, end: string): number {
  const diff = parseDate(end).getTime() - parseDate(start).getTime()
  return Math.floor(diff / MS_PER_DAY)
}

function startOfWeek(date: Date): Date {
  const copy = new Date(date)
  const day = copy.getDay() || 7
  copy.setDate(copy.getDate() - day + 1)
  copy.setHours(0, 0, 0, 0)
  return copy
}

function unique<T>(values: T[]): T[] {
  return [...new Set(values)]
}

function compareRygPriority(left: RygStatus, right: RygStatus): number {
  const rank: Record<RygStatus, number> = { RED: 0, YELLOW: 1, GREEN: 2 }
  return rank[left] - rank[right]
}

function calculateDateRyg(targetDate: string, settings: AppSettings): RygStatus {
  const now = isoDate(new Date())
  const remaining = daysBetween(now, targetDate)
  if (remaining < 0) {
    return 'RED'
  }
  if (remaining <= settings.warningDaysThreshold) {
    return 'YELLOW'
  }
  return 'GREEN'
}

function calculateWorkloadRyg(utilization: number, settings: AppSettings): RygStatus {
  const percentage = utilization * 100
  if (percentage >= settings.utilizationCriticalThreshold) {
    return 'RED'
  }
  if (percentage >= settings.utilizationWarningThreshold) {
    return 'YELLOW'
  }
  return 'GREEN'
}

function canManageProject(project: ProjectEntity, user: UserProfile): boolean {
  if (user.role === 'DIRECTOR') {
    return true
  }
  if (user.role === 'MANAGER') {
    return project.managerId === user.id
  }
  return project.projectManagerId === user.id
}

export function canAccessView(role: Role, view: ViewKey): boolean {
  const roleViews: Record<Role, ViewKey[]> = {
    PROJECT_MANAGER: ['pmStatus'],
    MANAGER: ['pmStatus', 'risks', 'workload'],
    DIRECTOR: ['pmStatus', 'risks', 'workload', 'config'],
  }

  return roleViews[role].includes(view)
}

function buildMilestoneTemplate(projectId: string, category: ProjectCategory, startDate: string): MilestoneEntity[] {
  const names = CATEGORY_MILESTONE_TEMPLATES[category]
  return names.map((name, index) => {
    const plannedDate = addWeeks(startDate, (index + 1) * 4)
    const milestoneId = `${projectId}-MS-${index + 1}`
    return {
      id: milestoneId,
      projectId,
      name,
      plannedDate,
      status: 'OPEN',
      ryg: 'GREEN',
      isCriticalPath: index === names.length - 1,
      dependsOnMilestoneIds: index === 0 ? [] : [`${projectId}-MS-${index}`],
    }
  })
}

function calculateCriticalPathActivities(activities: ActivityEntity[]): Set<string> {
  const byId = new Map(activities.map((activity) => [activity.id, activity]))
  const memo = new Map<string, number>()

  function chainDuration(id: string): number {
    if (memo.has(id)) {
      return memo.get(id) ?? 0
    }

    const current = byId.get(id)
    if (!current) {
      return 0
    }

    const ownDuration = Math.max(1, daysBetween(current.startDate, current.endDate))
    const dependentDurations = current.dependsOnActivityIds.map((parentId) => chainDuration(parentId))
    const longestParent = dependentDurations.length > 0 ? Math.max(...dependentDurations) : 0
    const total = ownDuration + longestParent
    memo.set(id, total)
    return total
  }

  activities.forEach((activity) => chainDuration(activity.id))
  const maxDuration = Math.max(...memo.values())

  const critical = new Set<string>()
  for (const [activityId, duration] of memo) {
    if (duration === maxDuration) {
      critical.add(activityId)
      let cursor = byId.get(activityId)
      while (cursor && cursor.dependsOnActivityIds.length > 0) {
        const nextParent = cursor.dependsOnActivityIds
          .map((id) => ({ id, duration: memo.get(id) ?? 0 }))
          .sort((left, right) => right.duration - left.duration)[0]
        if (!nextParent) {
          break
        }
        critical.add(nextParent.id)
        cursor = byId.get(nextParent.id)
      }
    }
  }

  return critical
}

function recomputeProjectRyg(project: ProjectEntity, milestones: MilestoneEntity[], activities: ActivityEntity[]): RygStatus {
  const allStatuses = [
    ...milestones.map((milestone) => milestone.ryg),
    ...activities.map((activity) => activity.ryg),
  ]

  if (allStatuses.includes('RED')) {
    return 'RED'
  }
  if (allStatuses.includes('YELLOW')) {
    return 'YELLOW'
  }
  return project.timeStatus
}

function computeDataState(data: AppData): AppData {
  const milestones = data.milestones.map((milestone) => ({
    ...milestone,
    ryg: calculateDateRyg(milestone.plannedDate, data.settings),
    status:
      milestone.status === 'CLOSED'
        ? milestone.status
        : calculateDateRyg(milestone.plannedDate, data.settings) === 'RED'
          ? 'DELAYED'
          : milestone.status,
  }))

  const activities = data.activities.map((activity) => {
    const endReference = activity.status === 'CLOSED' ? activity.closedDate ?? activity.endDate : activity.endDate
    return {
      ...activity,
      ryg: calculateDateRyg(endReference, data.settings),
    }
  })

  const activitiesByProject = new Map<string, ActivityEntity[]>()
  for (const activity of activities) {
    const bucket = activitiesByProject.get(activity.projectId)
    if (bucket) {
      bucket.push(activity)
    } else {
      activitiesByProject.set(activity.projectId, [activity])
    }
  }

  const criticalByProject = new Map<string, Set<string>>()
  for (const [projectId, projectActivities] of activitiesByProject) {
    criticalByProject.set(projectId, calculateCriticalPathActivities(projectActivities))
  }

  const activitiesWithCriticalFlag = activities.map((activity) => ({
    ...activity,
    isCriticalPath: criticalByProject.get(activity.projectId)?.has(activity.id) ?? false,
  }))

  const milestonesWithCriticalFlag = milestones.map((milestone) => ({
    ...milestone,
    isCriticalPath:
      milestone.isCriticalPath ||
      activitiesWithCriticalFlag.some(
        (activity) => activity.projectId === milestone.projectId && activity.milestoneId === milestone.id && activity.isCriticalPath,
      ),
  }))

  const projects = data.projects.map((project) => {
    const projectMilestones = milestonesWithCriticalFlag.filter((milestone) => milestone.projectId === project.id)
    const projectActivities = activitiesWithCriticalFlag.filter((activity) => activity.projectId === project.id)
    const criticalPathVersion = projectActivities.filter((activity) => activity.isCriticalPath).length
    const changed = criticalPathVersion !== project.criticalPathVersion

    return {
      ...project,
      timeStatus: recomputeProjectRyg(project, projectMilestones, projectActivities),
      criticalPathChangedFlag: changed,
      criticalPathVersion,
    }
  })

  const enriched: AppData = {
    ...data,
    projects,
    milestones: milestonesWithCriticalFlag,
    activities: activitiesWithCriticalFlag,
  }

  return {
    ...enriched,
    notifications: buildNotifications(enriched),
  }
}

function buildNotifications(data: AppData): NotificationEntity[] {
  const notifications: NotificationEntity[] = []
  const now = isoDate(new Date())

  const projectById = new Map(data.projects.map((project) => [project.id, project]))
  const userById = new Map(data.users.map((user) => [user.id, user]))

  let counter = 0
  function pushNotification(notification: Omit<NotificationEntity, 'id'>): void {
    counter += 1
    notifications.push({
      id: `NTF-${counter}`,
      ...notification,
    })
  }

  for (const activity of data.activities) {
    const remaining = daysBetween(now, activity.endDate)
    if (remaining >= 0 && remaining <= data.settings.warningDaysThreshold && activity.status !== 'CLOSED') {
      pushNotification({
        createdAt: now,
        severity: 'YELLOW',
        type: 'EARLY_WARNING_ACTIVITY',
        message: `Activity "${activity.name}" is due in ${remaining} day(s).`,
        recipientUserIds: [activity.ownerUserId],
      })
    }
  }

  for (const milestone of data.milestones) {
    const remaining = daysBetween(now, milestone.plannedDate)
    if (remaining >= 0 && remaining <= data.settings.warningDaysThreshold && milestone.status !== 'CLOSED') {
      const project = projectById.get(milestone.projectId)
      if (!project) {
        continue
      }

      pushNotification({
        createdAt: now,
        severity: 'YELLOW',
        type: 'EARLY_WARNING_MILESTONE',
        message: `Milestone "${milestone.name}" for ${project.name} is due in ${remaining} day(s).`,
        recipientUserIds: [project.projectManagerId, project.managerId],
      })
    }
  }

  const workloadRows = getWorkloadRows(data, undefined)
  for (const row of workloadRows) {
    if (row.utilization >= data.settings.utilizationCriticalThreshold / 100) {
      const resource = data.resources.find((entry) => entry.personId === row.personId)
      pushNotification({
        createdAt: now,
        severity: 'RED',
        type: 'WORKLOAD_100',
        message: `${row.personName} reached ${toPercentage(row.utilization)} utilization for week ${row.weekStart}.`,
        recipientUserIds: unique([
          resource?.managerUserId ?? '',
          ...data.projects
            .filter((project) => row.relatedProjectNames.includes(project.name))
            .map((project) => project.projectManagerId),
        ]).filter(Boolean),
      })
      continue
    }

    if (row.utilization >= data.settings.utilizationWarningThreshold / 100) {
      const resource = data.resources.find((entry) => entry.personId === row.personId)
      pushNotification({
        createdAt: now,
        severity: 'YELLOW',
        type: 'WORKLOAD_90',
        message: `${row.personName} is at ${toPercentage(row.utilization)} utilization for week ${row.weekStart}.`,
        recipientUserIds: unique([
          resource?.managerUserId ?? '',
          ...data.projects
            .filter((project) => row.relatedProjectNames.includes(project.name))
            .map((project) => project.projectManagerId),
        ]).filter(Boolean),
      })
    }
  }

  const delayedMilestones = data.milestones.filter((milestone) => milestone.ryg === 'RED' && milestone.status !== 'CLOSED')
  if (delayedMilestones.length > 0) {
    pushNotification({
      createdAt: now,
      severity: 'RED',
      type: 'WEEKLY_DELAYED_MILESTONES',
      message: `Weekly delayed milestones digest: ${delayedMilestones.length} delayed milestone(s).`,
      recipientUserIds: data.users.map((user) => user.id),
    })
  }

  const delayedActivities = data.activities.filter(
    (activity) => activity.ryg === 'RED' && activity.status !== 'CLOSED',
  )
  if (delayedActivities.length > 0) {
    pushNotification({
      createdAt: now,
      severity: 'RED',
      type: 'WEEKLY_DELAYED_ACTIVITIES',
      message: `Weekly delayed activities digest: ${delayedActivities.length} delayed activity(s).`,
      recipientUserIds: unique(delayedActivities.map((activity) => activity.ownerUserId)),
    })
  }

  for (const project of data.projects) {
    if (!project.criticalPathChangedFlag) {
      continue
    }

    const projectManager = userById.get(project.projectManagerId)
    pushNotification({
      createdAt: now,
      severity: 'YELLOW',
      type: 'CRITICAL_PATH_CHANGED',
      message: `Critical path updated for project ${project.name}.`,
      recipientUserIds: unique([project.projectManagerId, project.managerId, projectManager?.managerUserId ?? '']).filter(Boolean),
    })
  }

  return notifications
    .sort((left, right) => (left.createdAt < right.createdAt ? 1 : -1))
    .slice(0, 250)
}

function createUsersAndResources(defaultWeeklyCapacity: number): {
  users: UserProfile[]
  resources: ResourceEntity[]
} {
  const director: UserProfile = {
    id: 'USR-DIR-001',
    name: 'Daniel Executive',
    role: 'DIRECTOR',
    department: 'Portfolio Office',
    site: 'Stuttgart',
  }

  const managers: UserProfile[] = Array.from({ length: 8 }, (_, index) => ({
    id: `USR-MGR-${String(index + 1).padStart(3, '0')}`,
    name: `Manager ${index + 1}`,
    role: 'MANAGER',
    managerUserId: director.id,
    department: index % 2 === 0 ? 'Engineering' : 'Operations',
    site: index % 2 === 0 ? 'Stuttgart' : 'Madrid',
  }))

  const projectManagers: UserProfile[] = Array.from({ length: 28 }, (_, index) => {
    const manager = managers[index % managers.length]
    return {
      id: `USR-PM-${String(index + 1).padStart(3, '0')}`,
      name: `Project Manager ${index + 1}`,
      role: 'PROJECT_MANAGER',
      managerUserId: manager.id,
      department: manager.department,
      site: manager.site,
    }
  })

  const specialistUsers: UserProfile[] = Array.from({ length: 80 }, (_, index) => {
    const manager = managers[index % managers.length]
    return {
      id: `USR-RSC-${String(index + 1).padStart(3, '0')}`,
      name: `Resource ${index + 1}`,
      role: 'PROJECT_MANAGER',
      managerUserId: manager.id,
      department: manager.department,
      site: manager.site,
    }
  })

  const users = [director, ...managers, ...projectManagers, ...specialistUsers]
  const resources: ResourceEntity[] = specialistUsers.map((resourceUser, index) => ({
    personId: resourceUser.id,
    name: resourceUser.name,
    role: index % 3 === 0 ? 'Engineer' : index % 3 === 1 ? 'Planner' : 'Analyst',
    department: resourceUser.department,
    weeklyCapacityHours: defaultWeeklyCapacity,
    managerUserId: resourceUser.managerUserId ?? managers[0].id,
  }))

  return { users, resources }
}

export function initializeDemoData(): AppData {
  const settings: AppSettings = {
    warningDaysThreshold: 10,
    reportWeekWindow: 8,
    defaultWeeklyCapacity: 40,
    utilizationWarningThreshold: 90,
    utilizationCriticalThreshold: 100,
  }

  const { users, resources } = createUsersAndResources(settings.defaultWeeklyCapacity)
  const managers = users.filter((user) => user.role === 'MANAGER')
  const pms = users.filter((user) => user.role === 'PROJECT_MANAGER').slice(0, 28)

  const categories: ProjectCategory[] = ['ECR', 'Path Forward', 'CIP', 'New Programs']
  const sites = ['Stuttgart', 'Madrid', 'Bangalore', 'Cluj']

  const today = isoDate(new Date())
  const projects: ProjectEntity[] = []
  const milestones: MilestoneEntity[] = []
  const activities: ActivityEntity[] = []
  const assignments: AssignmentEntity[] = []

  const DEMO_NAMES: Array<{ name: string; objective: string; category: ProjectCategory; site: string }> = [
    { name: 'Assembly Line CIP \u2014 TIP Station 12', objective: 'Reduce cycle time and eliminate waste at Assembly Station 12 through Lean/Kaizen activities. Target: 15% OEE improvement.', category: 'CIP', site: 'TIP' },
    { name: 'ESC Sensor New Program \u2014 VW MQB', objective: 'Launch new Electronic Stability Control sensor program for Volkswagen MQB platform. Includes full DFMEA, DVP, and SOP.', category: 'New Programs', site: 'StgP' },
    { name: 'Brake Caliper Path Forward \u2014 GM Recall', objective: 'Resolve root cause of brake caliper corrosion reported in GM field returns. Define and implement permanent corrective actions to prevent recurrence.', category: 'Path Forward', site: 'TIP, StgP' },
    { name: 'ABS Module ECR \u2014 Ford F-Series', objective: 'Implement engineering change request to update ABS module software to comply with new brake safety standards for F-Series trucks.', category: 'ECR', site: 'TIP' },
    { name: 'Sparkplugs', objective: 'Optimize sparkplug manufacturing line throughput and reduce scrap rate. Target: 20% yield improvement.', category: 'ECR', site: 'TIP' },
    { name: 'Fuel Injector CIP \u2014 Line 7', objective: 'Continuous improvement of fuel injector assembly line 7 focusing on takt time reduction and quality defect elimination.', category: 'CIP', site: 'StgP' },
    { name: 'EPS Motor New Program \u2014 BMW G70', objective: 'Electric Power Steering motor launch for BMW G70 platform. Full APQP and PPAP deliverables required.', category: 'New Programs', site: 'TIP' },
    { name: 'Wiper System Path Forward \u2014 Warranty', objective: 'Address field warranty claims on wiper motor assemblies. Root cause analysis and design revision for improved durability.', category: 'Path Forward', site: 'StgP' },
    { name: 'Oxygen Sensor ECR \u2014 Stellantis', objective: 'Engineering change to oxygen sensor calibration tables for Stellantis Euro 7 compliance requirements.', category: 'ECR', site: 'TIP' },
    { name: 'Hydraulic Pump CIP \u2014 Noise Reduction', objective: 'Reduce hydraulic pump assembly noise levels by 3dB through component tolerance optimization and damping improvements.', category: 'CIP', site: 'StgP' },
  ]

  for (let index = 0; index < 520; index += 1) {
    const id = index < DEMO_NAMES.length
      ? `${DEMO_NAMES[index].category === 'CIP' ? 'CIP' : DEMO_NAMES[index].category === 'New Programs' ? 'NP' : DEMO_NAMES[index].category === 'Path Forward' ? 'PF' : 'ECR'}-2026-${String(index + 1).padStart(3, '0')}`
      : `PRJ-${String(index + 1).padStart(4, '0')}`
    const demo = index < DEMO_NAMES.length ? DEMO_NAMES[index] : undefined
    const category = demo?.category ?? categories[index % categories.length]
    const manager = managers[index % managers.length]
    const pm = pms[index % pms.length]
    const projectStart = addDays(today, -1 * (index % 30))

    projects.push({
      id,
      name: demo?.name ?? `Portfolio Project ${index + 1}`,
      category,
      site: demo?.site ?? sites[index % sites.length],
      objective: demo?.objective ?? `Drive ${category} outcomes for plant line ${(index % 12) + 1}`,
      sponsorExecutive: `Executive ${(index % 6) + 1}`,
      projectManagerId: pm.id,
      managerId: manager.id,
      timeStatus: 'GREEN',
      criticalPathChangedFlag: index % 9 === 0,
      dependencyVersion: 1,
      criticalPathVersion: 0,
    })

    const projectMilestones = buildMilestoneTemplate(id, category, projectStart)
    milestones.push(...projectMilestones)

    for (let activityIndex = 0; activityIndex < 20; activityIndex += 1) {
      const owner = resources[(index * 3 + activityIndex) % resources.length]
      const startDate = addDays(projectStart, activityIndex * 2)
      const endDate = addDays(startDate, 5 + (activityIndex % 5))
      const plannedEndDate = addDays(endDate, activityIndex % 4 === 0 ? -2 : 0)
      const isClosed = activityIndex % 7 === 0
      const idActivity = `${id}-ACT-${String(activityIndex + 1).padStart(2, '0')}`
      const weekStart = isoDate(startOfWeek(parseDate(startDate)))

      activities.push({
        id: idActivity,
        projectId: id,
        milestoneId: projectMilestones[Math.min(projectMilestones.length - 1, Math.floor(activityIndex / 5))]?.id,
        name: `Activity ${activityIndex + 1}`,
        ownerUserId: owner.personId,
        startDate,
        endDate,
        plannedEndDate,
        closedDate: isClosed ? addDays(endDate, 1) : undefined,
        status: isClosed ? 'CLOSED' : activityIndex % 11 === 0 ? 'BLOCKED' : 'IN_PROGRESS',
        ryg: 'GREEN',
        isCriticalPath: false,
        dependsOnActivityIds: activityIndex === 0 ? [] : [`${id}-ACT-${String(activityIndex).padStart(2, '0')}`],
      })

      assignments.push({
        id: `ASN-${id}-${String(activityIndex + 1).padStart(2, '0')}`,
        activityId: idActivity,
        personId: owner.personId,
        assignedHours: 8 + (activityIndex % 5) * 4,
        weekStart,
      })
    }
  }

  return computeDataState({
    users,
    projects,
    milestones,
    activities,
    resources,
    assignments,
    notifications: [],
    settings,
  })
}

export function recalculateData(next: AppData): AppData {
  return computeDataState(next)
}

export function createProjectWithTemplate(data: AppData, input: ProjectCreateInput): AppData {
  const projectId = `PRJ-${String(data.projects.length + 1).padStart(4, '0')}`
  const today = isoDate(new Date())

  const project: ProjectEntity = {
    id: projectId,
    name: input.name,
    category: input.category,
    site: input.site,
    objective: input.objective,
    sponsorExecutive: input.sponsorExecutive,
    projectManagerId: input.projectManagerId,
    managerId: input.managerId,
    timeStatus: 'GREEN',
    criticalPathChangedFlag: false,
    dependencyVersion: 1,
    criticalPathVersion: 0,
  }

  const templateMilestones = buildMilestoneTemplate(projectId, input.category, today)

  return computeDataState({
    ...data,
    projects: [...data.projects, project],
    milestones: [...data.milestones, ...templateMilestones],
  })
}

export function getPmProjects(data: AppData, currentUser: UserProfile): ProjectEntity[] {
  return data.projects.filter((project) => canManageProject(project, currentUser))
}

export function getProjectSummary(data: AppData, currentUser: UserProfile): ProjectSummaryRow[] {
  return getPmProjects(data, currentUser)
    .map((project) => {
      const managerName = data.users.find((user) => user.id === project.managerId)?.name ?? 'Unknown'
      const pmName = data.users.find((user) => user.id === project.projectManagerId)?.name ?? 'Unknown'

      const projectMilestones = data.milestones.filter((milestone) => milestone.projectId === project.id)
      const projectActivities = data.activities.filter((activity) => activity.projectId === project.id)

      const endDate = projectMilestones.length > 0
        ? projectMilestones.reduce((latest, milestone) => milestone.plannedDate > latest ? milestone.plannedDate : latest, projectMilestones[0].plannedDate)
        : ''

      const costRyg: RygStatus = projectMilestones.some((milestone) => milestone.ryg === 'RED')
        ? 'RED'
        : projectMilestones.some((milestone) => milestone.ryg === 'YELLOW') ? 'YELLOW' : 'GREEN'

      const openActivities = projectActivities.filter((activity) => activity.status !== 'CLOSED')
      const qualityRyg: RygStatus = openActivities.some((activity) => activity.ryg === 'RED')
        ? 'RED'
        : openActivities.some((activity) => activity.ryg === 'YELLOW') ? 'YELLOW' : 'GREEN'

      return {
        project,
        managerName,
        pmName,
        overallRyg: project.timeStatus,
        timeRyg: project.timeStatus,
        costRyg,
        qualityRyg,
        criticalPathChanged: project.criticalPathChangedFlag,
        endDate,
      }
    })
    .sort((left, right) => compareRygPriority(left.overallRyg, right.overallRyg))
}

export function getProjectTimelineItems(data: AppData, projectId: string): TimelineRow[] {
  const usersById = new Map(data.users.map((user) => [user.id, user]))

  return data.activities
    .filter((activity) => activity.projectId === projectId)
    .sort((left, right) => (left.startDate < right.startDate ? -1 : 1))
    .map((activity) => ({
      id: activity.id,
      name: activity.name,
      ownerName: usersById.get(activity.ownerUserId)?.name ?? 'Unassigned',
      startDate: activity.startDate,
      endDate: activity.endDate,
      isCriticalPath: activity.isCriticalPath,
      ryg: activity.ryg,
    }))
}

export function getUpcomingActivitiesForProject(
  data: AppData,
  projectId: string,
  weekWindow: number,
): Array<{
  id: string
  name: string
  ownerName: string
  startDate: string
  endDate: string
  status: ActivityEntity['status']
  ryg: RygStatus
}> {
  const today = isoDate(new Date())
  const maxDate = addWeeks(today, weekWindow)
  const usersById = new Map(data.users.map((user) => [user.id, user]))

  return data.activities
    .filter(
      (activity) =>
        activity.projectId === projectId &&
        activity.startDate >= today &&
        activity.startDate <= maxDate &&
        activity.status !== 'CLOSED',
    )
    .sort((left, right) => (left.startDate < right.startDate ? -1 : 1))
    .map((activity) => ({
      id: activity.id,
      name: activity.name,
      ownerName: usersById.get(activity.ownerUserId)?.name ?? 'Unknown',
      startDate: activity.startDate,
      endDate: activity.endDate,
      status: activity.status,
      ryg: activity.ryg,
    }))
}

export function getWeeklyClosedActivities(
  data: AppData,
  projectId: string,
  weekWindow: number,
): Array<{
  id: string
  name: string
  ownerName: string
  closedDate?: string
  endDate: string
}> {
  const minDate = addWeeks(isoDate(new Date()), -1 * weekWindow)
  const usersById = new Map(data.users.map((user) => [user.id, user]))

  return data.activities
    .filter(
      (activity) =>
        activity.projectId === projectId &&
        activity.status === 'CLOSED' &&
        (activity.closedDate ?? activity.endDate) >= minDate,
    )
    .sort((left, right) => ((left.closedDate ?? left.endDate) > (right.closedDate ?? right.endDate) ? -1 : 1))
    .map((activity) => ({
      id: activity.id,
      name: activity.name,
      ownerName: usersById.get(activity.ownerUserId)?.name ?? 'Unknown',
      closedDate: activity.closedDate,
      endDate: activity.endDate,
    }))
}

export function getDelayedActivitiesByProject(data: AppData, projectId: string): DelayedActivityRow[] {
  const usersById = new Map(data.users.map((user) => [user.id, user]))
  const today = isoDate(new Date())

  return data.activities
    .filter((activity) => activity.projectId === projectId && activity.status !== 'CLOSED')
    .map((activity) => ({
      activity,
      delayDays: daysBetween(activity.plannedEndDate, today),
    }))
    .filter((row) => row.delayDays > 0)
    .sort((left, right) => right.delayDays - left.delayDays)
    .map((row) => ({
      id: row.activity.id,
      name: row.activity.name,
      ownerName: usersById.get(row.activity.ownerUserId)?.name ?? 'Unknown',
      plannedEndDate: row.activity.plannedEndDate,
      delayDays: row.delayDays,
      status: row.activity.status,
      ryg: 'RED',
      isCriticalPath: row.activity.isCriticalPath,
    }))
}

export function getPrioritizationMilestones(
  data: AppData,
  currentUser: UserProfile,
  weekWindow: number,
): MilestonePrioritizationRow[] {
  const maxDate = addWeeks(isoDate(new Date()), weekWindow)
  const projectById = new Map(data.projects.map((project) => [project.id, project]))

  return data.milestones
    .filter((milestone) => {
      const project = projectById.get(milestone.projectId)
      if (!project || milestone.plannedDate > maxDate) {
        return false
      }

      if (currentUser.role === 'DIRECTOR') {
        return true
      }

      if (currentUser.role === 'MANAGER') {
        return project.managerId === currentUser.id
      }

      return project.projectManagerId === currentUser.id
    })
    .map((milestone) => {
      const project = projectById.get(milestone.projectId)
      return {
        id: milestone.id,
        projectName: project?.name ?? 'Unknown',
        name: milestone.name,
        plannedDate: milestone.plannedDate,
        daysRemaining: daysBetween(isoDate(new Date()), milestone.plannedDate),
        ryg: milestone.ryg,
      }
    })
    .sort((left, right) => {
      const byRyg = compareRygPriority(left.ryg, right.ryg)
      if (byRyg !== 0) {
        return byRyg
      }

      return left.plannedDate < right.plannedDate ? -1 : 1
    })
}

export function getWorkloadRows(data: AppData, currentUser?: UserProfile): WorkloadRow[] {
  const activityById = new Map(data.activities.map((activity) => [activity.id, activity]))
  const projectById = new Map(data.projects.map((project) => [project.id, project]))
  const usersById = new Map(data.users.map((user) => [user.id, user]))

  const rows = new Map<string, WorkloadRow>()

  for (const assignment of data.assignments) {
    const activity = activityById.get(assignment.activityId)
    if (!activity) {
      continue
    }

    const project = projectById.get(activity.projectId)
    const resource = data.resources.find((entry) => entry.personId === assignment.personId)
    if (!project || !resource) {
      continue
    }

    if (currentUser) {
      if (currentUser.role === 'PROJECT_MANAGER' && project.projectManagerId !== currentUser.id) {
        continue
      }
      if (currentUser.role === 'MANAGER' && resource.managerUserId !== currentUser.id) {
        continue
      }
    }

    const key = `${assignment.weekStart}_${assignment.personId}`
    const existing = rows.get(key)

    if (!existing) {
      rows.set(key, {
        weekStart: assignment.weekStart,
        personId: assignment.personId,
        personName: usersById.get(assignment.personId)?.name ?? resource.name,
        capacityHours: resource.weeklyCapacityHours,
        assignedHours: assignment.assignedHours,
        utilization: assignment.assignedHours / resource.weeklyCapacityHours,
        utilizationRyg: calculateWorkloadRyg(assignment.assignedHours / resource.weeklyCapacityHours, data.settings),
        relatedProjectNames: [project.name],
      })
      continue
    }

    const nextAssigned = existing.assignedHours + assignment.assignedHours
    existing.assignedHours = nextAssigned
    existing.utilization = nextAssigned / existing.capacityHours
    existing.utilizationRyg = calculateWorkloadRyg(existing.utilization, data.settings)
    existing.relatedProjectNames = unique([...existing.relatedProjectNames, project.name])
  }

  return [...rows.values()].sort((left, right) => {
    if (left.weekStart !== right.weekStart) {
      return left.weekStart < right.weekStart ? -1 : 1
    }

    return left.personName < right.personName ? -1 : 1
  })
}

export function getNextMilestoneForProject(
  data: AppData,
  projectId: string,
):
  | {
      id: string
      name: string
      plannedDate: string
      daysRemaining: number
    }
  | undefined {
  const today = isoDate(new Date())

  const next = data.milestones
    .filter((milestone) => milestone.projectId === projectId && milestone.status !== 'CLOSED' && milestone.plannedDate >= today)
    .sort((left, right) => (left.plannedDate < right.plannedDate ? -1 : 1))[0]

  if (!next) {
    return undefined
  }

  return {
    id: next.id,
    name: next.name,
    plannedDate: next.plannedDate,
    daysRemaining: daysBetween(today, next.plannedDate),
  }
}

export function formatDate(value?: string): string {
  if (!value) {
    return '-'
  }

  const date = parseDate(value)
  return date.toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
  })
}

export function toPercentage(value: number): string {
  return `${Math.round(value * 100)}%`
}
