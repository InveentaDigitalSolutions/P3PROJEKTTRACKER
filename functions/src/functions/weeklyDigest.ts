/**
 * weeklyDigest — Timer trigger (every Monday at 08:00 UTC).
 *
 * Produces a full portfolio status digest with KPIs, project-level
 * summaries, and overdue / upcoming highlights.  Sent to configured
 * digest recipients (typically leadership / PMO).
 */

import { app, type InvocationContext } from '@azure/functions'
import {
  fetchProjects,
  fetchActivities,
  fetchMilestones,
  type DvProject,
  type DvActivity,
  type DvMilestone,
} from '../shared/dataverse.js'
import { sendEmail } from '../shared/graph.js'
import { config } from '../shared/config.js'
import {
  emailLayout,
  kpiRow,
  sectionHeading,
  statusDot,
  progressBar,
  statusBadge,
  tableHead,
  tableRow,
  tableEnd,
  C,
} from '../shared/emailLayout.js'

// ── Helpers ──────────────────────────────────────────────────────────────

function todayISO(): string {
  return new Date().toISOString().slice(0, 10)
}

function daysUntil(target: string, today: string): number {
  return Math.floor((new Date(target).getTime() - new Date(today).getTime()) / 86_400_000)
}

interface ProjectSummary {
  project: DvProject
  totalActivities: number
  doneActivities: number
  overdueActivities: number
  progressPct: number
  nextMilestone: DvMilestone | null
  milestonesDue: number
}

// ── HTML builder ─────────────────────────────────────────────────────────

function buildDigestBody(
  summaries: ProjectSummary[],
  kpis: { total: number; green: number; yellow: number; red: number; completed: number },
  today: string,
): string {
  let body = ''

  // KPI cards row
  body += kpiRow([
    { label: 'Total', value: kpis.total, color: C.blue },
    { label: 'On Track', value: kpis.green, color: C.teal },
    { label: 'At Risk', value: kpis.yellow, color: C.yellow },
    { label: 'Critical', value: kpis.red, color: C.red },
    { label: 'Completed', value: kpis.completed, color: C.blue },
  ])

  // Project status table
  body += sectionHeading('Project Status Summary', C.blue)
  body += tableHead(['Project', 'PM', 'Status', 'Progress', 'Overdue', 'Next Milestone'])

  for (const s of summaries) {
    const dot = statusDot(s.project.timeStatus)
    const bar = progressBar(s.progressPct)
    const overdueVal = s.overdueActivities > 0
      ? `<strong style="color:${C.red}">${s.overdueActivities}</strong>`
      : `<span style="color:${C.muted}">0</span>`
    const nextMs = s.nextMilestone
      ? `${s.nextMilestone.name} <span style="color:${C.muted};font-size:11px">(${s.nextMilestone.plannedDate})</span>`
      : `<span style="color:${C.muted}">—</span>`

    body += tableRow(
      [
        `<strong>${s.project.name}</strong><br><span style="font-size:11px;color:${C.muted}">${s.project.projectCode} · ${s.project.category}</span>`,
        s.project.projectManagerName,
        dot,
        bar,
        overdueVal,
        nextMs,
      ],
      s.overdueActivities > 0,
    )
  }
  body += tableEnd()

  return emailLayout({
    icon: '📊',
    title: 'Weekly Portfolio Digest',
    subtitle: `Week of ${today} · ${kpis.total} projects tracked`,
    body,
    headerColor: C.sidebar,
  })
}

// ── Function handler ─────────────────────────────────────────────────────

async function weeklyDigest(_timer: unknown, context: InvocationContext): Promise<void> {
  context.log('weeklyDigest: starting…')

  const [projects, activities, milestones] = await Promise.all([
    fetchProjects(),
    fetchActivities(),
    fetchMilestones(),
  ])
  const today = todayISO()
  const threshold = config.warningDaysThreshold()

  // Build per-project summaries
  const summaries: ProjectSummary[] = projects.map((project) => {
    const pActs = activities.filter((a) => a.projectId === project.id)
    const pMss = milestones.filter((ms) => ms.projectId === project.id)
    const doneActs = pActs.filter((a) => a.status === 'DONE')
    const overdueActs = pActs.filter((a) => a.status !== 'DONE' && a.endDate < today)
    const progressPct = pActs.length > 0 ? Math.round((doneActs.length / pActs.length) * 100) : 0

    // Next upcoming milestone (not closed, nearest planned date in the future)
    const upcoming = pMss
      .filter((ms) => ms.status !== 'CLOSED' && ms.plannedDate >= today)
      .sort((a, b) => a.plannedDate.localeCompare(b.plannedDate))
    const nextMilestone = upcoming[0] ?? null

    // Milestones due within threshold
    const milestonesDue = pMss.filter(
      (ms) => ms.status !== 'CLOSED' && daysUntil(ms.plannedDate, today) >= 0 && daysUntil(ms.plannedDate, today) <= threshold,
    ).length

    return {
      project,
      totalActivities: pActs.length,
      doneActivities: doneActs.length,
      overdueActivities: overdueActs.length,
      progressPct,
      nextMilestone,
      milestonesDue,
    }
  })

  // Sort: RED first, then YELLOW, then GREEN; within same status by overdue count desc
  const statusRank = { RED: 0, YELLOW: 1, GREEN: 2 }
  summaries.sort((a, b) => {
    const ra = statusRank[a.project.timeStatus] ?? 3
    const rb = statusRank[b.project.timeStatus] ?? 3
    if (ra !== rb) return ra - rb
    return b.overdueActivities - a.overdueActivities
  })

  // KPIs
  const completed = summaries.filter((s) => s.progressPct === 100).length
  const kpis = {
    total: projects.length,
    green: projects.filter((p) => p.timeStatus === 'GREEN').length,
    yellow: projects.filter((p) => p.timeStatus === 'YELLOW').length,
    red: projects.filter((p) => p.timeStatus === 'RED').length,
    completed,
  }

  context.log(`weeklyDigest: ${kpis.total} projects — ${kpis.green} green, ${kpis.yellow} yellow, ${kpis.red} red, ${completed} completed`)

  // Determine recipients
  const digestRecipients = config.digestRecipients()
  const recipients = digestRecipients
    ? digestRecipients.split(',').map((e) => e.trim()).filter(Boolean)
    : []

  if (recipients.length === 0) {
    context.warn('weeklyDigest: no NOTIFICATION_DIGEST_RECIPIENTS configured. Skipping email.')
    return
  }

  const body = buildDigestBody(summaries, kpis, today)
  await sendEmail({
    to: recipients,
    subject: `[PTH] 📊 Weekly Portfolio Digest — ${today} — ${kpis.total} Projects`,
    body,
  })

  context.log(`weeklyDigest: sent to ${recipients.join(', ')}`)
}

// ── Registration ─────────────────────────────────────────────────────────

app.timer('weeklyDigest', {
  // Every Monday at 08:00 UTC
  schedule: '0 0 8 * * 1',
  handler: weeklyDigest,
})
