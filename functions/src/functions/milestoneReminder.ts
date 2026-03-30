/**
 * milestoneReminder — Timer trigger (every day at 07:30 UTC).
 *
 * Checks for milestones whose planned date falls within the configured
 * warning-days threshold and that are still OPEN or IN_PROGRESS.
 * Sends a reminder email per project.
 */

import { app, type InvocationContext } from '@azure/functions'
import { fetchProjects, fetchMilestones, type DvProject, type DvMilestone } from '../shared/dataverse.js'
import { sendEmail } from '../shared/graph.js'
import { config } from '../shared/config.js'
import {
  emailLayout, sectionHeading, statusBadge, tableHead, tableRow, tableEnd, C,
} from '../shared/emailLayout.js'

// ── Helpers ──────────────────────────────────────────────────────────────

function todayISO(): string {
  return new Date().toISOString().slice(0, 10)
}

function daysUntil(targetDate: string, today: string): number {
  const ms = new Date(targetDate).getTime() - new Date(today).getTime()
  return Math.floor(ms / 86_400_000)
}

interface ApproachingGroup {
  project: DvProject
  milestones: (DvMilestone & { daysLeft: number })[]
}

// ── HTML builder ─────────────────────────────────────────────────────────

function buildEmailBody(groups: ApproachingGroup[]): string {
  const total = groups.reduce((sum, g) => sum + g.milestones.length, 0)

  let body = ''
  for (const g of groups) {
    body += sectionHeading(
      `${g.project.name} <span style="font-weight:normal;color:${C.muted};font-size:12px">(${g.project.projectCode})</span>`,
      C.yellow,
    )
    body += `<p style="margin:0 0 10px;font-size:12px;color:${C.muted}">PM: ${g.project.projectManagerName} &middot; Sponsor: ${g.project.sponsorExecutive}</p>`
    body += tableHead(['Milestone', 'Target Date', 'Days Left', 'Status', 'Critical Path'])
    for (const ms of g.milestones) {
      const daysColor = ms.daysLeft <= 2 ? C.red : ms.daysLeft <= 5 ? C.yellow : C.text
      const cpBadge = ms.isCriticalPath
        ? `<span style="color:${C.red};font-weight:bold">●</span>`
        : `<span style="color:${C.muted}">—</span>`
      body += tableRow(
        [
          ms.name,
          ms.plannedDate,
          `<strong style="color:${daysColor}">${ms.daysLeft}d</strong>`,
          statusBadge(ms.status),
          cpBadge,
        ],
        ms.daysLeft <= 2,
      )
    }
    body += tableEnd()
  }

  return emailLayout({
    icon: '📅',
    title: 'Upcoming Milestones',
    subtitle: `${todayISO()} · ${total} milestone(s) approaching across ${groups.length} project(s)`,
    body,
    headerColor: '#B45309',
  })
}

// ── Function handler ─────────────────────────────────────────────────────

async function milestoneReminder(_timer: unknown, context: InvocationContext): Promise<void> {
  context.log('milestoneReminder: starting…')

  const threshold = config.warningDaysThreshold()
  const [projects, milestones] = await Promise.all([fetchProjects(), fetchMilestones()])
  const today = todayISO()

  // Find milestones approaching within threshold that aren't closed
  const approaching = milestones.filter((ms) => {
    if (ms.status === 'CLOSED') return false
    const daysLeft = daysUntil(ms.plannedDate, today)
    return daysLeft >= 0 && daysLeft <= threshold
  })

  if (approaching.length === 0) {
    context.log('milestoneReminder: no approaching milestones found.')
    return
  }

  // Group by project
  const projectMap = new Map(projects.map((p) => [p.id, p]))
  const grouped = new Map<string, ApproachingGroup>()

  for (const ms of approaching) {
    const project = projectMap.get(ms.projectId)
    if (!project) continue
    if (!grouped.has(project.id)) grouped.set(project.id, { project, milestones: [] })
    grouped.get(project.id)!.milestones.push({
      ...ms,
      daysLeft: daysUntil(ms.plannedDate, today),
    })
  }

  const groups = [...grouped.values()].sort((a, b) =>
    Math.min(...a.milestones.map((m) => m.daysLeft)) - Math.min(...b.milestones.map((m) => m.daysLeft)),
  )

  context.log(`milestoneReminder: ${approaching.length} approaching milestones across ${groups.length} projects`)

  // Send notification
  const overrideRecipients = config.overdueRecipients()
  const recipients = overrideRecipients
    ? overrideRecipients.split(',').map((e) => e.trim()).filter(Boolean)
    : []

  if (recipients.length === 0) {
    context.warn('milestoneReminder: no recipients configured (NOTIFICATION_OVERDUE_RECIPIENTS). Skipping email.')
    return
  }

  const body = buildEmailBody(groups)
  await sendEmail({
    to: recipients,
    subject: `[PTH] 📅 ${approaching.length} Milestones Approaching — ${today}`,
    body,
  })

  context.log(`milestoneReminder: sent to ${recipients.join(', ')}`)
}

// ── Registration ─────────────────────────────────────────────────────────

app.timer('milestoneReminder', {
  // Every day at 07:30 UTC
  schedule: '0 30 7 * * *',
  handler: milestoneReminder,
})
