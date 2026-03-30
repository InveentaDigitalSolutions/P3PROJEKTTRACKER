/**
 * dailyOverdueScan — Timer trigger (every day at 07:00 UTC).
 *
 * Queries Dataverse for activities whose end date has passed but that
 * are not yet DONE.  Groups them by project / project manager and
 * sends an email digest to each PM (or to the configured override
 * recipients).
 */

import { app, type InvocationContext } from '@azure/functions'
import { fetchProjects, fetchActivities, type DvProject, type DvActivity } from '../shared/dataverse.js'
import { sendEmail } from '../shared/graph.js'
import { config } from '../shared/config.js'
import {
  emailLayout, sectionHeading, tableHead, tableRow, tableEnd, C,
} from '../shared/emailLayout.js'

// ── Helpers ──────────────────────────────────────────────────────────────

function todayISO(): string {
  return new Date().toISOString().slice(0, 10)
}

function daysOverdue(endDate: string, today: string): number {
  const ms = new Date(today).getTime() - new Date(endDate).getTime()
  return Math.floor(ms / 86_400_000)
}

interface OverdueGroup {
  project: DvProject
  activities: (DvActivity & { daysLate: number })[]
}

// ── HTML builder ─────────────────────────────────────────────────────────

function buildEmailBody(groups: OverdueGroup[]): string {
  const totalActivities = groups.reduce((sum, g) => sum + g.activities.length, 0)

  let body = ''
  for (const g of groups) {
    body += sectionHeading(
      `${g.project.name} <span style="font-weight:normal;color:${C.muted};font-size:12px">(${g.project.projectCode})</span>`,
      C.red,
    )
    body += `<p style="margin:0 0 10px;font-size:12px;color:${C.muted}">PM: ${g.project.projectManagerName} &middot; Site: ${g.project.site}</p>`
    body += tableHead(['Activity', 'Owner', 'Due Date', 'Days Late', 'Critical'])
    for (const act of g.activities) {
      const cpBadge = act.isCriticalPath
        ? `<span style="color:${C.red};font-weight:bold">●</span>`
        : `<span style="color:${C.muted}">—</span>`
      const daysColor = act.daysLate > 14 ? C.red : act.daysLate > 7 ? C.yellow : C.text
      body += tableRow(
        [
          act.name,
          act.owner,
          act.endDate,
          `<strong style="color:${daysColor}">${act.daysLate}d</strong>`,
          cpBadge,
        ],
        act.daysLate > 14,
      )
    }
    body += tableEnd()
  }

  return emailLayout({
    icon: '⚠️',
    title: 'Overdue Activities Report',
    subtitle: `${todayISO()} · ${totalActivities} overdue across ${groups.length} project(s)`,
    body,
    headerColor: C.red,
  })
}

// ── Function handler ─────────────────────────────────────────────────────

async function dailyOverdueScan(_timer: unknown, context: InvocationContext): Promise<void> {
  context.log('dailyOverdueScan: starting…')

  const [projects, activities] = await Promise.all([fetchProjects(), fetchActivities()])
  const today = todayISO()

  // Find overdue: endDate < today AND not DONE
  const overdue = activities.filter(
    (a) => a.status !== 'DONE' && a.endDate < today,
  )

  if (overdue.length === 0) {
    context.log('dailyOverdueScan: no overdue activities found.')
    return
  }

  // Group by project
  const projectMap = new Map(projects.map((p) => [p.id, p]))
  const grouped = new Map<string, OverdueGroup>()

  for (const act of overdue) {
    const project = projectMap.get(act.projectId)
    if (!project) continue
    if (!grouped.has(project.id)) grouped.set(project.id, { project, activities: [] })
    grouped.get(project.id)!.activities.push({
      ...act,
      daysLate: daysOverdue(act.endDate, today),
    })
  }

  const groups = [...grouped.values()].sort((a, b) => b.activities.length - a.activities.length)
  context.log(`dailyOverdueScan: ${overdue.length} overdue activities across ${groups.length} projects`)

  // Determine recipients
  const overrideRecipients = config.overdueRecipients()
  if (overrideRecipients) {
    // Single digest to configured addresses
    const recipients = overrideRecipients.split(',').map((e) => e.trim()).filter(Boolean)
    const body = buildEmailBody(groups)
    await sendEmail({
      to: recipients,
      subject: `[PTH] ⚠️ ${overdue.length} Overdue Activities — ${today}`,
      body,
    })
    context.log(`dailyOverdueScan: sent combined digest to ${recipients.join(', ')}`)
  } else {
    // Per-PM emails
    const pmGroups = new Map<string, OverdueGroup[]>()
    for (const g of groups) {
      const pmName = g.project.projectManagerName
      if (!pmGroups.has(pmName)) pmGroups.set(pmName, [])
      pmGroups.get(pmName)!.push(g)
    }

    for (const [pmName, pmProjectGroups] of pmGroups) {
      // Try to find PM email from project manager name
      // Fallback: skip if no email found
      const pmProject = pmProjectGroups[0]?.project
      if (!pmProject) continue

      // We'll attempt to send to PM — the name is stored as text, so
      // in real usage you'd resolve the name to a systemuser email.
      // For now we log a warning and skip when no override is set.
      context.warn(
        `dailyOverdueScan: no NOTIFICATION_OVERDUE_RECIPIENTS configured and cannot resolve email for PM "${pmName}". ` +
        `Set NOTIFICATION_OVERDUE_RECIPIENTS or implement user lookup.`,
      )
    }
  }
}

// ── Registration ─────────────────────────────────────────────────────────

app.timer('dailyOverdueScan', {
  // Every day at 07:00 UTC
  schedule: '0 0 7 * * *',
  handler: dailyOverdueScan,
})
