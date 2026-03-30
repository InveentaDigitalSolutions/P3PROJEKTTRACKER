/**
 * onProjectCreated — HTTP-triggered function.
 *
 * Called by the PTH React app immediately after a new project is
 * created in Dataverse.  Sends a rich HTML notification email to
 * the configured recipients (or the project manager / sponsor)
 * via Microsoft Graph.
 *
 * POST /api/onProjectCreated
 * Body: { projectName, projectCode, category, objective, siteName,
 *         sponsorExecutive, projectManager, budgetAllocated,
 *         plannedStartDate, plannedEndDate, boschCode, milestones,
 *         creatorName, creatorEmail }
 */

import { app, type HttpRequest, type HttpResponseInit, type InvocationContext } from '@azure/functions'
import { sendEmail } from '../shared/graph.js'
import { config } from '../shared/config.js'
import {
  emailLayout, infoTable, sectionHeading, categoryBadge, calloutBox,
  tableHead, tableRow, tableEnd, C,
} from '../shared/emailLayout.js'

// ── Request body shape ───────────────────────────────────────────────────

interface ProjectCreatedPayload {
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
  boschCode: string
  milestones: { name: string; targetDate: string }[]
  creatorName: string
  creatorEmail: string
}

// ── HTML builder ─────────────────────────────────────────────────────────

function buildEmailBody(p: ProjectCreatedPayload): string {
  const budget = p.budgetAllocated
    ? `€${p.budgetAllocated.toLocaleString('de-DE', { minimumFractionDigits: 0 })}`
    : '—'

  // Project info table
  const rows: [string, string][] = [
    ['Project Name', `<strong>${p.projectName}</strong>`],
    ['Project Code', p.projectCode],
  ]
  if (p.boschCode) rows.push(['Bosch Code', p.boschCode])
  rows.push(
    ['Category', categoryBadge(p.category)],
    ['Site', p.siteName || '—'],
    ['Executive Sponsor', p.sponsorExecutive || '—'],
    ['Project Manager', p.projectManager || '—'],
    ['Budget', budget],
    ['Planned Start', p.plannedStartDate || '—'],
    ['Planned End', p.plannedEndDate || '—'],
  )

  // Milestones table
  let milestonesHtml = ''
  if (p.milestones.length > 0) {
    milestonesHtml =
      sectionHeading('Milestones', C.teal) +
      tableHead(['#', 'Name', 'Target Date']) +
      p.milestones
        .map((ms, i) =>
          tableRow([
            `<span style="color:${C.muted}">${i + 1}</span>`,
            ms.name,
            ms.targetDate,
          ]),
        )
        .join('') +
      tableEnd()
  }

  const body =
    `<p style="margin:0 0 20px;font-size:14px;color:${C.text};line-height:1.6">
      A new project has been created in <strong>Project Tracker Hub</strong> by <strong>${p.creatorName}</strong>.
    </p>` +
    infoTable(rows) +
    (p.objective ? calloutBox('Objective', p.objective) : '') +
    milestonesHtml

  return emailLayout({
    icon: '🚀',
    title: 'New Project Created',
    subtitle: `${p.projectName} (${p.projectCode})`,
    body,
  })
}

// ── Function handler ─────────────────────────────────────────────────────

async function onProjectCreated(req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  context.log('onProjectCreated: received request')

  // Parse body
  let payload: ProjectCreatedPayload
  try {
    payload = (await req.json()) as ProjectCreatedPayload
  } catch {
    return { status: 400, body: JSON.stringify({ error: 'Invalid JSON body' }) }
  }

  if (!payload.projectName || !payload.projectCode) {
    return { status: 400, body: JSON.stringify({ error: 'projectName and projectCode are required' }) }
  }

  // Build recipient list: configured override → creator email
  const overrideRecipients = config.overdueRecipients()
  const recipients: string[] = overrideRecipients
    ? overrideRecipients.split(',').map((e) => e.trim()).filter(Boolean)
    : []

  // Always include the creator if they have an email
  if (payload.creatorEmail && !recipients.includes(payload.creatorEmail)) {
    recipients.push(payload.creatorEmail)
  }

  if (recipients.length === 0) {
    context.warn('onProjectCreated: no recipients available, skipping email.')
    return {
      status: 200,
      jsonBody: { sent: false, reason: 'no recipients' },
    }
  }

  // Build and send
  const body = buildEmailBody(payload)
  try {
    await sendEmail({
      to: recipients,
      subject: `[PTH] 🚀 New Project: ${payload.projectName} (${payload.projectCode})`,
      body,
    })
    context.log(`onProjectCreated: email sent to ${recipients.join(', ')}`)
    return {
      status: 200,
      jsonBody: { sent: true, recipients },
    }
  } catch (err) {
    context.error('onProjectCreated: failed to send email', err)
    return {
      status: 500,
      jsonBody: { sent: false, error: String(err) },
    }
  }
}

// ── Registration ─────────────────────────────────────────────────────────

app.http('onProjectCreated', {
  methods: ['POST'],
  authLevel: 'function',
  handler: onProjectCreated,
})
