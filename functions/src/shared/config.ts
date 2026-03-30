/**
 * config.ts — Centralised environment variable access.
 *
 * All Azure Function app settings are read here so functions only
 * import typed helpers instead of raw `process.env` lookups.
 */

function required(name: string): string {
  const v = process.env[name]
  if (!v) throw new Error(`Missing required app setting: ${name}`)
  return v
}

function optional(name: string, fallback = ''): string {
  return process.env[name] ?? fallback
}

export const config = {
  /** Dataverse org URL, e.g. https://org2d99840c.crm.dynamics.com */
  dataverseUrl: () => required('DATAVERSE_URL'),

  /** Entra ID (AAD) tenant */
  tenantId: () => required('AZURE_TENANT_ID'),
  /** App-registration client ID */
  clientId: () => required('AZURE_CLIENT_ID'),
  /** App-registration client secret */
  clientSecret: () => required('AZURE_CLIENT_SECRET'),

  /** "From" address for Graph API mail sends */
  notificationSender: () => optional('NOTIFICATION_SENDER', 'no-reply@inveenta.com'),

  /**
   * Comma-separated email addresses that receive the overdue / digest
   * mails.  If empty, the Project Manager email from Dataverse is used.
   */
  overdueRecipients: () => optional('NOTIFICATION_OVERDUE_RECIPIENTS'),
  digestRecipients: () => optional('NOTIFICATION_DIGEST_RECIPIENTS'),

  /** Days before a due date is considered "approaching" (default 7) */
  warningDaysThreshold: () => Number(optional('WARNING_DAYS_THRESHOLD', '7')),
} as const
