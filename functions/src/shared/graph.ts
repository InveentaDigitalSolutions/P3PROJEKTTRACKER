/**
 * graph.ts — Microsoft Graph API client for sending emails and Teams messages.
 *
 * Uses the same ClientSecretCredential as the Dataverse client but scoped
 * to `https://graph.microsoft.com/.default`.
 */

import { ClientSecretCredential } from '@azure/identity'
import { Client } from '@microsoft/microsoft-graph-client'
import { config } from './config.js'

// ── Token provider ───────────────────────────────────────────────────────

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

async function getGraphToken(): Promise<string> {
  const cred = getCredential()
  const res = await cred.getToken('https://graph.microsoft.com/.default')
  return res.token
}

function getGraphClient(): Client {
  return Client.init({
    authProvider: async (done) => {
      try {
        const token = await getGraphToken()
        done(null, token)
      } catch (err) {
        done(err as Error, null)
      }
    },
  })
}

// ── Email ────────────────────────────────────────────────────────────────

export interface EmailOptions {
  /** Recipient email addresses */
  to: string[]
  subject: string
  /** HTML body */
  body: string
  /** Optional CC addresses */
  cc?: string[]
}

/**
 * Send an email via Microsoft Graph using the configured sender mailbox.
 *
 * Requires the app registration to have **Mail.Send** application
 * permission (or be authorised as a send-as delegate for the sender).
 */
export async function sendEmail(options: EmailOptions): Promise<void> {
  const client = getGraphClient()
  const sender = config.notificationSender()

  const message = {
    subject: options.subject,
    body: {
      contentType: 'HTML',
      content: options.body,
    },
    toRecipients: options.to.map((email) => ({
      emailAddress: { address: email },
    })),
    ...(options.cc?.length
      ? {
          ccRecipients: options.cc.map((email) => ({
            emailAddress: { address: email },
          })),
        }
      : {}),
  }

  await client.api(`/users/${sender}/sendMail`).post({
    message,
    saveToSentItems: false,
  })
}

// ── Teams chat message (1-on-1) ──────────────────────────────────────────

/**
 * Send a 1-on-1 Teams chat message to a user.
 *
 * Requires **Chat.Create** + **ChatMessage.Send** application permissions.
 * Falls back gracefully if the user has no Teams licence.
 */
export async function sendTeamsChat(
  recipientUserId: string,
  htmlContent: string,
): Promise<void> {
  const client = getGraphClient()

  // 1. Create (or get existing) 1:1 chat between the app and the user
  const chat = await client.api('/chats').post({
    chatType: 'oneOnOne',
    members: [
      {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${recipientUserId}')`,
      },
    ],
  })

  // 2. Post message
  await client.api(`/chats/${chat.id}/messages`).post({
    body: {
      contentType: 'html',
      content: htmlContent,
    },
  })
}
