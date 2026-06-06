/**
 * Dev-only Dataverse token provider for the Vite proxy.
 *
 * Keeps a fresh access token in memory and silently refreshes it from the
 * MSAL on-disk cache (refresh token, ~90-day life) before it expires. The
 * Vite dev proxy reads `getToken()` synchronously on every `/api/data`
 * request and injects it as the Authorization header — so the browser never
 * holds a token and live data keeps showing without manual `dev-token` reruns.
 *
 * Only the first run (or after the ~90-day refresh token lapses) prints a
 * device-code prompt in the dev-server terminal.
 */
import { resolveToken } from './dataverse/auth.mjs'

function decodeExpMs(tok) {
  try {
    const payload = JSON.parse(Buffer.from(tok.split('.')[1], 'base64').toString())
    return (payload.exp || 0) * 1000
  } catch {
    return 0
  }
}

export function startDevTokenProvider(dataverseUrl) {
  let current = ''
  let expMs = 0
  let inFlight = null

  async function refresh() {
    if (inFlight) return inFlight
    inFlight = (async () => {
      try {
        const tok = await resolveToken(dataverseUrl)
        if (tok) {
          current = tok
          expMs = decodeExpMs(tok)
          const mins = Math.round((expMs - Date.now()) / 60000)
          console.log(`[dev-token] Dataverse token refreshed (valid ~${mins} min)`)
        } else {
          console.error('[dev-token] could not resolve a Dataverse token')
        }
      } catch (err) {
        console.error('[dev-token] refresh failed:', err.message)
      } finally {
        inFlight = null
      }
      return current
    })()
    return inFlight
  }

  // Warm immediately, then proactively refresh every 20 min.
  refresh()
  const timer = setInterval(refresh, 20 * 60 * 1000)
  timer.unref?.()

  return {
    /** Synchronous read for the proxy; kicks a non-blocking refresh if stale. */
    getToken() {
      if (!current || Date.now() > expMs - 5 * 60 * 1000) refresh()
      return current
    },
  }
}
