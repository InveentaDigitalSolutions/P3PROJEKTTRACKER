/**
 * Shared Dataverse auth helper — resolves an access token via:
 *   1. DATAVERSE_TOKEN env var
 *   2. Azure CLI (`az`)
 *   3. MSAL device-code flow with persistent on-disk cache
 */
import { readFileSync, writeFileSync, existsSync } from 'node:fs'
import { execSync } from 'node:child_process'
import { createRequire } from 'node:module'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const require = createRequire(import.meta.url)
const TOKEN_CACHE_FILE = join(tmpdir(), 'pth-msal-cache.json')

export async function resolveToken(dataverseUrl) {
  if (process.env.DATAVERSE_TOKEN) return process.env.DATAVERSE_TOKEN

  // Try Azure CLI
  try {
    const cmd = `az account get-access-token --resource ${dataverseUrl} --query accessToken -o tsv`
    const token = execSync(cmd, { encoding: 'utf8', stdio: ['pipe', 'pipe', 'ignore'] }).trim()
    if (token) return token
  } catch { /* fall through */ }

  // MSAL device-code with persistent cache
  try {
    const msal = require('@azure/msal-node')
    const envHost = dataverseUrl.replace(/\/+$/, '')

    let cacheData = ''
    if (existsSync(TOKEN_CACHE_FILE)) {
      cacheData = readFileSync(TOKEN_CACHE_FILE, 'utf8')
    }

    const cachePlugin = {
      beforeCacheAccess: async (ctx) => { if (cacheData) ctx.tokenCache.deserialize(cacheData) },
      afterCacheAccess: async (ctx) => {
        if (ctx.cacheHasChanged) {
          cacheData = ctx.tokenCache.serialize()
          writeFileSync(TOKEN_CACHE_FILE, cacheData, 'utf8')
        }
      },
    }

    const pca = new msal.PublicClientApplication({
      auth: {
        clientId: '51f81489-12ee-4a9e-aaae-a2591f45987d',
        authority: 'https://login.microsoftonline.com/organizations',
      },
      cache: { cachePlugin },
    })

    // Silent first (from cache)
    const accounts = await pca.getTokenCache().getAllAccounts()
    if (accounts.length > 0) {
      try {
        const silent = await pca.acquireTokenSilent({
          scopes: [`${envHost}/.default`],
          account: accounts[0],
        })
        if (silent?.accessToken) {
          console.log(`Token acquired silently for ${accounts[0].username}`)
          return silent.accessToken
        }
      } catch { /* fall through to device code */ }
    }

    console.log('Authenticating via device code flow…')
    const result = await pca.acquireTokenByDeviceCode({
      scopes: [`${envHost}/.default`],
      deviceCodeCallback: (r) => console.log(r.message),
    })
    if (result?.accessToken) {
      console.log('Token acquired via MSAL device-code flow')
      return result.accessToken
    }
  } catch (err) {
    console.error('MSAL authentication failed:', err.message)
  }

  return ''
}
