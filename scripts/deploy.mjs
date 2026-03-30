#!/usr/bin/env node
/**
 * deploy.mjs – Push code-first Power App to Power Platform
 *
 * Replicates the exact 6-step workflow that `pac code push` uses internally:
 *   1. Generate an ephemeral Azure Blob Storage SAS URL
 *   2. Upload every file under buildPath as BlockBlobs
 *   3. Create (or update) the Power App
 *   4. (optional) Make solution-aware
 *   5. Publish the app
 *   6. Write the appId back to power.config.json
 *
 * Auth: MSAL device-code flow targeting https://api.powerplatform.com/.default
 *       Uses the PAC CLI first-party app reg (9cee029c-6210-4654-90bb-17e6e9d36617)
 *
 * Usage:
 *   node scripts/deploy.mjs                      # create or update
 *   node scripts/deploy.mjs --solution <name>     # also add to solution
 */
import { readFileSync, writeFileSync, existsSync, readdirSync, statSync } from 'node:fs'
import { join, relative, extname } from 'node:path'
import { createRequire } from 'node:module'
import { tmpdir } from 'node:os'
import { randomUUID } from 'node:crypto'
import { fileURLToPath } from 'node:url'

const require = createRequire(import.meta.url)
const __dirname = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(__dirname, '..')

/* ------------------------------------------------------------------ */
/*  Config                                                             */
/* ------------------------------------------------------------------ */
const configPath = join(ROOT, 'power.config.json')
const config = JSON.parse(readFileSync(configPath, 'utf8'))
const {
  appId,
  appDisplayName,
  environmentId,
  buildPath,
  buildEntryPoint,
  description,
  databaseReferences,
  connectionReferences,
} = config

const absDistDir = join(ROOT, buildPath)
if (!existsSync(absDistDir)) {
  console.error(`Build directory "${absDistDir}" not found. Run 'npm run build' first.`)
  process.exit(1)
}

/* ------------------------------------------------------------------ */
/*  Derive Power Platform API endpoint from environmentId              */
/* ------------------------------------------------------------------ */
const envNorm = environmentId.replace(/-/g, '')           // remove dashes
const hexSuffix = envNorm.slice(-2)                        // last 2 hex chars
const hexPrefix = envNorm.slice(0, -2)                     // everything else
const API_BASE = `https://${hexPrefix}.${hexSuffix}.environment.api.powerplatform.com`

console.log(`Environment : ${environmentId}`)
console.log(`API endpoint: ${API_BASE}`)

/* ------------------------------------------------------------------ */
/*  Auth – acquire token for https://api.powerplatform.com             */
/* ------------------------------------------------------------------ */
const MSAL_CACHE_PATH = join(tmpdir(), 'pth-deploy-msal-cache.json')
const PAC_CLIENT_ID = '9cee029c-6210-4654-90bb-17e6e9d36617' // PAC CLI first-party
const SCOPE = 'https://api.powerplatform.com/.default'

async function acquireToken() {
  // 1. Check env-var shortcut
  if (process.env.PP_TOKEN) return process.env.PP_TOKEN

  const msal = require('@azure/msal-node')
  let cacheData = ''
  if (existsSync(MSAL_CACHE_PATH)) {
    cacheData = readFileSync(MSAL_CACHE_PATH, 'utf8')
  }

  const cachePlugin = {
    beforeCacheAccess: async (ctx) => { if (cacheData) ctx.tokenCache.deserialize(cacheData) },
    afterCacheAccess: async (ctx) => {
      if (ctx.cacheHasChanged) {
        cacheData = ctx.tokenCache.serialize()
        writeFileSync(MSAL_CACHE_PATH, cacheData, 'utf8')
      }
    },
  }

  const pca = new msal.PublicClientApplication({
    auth: {
      clientId: PAC_CLIENT_ID,
      authority: 'https://login.microsoftonline.com/organizations',
    },
    cache: { cachePlugin },
  })

  // Try silent first
  const accounts = await pca.getTokenCache().getAllAccounts()
  if (accounts.length > 0) {
    try {
      const silent = await pca.acquireTokenSilent({
        scopes: [SCOPE],
        account: accounts[0],
      })
      if (silent?.accessToken) {
        console.log(`✓ Token (silent) for ${accounts[0].username}`)
        return silent.accessToken
      }
    } catch { /* fall through */ }
  }

  console.log('Authenticating via device code…')
  const result = await pca.acquireTokenByDeviceCode({
    scopes: [SCOPE],
    deviceCodeCallback: (r) => console.log(r.message),
  })
  if (result?.accessToken) {
    console.log('✓ Token acquired')
    return result.accessToken
  }
  throw new Error('Authentication failed')
}

/* ------------------------------------------------------------------ */
/*  Helpers                                                            */
/* ------------------------------------------------------------------ */
/** Call Power Platform API */
async function ppApi(method, path, token, body) {
  const url = `${API_BASE}${path}${path.includes('?') ? '&' : '?'}api-version=1`
  const opts = {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      'x-ms-client-request-id': randomUUID(),
    },
  }
  if (body !== undefined) opts.body = JSON.stringify(body)
  const res = await fetch(url, opts)
  const text = await res.text()
  if (!res.ok) {
    throw new Error(`PP API ${method} ${path} → ${res.status}\n${text}`)
  }
  return text ? JSON.parse(text) : null
}

/** Recursively list all files under a dir */
function walkDir(dir) {
  const result = []
  for (const name of readdirSync(dir)) {
    const full = join(dir, name)
    if (statSync(full).isDirectory()) {
      result.push(...walkDir(full))
    } else {
      result.push(full)
    }
  }
  return result
}

/** Ext → MIME type (matching PAC CLI behavior) */
function mimeFor(ext) {
  switch (ext.toLowerCase()) {
    case '.html': return 'text/html'
    case '.js':   return 'application/javascript'
    case '.css':  return 'text/css'
    case '.json': return 'application/json'
    case '.svg':  return 'image/svg+xml'
    case '.png':  return 'image/png'
    case '.jpg':
    case '.jpeg': return 'image/jpeg'
    case '.woff': return 'font/woff'
    case '.woff2':return 'font/woff2'
    default:      return 'application/octet-stream'
  }
}

/* ------------------------------------------------------------------ */
/*  Step 1: Generate Blob Storage SAS                                  */
/* ------------------------------------------------------------------ */
async function generateSas(token) {
  console.log('\n[1/6] Generating blob storage SAS …')
  const data = await ppApi('POST', '/powerapps/generateResourceStorage', token, {})
  const sas = data.sharedAccessSignature
  if (!sas) throw new Error('No SAS URL returned')
  console.log(`  SAS URL generated (${sas.substring(0, 60)}…)`)
  return sas
}

/* ------------------------------------------------------------------ */
/*  Step 2: Upload build files as BlockBlobs                           */
/* ------------------------------------------------------------------ */
async function uploadFiles(sasBaseUrl, distDir) {
  console.log('\n[2/6] Uploading build files …')
  const files = walkDir(distDir)
  let entryBlobUrl = ''

  // Parse SAS URL: container base + query string
  const sasUrl = new URL(sasBaseUrl)
  const basePath = sasUrl.origin + sasUrl.pathname   // e.g. https://xyz.blob.core.windows.net/container
  const sasQuery = sasUrl.search                      // e.g. ?sv=...&sig=...

  for (const filePath of files) {
    const relPath = relative(distDir, filePath).replace(/\\/g, '/')
    const blobUrl = `${basePath}/${relPath}${sasQuery}`
    const content = readFileSync(filePath)
    const mime = mimeFor(extname(filePath))

    const res = await fetch(blobUrl, {
      method: 'PUT',
      headers: {
        'Content-Type': mime,
        'x-ms-blob-type': 'BlockBlob',
      },
      body: content,
    })
    if (!res.ok) {
      const errText = await res.text()
      throw new Error(`Blob upload failed for ${relPath}: ${res.status}\n${errText}`)
    }
    console.log(`  ✓ ${relPath}  (${(content.length / 1024).toFixed(1)} KB, ${mime})`)

    if (relPath === buildEntryPoint) {
      entryBlobUrl = blobUrl
    }
  }

  if (!entryBlobUrl) {
    throw new Error(`Entry point "${buildEntryPoint}" not found in build output`)
  }
  return entryBlobUrl
}

/* ------------------------------------------------------------------ */
/*  Step 3: Create or Update App                                       */
/* ------------------------------------------------------------------ */
function buildDbRefs() {
  // Replicate the constructDatabaseReferences logic from PushApp.js
  // power.config.json databaseReferences is keyed by e.g. "default.cds"
  // Each entry has { dataSources: {...}, environmentVariableName: "" }
  const result = {}
  if (!databaseReferences || Object.keys(databaseReferences).length === 0) return result

  for (const [key, value] of Object.entries(databaseReferences)) {
    result[key] = {
      databaseDetails: {
        environmentName: key,
        overrideValues: {
          environmentVariableName: value?.environmentVariableName ?? '',
        },
      },
      dataSources: value?.dataSources,
    }
  }
  return result
}

function buildAppMetadata(entryBlobUrl) {
  return {
    appType: 'CodeApp',
    appSubtype: 'BYOCApp',
    properties: {
      lifeCycleId: 'Draft',
      displayName: appDisplayName || 'PTH-App',
      description: description || '',
      backgroundColor: 'RGBA(255,255,255,1)',
      backgroundImageUri: '',
      createdByClientVersion: { major: 1 },
      appUris: {
        codeAppPackageUri: {
          value: entryBlobUrl,
          readonlyValue: entryBlobUrl,
        },
      },
      environment: {
        name: environmentId,
        id: `/providers/Microsoft.PowerApps/environments/${environmentId}`,
      },
      connectionReferences: connectionReferences || {},
      databaseReferences: buildDbRefs(),
    },
  }
}

async function createApp(token, entryBlobUrl) {
  console.log('\n[3/6] Creating new Power App …')
  const meta = buildAppMetadata(entryBlobUrl)
  const result = await ppApi('POST', '/powerapps/apps', token, meta)
  const newId = result?.name || result?.appId || result?.id
  console.log(`  ✓ App created: ${newId}`)
  return newId
}

async function updateApp(token, existingAppId, entryBlobUrl) {
  console.log('\n[3/6] Updating existing Power App …')

  // Start writer session
  console.log('  Starting writer session …')
  const session = await ppApi('POST',
    `/powerapps/apps/${existingAppId}/startSession?forceStartNewSession=true`,
    token, { sessionType: 'Writer' })
  const sessionId = session?.sessionId || session?.id
  console.log(`  Session: ${sessionId}`)

  // Save
  const meta = buildAppMetadata(entryBlobUrl)
  await ppApi('PUT', `/powerapps/apps/${existingAppId}`, token, meta)
  console.log('  ✓ App metadata saved')

  // End session
  if (sessionId) {
    await ppApi('POST',
      `/powerapps/apps/${existingAppId}/sessions/${sessionId}/endSession`,
      token)
    console.log('  ✓ Session closed')
  }
  return existingAppId
}

/* ------------------------------------------------------------------ */
/*  Step 4: Make solution-aware (optional)                             */
/* ------------------------------------------------------------------ */
async function makeSolutionAware(token, appName, solutionName) {
  if (!solutionName) return
  console.log(`\n[4/6] Making solution-aware: ${solutionName} …`)
  await ppApi('POST', `/powerapps/apps/${appName}/makeSolutionAware`, token, {
    solutionId: solutionName,
  })
  console.log('  ✓ Done')
}

/* ------------------------------------------------------------------ */
/*  Step 5: Publish                                                    */
/* ------------------------------------------------------------------ */
async function publishApp(token, appName) {
  console.log('\n[5/6] Publishing app …')
  await ppApi('POST', `/powerapps/apps/${appName}/publish`, token)
  console.log('  ✓ App published')
}

/* ------------------------------------------------------------------ */
/*  Step 6: Update power.config.json                                   */
/* ------------------------------------------------------------------ */
function saveAppId(newAppId) {
  console.log('\n[6/6] Saving appId to power.config.json …')
  config.appId = newAppId
  writeFileSync(configPath, JSON.stringify(config, null, 2) + '\n', 'utf8')
  console.log(`  ✓ appId = ${newAppId}`)
}

/* ------------------------------------------------------------------ */
/*  Main                                                               */
/* ------------------------------------------------------------------ */
;(async () => {
  try {
    // Parse args
    const args = process.argv.slice(2)
    const solIdx = args.indexOf('--solution')
    const solutionName = solIdx >= 0 ? args[solIdx + 1] : null

    const token = await acquireToken()

    const sasUrl = await generateSas(token)
    const entryBlobUrl = await uploadFiles(sasUrl, absDistDir)

    let finalAppId
    if (appId) {
      finalAppId = await updateApp(token, appId, entryBlobUrl)
    } else {
      finalAppId = await createApp(token, entryBlobUrl)
    }

    await makeSolutionAware(token, finalAppId, solutionName)
    await publishApp(token, finalAppId)
    saveAppId(finalAppId)

    console.log('\n🎉 Deploy complete!')
    console.log(`   App: ${appDisplayName}`)
    console.log(`   ID : ${finalAppId}`)
    console.log(`   Env: ${environmentId}`)
  } catch (err) {
    console.error('\n❌ Deploy failed:', err.message)
    if (err.cause) console.error('  Cause:', err.cause)
    process.exit(1)
  }
})()
