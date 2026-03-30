import { readFileSync, writeFileSync, existsSync } from 'node:fs'
import { execSync } from 'node:child_process'
import { createRequire } from 'node:module'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import process from 'node:process'

const require = createRequire(import.meta.url)

const schemaPath = new URL('../../dataverse/ppm.dataverse.schema.json', import.meta.url)
const schema = JSON.parse(readFileSync(schemaPath, 'utf8'))

const dryRun = process.argv.includes('--dry-run')
const dataverseUrl = process.env.DATAVERSE_URL

if (!dataverseUrl) {
  console.error('Missing DATAVERSE_URL. Example: https://orgname.crm.dynamics.com')
  process.exit(1)
}

const TOKEN_CACHE_FILE = join(tmpdir(), 'pth-msal-cache.json')

async function resolveToken() {
  if (process.env.DATAVERSE_TOKEN) {
    return process.env.DATAVERSE_TOKEN
  }

  // Try Azure CLI
  try {
    const command = `az account get-access-token --resource ${dataverseUrl} --query accessToken -o tsv`
    const token = execSync(command, { encoding: 'utf8', stdio: ['pipe', 'pipe', 'ignore'] }).trim()
    if (token) {
      return token
    }
  } catch {
    // Fall through
  }

  // Try MSAL with persistent token cache (avoids repeated device-code auth)
  try {
    const msal = require('@azure/msal-node')
    const envHost = dataverseUrl.replace(/\/+$/, '')

    // Load cache from disk if it exists
    let cacheData = ''
    if (existsSync(TOKEN_CACHE_FILE)) {
      cacheData = readFileSync(TOKEN_CACHE_FILE, 'utf8')
    }

    const cachePlugin = {
      beforeCacheAccess: async (context) => {
        if (cacheData) context.tokenCache.deserialize(cacheData)
      },
      afterCacheAccess: async (context) => {
        if (context.cacheHasChanged) {
          cacheData = context.tokenCache.serialize()
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

    // Try silent acquisition first (from cache)
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
      } catch {
        // Silent failed — fall through to device code
      }
    }

    // Fall back to device code flow
    console.log('Authenticating via device code flow…')
    const result = await pca.acquireTokenByDeviceCode({
      scopes: [`${envHost}/.default`],
      deviceCodeCallback: (response) => {
        console.log(response.message)
      },
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

const token = await resolveToken()
if (!token && !dryRun) {
  console.error('Missing DATAVERSE_TOKEN and unable to get token from Azure CLI or MSAL.')
  console.error('Set DATAVERSE_TOKEN manually or login with az / pac and ensure access to Dataverse environment.')
  process.exit(1)
}

function labelObject(label) {
  return {
    LocalizedLabels: [{ Label: label, LanguageCode: 1033 }],
  }
}

async function dataverseRequest(path, method, body) {
  const endpoint = `${dataverseUrl}/api/data/v9.2/${path}`
  const headers = {
    Accept: 'application/json',
    'Content-Type': 'application/json; charset=utf-8',
    'OData-MaxVersion': '4.0',
    'OData-Version': '4.0',
    Authorization: `Bearer ${token}`,
  }

  const response = await fetch(endpoint, {
    method,
    headers,
    body: body ? JSON.stringify(body) : undefined,
  })

  if (response.ok) {
    if (response.status === 204) {
      return null
    }

    const text = await response.text()
    return text ? JSON.parse(text) : null
  }

  const errorText = await response.text()
  throw new Error(`${method} ${path} failed (${response.status}) ${errorText}`)
}

function buildColumnPayload(column) {
  const required = column.required
    ? { RequiredLevel: { Value: 'ApplicationRequired' } }
    : { RequiredLevel: { Value: 'None' } }

  switch (column.type) {
    case 'string':
      return {
        '@odata.type': 'Microsoft.Dynamics.CRM.StringAttributeMetadata',
        SchemaName: column.schemaName,
        DisplayName: labelObject(column.displayName),
        Description: labelObject(column.displayName),
        MaxLength: column.maxLength ?? 100,
        FormatName: { Value: 'Text' },
        ...required,
      }
    case 'memo':
      return {
        '@odata.type': 'Microsoft.Dynamics.CRM.MemoAttributeMetadata',
        SchemaName: column.schemaName,
        DisplayName: labelObject(column.displayName),
        Description: labelObject(column.displayName),
        MaxLength: column.maxLength ?? 2000,
        FormatName: { Value: 'TextArea' },
        ...required,
      }
    case 'wholeNumber':
      return {
        '@odata.type': 'Microsoft.Dynamics.CRM.IntegerAttributeMetadata',
        SchemaName: column.schemaName,
        DisplayName: labelObject(column.displayName),
        Description: labelObject(column.displayName),
        MinValue: column.min ?? 0,
        MaxValue: column.max ?? 2147483647,
        ...required,
      }
    case 'date':
      return {
        '@odata.type': 'Microsoft.Dynamics.CRM.DateTimeAttributeMetadata',
        SchemaName: column.schemaName,
        DisplayName: labelObject(column.displayName),
        Description: labelObject(column.displayName),
        Format: 'DateOnly',
        ImeMode: 'Auto',
        ...required,
      }
    case 'boolean':
      return {
        '@odata.type': 'Microsoft.Dynamics.CRM.BooleanAttributeMetadata',
        SchemaName: column.schemaName,
        DisplayName: labelObject(column.displayName),
        Description: labelObject(column.displayName),
        DefaultValue: column.defaultValue ?? false,
        OptionSet: {
          TrueOption: {
            Value: 1,
            Label: labelObject('Yes'),
          },
          FalseOption: {
            Value: 0,
            Label: labelObject('No'),
          },
        },
        ...required,
      }
    case 'choice':
      return {
        '@odata.type': 'Microsoft.Dynamics.CRM.PicklistAttributeMetadata',
        SchemaName: column.schemaName,
        DisplayName: labelObject(column.displayName),
        Description: labelObject(column.displayName),
        OptionSet: {
          IsGlobal: false,
          OptionSetType: 'Picklist',
          Options: (column.options ?? []).map((option) => ({
            Value: option.value,
            Label: labelObject(option.label),
          })),
        },
        ...required,
      }
    case 'lookup':
      return {
        '@odata.type': 'Microsoft.Dynamics.CRM.LookupAttributeMetadata',
        SchemaName: column.schemaName,
        DisplayName: labelObject(column.displayName),
        Description: labelObject(column.displayName),
        Targets: [column.targetLogicalName],
        ...required,
      }
    default:
      throw new Error(`Unsupported column type: ${column.type}`)
  }
}

async function resolveLogicalName(schemaName) {
  try {
    const result = await dataverseRequest(`EntityDefinitions(LogicalName='${schemaName.toLowerCase()}')?$select=LogicalName`, 'GET')
    return result?.LogicalName ?? schemaName.toLowerCase()
  } catch {
    // Dataverse derives logical name from schema name as lowercase
    return schemaName.toLowerCase()
  }
}

async function tableExists(logicalName) {
  try {
    await dataverseRequest(`EntityDefinitions(LogicalName='${logicalName}')?$select=LogicalName`, 'GET')
    return true
  } catch {
    return false
  }
}

async function columnExists(tableLogicalName, columnLogicalName) {
  try {
    await dataverseRequest(
      `EntityDefinitions(LogicalName='${tableLogicalName}')/Attributes(LogicalName='${columnLogicalName}')?$select=LogicalName`,
      'GET',
    )
    return true
  } catch {
    return false
  }
}

async function createTable(table) {
  // Determine the primary name column — use the first string column or a default
  const primaryCol = table.columns.find((c) => c.type === 'string') ?? { schemaName: `${table.schemaName}Name`, displayName: 'Name', maxLength: 200 }
  const primaryLogical = primaryCol.schemaName.charAt(0).toLowerCase() + primaryCol.schemaName.slice(1)

  // Dataverse ignores LogicalName on create — it derives it from SchemaName (lowercase)
  const payload = {
    SchemaName: table.schemaName,
    DisplayName: labelObject(table.displayName),
    DisplayCollectionName: labelObject(table.displayCollectionName),
    Description: labelObject(table.description),
    OwnershipType: table.ownershipType,
    IsActivity: false,
    HasActivities: false,
    HasNotes: false,
    PrimaryNameAttribute: primaryLogical.toLowerCase(),
    Attributes: [
      {
        '@odata.type': 'Microsoft.Dynamics.CRM.StringAttributeMetadata',
        SchemaName: primaryCol.schemaName,
        LogicalName: primaryLogical.toLowerCase(),
        DisplayName: labelObject(primaryCol.displayName),
        Description: labelObject(primaryCol.displayName),
        IsPrimaryName: true,
        MaxLength: primaryCol.maxLength ?? 200,
        FormatName: { Value: 'Text' },
        RequiredLevel: { Value: 'ApplicationRequired' },
      },
    ],
  }

  if (dryRun) {
    console.log(`[DRY RUN] Create table ${table.schemaName}`)
    return table.schemaName.toLowerCase()
  }

  await dataverseRequest('EntityDefinitions', 'POST', payload)
  const logicalName = await resolveLogicalName(table.schemaName)
  console.log(`Created table ${logicalName} (schema: ${table.schemaName})`)
  return logicalName
}

async function createColumn(tableLogicalName, column) {
  if (dryRun) {
    console.log(`[DRY RUN] Create column ${column.logicalName} on ${tableLogicalName}`)
    return
  }

  const payload = buildColumnPayload(column)
  await dataverseRequest(`EntityDefinitions(LogicalName='${tableLogicalName}')/Attributes`, 'POST', payload)
  console.log(`Created column ${column.logicalName} on ${tableLogicalName}`)
}

async function createLookupRelationship(referencingTable, column, targetActualLogical) {
  // Lookup columns must be created as OneToMany relationships
  const referencedEntity = targetActualLogical
  const referencingEntity = referencingTable
  const relationshipSchemaName = `${column.schemaName}_${referencingEntity}_${referencedEntity}`

  const required = column.required
    ? { Value: 'ApplicationRequired' }
    : { Value: 'None' }

  const payload = {
    '@odata.type': 'Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata',
    SchemaName: relationshipSchemaName,
    ReferencedEntity: referencedEntity,
    ReferencingEntity: referencingEntity,
    CascadeConfiguration: {
      Assign: 'NoCascade',
      Delete: 'RemoveLink',
      Merge: 'NoCascade',
      Reparent: 'NoCascade',
      Share: 'NoCascade',
      Unshare: 'NoCascade',
    },
    Lookup: {
      SchemaName: column.schemaName,
      DisplayName: labelObject(column.displayName),
      Description: labelObject(column.displayName),
      RequiredLevel: required,
    },
  }

  if (dryRun) {
    console.log(`[DRY RUN] Create lookup ${column.schemaName} on ${referencingEntity} → ${referencedEntity}`)
    return
  }

  await dataverseRequest('RelationshipDefinitions', 'POST', payload)
  console.log(`Created lookup ${column.schemaName} on ${referencingEntity} → ${referencedEntity}`)
}

async function publishAllCustomizations() {
  if (dryRun) {
    console.log('[DRY RUN] Publish all customizations')
    return
  }

  await dataverseRequest('PublishAllXml', 'POST', {})
  console.log('Published all customizations')
}

async function run() {
  console.log(`Applying Dataverse schema: ${schema.solutionName}`)

  // Build a mapping from schema logicalName (e.g. pth_projects) to actual Dataverse logical name (e.g. pth_project)
  const schemaToActual = {}
  for (const table of schema.tables) {
    const derivedLogical = table.schemaName.toLowerCase()
    schemaToActual[table.logicalName] = derivedLogical
    schemaToActual[derivedLogical] = derivedLogical
  }

  // Collect lookup columns to create after all tables exist
  const pendingLookups = []

  // Phase 1: Create tables and non-lookup columns
  for (const table of schema.tables) {
    const primaryCol = table.columns.find((c) => c.type === 'string')
    const primaryLogical = primaryCol ? primaryCol.logicalName : null

    const derivedLogical = table.schemaName.toLowerCase()
    const exists = dryRun ? false : await tableExists(derivedLogical)
    let actualLogical = derivedLogical
    if (!exists) {
      actualLogical = await createTable(table)
    } else {
      console.log(`Table ${derivedLogical} already exists`)
    }

    for (const column of table.columns) {
      if (column.logicalName === primaryLogical) continue

      // Defer lookup columns to Phase 2
      if (column.type === 'lookup') {
        pendingLookups.push({ tableLogical: actualLogical, column })
        continue
      }

      const existsColumn = dryRun ? false : await columnExists(actualLogical, column.logicalName)
      if (existsColumn) {
        console.log(`Column ${column.logicalName} already exists on ${actualLogical}`)
        continue
      }
      await createColumn(actualLogical, column)
    }
  }

  // Phase 2: Create lookup columns via relationships (all target tables now exist)
  for (const { tableLogical, column } of pendingLookups) {
    // Check if the lookup column already exists
    const lookupLogical = column.schemaName.toLowerCase()
    const existsCol = dryRun ? false : await columnExists(tableLogical, lookupLogical)
    if (existsCol) {
      console.log(`Lookup ${lookupLogical} already exists on ${tableLogical}`)
      continue
    }

    // Resolve the target table's actual Dataverse logical name
    const targetActual = schemaToActual[column.targetLogicalName] ?? column.targetLogicalName
    await createLookupRelationship(tableLogical, column, targetActual)
  }

  await publishAllCustomizations()
  console.log('Dataverse schema apply completed')
}

run().catch((error) => {
  console.error(error.message)
  process.exit(1)
})
