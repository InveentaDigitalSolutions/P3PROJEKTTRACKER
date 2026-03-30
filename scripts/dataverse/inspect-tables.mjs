import { createRequire } from 'node:module'
import process from 'node:process'

const require = createRequire(import.meta.url)
const msal = require('@azure/msal-node')

const dataverseUrl = process.env.DATAVERSE_URL || 'https://org2d99840c.crm.dynamics.com'

const pca = new msal.PublicClientApplication({
  auth: {
    clientId: '51f81489-12ee-4a9e-aaae-a2591f45987d',
    authority: 'https://login.microsoftonline.com/organizations',
  },
})

const result = await pca.acquireTokenByDeviceCode({
  scopes: [`${dataverseUrl}/.default`],
  deviceCodeCallback: (r) => console.log(r.message),
})

const token = result.accessToken
const url = `${dataverseUrl}/api/data/v9.2/EntityDefinitions?$filter=startswith(SchemaName,'pth_')&$select=LogicalName,SchemaName`
const res = await fetch(url, {
  headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
})
const data = await res.json()
console.log('Tables with pth_ prefix:')
for (const t of data.value ?? []) {
  console.log(`  schema=${t.SchemaName}  logical=${t.LogicalName}`)
}
