import { resolveToken } from './auth.mjs'
const token = await resolveToken('https://org2d99840c.crm.dynamics.com')
const res = await fetch("https://org2d99840c.crm.dynamics.com/api/data/v9.2/EntityDefinitions?$select=LogicalName,SchemaName,LogicalCollectionName", { headers: { Authorization: 'Bearer ' + token } })
const data = await res.json()
console.log('Status:', res.status)
if (data.error) console.log('Error:', data.error.message)
;(data.value || []).filter(e => e.SchemaName.startsWith('pth_')).forEach(e => console.log(e.SchemaName, '->', e.LogicalName, '->', e.LogicalCollectionName))
