/**
 * Provisions the 6 missing columns in the pth_project Dataverse table.
 */
import { resolveToken } from './auth.mjs'

const dvUrl = process.env.DATAVERSE_URL || 'https://org2d99840c.crm.dynamics.com'
const token = await resolveToken(dvUrl)
const api = `${dvUrl.replace(/\/+$/, '')}/api/data/v9.2`

const headers = {
  Authorization: `Bearer ${token}`,
  'Content-Type': 'application/json',
  Accept: 'application/json',
  'OData-MaxVersion': '4.0',
  'OData-Version': '4.0',
}

function label(text) {
  return {
    '@odata.type': 'Microsoft.Dynamics.CRM.Label',
    LocalizedLabels: [
      { '@odata.type': 'Microsoft.Dynamics.CRM.LocalizedLabel', Label: text, LanguageCode: 1033 },
    ],
  }
}

const columns = [
  {
    '@odata.type': '#Microsoft.Dynamics.CRM.StringAttributeMetadata',
    SchemaName: 'pth_BoschCode',
    LogicalName: 'pth_boschcode',
    DisplayName: label('Bosch Code'),
    Description: label('Bosch project code'),
    RequiredLevel: { Value: 'None' },
    MaxLength: 50,
    FormatName: { Value: 'Text' },
  },
  {
    '@odata.type': '#Microsoft.Dynamics.CRM.DecimalAttributeMetadata',
    SchemaName: 'pth_BudgetAllocated',
    LogicalName: 'pth_budgetallocated',
    DisplayName: label('Budget Allocated'),
    Description: label('Budget allocated'),
    RequiredLevel: { Value: 'None' },
    Precision: 2,
    MinValue: 0,
    MaxValue: 999999999,
  },
  {
    '@odata.type': '#Microsoft.Dynamics.CRM.MemoAttributeMetadata',
    SchemaName: 'pth_KpisImpacted',
    LogicalName: 'pth_kpisimpacted',
    DisplayName: label('KPIs Impacted'),
    Description: label('KPIs impacted by this project'),
    RequiredLevel: { Value: 'None' },
    MaxLength: 1000,
    Format: 'Text',
  },
  {
    '@odata.type': '#Microsoft.Dynamics.CRM.MemoAttributeMetadata',
    SchemaName: 'pth_MonitoringCriteria',
    LogicalName: 'pth_monitoringcriteria',
    DisplayName: label('Monitoring Criteria'),
    Description: label('Monitoring criteria'),
    RequiredLevel: { Value: 'None' },
    MaxLength: 500,
    Format: 'Text',
  },
  {
    '@odata.type': '#Microsoft.Dynamics.CRM.DateTimeAttributeMetadata',
    SchemaName: 'pth_PlannedStartDate',
    LogicalName: 'pth_plannedstartdate',
    DisplayName: label('Planned Start Date'),
    Description: label('Planned start date'),
    RequiredLevel: { Value: 'None' },
    Format: 'DateOnly',
    DateTimeBehavior: { Value: 'DateOnly' },
  },
  {
    '@odata.type': '#Microsoft.Dynamics.CRM.DateTimeAttributeMetadata',
    SchemaName: 'pth_PlannedEndDate',
    LogicalName: 'pth_plannedenddate',
    DisplayName: label('Planned End Date'),
    Description: label('Planned end date'),
    RequiredLevel: { Value: 'None' },
    Format: 'DateOnly',
    DateTimeBehavior: { Value: 'DateOnly' },
  },
]

const endpoint = `${api}/EntityDefinitions(LogicalName='pth_project')/Attributes`

for (const col of columns) {
  console.log(`Creating ${col.LogicalName} ...`)
  const res = await fetch(endpoint, {
    method: 'POST',
    headers,
    body: JSON.stringify(col),
  })
  if (res.ok || res.status === 204) {
    console.log(`  ✓ ${col.LogicalName}`)
  } else {
    const errText = await res.text()
    console.error(`  ✗ ${col.LogicalName} (${res.status}):`, errText.slice(0, 300))
  }
}

// Publish the table so new columns are available
console.log('\nPublishing pth_project table ...')
const pubRes = await fetch(`${api}/PublishXml`, {
  method: 'POST',
  headers,
  body: JSON.stringify({
    ParameterXml: '<importexportxml><entities><entity>pth_project</entity></entities></importexportxml>',
  }),
})
if (pubRes.ok || pubRes.status === 204) {
  console.log('  ✓ Published')
} else {
  console.error('  ✗ Publish failed:', pubRes.status, (await pubRes.text()).slice(0, 200))
}

console.log('\nDone.')
