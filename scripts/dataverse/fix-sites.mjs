/**
 * fix-sites.mjs — Update the pth_site column on all existing project records.
 *
 * Maps old city names to "TCA" (Toluca) or "SLP" (San Luis Potosí).
 *
 * Usage:
 *   DATAVERSE_URL=https://orgname.crm.dynamics.com node scripts/dataverse/fix-sites.mjs
 */
import { resolveToken } from './auth.mjs'
import process from 'node:process'

const dataverseUrl = process.env.DATAVERSE_URL
if (!dataverseUrl) {
  console.error('Missing DATAVERSE_URL env var')
  process.exit(1)
}

const token = await resolveToken(dataverseUrl)
if (!token) {
  console.error('Could not obtain a Dataverse access token.')
  process.exit(1)
}

const API = `${dataverseUrl.replace(/\/+$/, '')}/api/data/v9.2`

/** Map old site values → new abbreviations */
const SITE_FIX = {
  stuttgart: 'TCA',
  turin:    'SLP',
  madrid:   'TCA',
  wroclaw:  'SLP',
  tip:      'TCA',
  slpp:     'SLP',
}

async function run() {
  // Fetch all projects
  const res = await fetch(`${API}/pth_projects?$select=pth_projectid,pth_projectname,pth_site`, {
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
  })
  if (!res.ok) {
    console.error('Failed to fetch projects:', res.status, await res.text())
    process.exit(1)
  }
  const { value: projects } = await res.json()
  console.log(`Found ${projects.length} project(s)\n`)

  let updated = 0
  for (const p of projects) {
    const current = (p.pth_site ?? '').toString().trim()
    const key = current.toLowerCase().replace(/\s+/g, '_')
    const newSite = SITE_FIX[key]

    if (!newSite) {
      if (current === 'TCA' || current === 'SLP') {
        console.log(`  ✓ "${p.pth_projectname}" — already "${current}", skipping`)
      } else {
        console.warn(`  ? "${p.pth_projectname}" — unknown site "${current}", skipping`)
      }
      continue
    }

    console.log(`  → "${p.pth_projectname}": "${current}" → "${newSite}"`)
    const patchRes = await fetch(`${API}/pth_projects(${p.pth_projectid})`, {
      method: 'PATCH',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
      },
      body: JSON.stringify({ pth_site: newSite }),
    })
    if (patchRes.ok || patchRes.status === 204) {
      updated++
    } else {
      console.error(`    ✗ PATCH failed (${patchRes.status}):`, await patchRes.text())
    }
  }

  console.log(`\n✅ Updated ${updated} project(s)`)
}

run().catch((err) => { console.error(err); process.exit(1) })
