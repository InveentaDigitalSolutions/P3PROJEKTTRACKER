/**
 * Shared email layout — matches the Project Tracker Hub design system.
 *
 * Design tokens:
 *   --pth-bg:       #F5F5F7     --pth-card:     #FFFFFF
 *   --pth-panel:    #E8E8ED     --pth-border:   rgba(0,0,0,0.08)
 *   --pth-text:     #1D1D1F     --pth-muted:    #86868B
 *   --pth-blue:     #005691     --pth-red:      #DC2626
 *   --pth-teal:     #18837E     --pth-sidebar:  #31343A
 *   --pth-btn:      #31343A     --pth-accent:   #FF4D5E
 *
 * All emails use this wrapper so they look consistent and branded.
 */

const LOGO_URL = 'https://pthnotifstore.blob.core.windows.net/assets/logo.png'

// ── Design tokens (inline-safe for email) ────────────────────────────────
export const C = {
  bg:        '#F5F5F7',
  card:      '#FFFFFF',
  panel:     '#E8E8ED',
  border:    '#E0E0E2',
  text:      '#1D1D1F',
  muted:     '#86868B',
  blue:      '#005691',
  red:       '#DC2626',
  teal:      '#18837E',
  sidebar:   '#31343A',
  btn:       '#31343A',
  accent:    '#FF4D5E',
  green:     '#18837E',
  yellow:    '#F59E0B',
  cardHover: '#F9FAFB',
} as const

// ── Shared helpers ───────────────────────────────────────────────────────

/** Status pill like the app sidebar badges */
export function statusBadge(status: string): string {
  const map: Record<string, { bg: string; fg: string }> = {
    GREEN:  { bg: '#E6F4F0', fg: C.teal },
    YELLOW: { bg: '#FEF3C7', fg: '#B45309' },
    RED:    { bg: '#FEE2E2', fg: C.red },
    DONE:   { bg: '#E6F4F0', fg: C.teal },
    OPEN:   { bg: '#EFF6FF', fg: C.blue },
    IN_PROGRESS: { bg: '#FEF3C7', fg: '#B45309' },
    CLOSED: { bg: '#F3F4F6', fg: C.muted },
  }
  const s = map[status] ?? { bg: '#F3F4F6', fg: C.muted }
  return `<span style="display:inline-block;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:600;letter-spacing:.3px;background:${s.bg};color:${s.fg}">${status}</span>`
}

/** Category pill matching the app's colour scheme */
export function categoryBadge(cat: string): string {
  const map: Record<string, { bg: string; fg: string }> = {
    ECR:          { bg: '#EFF6FF', fg: '#1565C0' },
    CIP:          { bg: '#E6F4F0', fg: C.teal },
    'Path Forward': { bg: '#FEF3C7', fg: '#92400E' },
  }
  const s = map[cat] ?? { bg: '#F5E6FB', fg: '#7B1FA2' }
  return `<span style="display:inline-block;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:600;letter-spacing:.3px;background:${s.bg};color:${s.fg}">${cat}</span>`
}

/** Small status dot used in tables */
export function statusDot(status: string): string {
  const clr = status === 'RED' ? C.red : status === 'YELLOW' ? C.yellow : C.teal
  return `<span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${clr}"></span>`
}

/** Progress bar matching the app's rounded bars */
export function progressBar(pct: number, width = 80): string {
  const fill = pct === 100 ? C.teal : C.blue
  return `
    <div style="display:inline-block;vertical-align:middle;width:${width}px;background:${C.panel};border-radius:6px;height:6px;overflow:hidden">
      <div style="width:${pct}%;height:100%;border-radius:6px;background:${fill}"></div>
    </div>
    <span style="font-size:11px;color:${C.muted};margin-left:5px">${pct}%</span>`
}

// ── Main layout wrapper ──────────────────────────────────────────────────

export interface EmailLayoutOptions {
  /** Icon emoji for the header (e.g., "🚀") */
  icon: string
  /** Main heading text */
  title: string
  /** Optional subtitle line */
  subtitle?: string
  /** The inner HTML content */
  body: string
  /** Header accent colour override (defaults to sidebar dark: #31343A) */
  headerColor?: string
}

/**
 * Wraps content in the branded PTH email frame.
 * Produces a full email-safe HTML string (inline styles only).
 */
export function emailLayout(opts: EmailLayoutOptions): string {
  const hc = opts.headerColor ?? C.sidebar

  return `
<!DOCTYPE html>
<html lang="en">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:${C.bg};-webkit-font-smoothing:antialiased">
<div style="font-family:'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif;max-width:680px;margin:0 auto;padding:24px 16px">

  <!-- Card wrapper -->
  <div style="background:${C.card};border-radius:16px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.06)">

    <!-- Header banner -->
    <div style="background:${hc};padding:28px 32px 24px;text-align:left">
      <table cellpadding="0" cellspacing="0" border="0" style="width:100%">
        <tr>
          <td style="vertical-align:middle">
            <img src="${LOGO_URL}" alt="Project Tracker Hub" width="48" height="32" style="display:block;border:0;width:48px;height:auto;margin-bottom:12px" />
            <h1 style="margin:0;font-size:20px;font-weight:700;color:#FFFFFF;line-height:1.3">${opts.icon} ${opts.title}</h1>
            ${opts.subtitle ? `<p style="margin:8px 0 0;font-size:13px;color:rgba(255,255,255,0.72);line-height:1.4">${opts.subtitle}</p>` : ''}
          </td>
        </tr>
      </table>
    </div>

    <!-- Body -->
    <div style="padding:28px 32px">
      ${opts.body}
    </div>

  </div>

  <!-- Footer -->
  <div style="padding:20px 32px 8px;text-align:center">
    <p style="margin:0;font-size:11px;color:${C.muted};line-height:1.6">
      This is an automated notification from <strong style="color:${C.text}">Project Tracker Hub</strong>.<br>
      You are receiving this because you are a notification recipient in PTH.
    </p>
    <p style="margin:8px 0 0;font-size:10px;color:${C.muted}">
      &copy; ${new Date().getFullYear()} Inveenta &middot; Powered by Azure Functions + Microsoft Graph
    </p>
  </div>

</div>
</body>
</html>`
}

// ── Table helpers (matching app's clean card tables) ─────────────────────

export function tableHead(columns: string[]): string {
  return `<table cellpadding="0" cellspacing="0" border="0" style="width:100%;border-collapse:separate;border-spacing:0;font-size:13px;margin-bottom:20px">
    <tr>
      ${columns.map((col) => `<th style="padding:10px 14px;text-align:left;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.5px;color:${C.muted};border-bottom:2px solid ${C.panel};background:${C.cardHover}">${col}</th>`).join('')}
    </tr>`
}

export function tableRow(cells: string[], highlight = false): string {
  const bg = highlight ? '#FFF5F5' : C.card
  return `<tr>
    ${cells.map((cell) => `<td style="padding:10px 14px;border-bottom:1px solid ${C.panel};background:${bg};color:${C.text};font-size:13px;line-height:1.4">${cell}</td>`).join('')}
  </tr>`
}

export function tableEnd(): string {
  return '</table>'
}

// ── Info row (key-value) ─────────────────────────────────────────────────

export function infoRow(label: string, value: string): string {
  return `<tr>
    <td style="padding:10px 14px;width:40%;font-size:12px;font-weight:600;color:${C.muted};text-transform:uppercase;letter-spacing:.4px;border-bottom:1px solid ${C.panel};background:${C.cardHover}">${label}</td>
    <td style="padding:10px 14px;font-size:13px;color:${C.text};border-bottom:1px solid ${C.panel}">${value}</td>
  </tr>`
}

export function infoTable(rows: [string, string][]): string {
  return `<table cellpadding="0" cellspacing="0" border="0" style="width:100%;border-collapse:separate;border-spacing:0;margin-bottom:20px;border-radius:10px;overflow:hidden;border:1px solid ${C.panel}">
    ${rows.map(([l, v]) => infoRow(l, v)).join('')}
  </table>`
}

// ── Section heading ──────────────────────────────────────────────────────

export function sectionHeading(text: string, accent: string = C.blue): string {
  return `<h2 style="margin:24px 0 12px;font-size:14px;font-weight:700;color:${C.text};padding-bottom:8px;border-bottom:2px solid ${accent}">${text}</h2>`
}

// ── KPI card row (for weekly digest) ─────────────────────────────────────

export interface KpiCard { label: string; value: string | number; color: string }

export function kpiRow(cards: KpiCard[]): string {
  // Use table-based layout for email client compat
  return `<table cellpadding="0" cellspacing="0" border="0" style="width:100%;margin-bottom:24px">
    <tr>
      ${cards.map((c) => `
        <td style="text-align:center;padding:8px 4px;width:${Math.floor(100 / cards.length)}%">
          <div style="background:${C.cardHover};border:1px solid ${C.panel};border-radius:12px;padding:16px 8px">
            <div style="font-size:28px;font-weight:700;color:${c.color};line-height:1">${c.value}</div>
            <div style="font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.6px;color:${C.muted};margin-top:6px">${c.label}</div>
          </div>
        </td>`).join('')}
    </tr>
  </table>`
}

// ── Callout box (for objective, notes) ───────────────────────────────────

export function calloutBox(heading: string, text: string): string {
  return `<div style="background:${C.cardHover};border:1px solid ${C.panel};border-radius:12px;padding:16px 18px;margin-bottom:20px">
    <h3 style="margin:0 0 6px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:${C.blue}">${heading}</h3>
    <p style="margin:0;font-size:13px;line-height:1.6;color:${C.text}">${text}</p>
  </div>`
}
