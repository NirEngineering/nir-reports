// Read-only lookup against the Google Sheet nir-dashboard-new (the document
// validity-tracking dashboard) uses as its data source — lets the WhatsApp
// wizard offer "is this a renewal of an existing site?" when the field
// engineer names a known client, instead of starting every report cold.
//
// Uses the exact same spreadsheet + read-only API key the dashboard's own
// dashboard.html reads with — a key already public and committed there
// (Google API keys scoped to read-only Sheets access are meant to ship in
// client-side code), so no new credential was introduced for this.
const SHEET_ID = process.env.DASHBOARD_SHEET_ID || '1mZZq0QrQVqzNJ66ErJ-h81FSkO_5KIu3orePE5A9BqE';
const API_KEY = process.env.DASHBOARD_API_KEY || 'AIzaSyDJy8tolZu8z-IuMFXSmRfsFFrM_rJkK8w';
const SHEET_NAME = process.env.DASHBOARD_SHEET_NAME || 'מסמכים';

// Strips the RTL/LTR mark characters Google Sheets sometimes injects into
// Hebrew cell text — same cleanup dashboard.html's own cl() does.
function clean(s) {
  if (!s) return '';
  let r = String(s);
  for (const c of '‏‎‪‫‬﻿​') r = r.split(c).join('');
  return r.trim();
}

let cache = null; // { rows, fetchedAt }
const CACHE_MS = 2 * 60 * 1000; // avoid re-fetching the whole sheet on every keystroke of a wizard run

async function fetchAllRows() {
  if (cache && Date.now() - cache.fetchedAt < CACHE_MS) return cache.rows;

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(SHEET_NAME)}?key=${API_KEY}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Sheets API error: ${res.status}`);
  const data = await res.json();
  const values = data.values || [];
  if (values.length < 2) { cache = { rows: [], fetchedAt: Date.now() }; return []; }

  const header = values[0];
  const colIndex = (name) => header.findIndex((h) => h && h.includes(name));
  const iClient = colIndex('לקוח'), iLocation = colIndex('מיקום'), iType = colIndex('סוג'),
    iDocDate = colIndex('תאריך מסמך'), iExpiry = colIndex('תוקף'), iArchived = colIndex('ארכיון');

  const rows = values.slice(1)
    .map((r) => ({
      client: clean(r[iClient]),
      location: clean(r[iLocation]),
      docType: clean(r[iType]),
      docDate: clean(r[iDocDate]),
      expiry: clean(r[iExpiry]),
      archived: iArchived >= 0 && clean(r[iArchived]) === 'כן',
    }))
    .filter((r) => r.client && !r.archived);

  cache = { rows, fetchedAt: Date.now() };
  return rows;
}

// Parses the dashboard's own dd.mm.yyyy-ish date formats loosely enough to
// sort by recency; falls back to "oldest" (never first) if unparseable.
function parseDate(s) {
  if (!s) return null;
  const m = String(s).match(/(\d{1,2})[./](\d{1,2})[./](\d{2,4})/);
  if (!m) return null;
  const [, d, mo, y] = m;
  const year = y.length === 2 ? `20${y}` : y;
  const date = new Date(`${year}-${mo.padStart(2, '0')}-${d.padStart(2, '0')}`);
  return isNaN(date.getTime()) ? null : date;
}

/**
 * Finds past (non-archived) report rows whose client name roughly matches
 * the given text, most recent first. Returns at most `limit` rows.
 * Never throws to the caller — a lookup failure should degrade to "nothing
 * found", not block the wizard, since the sheet is a convenience, not a
 * dependency of report generation itself.
 */
export async function findPastReportsByClient(clientText, limit = 5) {
  const needle = clean(clientText).toLowerCase();
  if (!needle) return [];
  try {
    const rows = await fetchAllRows();
    return rows
      .filter((r) => r.client.toLowerCase().includes(needle) || needle.includes(r.client.toLowerCase()))
      .sort((a, b) => (parseDate(b.docDate)?.getTime() || 0) - (parseDate(a.docDate)?.getTime() || 0))
      .slice(0, limit);
  } catch (e) {
    console.error('Dashboard sheet lookup failed:', e.message);
    return [];
  }
}
