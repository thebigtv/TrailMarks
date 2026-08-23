/**
 * TRAILMARKS — APPS SCRIPT BACKEND (Phase 2: read-only)
 * ============================================================
 * SETUP
 * 1. Create a Google Sheet with tabs named exactly: WALKS, SECTIONS
 *    (see /sheet-templates/*.csv for the column headers + starter
 *    data — import each CSV into a tab of the same name).
 * 2. Extensions → Apps Script, paste this file in as Code.gs.
 * 3. Deploy → New deployment → Web app.
 *      Execute as: Me
 *      Who has access: Anyone
 * 4. Copy the deployment URL into API_BASE in index.html.
 *
 * This deliberately only exposes GET endpoints for now — no writes,
 * no accounts, no orders. Those are Phases 3–6 per the architecture
 * doc and get their own Apps Script files when we build them, so
 * this one file doesn't grow into an unreviewable mess.
 * ============================================================
 */

function doGet(e) {
  const resource = (e.parameter.resource || '').toLowerCase();

  try {
    if (resource === 'sections') {
      return jsonResponse(getSections(e.parameter.walk_id || 'HW'));
    }
    if (resource === 'walks') {
      return jsonResponse(getWalks());
    }
    return jsonResponse({ error: 'Unknown resource. Try ?resource=sections or ?resource=walks' }, 400);
  } catch (err) {
    return jsonResponse({ error: String(err) }, 500);
  }
}

/* ---------- data access ---------- */

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`Sheet tab "${name}" not found`);
  return sheet;
}

/* Generic: read a sheet into an array of objects keyed by header row */
function getRows(sheetName) {
  const sheet = getSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1)
    .filter(row => row.some(cell => cell !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function getWalks() {
  return getRows('WALKS');
}

function getSections(walkId) {
  return getRows('SECTIONS')
    .filter(row => row.walk_id === walkId)
    .filter(row => row.active === true || row.active === 'TRUE' || row.active === 1)
    .sort((a, b) => Number(a.mile) - Number(b.mile));
}

/* ---------- response helper ---------- */

function jsonResponse(data, status) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
