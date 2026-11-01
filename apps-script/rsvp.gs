/**
 * RSVP receiver for the wedding site. Each form submission becomes a row in the "RSVP" tab.
 *
 * Setup (once):
 *  1. Create a Google Sheet, then open Extensions > Apps Script.
 *  2. Replace the editor contents with this file and save.
 *  3. Deploy > New deployment > type "Web app".
 *       Execute as: Me
 *       Who has access: Anyone
 *     Approve the permission prompt (Google warns the app is unverified: Advanced > Go to ...).
 *  4. Copy the web app URL (ends in /exec) into RSVP_URL in index.html.
 *
 * After changing this script: Deploy > Manage deployments > edit (pencil) > Version: New version.
 * That keeps the same URL. A "New deployment" would create a new URL instead.
 */

const SHEET_NAME = 'RSVP';
const HEADERS = ['Data', 'Nome', 'Cognome', 'Presenza', 'Intolleranze', 'Lingua'];
// Optional: an address that gets an email for every RSVP. Leave empty to turn off.
const NOTIFY_EMAIL = '';
// Replies sent sooner than this after the page loaded are treated as bots.
const MIN_FILL_MS = 3000;
// At most RATE_LIMIT replies per RATE_WINDOW_S seconds, across all visitors.
const RATE_LIMIT = 20;
const RATE_WINDOW_S = 600;

function doPost(e) {
  const p = (e && e.parameter) || {};

  // Honeypot: real visitors never fill this hidden field, bots often do. Pretend success.
  if (p.website) return json({ ok: true });

  // Too fast: the page sends how long it was open (ms). Bots submit instantly. Pretend success.
  if (!(Number(p.t) >= MIN_FILL_MS)) return json({ ok: true });

  const nome = clean(p.nome, 100);
  const cognome = clean(p.cognome, 100);
  const presenza = p.presenza === 'si' ? 'Sì' : p.presenza === 'no' ? 'No' : '';
  const intolleranze = clean(p.intolleranze, 1000);
  const lingua = p.lingua === 'de' ? 'DE' : 'IT';
  if (!nome || !cognome || !presenza) return json({ ok: false, error: 'invalid' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!underRateLimit()) return json({ ok: false, error: 'busy' });
    getSheet().appendRow([new Date(), nome, cognome, presenza, intolleranze, lingua].map(safe));
  } finally {
    lock.releaseLock();
  }

  if (NOTIFY_EMAIL) {
    MailApp.sendEmail(
      NOTIFY_EMAIL,
      `RSVP: ${nome} ${cognome} (${presenza})`,
      `${nome} ${cognome}\nPresenza: ${presenza}\nIntolleranze: ${intolleranze || '-'}\nLingua: ${lingua}`
    );
  }
  return json({ ok: true });
}

// Opening the /exec URL in a browser shows this, handy for checking the deployment works.
function doGet() {
  return json({ ok: true, status: 'RSVP endpoint is running' });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

// Counts replies in fixed time windows. Call only while holding the script lock.
function underRateLimit() {
  const cache = CacheService.getScriptCache();
  const key = 'rsvp-count-' + Math.floor(Date.now() / (RATE_WINDOW_S * 1000));
  const count = Number(cache.get(key) || 0);
  if (count >= RATE_LIMIT) return false;
  cache.put(key, String(count + 1), RATE_WINDOW_S);
  return true;
}

function clean(value, max) {
  return String(value || '').trim().slice(0, max);
}

// Stop text like "=IMPORTXML(...)" from being run as a spreadsheet formula.
function safe(value) {
  return typeof value === 'string' && /^[=+\-@]/.test(value) ? "'" + value : value;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
