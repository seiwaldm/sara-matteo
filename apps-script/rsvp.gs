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
// If you change the columns, rename or delete the old RSVP tab: headers are only written to a new tab.
const HEADERS = [
  'Data', 'Nome', 'Cognome', 'Presenza', 'Adulti', 'Bambini', 'Ospitalità', 'Messaggio', 'Lingua',
];
// Highest number of adults or children accepted in one reply.
const MAX_GUESTS = 10;
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
  const viene = presenza === 'Sì';
  // Guest counts and accommodation are only asked when attending; otherwise counts are 0.
  const adulti = viene ? count(p.adulti) : 0;
  const bambini = viene && p.bambini ? count(p.bambini) : 0; // optional, empty counts as 0
  const ospitalita = !viene ? '' : p.ospitalita === 'si' ? 'Sì' : p.ospitalita === 'no' ? 'No' : '';
  const messaggio = clean(p.messaggio, 1000);
  const lingua = p.lingua === 'de' ? 'DE' : 'IT';
  if (!nome || !cognome || !presenza || adulti === null || bambini === null || (viene && adulti < 1)) {
    return json({ ok: false, error: 'invalid' });
  }

  // The lock only guards the quick rate-limit counter. The sheet write happens after releasing it,
  // so a slow or stuck Sheets call can't block every other reply.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json({ ok: false, error: 'busy' });
  let allowed;
  try {
    allowed = underRateLimit();
  } finally {
    lock.releaseLock();
  }
  if (!allowed) return json({ ok: false, error: 'busy' });
  getSheet().appendRow(
    [new Date(), nome, cognome, presenza, adulti, bambini, ospitalita, messaggio, lingua].map(safe)
  );

  if (NOTIFY_EMAIL) {
    const summary = viene ? `${adulti} adulti, ${bambini} bambini` : 'non partecipa';
    MailApp.sendEmail(
      NOTIFY_EMAIL,
      `RSVP: ${nome} ${cognome} (${summary})`,
      [
        `${nome} ${cognome}`,
        `Presenza: ${presenza}`,
        `Adulti: ${adulti}`,
        `Bambini: ${bambini}`,
        `Ospitalità: ${ospitalita || '-'}`,
        `Messaggio: ${messaggio || '-'}`,
        `Lingua: ${lingua}`,
      ].join('\n')
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

// Whole number from 0 to MAX_GUESTS, or null if the value is anything else.
function count(value) {
  const s = String(value || '').trim();
  if (!/^\d+$/.test(s)) return null;
  const n = Number(s);
  return n <= MAX_GUESTS ? n : null;
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
