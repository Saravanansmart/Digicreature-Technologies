/**
 * Digicreature Technologies — lead intake
 *
 * Receives leads from BOTH the contact form and the "Digi" chatbot on
 * index.html and appends one row per lead to the Leads tab.
 *
 * Sheet: https://docs.google.com/spreadsheets/d/1Xrdgd0gTjNPJYHo_K1h3WN766dGqaxXopXUSE3lz1Go/edit
 *
 * DEPLOY
 *   1. Extensions → Apps Script in that spreadsheet, paste this file.
 *   2. Deploy → New deployment → type "Web app".
 *        Execute as:      Me
 *        Who has access:  Anyone
 *   3. Copy the /exec URL and set LEAD_ENDPOINT in index.html to it.
 *   4. Run testAppend() once to create headers and approve the auth scopes.
 *
 * The site posts with mode:'no-cors', so the browser cannot read this
 * response — it is returned for manual testing and future CORS use.
 */

var SHEET_ID  = '1Xrdgd0gTjNPJYHo_K1h3WN766dGqaxXopXUSE3lz1Go';
var TAB_NAME  = 'Leads';

// Column order — must stay in sync with buildLead() in index.html.
var HEADERS = [
  'Timestamp',   // A — ISO 8601 from the browser
  'Source',      // B — "Contact form" | "Chatbot (Digi)"
  'Name',        // C
  'Email',       // D
  'Phone',       // E
  'Company',     // F — blank for chatbot leads
  'Message',     // G
  'Page',        // H — path + query the lead came from
  'Referrer'     // I — document.referrer, or "direct"
];

function doPost(e) {
  try {
    var payload = parseBody(e);
    if (!payload) return json({ ok: false, error: 'Empty or unparseable body' });

    // Minimum viable lead: we need a way to reply.
    if (!payload.email && !payload.phone) {
      return json({ ok: false, error: 'Lead needs at least an email or a phone' });
    }

    var sheet = getSheet();
    sheet.appendRow([
      payload.timestamp || new Date().toISOString(),
      payload.source    || 'Unknown',
      payload.name      || '',
      payload.email     || '',
      payload.phone     || '',
      payload.company   || '',
      payload.message   || '',
      payload.page      || '',
      payload.referrer  || ''
    ]);

    return json({ ok: true, row: sheet.getLastRow() });
  } catch (err) {
    // Keep a trace in the Apps Script execution log — the browser can't see this.
    console.error('doPost failed: ' + err);
    return json({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json({ ok: true, service: 'Digicreature lead intake', tab: TAB_NAME });
}

/**
 * The site posts with mode:'no-cors', which forces the request body to be sent
 * as text/plain — so read e.postData.contents directly rather than trusting the
 * declared type. Falls back to form-encoded params if posted that way.
 */
function parseBody(e) {
  if (!e) return null;
  if (e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (err) {
      // not JSON — fall through to parameters
    }
  }
  if (e.parameter && Object.keys(e.parameter).length) return e.parameter;
  return null;
}

function getSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) sheet = ss.insertSheet(TAB_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
         .setFontWeight('bold')
         .setBackground('#FFE83D');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 180); // Timestamp
    sheet.setColumnWidth(7, 380); // Message
  }
  return sheet;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the editor to create headers and grant permissions. */
function testAppend() {
  var res = doPost({
    postData: {
      contents: JSON.stringify({
        timestamp: new Date().toISOString(),
        source: 'Contact form',
        name: 'Test Lead',
        email: 'test@example.com',
        phone: '+91 90000 00000',
        company: 'Test Co.',
        message: 'Verifying the sheet mapping.',
        page: '/',
        referrer: 'direct'
      })
    }
  });
  Logger.log(res.getContent());
}
