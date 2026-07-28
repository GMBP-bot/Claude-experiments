/**
 * CONTACT FORM → GOOGLE SHEET
 *
 * Setup:
 *   1. Fill in SHEET_ID below (see its comment).
 *   2. Run the testAppend function once from the editor toolbar
 *      (Run ▸ testAppend). Approve permissions when asked. A test
 *      row should appear in the "Leads" tab of your sheet. If this
 *      fails, the deployment can never work — fix this first.
 *   3. Deploy → Manage deployments → pencil → Version: NEW VERSION
 *      → Deploy. (Editing code does NOT update a live deployment —
 *      a deployment is frozen at the version it was created with.
 *      This is the second most common reason rows don't appear.)
 *
 * Deployment settings (Web app):
 *      Execute as:      Me
 *      Who has access:  Anyone   ← anything else silently rejects
 *                                  the form's anonymous posts
 */

// ▼▼ REQUIRED ▼▼
// The long ID from your sheet's URL:
//   https://docs.google.com/spreadsheets/d/  THIS_PART  /edit
// Opening by ID works whether this script lives inside the sheet
// or was created standalone at script.google.com — the
// getActiveSpreadsheet() approach only works for the former, and
// failing that is the most common reason submissions vanish.
var SHEET_ID = '1RX4bOsCPT3iCd0Ig-duTpAUneKljzfUzbyxXLE2XPFs';

// Tab rows are written to. Created (with headers) if missing.
var SHEET_NAME = 'Leads';

// Set to an email address (e.g. 'info@saeragroup.com') to also get
// each submission by email. Leave '' to disable.
var NOTIFY_EMAIL = '';

function getBook() {
  if (SHEET_ID && SHEET_ID !== 'PASTE_YOUR_SHEET_ID_HERE') {
    return SpreadsheetApp.openById(SHEET_ID);
  }
  // Fallback for container-bound scripts with no ID filled in.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error(
      'No spreadsheet: fill in SHEET_ID at the top of this script. ' +
      'getActiveSpreadsheet() only works when the script was created ' +
      'from inside the sheet via Extensions → Apps Script.'
    );
  }
  return ss;
}

function getSheet() {
  var ss = getBook();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'First name', 'Last name', 'Email', 'Subject', 'Message']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Everything is coerced to a trimmed string. A cell can hold
// anything, but a formula-looking payload ("=IMPORTRANGE...")
// must not execute — the leading apostrophe neutralises it.
function clean(v) {
  v = (v == null ? '' : String(v)).trim();
  if (/^[=+\-@]/.test(v)) v = "'" + v;
  return v;
}

function doPost(e) {
  try {
    var lock = LockService.getScriptLock();
    // Two visitors submitting in the same second would otherwise
    // race appendRow and interleave columns.
    lock.waitLock(10000);

    var sheet = getSheet();
    var p = (e && e.parameter) || {};

    sheet.appendRow([
      new Date(),
      clean(p.first),
      clean(p.last),
      clean(p.email),
      clean(p.subject),
      clean(p.message)
    ]);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: 'New website enquiry' + (p.subject ? ': ' + clean(p.subject) : ''),
        body:
          'A new enquiry arrived via the website contact form.\n\n' +
          'Name:    ' + clean(p.first) + ' ' + clean(p.last) + '\n' +
          'Email:   ' + clean(p.email) + '\n' +
          'Subject: ' + clean(p.subject) + '\n\n' +
          clean(p.message)
      });
    }

    lock.releaseLock();

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    // The form posts in no-cors mode and can't read this reply;
    // it exists for manual testing. Errors also land in the
    // editor under Executions (left sidebar, clock icon) — check
    // there when rows are not appearing.
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Visiting the /exec URL in a browser hits doGet. It now performs
// a real write-permission check, so one visit tells you whether
// the deployment can reach the sheet at all.
function doGet() {
  try {
    var ss = getBook();
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        service: 'contact-form',
        sheet: ss.getName()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Run this once from the editor (Run ▸ testAppend) after filling
// in SHEET_ID. It writes a visible test row and triggers the
// permission prompt. If this works but the form still doesn't,
// the problem is the deployment (version or access), not the code.
function testAppend() {
  getSheet().appendRow([new Date(), 'Test', 'Row', 'test@example.com', 'testAppend()', 'If you can read this, the script can write to the sheet.']);
  Logger.log('Row written to "' + SHEET_NAME + '" in "' + getBook().getName() + '".');
}
