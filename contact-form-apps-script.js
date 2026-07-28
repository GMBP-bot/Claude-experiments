/**
 * CONTACT FORM → GOOGLE SHEET
 * Paste this whole file into Apps Script (Extensions → Apps Script
 * from your sheet) and deploy as a Web App. Full steps live in the
 * conversation / README; the short version:
 *
 *   Deploy → New deployment → Web app
 *     Execute as:      Me
 *     Who has access:  Anyone
 *
 * Copy the /exec URL it gives you into CF_SHEET_ENDPOINT in
 * contact.html.
 *
 * Redeploying after edits: Deploy → Manage deployments → pencil →
 * Version: New version. Creating a brand-new deployment instead
 * changes the URL and silently orphans the form.
 */

// Name of the tab rows are written to. Created on first submission
// if it doesn't exist, headers included.
var SHEET_NAME = 'Leads';

// Set to an email address (e.g. 'info@saeragroup.com') to also get
// each submission by email. Leave '' to disable.
var NOTIFY_EMAIL = '';

function doPost(e) {
  try {
    var lock = LockService.getScriptLock();
    // Two visitors submitting in the same second would otherwise
    // race appendRow and interleave columns.
    lock.waitLock(10000);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Timestamp', 'First name', 'Last name', 'Email', 'Subject', 'Message']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    var p = (e && e.parameter) || {};

    // Everything is coerced to a trimmed string. A cell can hold
    // anything, but a formula-looking payload ("=IMPORTRANGE...")
    // must not execute — the leading apostrophe neutralises it.
    function clean(v) {
      v = (v == null ? '' : String(v)).trim();
      if (/^[=+\-@]/.test(v)) v = "'" + v;
      return v;
    }

    var row = [
      new Date(),
      clean(p.first),
      clean(p.last),
      clean(p.email),
      clean(p.subject),
      clean(p.message)
    ];

    sheet.appendRow(row);

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
    // it exists for manual testing with curl or the editor.
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Visiting the /exec URL in a browser hits doGet — answering it
// makes "is my deployment alive?" a one-click check.
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: 'contact-form' }))
    .setMimeType(ContentService.MimeType.JSON);
}
