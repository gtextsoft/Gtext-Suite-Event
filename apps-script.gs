/**
 * High Ticket Sales Mastery — registration backend
 * Paste this into Extensions → Apps Script on the Google Sheet, then Deploy → New deployment → Web app.
 *   Execute as: Me
 *   Who has access: Anyone
 * Copy the Web app URL (ends in /exec) into SCRIPT_URL in index.html.
 */
const SHEET_ID   = '1xoJlRV_2ZeaK2HyMRNcAkAqA6yrMhmVnqDqmsWgaEng';
const SHEET_NAME = 'Registrations';
const HEADERS = ['Timestamp', 'Full name', 'Phone', 'Email', 'Company / business', 'Role',
                 'City / country', 'How they heard', 'What they want to learn'];

function sheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#d71920').setFontColor('#ffffff');
  }
  return sh;
}

// Stops values like "=SUM(...)" being read as formulas
function clean_(v, max) {
  v = String(v == null ? '' : v).trim().slice(0, max || 200);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_({ ok: true });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    if (d.website) return json_({ ok: true }); // spam trap

    const name = String(d.name || '').trim();
    const phone = String(d.phone || '').trim();
    const email = String(d.email || '').trim().toLowerCase();
    if (name.length < 2) return json_({ ok: false, error: 'Enter your full name.' });
    if (phone.replace(/\D/g, '').length < 7) return json_({ ok: false, error: 'Enter a valid phone number.' });
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return json_({ ok: false, error: 'Enter a valid email address.' });

    const sh = sheet_();
    const n = sh.getLastRow() - 1;
    if (n > 0) {
      const emails = sh.getRange(2, 4, n, 1).getValues();
      if (emails.some(r => String(r[0]).toLowerCase() === email)) return json_({ ok: true, duplicate: true });
    }

    sh.appendRow([
      new Date(), clean_(name, 120), "'" + phone.slice(0, 30), clean_(email, 160),
      clean_(d.company), clean_(d.role), clean_(d.location), clean_(d.source), clean_(d.goal, 1000)
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'Something went wrong on our side. Please try again.' });
  } finally {
    lock.releaseLock();
  }
}
