/**
 * PIA NA SIA 21 — Lead Capture + Live Counter Google Apps Script (Zero-Configuration Edition)
 * ==============================================================
 * HOW TO DEPLOY:
 * 1. Go to script.google.com — create a new project.
 * 2. Delete any existing code and paste this entire script.
 * 3. Click the Save icon (floppy disk).
 * 4. Deploy > New Deployment > Web App.
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Click Deploy. Authorize the permissions when prompted.
 * 6. Copy the Web App URL (ends with /exec).
 */

const NOTIFY_EMAIL = 'womantowomannetwork21@gmail.com';
const SHEET_NAME   = 'Leads';
const LGA_SHEET    = 'By LGA';
const SOURCE_SHEET = 'By Source';

const HEADERS = [
  'Timestamp','Name','Phone','LGA','WhatsApp Group',
  'Traffic Source','Medium','Campaign','Referrer','Full URL','Row #'
];

/* 
 * Auto-detect spreadsheet:
 * 1. Works automatically if created inside a spreadsheet (bound)
 * 2. If stand-alone, searches your Google Drive for "Pia na Sia - Project Execution"
 * 3. If still not found, automatically creates a new Google Sheet named "Pia na Sia - Project Execution" in your Drive!
 */
function getActiveSpreadsheet() {
  try {
    // 1. Try to get active spreadsheet (if bound)
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss && ss.getUrl()) return ss;
  } catch (err) {}
  
  try {
    // 2. Search for the sheet by name in Google Drive
    var files = DriveApp.getFilesByName("Pia na Sia - Project Execution");
    if (files.hasNext()) {
      return SpreadsheetApp.open(files.next());
    }
  } catch (err) {}
  
  try {
    // 3. Fallback search for any sheet containing "Pia na Sia"
    var filesFallback = DriveApp.searchFiles("mimeType = 'application/vnd.google-apps.spreadsheet' and name contains 'Pia na Sia'");
    if (filesFallback.hasNext()) {
      return SpreadsheetApp.open(filesFallback.next());
    }
  } catch (err) {}
  
  try {
    // 4. If no sheet exists anywhere, create a brand new one!
    var newSS = SpreadsheetApp.create("Pia na Sia - Project Execution");
    return newSS;
  } catch (err) {}
  
  return null;
}

/* ── CORS HEADERS (required for live counter fetch from browser) ── */
function setCORS(output) {
  return output
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

/* ── POST: receive a new lead from the landing page ── */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = getActiveSpreadsheet();
    if (!ss) throw new Error("Could not locate or create your Google Spreadsheet named 'Pia na Sia - Project Execution'.");

    const leadsSheet = getOrCreateSheet(ss, SHEET_NAME, HEADERS);
    const rowNum     = leadsSheet.getLastRow();

    leadsSheet.appendRow([
      data.timestamp      || new Date().toISOString(),
      data.name           || '',
      data.phone          || '',
      data.lga            || '',
      data.whatsapp_group || '',
      data.source         || 'direct',
      data.medium         || '',
      data.campaign       || '',
      data.referrer       || '',
      data.full_url       || '',
      rowNum
    ]);

    updateSummary(ss, LGA_SHEET,    data.lga    || 'Unknown');
    updateSummary(ss, SOURCE_SHEET, data.source || 'direct');
    sendAlert(data, rowNum);

    return setCORS(
      ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', row: rowNum }))
        .setMimeType(ContentService.MimeType.JSON)
    );
  } catch (err) {
    return setCORS(
      ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON)
    );
  }
}

/* ── GET: return live summary for the counter on the landing page ── */
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) ? e.parameter.action : '';
    const ss     = getActiveSpreadsheet();
    if (!ss) throw new Error("Could not locate your Google Spreadsheet.");

    const leads  = ss.getSheetByName(SHEET_NAME);
    const total  = leads ? Math.max(0, leads.getLastRow() - 1) : 0;

    if (action === 'summary') {
      const byLGA  = ss.getSheetByName(LGA_SHEET);
      const bySrc  = ss.getSheetByName(SOURCE_SHEET);
      const lgaData = byLGA ? byLGA.getDataRange().getValues().slice(1) : [];
      const srcData = bySrc ? bySrc.getDataRange().getValues().slice(1) : [];

      return setCORS(
        ContentService.createTextOutput(JSON.stringify({
          total_leads: total,
          by_lga:    lgaData.map(r => ({ lga: r[0], count: r[1] })),
          by_source: srcData.map(r => ({ source: r[0], count: r[1] }))
        })).setMimeType(ContentService.MimeType.JSON)
      );
    }

    // default: just the count (lightweight, fast)
    return setCORS(
      ContentService
        .createTextOutput(JSON.stringify({ total_leads: total }))
        .setMimeType(ContentService.MimeType.JSON)
    );
  } catch (err) {
    return setCORS(
      ContentService
        .createTextOutput(JSON.stringify({ total_leads: 0, error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON)
    );
  }
}

/* ── EMAIL ALERT ── */
function sendAlert(data, rowNum) {
  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: `New Sister Joined: ${data.lga} LGA — Pia Na Sia 21`,
    body: `
A new woman has joined the network.

NAME       : ${data.name}
PHONE      : ${data.phone || 'Not provided'}
LGA        : ${data.lga}
SOURCE     : ${data.source || 'direct'}
CAMPAIGN   : ${data.campaign || 'none'}
TIME       : ${data.timestamp}
ROW #      : ${rowNum}

WhatsApp Group : ${data.whatsapp_group}

Mun kasance tare. Mu ci gaba tare.
Pia Na Sia 21 Digital Team
    `.trim()
  });
}

/* ── SUMMARY SHEET UPSERT ── */
function updateSummary(ss, sheetName, key) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['Category', 'Count', 'Last Updated']);
  }
  const data  = sheet.getDataRange().getValues();
  let found   = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(data[i][1] + 1);
      sheet.getRange(i + 1, 3).setValue(new Date().toISOString());
      found = true;
      break;
    }
  }
  if (!found) sheet.appendRow([key, 1, new Date().toISOString()]);
}

/* ── CREATE SHEET IF MISSING ── */
function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
         .setBackground('#0A6B2F')
         .setFontColor('#ffffff')
         .setFontWeight('bold');
  }
  return sheet;
}
