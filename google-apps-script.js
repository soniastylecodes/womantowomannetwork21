/**
 * PIA NA SIA 21 — Lead Capture + Live Counter Google Apps Script (Stand-Alone Friendly)
 * ==============================================================
 * HOW TO DEPLOY:
 * 1. Open your Google Sheet, and copy its browser URL.
 * 2. Paste it in the SPREADSHEET_URL variable below.
 * 3. Go to script.google.com — create new project, paste this whole file.
 * 4. Deploy > New Deployment > Web App.
 * 5. Execute as: Me  |  Who has access: Anyone.
 * 6. Copy the Web App URL.
 * 7. In index.html replace SHEET_URL with that URL.
 */

// 1. paste your Google Sheet link between the quotes below:
const SPREADSHEET_URL = 'YOUR_SPREADSHEET_URL_HERE'; 

const NOTIFY_EMAIL = 'womantowomannetwork21@gmail.com';
const SHEET_NAME   = 'Leads';
const LGA_SHEET    = 'By LGA';
const SOURCE_SHEET = 'By Source';

const HEADERS = [
  'Timestamp','Name','Phone','LGA','WhatsApp Group',
  'Traffic Source','Medium','Campaign','Referrer','Full URL','Row #'
];

/* Helper to get the correct spreadsheet whether bound or standalone */
function getActiveSpreadsheet() {
  if (SPREADSHEET_URL && SPREADSHEET_URL !== 'YOUR_SPREADSHEET_URL_HERE') {
    return SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
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
    if (!ss) throw new Error("Could not access Google Spreadsheet. Please verify your SPREADSHEET_URL in the script.");

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
    if (!ss) throw new Error("Could not access Google Spreadsheet.");

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
