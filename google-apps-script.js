/**
 * PIA NA SIA 21 — Lead Capture + Live Counter Google Apps Script
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
 *
 * SHEET STRUCTURE:
 *   "Leads"           — master list of ALL registrations across all LGAs
 *   "LGA - Demsa"     — registrations for Demsa LGA only
 *   "LGA - Fufore"    — registrations for Fufore LGA only
 *   ...               — one sheet per LGA (auto-created on first signup)
 *   "By LGA"          — summary count per LGA
 *   "By Source"       — summary count per traffic source
 */

const NOTIFY_EMAIL = 'womantowomannetwork21@gmail.com';
const SHEET_NAME   = 'Leads';
const LGA_SHEET    = 'By LGA';
const SOURCE_SHEET = 'By Source';

/* Master sheet columns */
const HEADERS = [
  'Timestamp','Name','Phone','Ward','LGA','WhatsApp Group',
  'Traffic Source','Medium','Campaign','Referrer','Full URL','Row #'
];

/* LGA-specific sheet columns — simpler, focused on the member */
const LGA_HEADERS = [
  '#','Timestamp','Name','Phone','Ward','Traffic Source','Campaign'
];

/* Colour palette */
const COLOR_GREEN = '#0A6B2F';
const COLOR_WHITE = '#FFFFFF';

/* ── LGA SHEET NAME HELPER ── */
function lgaSheetName(lgaName) {
  // Truncate to 99 chars to stay within Google Sheets 100-char sheet name limit
  return ('LGA - ' + lgaName).substring(0, 99);
}

/*
 * MANUAL TEST FUNCTION:
 * Select "testConnection" from the dropdown list at the top and click "Run"
 * to verify the spreadsheet connection without errors!
 */
function testConnection() {
  var ss = getActiveSpreadsheet();
  if (ss) {
    Logger.log('🎉 SUCCESS! Connected to Google Sheet: ' + ss.getName());
    Logger.log('Link to your Google Sheet: ' + ss.getUrl());
    var sheet = getOrCreateSheet(ss, SHEET_NAME, HEADERS);
    Logger.log('Master Leads sheet ready. You are ready to receive signups!');
  } else {
    Logger.log('❌ ERROR: Could not locate or create the Google Sheet.');
  }
}

/*
 * Auto-detect spreadsheet:
 * 1. Works automatically if created inside a spreadsheet (bound)
 * 2. If stand-alone, searches your Google Drive for "Pia na Sia - Project Execution"
 * 3. If still not found, automatically creates a new Google Sheet named
 *    "Pia na Sia - Project Execution" in your Drive!
 */
function getActiveSpreadsheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss && ss.getUrl()) return ss;
  } catch (err) {}

  try {
    var files = DriveApp.getFilesByName('Pia na Sia - Project Execution');
    if (files.hasNext()) return SpreadsheetApp.open(files.next());
  } catch (err) {}

  try {
    var filesFallback = DriveApp.searchFiles(
      "mimeType = 'application/vnd.google-apps.spreadsheet' and name contains 'Pia na Sia'"
    );
    if (filesFallback.hasNext()) return SpreadsheetApp.open(filesFallback.next());
  } catch (err) {}

  try {
    return SpreadsheetApp.create('Pia na Sia - Project Execution');
  } catch (err) {}

  return null;
}

/* ── CORS (Google Apps Script handles CORS redirects automatically) ── */
function setCORS(output) {
  return output;
}

// ── POST: save a new registration ──
function doPost(e) {
  try {
    // Support both URLSearchParams (from browser no-cors) and JSON (from direct API calls)
    let data = {};
    if (e.postData && e.postData.type && e.postData.type.indexOf('application/x-www-form-urlencoded') !== -1) {
      // Form-encoded — parse from e.parameter
      data = e.parameter || {};
    } else if (e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); } catch(x) { data = e.parameter || {}; }
    } else {
      data = e.parameter || {};
    }

    const ss = getActiveSpreadsheet();
    if (!ss) throw new Error('Could not find or create the Google Spreadsheet.');

    // 1. Master Leads sheet
    const leadsSheet = getOrCreateSheet(ss, SHEET_NAME, HEADERS);
    const rowNum     = leadsSheet.getLastRow();
    leadsSheet.appendRow([
      data.timestamp      || new Date().toISOString(),
      data.name           || '',
      data.phone          || '',
      data.ward           || '',
      data.lga            || '',
      data.whatsapp_group || '',
      data.source         || 'direct',
      data.medium         || '',
      data.campaign       || '',
      data.referrer       || '',
      data.full_url       || '',
      rowNum
    ]);

    // 2. Dedicated LGA sheet (auto-created)
    const lgaName = data.lga || 'Unknown';
    writeToLgaSheet(ss, lgaName, data);

    // 3. Summary counts
    updateSummary(ss, LGA_SHEET,    lgaName);
    updateSummary(ss, SOURCE_SHEET, data.source || 'direct');

    // 4. Email alert
    sendAlert(data, rowNum);

    return setCORS(
      ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', row: rowNum, lga_sheet: lgaSheetName(lgaName) }))
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

/* ── WRITE ONE ROW TO THE DEDICATED LGA SHEET ── */
function writeToLgaSheet(ss, lgaName, data) {
  const sheetName = lgaSheetName(lgaName);
  const sheet     = getOrCreateSheet(ss, sheetName, LGA_HEADERS, lgaName);

  // Registration # within this LGA = rows already written (minus header)
  const lgaRegNum = sheet.getLastRow(); // before append, so header row = 1 → first member = #1

  sheet.appendRow([
    lgaRegNum,                              // # (sequential within LGA)
    data.timestamp || new Date().toISOString(),
    data.name      || '',
    data.phone     || '',
    data.ward      || '',
    data.source    || 'direct',
    data.campaign  || ''
  ]);
}

/* ── GET: return live summary for the counter on the landing page ── */
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) ? e.parameter.action : '';
    const ss     = getActiveSpreadsheet();
    if (!ss) throw new Error('Could not locate your Google Spreadsheet.');

    const leads = ss.getSheetByName(SHEET_NAME);
    const total = leads ? Math.max(0, leads.getLastRow() - 1) : 0;

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

Sheet tabs updated:
  ✅ Leads (master)
  ✅ ${lgaSheetName(data.lga || 'Unknown')}

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
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 3)
         .setBackground(COLOR_GREEN)
         .setFontColor(COLOR_WHITE)
         .setFontWeight('bold');
  }
  const vals  = sheet.getDataRange().getValues();
  let found   = false;
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(vals[i][1] + 1);
      sheet.getRange(i + 1, 3).setValue(new Date().toISOString());
      found = true;
      break;
    }
  }
  if (!found) sheet.appendRow([key, 1, new Date().toISOString()]);
}

/* ── CREATE SHEET IF MISSING ── */
/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} name  - Sheet tab name
 * @param {string[]} headers
 * @param {string} [lgaLabel] - When set, adds an LGA banner row above the header
 */
function getOrCreateSheet(ss, name, headers, lgaLabel) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);

    if (lgaLabel) {
      // Row 1: LGA title banner
      sheet.appendRow([lgaLabel + ' LGA — Registrations']);
      sheet.getRange(1, 1, 1, headers.length)
           .merge()
           .setBackground(COLOR_GREEN)
           .setFontColor(COLOR_WHITE)
           .setFontWeight('bold')
           .setFontSize(13)
           .setHorizontalAlignment('center');

      // Row 2: column headers
      sheet.appendRow(headers);
      sheet.getRange(2, 1, 1, headers.length)
           .setBackground('#1D7A3A')
           .setFontColor(COLOR_WHITE)
           .setFontWeight('bold');
      sheet.setFrozenRows(2);
    } else {
      // Standard header (row 1 only)
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setBackground(COLOR_GREEN)
           .setFontColor(COLOR_WHITE)
           .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // Auto-resize all columns for readability
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}
