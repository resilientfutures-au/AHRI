// ACO Strategic Direction Surveys — Google Apps Script (combined handler)
// Handles both internal and external survey submissions into separate tabs
// of the same Google Spreadsheet.
//
// Deploy as: Web app → Execute as: Me → Who has access: Anyone
// Use this single deployment URL in both survey index.html files.

// ── SHEET CONFIGURATION ───────────────────────────────────────────────────────

var CONFIG = {
  internal: {
    sheetName: 'Internal Responses',
    keyOrder: [
      'q1', 'q2', 'q3',
      'q4', 'q4_comment',
      'q5', 'q6',
      'q7', 'q8', 'q9', 'q10',
      'q11', 'q12', 'q13', 'q14', 'q15', 'q16', 'q17',
      'q18', 'q19',
      'q20', 'q20_comment',
      'q21', 'q22',
      'q23', 'q24',
    ],
    headers: [
      'Timestamp',
      'Q01 — Role',
      'Q02 — Area of work',
      'Q03 — Years associated with ACO',
      'Q04 — Rating: current strategic direction',
      'Q04 — Comments on rating',
      'Q05 — What is working particularly well',
      'Q06 — Where could approach be stronger',
      'Q07 — Primary value delivered today',
      'Q08 — Where value could be stronger',
      'Q09 — Value expected in 5–10 years',
      'Q10 — Known for in five years',
      'Q11 — Future of Work',
      'Q12 — Technology & Innovation',
      'Q13 — Governance & Regulation',
      'Q14 — Healthy Communities',
      'Q15 — Economic & Funding Environment',
      'Q16 — Social Expectations & Community Needs',
      'Q17 — Emerging Conditions',
      'Q18 — Greatest strategic opportunities (3–5 years)',
      'Q19 — Most significant strategic risks (3–5 years)',
      'Q20 — Rating: organisational readiness',
      'Q20 — Comments on rating',
      'Q21 — Capabilities to develop or strengthen',
      'Q22 — What gets in the way',
      'Q23 — Single most important strategic question',
      'Q24 — Anything else to share',
    ],
  },
  external: {
    sheetName: 'External Responses',
    keyOrder: [
      'q1', 'q2',
      'q3', 'q4', 'q5', 'q6',
      'q7', 'q8', 'q9', 'q10', 'q11', 'q12',
      'q13', 'q14', 'q15',
      'q16',
    ],
    headers: [
      'Timestamp',
      'Q01 — Relationship with ACO',
      'Q02 — Connection length',
      'Q03 — Primary value ACO provides today',
      'Q04 — Where contribution could be stronger',
      'Q05 — What to deliver in 5–10 years',
      'Q06 — Biggest gap between current and needed',
      'Q07 — Future of Work',
      'Q08 — Technology & Innovation',
      'Q09 — Governance & Policy',
      'Q10 — Community Health Needs',
      'Q11 — Geopolitical & Broader System Conditions',
      'Q12 — Emerging Conditions',
      'Q13 — Greatest strategic opportunities (3–5 years)',
      'Q14 — Strategic risks to manage',
      'Q15 — One thing done exceptionally well',
      'Q16 — Anything else to share',
    ],
  },
};

// ── HANDLERS ──────────────────────────────────────────────────────────────────

// Run this once manually after pasting the script to pre-create both sheets.
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(CONFIG).forEach(function(type) {
    var cfg = CONFIG[type];
    var sheet = ss.getSheetByName(cfg.sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(cfg.sheetName);
      sheet.appendRow(cfg.headers);
      formatHeaderRow(sheet, cfg.headers.length);
      Logger.log('Created sheet: ' + cfg.sheetName);
    } else {
      Logger.log('Sheet already exists: ' + cfg.sheetName);
    }
  });
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var surveyType = payload.surveyType;

    if (!CONFIG[surveyType]) {
      return jsonResponse({ status: 'ignored', reason: 'unknown surveyType: ' + surveyType });
    }

    var cfg = CONFIG[surveyType];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(cfg.sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(cfg.sheetName);
      sheet.appendRow(cfg.headers);
      formatHeaderRow(sheet, cfg.headers.length);
    }

    var answers = payload.answers || {};
    var timestamp = payload.submittedAt
      ? new Date(payload.submittedAt).toLocaleString('en-AU', { timeZone: 'Australia/Melbourne' })
      : new Date().toLocaleString('en-AU', { timeZone: 'Australia/Melbourne' });

    var row = [timestamp].concat(cfg.keyOrder.map(function(key) {
      return answers[key] || '';
    }));

    sheet.appendRow(row);

    return jsonResponse({ status: 'ok' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function doGet(e) {
  return jsonResponse({ status: 'ok', message: 'ACO Survey endpoint active' });
}

// ── HELPERS ───────────────────────────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatHeaderRow(sheet, colCount) {
  var range = sheet.getRange(1, 1, 1, colCount);
  range.setFontWeight('bold');
  range.setBackground('#1a1a1a');
  range.setFontColor('#ffffff');
  range.setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidths(2, 3, 220);
  sheet.setColumnWidths(5, colCount - 4, 300);
}
