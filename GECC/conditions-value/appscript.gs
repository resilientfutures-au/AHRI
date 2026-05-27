// ════════════════════════════════════════════════════════════════════════════
// Glen Eira City Council — Conditions & Value Capture Survey
// Google Apps Script (v1)
// ════════════════════════════════════════════════════════════════════════════
//
// SETUP:
//   1. Create a new Google Sheet (e.g. "GECC Conditions & Value Responses").
//   2. Extensions > Apps Script — paste this file.
//   3. Run setupSheets() once to create/refresh all view sheets.
//      (You will be prompted to authorise; accept the permissions.)
//   4. Deploy as Web App:
//        - Execute as: "Me"
//        - Who has access: "Anyone"
//      Copy the resulting deployment URL.
//   5. Paste the URL into APPS_SCRIPT_URL in
//      GECC/conditions-value/index.html (replacing REPLACE_WITH_GECC_APPS_SCRIPT_URL).
//
// SHEETS CREATED:
//   • Responses   Raw response data — one row per respondent
//   • Dashboard   KPIs: respondent count, role/tenure breakdown, completion
//   • By Lens     The 12 strategic lenses (Section 4) + Emerging (Section 5)
//                 with all responses stacked per lens for thematic review
//   • By Theme    Value Today, Future Value, Strategic Opp/Risk, Final Thoughts
//
// MAINTENANCE:
//   • Run setupSheets() any time you change the QUESTIONS/LENSES constants.
//   • Run refreshAggregates() any time you want the By Lens / By Theme sheets
//     re-built from the current Responses data.
//
// ════════════════════════════════════════════════════════════════════════════

// ── CONFIGURATION ──────────────────────────────────────────────────────────

var RESPONSES_SHEET = 'Responses';
var DASHBOARD_SHEET = 'Dashboard';
var BY_LENS_SHEET   = 'By Lens';
var BY_THEME_SHEET  = 'By Theme';
var TIMEZONE        = 'Australia/Melbourne';

// Brand colours — Resilient Futures palette (matches AHRI / ACO sheets)
var C_GREEN       = '#a44f43';   // (kept as var name C_GREEN for code reach; value is RF terracotta)
var C_GREEN_LT    = '#f5e0dd';
var C_INK         = '#1a1a1a';
var C_INK_HEADER  = '#2d2d2d';
var C_GREY_ROW    = '#fafafa';
var C_GREEN_HI    = '#d4edda';
var C_AMBER_MID   = '#fff3cd';
var C_RED_LO      = '#f8d7da';

// Role + tenure value labels (mirrors the survey HTML data-* attributes)
var ROLE_LABELS = {
  'senior_exec':            'Senior Executive',
  'senior_leader':          'Senior Leader',
  'service_stream_manager': 'Service Stream Manager',
  'other':                  'Other'
};

var TENURE_LABELS = {
  'lt1': 'Less than 1 year',
  '1-3': '1–3 years',
  '3-7': '3–7 years',
  '7+':  '7+ years'
};

// Every textarea field on the survey, in the order they appear.
// Each item: { id, section, label, prompt }
var QUESTIONS = [
  // Section 2 — Value Delivered Today
  { id: 's2q1', section: 'Value Today',          label: 'Primary forms of value delivered today & to whom',
    prompt: 'What are the primary forms of value that Glen Eira delivers today and to whom?' },
  { id: 's2q2', section: 'Value Today',          label: 'Where value could be stronger / more relevant',
    prompt: 'Where could the value Council delivers be stronger or more relevant?' },

  // Section 3 — Future Value Expectations
  { id: 's3q1', section: 'Future Value',         label: '5–10 yr expectations of these stakeholders',
    prompt: 'Looking ahead 5–10 years, what forms of value will these same stakeholders expect from Council?' },
  { id: 's3q2', section: 'Future Value',         label: 'Where expectations of value may already be changing',
    prompt: 'Where do you believe that expectations of value may already be changing or evolving?' },

  // Section 4 — Immediate Conditions (the 12 lenses)
  { id: 's4q1',  section: 'Immediate Lens',      label: 'Lens 1 · Planet Reliability',                            prompt: 'Planet Reliability — conditions on the strategic radar' },
  { id: 's4q2',  section: 'Immediate Lens',      label: 'Lens 2 · Future of Work',                                prompt: 'Future of Work — conditions on the strategic radar' },
  { id: 's4q3',  section: 'Immediate Lens',      label: 'Lens 3 · Technology Alternative to Human Work',          prompt: 'Technology Alternative to Human Work & Systems — conditions on the strategic radar' },
  { id: 's4q4',  section: 'Immediate Lens',      label: 'Lens 4 · Governance & Institutional Trust',              prompt: 'Governance & Institutional Trust — conditions on the strategic radar' },
  { id: 's4q5',  section: 'Immediate Lens',      label: 'Lens 5 · Healthy Humans',                                prompt: 'Healthy Humans — conditions on the strategic radar' },
  { id: 's4q6',  section: 'Immediate Lens',      label: 'Lens 6 · Sustainable Social & Economic Growth',          prompt: 'Sustainable Social & Economic Growth — conditions on the strategic radar' },
  { id: 's4q7',  section: 'Immediate Lens',      label: 'Lens 7 · Energy Systems & Transition',                   prompt: 'Energy Systems & Transition — conditions on the strategic radar' },
  { id: 's4q8',  section: 'Immediate Lens',      label: 'Lens 8 · Social Orientation & Societal Expectations',    prompt: 'Social Orientation & Societal Expectations — conditions on the strategic radar' },
  { id: 's4q9',  section: 'Immediate Lens',      label: 'Lens 9 · Infrastructure Transformation',                 prompt: 'Infrastructure Transformation — conditions on the strategic radar' },
  { id: 's4q10', section: 'Immediate Lens',      label: 'Lens 10 · Food & Water Security',                        prompt: 'Food & Water Security — conditions on the strategic radar' },
  { id: 's4q11', section: 'Immediate Lens',      label: 'Lens 11 · Geopolitical',                                 prompt: 'Geopolitical — conditions on the strategic radar' },
  { id: 's4q12', section: 'Immediate Lens',      label: 'Lens 12 · Wildcards & Unexpected Disruptions',           prompt: 'Wildcards & Unexpected Disruptions — conditions on the strategic radar' },

  // Section 5 — Emerging Conditions
  { id: 's5q1',  section: 'Emerging',            label: 'Emerging conditions (3–10 yr signals across the lenses)',
    prompt: 'What emerging conditions do you see across any of the 12 lenses that could shape the future for Council?' },

  // Section 6 — Other Conditions
  { id: 's6q1',  section: 'Other',               label: 'Other conditions affecting Council over next 5–10 yrs',
    prompt: 'Are there any other conditions that you believe will affect Council and the community over the next 5–10 years?' },

  // Section 7 — Strategic Opportunity-Risk (SOR)
  { id: 's7q1',  section: 'Strategic Opp-Risk (SOR)',  label: 'Strategic opportunities for Council',
    prompt: 'What do you see as potential strategic opportunities for Council?' },
  { id: 's7q2',  section: 'Strategic Opp-Risk (SOR)',  label: 'Strategic risks for Council',
    prompt: 'What do you see as potential strategic risks for Council?' },

  // Section 8 — Final Thoughts
  { id: 's8q1',  section: 'Final Thoughts',      label: 'Anything else about value, conditions and SOR',
    prompt: 'Is there anything else about value, conditions and SOR that you would like to share at this point in time?' }
];

// Helper — convenience filters
function lensQuestions()  { return QUESTIONS.filter(function(q) { return q.section === 'Immediate Lens'; }); }
function themeQuestions() { return QUESTIONS.filter(function(q) { return q.section !== 'Immediate Lens'; }); }

// ── HELPERS ─────────────────────────────────────────────────────────────────

function colLetter(n) {
  var s = '';
  while (n > 0) {
    var r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function escapeHtml(s) {
  if (s === null || s === undefined) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── SETUP ───────────────────────────────────────────────────────────────────

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupResponsesSheet(ss);
  setupDashboardSheet(ss);
  setupByLensSheet(ss);
  setupByThemeSheet(ss);
  Logger.log('All sheets configured.');
}

// ── RESPONSES SHEET ────────────────────────────────────────────────────────

function setupResponsesSheet(ss) {
  var sheet = ss.getSheetByName(RESPONSES_SHEET);
  if (!sheet) sheet = ss.insertSheet(RESPONSES_SHEET);

  var headers = ['Timestamp', 'Role', 'Tenure'];
  QUESTIONS.forEach(function(q) {
    headers.push(q.label);
  });

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold')
             .setBackground(C_INK_HEADER)
             .setFontColor('#ffffff')
             .setFontSize(10)
             .setWrap(true)
             .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 56);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 170);
  sheet.setColumnWidth(3, 130);
  for (var i = 0; i < QUESTIONS.length; i++) {
    sheet.setColumnWidth(4 + i, 300);
  }

  // Highlight the 12 lens columns with civic-green tint
  QUESTIONS.forEach(function(q, i) {
    if (q.section === 'Immediate Lens') {
      sheet.getRange(1, 4 + i).setBackground(C_GREEN);
    }
  });

  Logger.log('Responses sheet configured.');
}

// ── DASHBOARD ──────────────────────────────────────────────────────────────

function setupDashboardSheet(ss) {
  var sheet = ss.getSheetByName(DASHBOARD_SHEET);
  if (sheet) { sheet.clear(); sheet.clearConditionalFormatRules(); }
  else sheet = ss.insertSheet(DASHBOARD_SHEET);

  var refR = "'" + RESPONSES_SHEET + "'";

  // Title banner
  sheet.getRange(1, 1).setValue('Conditions & Value Capture — Dashboard');
  sheet.getRange(1, 1, 1, 5).merge()
       .setBackground(C_GREEN).setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(18)
       .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 56);

  sheet.getRange(2, 1).setValue('Glen Eira City Council · live aggregation across all submitted responses');
  sheet.getRange(2, 1, 1, 5).merge().setFontStyle('italic').setFontColor('#666');

  // KPIs
  var kpis = [
    { label: 'Total respondents',           formula: '=IFERROR(COUNTA(' + refR + '!A2:A),0)' },
    { label: 'Senior Executive',            formula: '=IFERROR(COUNTIF(' + refR + '!B2:B,"' + ROLE_LABELS['senior_exec']            + '"),0)' },
    { label: 'Senior Leader',               formula: '=IFERROR(COUNTIF(' + refR + '!B2:B,"' + ROLE_LABELS['senior_leader']          + '"),0)' },
    { label: 'Service Stream Manager',      formula: '=IFERROR(COUNTIF(' + refR + '!B2:B,"' + ROLE_LABELS['service_stream_manager'] + '"),0)' },
    { label: 'Other',                       formula: '=IFERROR(COUNTIF(' + refR + '!B2:B,"' + ROLE_LABELS['other']                  + '"),0)' }
  ];
  kpis.forEach(function(kpi, i) {
    sheet.getRange(4, 1 + i).setValue(kpi.label);
    sheet.getRange(5, 1 + i).setFormula(kpi.formula);
  });
  sheet.getRange(4, 1, 1, kpis.length)
       .setBackground(C_INK_HEADER).setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center').setWrap(true);
  sheet.getRange(5, 1, 1, kpis.length)
       .setBackground(C_GREEN_LT).setFontColor(C_GREEN)
       .setFontWeight('bold').setFontSize(20).setHorizontalAlignment('center');
  sheet.setRowHeight(4, 28);
  sheet.setRowHeight(5, 52);

  // By role table
  sheet.getRange(7, 1).setValue('BREAKDOWN BY ROLE');
  sheet.getRange(7, 1, 1, 3).merge()
       .setBackground(C_INK).setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  sheet.getRange(8, 1, 1, 3).setValues([['Role', 'Count', 'Share']]);
  sheet.getRange(8, 1, 1, 3).setBackground('#444').setFontColor('#fff').setFontWeight('bold');

  var roleKeys = Object.keys(ROLE_LABELS);
  roleKeys.forEach(function(k, i) {
    var row = 9 + i;
    sheet.getRange(row, 1).setValue(ROLE_LABELS[k]);
    sheet.getRange(row, 2).setFormula('=IFERROR(COUNTIF(' + refR + '!B2:B,A' + row + '),0)');
    sheet.getRange(row, 3).setFormula('=IFERROR(B' + row + '/$B$5,0)');
  });
  sheet.getRange(9, 2, roleKeys.length, 1).setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange(9, 3, roleKeys.length, 1).setHorizontalAlignment('center').setNumberFormat('0%');

  // By tenure table
  var tenStartCol = 4;
  sheet.getRange(7, tenStartCol).setValue('BREAKDOWN BY TENURE');
  sheet.getRange(7, tenStartCol, 1, 3).merge()
       .setBackground(C_INK).setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  sheet.getRange(8, tenStartCol, 1, 3).setValues([['Tenure', 'Count', 'Share']]);
  sheet.getRange(8, tenStartCol, 1, 3).setBackground('#444').setFontColor('#fff').setFontWeight('bold');

  var tenureKeys = Object.keys(TENURE_LABELS);
  tenureKeys.forEach(function(k, i) {
    var row = 9 + i;
    sheet.getRange(row, tenStartCol).setValue(TENURE_LABELS[k]);
    sheet.getRange(row, tenStartCol + 1).setFormula('=IFERROR(COUNTIF(' + refR + '!C2:C,D' + row + '),0)');
    sheet.getRange(row, tenStartCol + 2).setFormula('=IFERROR(E' + row + '/$B$5,0)');
  });
  sheet.getRange(9, tenStartCol + 1, tenureKeys.length, 1).setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange(9, tenStartCol + 2, tenureKeys.length, 1).setHorizontalAlignment('center').setNumberFormat('0%');

  // Completion rate per question
  var compStartRow = 9 + Math.max(roleKeys.length, tenureKeys.length) + 2;
  sheet.getRange(compStartRow, 1).setValue('RESPONSE COMPLETION BY QUESTION');
  sheet.getRange(compStartRow, 1, 1, 5).merge()
       .setBackground(C_INK).setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);

  sheet.getRange(compStartRow + 1, 1, 1, 4).setValues([['Section', 'Question', 'Answered', 'Completion %']]);
  sheet.getRange(compStartRow + 1, 1, 1, 4).setBackground('#444').setFontColor('#fff').setFontWeight('bold');

  QUESTIONS.forEach(function(q, i) {
    var row = compStartRow + 2 + i;
    var qCol = 4 + i; // Responses column: D = 4
    var colL = colLetter(qCol);
    sheet.getRange(row, 1).setValue(q.section);
    sheet.getRange(row, 2).setValue(q.label);
    sheet.getRange(row, 3).setFormula('=IFERROR(COUNTIF(' + refR + '!' + colL + '2:' + colL + ',"<>"),0)');
    sheet.getRange(row, 4).setFormula('=IFERROR(C' + row + '/$B$5,0)');
  });
  var qBlock = sheet.getRange(compStartRow + 2, 1, QUESTIONS.length, 4);
  qBlock.setFontSize(10);
  sheet.getRange(compStartRow + 2, 3, QUESTIONS.length, 1).setHorizontalAlignment('center');
  sheet.getRange(compStartRow + 2, 4, QUESTIONS.length, 1).setHorizontalAlignment('center').setNumberFormat('0%');

  // Conditional formatting on completion %
  var compRange = sheet.getRange(compStartRow + 2, 4, QUESTIONS.length, 1);
  var rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0.5).setBackground(C_RED_LO).setRanges([compRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.5, 0.8).setBackground(C_AMBER_MID).setRanges([compRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(0.8).setBackground(C_GREEN_HI).setRanges([compRange]).build());
  sheet.setConditionalFormatRules(rules);

  // Tint lens rows green
  QUESTIONS.forEach(function(q, i) {
    if (q.section === 'Immediate Lens') {
      sheet.getRange(compStartRow + 2 + i, 1).setBackground(C_GREEN_LT).setFontColor(C_GREEN).setFontWeight('bold');
    }
  });

  // Column widths
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 320);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 130);

  sheet.setHiddenGridlines(true);

  Logger.log('Dashboard sheet configured.');
}

// ── BY LENS ────────────────────────────────────────────────────────────────
// Stacks all submitted responses for each of the 12 strategic lenses (plus the
// "Emerging" free-text), grouped together for thematic review.

function setupByLensSheet(ss) {
  var sheet = ss.getSheetByName(BY_LENS_SHEET);
  if (sheet) { sheet.clear(); sheet.clearConditionalFormatRules(); }
  else sheet = ss.insertSheet(BY_LENS_SHEET);

  sheet.getRange(1, 1).setValue('By Lens — Responses grouped by strategic lens');
  sheet.getRange(1, 1, 1, 4).merge()
       .setBackground(C_GREEN).setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(14);
  sheet.setRowHeight(1, 36);

  sheet.getRange(2, 1).setValue('Refresh after new submissions: Extensions > Apps Script > run refreshAggregates()');
  sheet.getRange(2, 1, 1, 4).merge().setFontStyle('italic').setFontColor('#666');

  var headers = ['Lens', 'Respondent #', 'Role', 'Response'];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(3, 1, 1, headers.length)
       .setBackground(C_INK_HEADER).setFontColor('#ffffff')
       .setFontWeight('bold').setHorizontalAlignment('center');

  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 700);

  sheet.setFrozenRows(3);

  populateByLens(sheet);
  Logger.log('By Lens sheet configured.');
}

function populateByLens(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BY_LENS_SHEET);
  if (!sheet) return;

  // Wipe existing data rows (keep title + header rows 1-3)
  var lastRow = sheet.getLastRow();
  if (lastRow > 3) {
    sheet.getRange(4, 1, lastRow - 3, 4).clearContent().clearFormat();
  }

  var responses = readResponses();
  if (!responses.length) return;

  // Section 4 lenses + Section 5 "Emerging" rolls in too
  var lensCols = QUESTIONS
    .map(function(q, i) { return { q: q, colIdx: 4 + i }; })
    .filter(function(item) { return item.q.section === 'Immediate Lens' || item.q.section === 'Emerging'; });

  var rowOut = 4;

  lensCols.forEach(function(item) {
    // Lens band row
    sheet.getRange(rowOut, 1, 1, 4).merge()
         .setValue(item.q.label)
         .setBackground(C_GREEN_LT).setFontColor(C_GREEN)
         .setFontWeight('bold').setFontSize(12).setVerticalAlignment('middle');
    sheet.setRowHeight(rowOut, 28);
    rowOut++;

    var any = false;
    responses.forEach(function(resp, ri) {
      var text = (resp.values[item.colIdx - 1] || '').toString().trim();
      if (!text) return;
      any = true;
      sheet.getRange(rowOut, 1).setValue('');
      sheet.getRange(rowOut, 2).setValue('R-' + (ri + 1)).setHorizontalAlignment('center');
      sheet.getRange(rowOut, 3).setValue(resp.role || '').setFontSize(10);
      sheet.getRange(rowOut, 4).setValue(text).setWrap(true).setFontSize(11).setVerticalAlignment('top');
      if (rowOut % 2 === 0) {
        sheet.getRange(rowOut, 1, 1, 4).setBackground(C_GREY_ROW);
      }
      rowOut++;
    });
    if (!any) {
      sheet.getRange(rowOut, 1, 1, 4).merge()
           .setValue('(no responses yet)')
           .setFontStyle('italic').setFontColor('#999')
           .setHorizontalAlignment('center');
      rowOut++;
    }
    rowOut++; // spacer row
  });
}

// ── BY THEME ───────────────────────────────────────────────────────────────
// Aggregates the non-lens open-text questions: Value Today, Future Value,
// Strategic Opp-Risk, Other, Final Thoughts.

function setupByThemeSheet(ss) {
  var sheet = ss.getSheetByName(BY_THEME_SHEET);
  if (sheet) { sheet.clear(); sheet.clearConditionalFormatRules(); }
  else sheet = ss.insertSheet(BY_THEME_SHEET);

  sheet.getRange(1, 1).setValue('By Theme — Open-text responses grouped by section');
  sheet.getRange(1, 1, 1, 4).merge()
       .setBackground(C_GREEN).setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(14);
  sheet.setRowHeight(1, 36);

  sheet.getRange(2, 1).setValue('Refresh after new submissions: Extensions > Apps Script > run refreshAggregates()');
  sheet.getRange(2, 1, 1, 4).merge().setFontStyle('italic').setFontColor('#666');

  var headers = ['Question', 'Respondent #', 'Role', 'Response'];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(3, 1, 1, headers.length)
       .setBackground(C_INK_HEADER).setFontColor('#ffffff')
       .setFontWeight('bold').setHorizontalAlignment('center');

  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 700);

  sheet.setFrozenRows(3);

  populateByTheme(sheet);
  Logger.log('By Theme sheet configured.');
}

function populateByTheme(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BY_THEME_SHEET);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow > 3) {
    sheet.getRange(4, 1, lastRow - 3, 4).clearContent().clearFormat();
  }

  var responses = readResponses();
  if (!responses.length) return;

  var themeCols = QUESTIONS
    .map(function(q, i) { return { q: q, colIdx: 4 + i }; })
    .filter(function(item) { return item.q.section !== 'Immediate Lens' && item.q.section !== 'Emerging'; });

  var rowOut = 4;
  var lastSection = null;

  themeCols.forEach(function(item) {
    // Section divider (e.g. "VALUE TODAY")
    if (item.q.section !== lastSection) {
      sheet.getRange(rowOut, 1, 1, 4).merge()
           .setValue(item.q.section.toUpperCase())
           .setBackground(C_INK).setFontColor('#ffffff')
           .setFontWeight('bold').setFontSize(11)
           .setHorizontalAlignment('left');
      sheet.setRowHeight(rowOut, 26);
      rowOut++;
      lastSection = item.q.section;
    }

    // Question band
    sheet.getRange(rowOut, 1, 1, 4).merge()
         .setValue(item.q.label + '  —  ' + item.q.prompt)
         .setBackground(C_GREEN_LT).setFontColor(C_GREEN)
         .setFontWeight('bold').setFontSize(11);
    sheet.setRowHeight(rowOut, 28);
    rowOut++;

    var any = false;
    responses.forEach(function(resp, ri) {
      var text = (resp.values[item.colIdx - 1] || '').toString().trim();
      if (!text) return;
      any = true;
      sheet.getRange(rowOut, 2).setValue('R-' + (ri + 1)).setHorizontalAlignment('center');
      sheet.getRange(rowOut, 3).setValue(resp.role || '').setFontSize(10);
      sheet.getRange(rowOut, 4).setValue(text).setWrap(true).setFontSize(11).setVerticalAlignment('top');
      if (rowOut % 2 === 0) {
        sheet.getRange(rowOut, 1, 1, 4).setBackground(C_GREY_ROW);
      }
      rowOut++;
    });
    if (!any) {
      sheet.getRange(rowOut, 1, 1, 4).merge()
           .setValue('(no responses yet)')
           .setFontStyle('italic').setFontColor('#999')
           .setHorizontalAlignment('center');
      rowOut++;
    }
    rowOut++; // spacer
  });
}

// ── DATA READERS ───────────────────────────────────────────────────────────

function readResponses() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RESPONSES_SHEET);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = 3 + QUESTIONS.length;
  var data    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return data.map(function(row) {
    return {
      timestamp: row[0],
      role:      row[1],
      tenure:    row[2],
      // values[0..QUESTIONS.length-1] aligns with QUESTIONS[]
      values:    row.slice(3)
    };
  });
}

// Manual refresh — run after new responses come in, to rebuild the
// By Lens / By Theme sheets from current data.
function refreshAggregates() {
  populateByLens();
  populateByTheme();
  Logger.log('Aggregates refreshed.');
}

// ── WEB APP HANDLERS ───────────────────────────────────────────────────────

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: 'error', message: 'no payload' });
    }

    var payload = JSON.parse(e.postData.contents);
    var responses = payload.responses || {};

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(RESPONSES_SHEET);
    if (!sheet) {
      setupSheets();
      sheet = ss.getSheetByName(RESPONSES_SHEET);
    }

    // Timestamp
    var when = payload.submittedAt ? new Date(payload.submittedAt) : new Date();
    var timestamp = Utilities.formatDate(when, TIMEZONE, 'dd/MM/yyyy HH:mm:ss');

    // Map role/tenure codes -> friendly labels
    var roleLabel   = ROLE_LABELS[payload.role]     || payload.role   || '';
    var tenureLabel = TENURE_LABELS[payload.tenure] || payload.tenure || '';

    // Build the row in the canonical question order
    var row = [timestamp, roleLabel, tenureLabel];
    QUESTIONS.forEach(function(q) {
      var v = responses[q.id];
      row.push((v === undefined || v === null) ? '' : String(v));
    });

    sheet.appendRow(row);

    // Light formatting on the new row — wrap long text + top-align
    var newRow = sheet.getLastRow();
    sheet.getRange(newRow, 4, 1, QUESTIONS.length)
         .setWrap(true).setVerticalAlignment('top').setFontSize(10);

    return jsonResponse({ status: 'ok' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function doGet(e) {
  return jsonResponse({
    status:  'ok',
    message: 'Glen Eira City Council — Conditions & Value Capture endpoint active'
  });
}

// ── TESTING ────────────────────────────────────────────────────────────────
// Run this manually after setupSheets() to insert a fake submission and
// verify the Responses / Dashboard / By Lens / By Theme views all behave.

function testSubmission() {
  var responses = {};
  QUESTIONS.forEach(function(q, i) {
    if (q.section === 'Immediate Lens') {
      responses[q.id] = 'Sample condition A for ' + q.label + '; Sample condition B';
    } else {
      responses[q.id] = 'Sample test response for ' + q.label + '.';
    }
  });
  var fakeEvent = {
    postData: {
      contents: JSON.stringify({
        role:        'manager',
        tenure:      '3-7',
        responses:   responses,
        submittedAt: new Date().toISOString()
      })
    }
  };
  var result = doPost(fakeEvent);
  Logger.log(result.getContent());
  refreshAggregates();
}
