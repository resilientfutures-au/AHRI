// AHRI SOR Prioritisation Survey — Google Apps Script (v2)
//
// SETUP:
//   1. Open your destination Google Spreadsheet.
//   2. Extensions > Apps Script, paste this file.
//   3. Run setupSheets() once to create/refresh all view sheets.
//   4. Deploy as Web App: Execute as "Me", access "Anyone".
//   5. Copy the deployment URL into AHRI/sor-prioritisation/index.html
//
// SHEETS CREATED:
//   • Responses    Raw response data (one row per respondent)
//   • Summary      Average score per sub-SOR per criterion
//   • Ranking      All 27 sub-SORs ranked by average total score
//   • By Primary SOR     Aggregated scores at Primary SOR level (6 Primary SORs)
//   • Dashboard    High-level metrics, top picks, leaderboard
//
// MAINTENANCE:
//   • Run setupSheets() any time you change ESORS or QUESTIONS constants.
//   • All summary sheets use formulas that auto-update on new submissions.
//   • Run sortRankingSheet() any time you want the Ranking sheet re-sorted.

// ── CONFIGURATION ─────────────────────────────────────────────────────────────

var RESPONSES_SHEET = 'Responses';
var SUMMARY_SHEET   = 'Summary';
var RANKING_SHEET   = 'Ranking';
var ESOR_SHEET      = 'By Primary SOR';
var DASHBOARD_SHEET = 'Dashboard';
var TIMEZONE        = 'Australia/Melbourne';

// Brand colours
var C_GOLD       = '#a44f43';
var C_GOLD_LT    = '#f5e0dd';
var C_INK        = '#1a1a1a';
var C_INK_HEADER = '#2d2d2d';
var C_GREY_ROW   = '#fafafa';
var C_GREEN_HI   = '#d4edda';
var C_RED_LO     = '#f8d7da';
var C_AMBER_MID  = '#fff3cd';

var ESORS = [
  {
    id: 1, title: 'Scope of Play',
    subSors: [
      { idx: 1, label: 'The Foundational Identity Choice' },
      { idx: 2, label: 'Beyond HR: Redefining the Professional Boundary' },
      { idx: 3, label: 'Whole-System Workforce Leadership > Filling the National Vacuum' },
      { idx: 4, label: 'Brand & Identity Evolution > From \'Human Resources\' Institute to Contemporary Professional Home' },
      { idx: 5, label: 'Strategic Letting Go > What AHRI Must Stop Doing' }
    ]
  },
  {
    id: 2, title: 'Certification and Professional Pathway',
    subSors: [
      { idx: 1, label: 'CPA/AICD Equivalence — From Aspirational Credential to Market-Making Standard' },
      { idx: 2, label: 'Certification as Long-Term Centrepiece' },
      { idx: 3, label: 'Employer Mandate & C-Suite Sponsorship as the Adoption Mechanism' },
      { idx: 4, label: 'Future-Fit Skills Credential: From Topic Knowledge to Capability-Based' },
      { idx: 5, label: 'University Integration & Career Pipeline — The Certification On-Ramp' }
    ]
  },
  {
    id: 3, title: 'AI and Workforce Transformation',
    subSors: [
      { idx: 1, label: 'Trusted AI Advisor & Closing the HR Credibility Gap on AI' },
      { idx: 2, label: 'Human-AI Hybrid Workforce Governance. Designing for Mixed Human-Machine Teams' },
      { idx: 3, label: 'Ethical AI in People Decisions & Building the Governance Framework' },
      { idx: 4, label: 'AI as Psychosocial Hazard & Managing the Wellbeing-Automation' },
      { idx: 5, label: 'Australia\'s Technology Lag — Leading From Behind in a Fast-Moving Global Context' },
      { idx: 6, label: 'Agentic HR Value-Add Elevator of Practice and Potential Professional Displacement' }
    ]
  },
  {
    id: 4, title: 'Advocacy and Strategic Voice',
    subSors: [
      { idx: 1, label: 'Government\'s Go-To Advisor: Filling the Policy Vacuum' },
      { idx: 2, label: 'Regulatory Translation Service: From Compliance Commentary to Trusted Rapid-Response' },
      { idx: 3, label: 'Research & Thought Leadership — From Under-Promoted Asset to National Intelligence Source' },
      { idx: 4, label: 'Navigating Social & Political Complexity — Purpose, DEI, and Values-Based Leadership' },
      { idx: 5, label: 'Visibility Amplification — Leveraging Existing Value to Build National Influence' }
    ]
  },
  {
    id: 5, title: 'Platform / Ecosystem Model',
    subSors: [
      { idx: 1, label: 'The Apple App Store Model > From In-House Producer to Curated Ecosystem' },
      { idx: 2, label: 'From Defensive Insularity to Collaborative Amplification' },
      { idx: 3, label: 'Certification as Platform Trust Architecture' },
      { idx: 4, label: 'Building vs Buying vs Partnering: The Capability Acquisition Framework' },
      { idx: 5, label: 'Revenue Diversification Through Ecosystem Economics' }
    ]
  },
  {
    id: 6, title: 'Membership and Member Economics',
    subSors: [
      { idx: 1, label: 'C-Suite Engagement as the Membership Force Multiplier' },
      { idx: 2, label: 'Guild Model Transformation: From 17,000 Members to 100,000+ Professional Home' },
      { idx: 3, label: 'Hyper-Personalisation — From One-Size-Fits-All to Contextually Relevant' },
      { idx: 4, label: 'Volunteer Workforce Redesign — From Event Burden to Meaningful Professional Contribution' },
      { idx: 5, label: 'Events as Key Components of Community and Connection' },
      { idx: 6, label: 'Communities of Practice for Relationship with Like-Minded / Like-Focussed Professionals' }
    ]
  }
];

var QUESTIONS = [
  { id: 1,  label: 'Scale of Impact' },
  { id: 2,  label: 'Conditions Volatility' },
  { id: 3,  label: 'Conditions Timing' },
  { id: 4,  label: 'Lag Time — Systemic Effect Speed' },
  { id: 5,  label: 'Speed of Activation' },
  { id: 6,  label: 'Systemic Responses' },
  { id: 7,  label: 'Strategic Advantage' },
  { id: 8,  label: 'Reversibility & Optionality' },
  { id: 9,  label: 'Capability Readiness' },
  { id: 10, label: 'Capability-Building Return' },
  { id: 11, label: 'Learning Velocity' },
  { id: 12, label: 'Value Network Alignment Effort' },
  { id: 13, label: 'Resource & Effort Consumption' },
  { id: 14, label: 'Strategic Fitness Contribution' }
];

// ── HELPERS ───────────────────────────────────────────────────────────────────

function makeKey(esorId, ssIdx, qId) {
  return 'esor' + esorId + '_ss' + ssIdx + '_q' + qId;
}

function colLetter(n) {
  var s = '';
  while (n > 0) {
    var r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function keyOrder() {
  var keys = [];
  ESORS.forEach(function(esor) {
    esor.subSors.forEach(function(ss) {
      QUESTIONS.forEach(function(q) {
        keys.push(makeKey(esor.id, ss.idx, q.id));
      });
    });
  });
  return keys;
}

function flatSubSors() {
  var out = [];
  ESORS.forEach(function(esor, esorIdx) {
    esor.subSors.forEach(function(subSor, ssIdx) {
      out.push({
        esor: esor, esorIdx: esorIdx,
        subSor: subSor, ssIdx: ssIdx,
        globalIdx: out.length
      });
    });
  });
  return out;
}

function responseCol(esorIdx, ssIdx, qIdx) {
  var col = 3;
  for (var i = 0; i < esorIdx; i++) {
    col += ESORS[i].subSors.length * QUESTIONS.length;
  }
  col += ssIdx * QUESTIONS.length;
  col += qIdx + 1;
  return col;
}

function subtotalCol(globalSsIdx) {
  var ratingsTotal = 0;
  ESORS.forEach(function(e) { ratingsTotal += e.subSors.length * QUESTIONS.length; });
  return 3 + ratingsTotal + globalSsIdx + 1;
}

function responseHeaders() {
  var headers = ['Timestamp', 'Name', 'Email'];
  ESORS.forEach(function(esor) {
    esor.subSors.forEach(function(ss) {
      QUESTIONS.forEach(function(q) {
        headers.push('Primary SOR ' + esor.id + ' · SS' + ss.idx + ' · Q' + q.id + ': ' + q.label);
      });
    });
  });
  ESORS.forEach(function(esor) {
    esor.subSors.forEach(function(ss) {
      headers.push('Primary SOR ' + esor.id + ' SS' + ss.idx + ' Total (/98)');
    });
  });
  return headers;
}

// ── SETUP ─────────────────────────────────────────────────────────────────────

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupResponsesSheet(ss);
  setupSummarySheet(ss);
  setupRankingSheet(ss);
  setupEsorRollupSheet(ss);
  setupDashboardSheet(ss);
  Logger.log('All sheets configured.');
}

// ── RESPONSES SHEET ──────────────────────────────────────────────────────────

function setupResponsesSheet(ss) {
  var sheet = ss.getSheetByName(RESPONSES_SHEET);
  if (!sheet) sheet = ss.insertSheet(RESPONSES_SHEET);

  var headers = responseHeaders();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold')
             .setBackground(C_INK_HEADER)
             .setFontColor('#ffffff')
             .setFontSize(9)
             .setWrap(true);
  sheet.setRowHeight(1, 80);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 180);

  var ratingStart  = 4;
  var totalRatings = keyOrder().length;
  sheet.setColumnWidths(ratingStart, totalRatings, 60);

  var subtotalStart = ratingStart + totalRatings;
  var totalSubSors  = flatSubSors().length;
  sheet.setColumnWidths(subtotalStart, totalSubSors, 100);

  sheet.getRange(1, subtotalStart, 1, totalSubSors).setBackground(C_GOLD);

  applyRatingConditionalFormat(sheet, ratingStart, totalRatings);

  Logger.log('Responses sheet configured.');
}

function applyRatingConditionalFormat(sheet, startCol, numCols) {
  var range = sheet.getRange(2, startCol, Math.max(1, sheet.getMaxRows() - 1), numCols);
  var rules = sheet.getConditionalFormatRules();

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(2)
    .setBackground(C_RED_LO)
    .setRanges([range]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(6)
    .setBackground(C_GREEN_HI)
    .setRanges([range]).build());

  sheet.setConditionalFormatRules(rules);
}

// ── SUMMARY SHEET (sub-SOR × criterion matrix) ───────────────────────────────

function setupSummarySheet(ss) {
  var sheet = ss.getSheetByName(SUMMARY_SHEET);
  if (sheet) { sheet.clear(); sheet.clearConditionalFormatRules(); }
  else sheet = ss.insertSheet(SUMMARY_SHEET);

  var refR  = "'" + RESPONSES_SHEET + "'";
  var flat  = flatSubSors();
  var nCols = flat.length;

  // Title row (not merged — would conflict with frozen column 1)
  sheet.getRange(1, 1).setValue('Summary — Average score per sub-SOR per criterion');
  sheet.getRange(1, 1, 1, nCols + 2)
       .setBackground(C_GOLD).setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('left');
  sheet.setRowHeight(1, 36);

  // Primary SOR band row
  var row2 = [''];
  flat.forEach(function(item) { row2.push('Primary SOR ' + item.esor.id + ': ' + item.esor.title); });
  row2.push('');
  sheet.getRange(2, 1, 1, row2.length).setValues([row2]);
  sheet.setRowHeight(2, 32);

  // Sub-SOR label row
  var row3 = ['Criterion'];
  flat.forEach(function(item) { row3.push(item.subSor.label); });
  row3.push('Avg across all');
  sheet.getRange(3, 1, 1, row3.length).setValues([row3]);
  sheet.setRowHeight(3, 90);

  // Criteria rows
  for (var qi = 0; qi < QUESTIONS.length; qi++) {
    var row = [QUESTIONS[qi].label];
    flat.forEach(function(item) {
      var col  = responseCol(item.esorIdx, item.ssIdx, qi);
      var colL = colLetter(col);
      row.push('=IFERROR(AVERAGEIF(' + refR + '!' + colL + ':' + colL + ',"<>",' + refR + '!' + colL + ':' + colL + '),"")');
    });
    var firstColL = colLetter(2);
    var lastColL  = colLetter(2 + nCols - 1);
    row.push('=IFERROR(AVERAGE(' + firstColL + (4 + qi) + ':' + lastColL + (4 + qi) + '),"")');
    sheet.getRange(4 + qi, 1, 1, row.length).setValues([row]);
  }

  // Total row
  var totalRow = ['Average Total /98'];
  for (var si = 0; si < nCols; si++) {
    var tcolL = colLetter(2 + si);
    totalRow.push('=IFERROR(SUM(' + tcolL + '4:' + tcolL + (3 + QUESTIONS.length) + '),"")');
  }
  totalRow.push('');
  sheet.getRange(4 + QUESTIONS.length, 1, 1, totalRow.length).setValues([totalRow]);

  // Respondent count
  sheet.getRange(4 + QUESTIONS.length + 2, 1).setValue('Respondents');
  sheet.getRange(4 + QUESTIONS.length + 2, 2)
       .setFormula('=IFERROR(COUNTA(' + refR + '!A2:A),0)');

  // Formatting
  formatSubSorBandRow(sheet, 2, nCols);

  var r3 = sheet.getRange(3, 1, 1, nCols + 2);
  r3.setFontWeight('bold').setBackground(C_INK_HEADER).setFontColor('#ffffff')
    .setFontSize(9).setWrap(true).setVerticalAlignment('bottom');
  sheet.getRange(3, nCols + 2).setBackground(C_GOLD);

  sheet.getRange(4, 1, QUESTIONS.length + 1, 1)
       .setFontWeight('bold').setBackground('#fafafa');

  var dataRange = sheet.getRange(4, 2, QUESTIONS.length, nCols + 1);
  dataRange.setHorizontalAlignment('center').setNumberFormat('0.00');

  applySummaryColourRules(sheet, sheet.getRange(4, 2, QUESTIONS.length, nCols));

  var totalRange = sheet.getRange(4 + QUESTIONS.length, 1, 1, nCols + 2);
  totalRange.setFontWeight('bold').setBackground(C_GOLD_LT).setFontColor(C_GOLD);
  sheet.getRange(4 + QUESTIONS.length, 2, 1, nCols)
       .setHorizontalAlignment('center').setNumberFormat('0.00');

  for (var qi2 = 0; qi2 < QUESTIONS.length; qi2++) {
    if (qi2 % 2 === 0) {
      sheet.getRange(4 + qi2, 1, 1, nCols + 2).setBackground(C_GREY_ROW);
    }
  }

  sheet.getRange(4, nCols + 2, QUESTIONS.length, 1).setBackground('#fffaf0');

  sheet.setColumnWidth(1, 240);
  for (var si2 = 0; si2 < nCols; si2++) sheet.setColumnWidth(2 + si2, 110);
  sheet.setColumnWidth(nCols + 2, 110);

  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(1);

  Logger.log('Summary sheet configured.');
}

function formatSubSorBandRow(sheet, rowNum, nCols) {
  var palette = ['#fbe9e7', '#f3e5f5', '#e3f2fd', '#e8f5e9', '#fff9c4', '#fce4ec'];
  var flat = flatSubSors();
  flat.forEach(function(item, i) {
    sheet.getRange(rowNum, 2 + i)
         .setBackground(palette[item.esorIdx])
         .setFontWeight('bold').setFontSize(9)
         .setWrap(true).setHorizontalAlignment('center')
         .setFontColor(C_INK);
  });
}

function applySummaryColourRules(sheet, range) {
  var rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(3).setBackground(C_RED_LO).setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(3, 5).setBackground(C_AMBER_MID).setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(5).setBackground(C_GREEN_HI).setRanges([range]).build());
  sheet.setConditionalFormatRules(rules);
}

// ── RANKING SHEET (all sub-SORs sorted by total score) ──────────────────────

function setupRankingSheet(ss) {
  var sheet = ss.getSheetByName(RANKING_SHEET);
  if (sheet) { sheet.clear(); sheet.clearConditionalFormatRules(); }
  else sheet = ss.insertSheet(RANKING_SHEET);

  var refR = "'" + RESPONSES_SHEET + "'";
  var flat = flatSubSors();

  sheet.getRange(1, 1).setValue('Ranking — All sub-SORs by average total score');
  sheet.getRange(1, 1, 1, 6).merge()
       .setBackground(C_GOLD).setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(13);
  sheet.setRowHeight(1, 36);

  sheet.getRange(2, 1).setValue('Max possible = 98 (14 criteria × 7). Run sortRankingSheet() to re-sort after new responses.');
  sheet.getRange(2, 1, 1, 6).merge().setFontStyle('italic').setFontColor('#666');
  sheet.setRowHeight(2, 22);

  var headers = ['Rank', 'Primary SOR', 'Sub-SOR', 'Avg Total /98', 'Avg per Criterion', 'Respondents'];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(3, 1, 1, headers.length)
       .setBackground(C_INK_HEADER).setFontColor('#ffffff')
       .setFontWeight('bold').setHorizontalAlignment('center');

  for (var i = 0; i < flat.length; i++) {
    var item    = flat[i];
    var subCol  = subtotalCol(item.globalIdx);
    var subColL = colLetter(subCol);
    var avgCell = 'AVERAGEIF(' + refR + '!' + subColL + ':' + subColL + ',"<>",' + refR + '!' + subColL + ':' + subColL + ')';
    var cntCell = 'COUNTIF(' + refR + '!' + subColL + ':' + subColL + ',"<>")';
    var row     = 4 + i;

    sheet.getRange(row, 1).setFormula('=IFERROR(RANK(D' + row + ',D$4:D$' + (3 + flat.length) + ',0),"")');
    sheet.getRange(row, 2).setValue('Primary SOR ' + item.esor.id + ': ' + item.esor.title);
    sheet.getRange(row, 3).setValue(item.subSor.label);
    sheet.getRange(row, 4).setFormula('=IFERROR(' + avgCell + ',"")');
    sheet.getRange(row, 5).setFormula('=IFERROR(D' + row + '/14,"")');
    sheet.getRange(row, 6).setFormula('=IFERROR(' + cntCell + ',0)');
  }

  // Formatting
  sheet.getRange(4, 1, flat.length, 1).setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange(4, 2, flat.length, 1).setFontSize(10);
  sheet.getRange(4, 3, flat.length, 1).setWrap(true).setFontSize(10);
  sheet.getRange(4, 4, flat.length, 1).setHorizontalAlignment('center').setNumberFormat('0.0').setFontWeight('bold');
  sheet.getRange(4, 5, flat.length, 1).setHorizontalAlignment('center').setNumberFormat('0.00');
  sheet.getRange(4, 6, flat.length, 1).setHorizontalAlignment('center');

  var palette = ['#fbe9e7', '#f3e5f5', '#e3f2fd', '#e8f5e9', '#fff9c4', '#fce4ec'];
  flat.forEach(function(item, i) {
    sheet.getRange(4 + i, 2).setBackground(palette[item.esorIdx]);
  });

  var avgRange = sheet.getRange(4, 4, flat.length, 1);
  var rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint('#2e7d32').setGradientMinpoint('#c62828')
    .setRanges([avgRange]).build());
  sheet.setConditionalFormatRules(rules);

  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 380);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 110);

  sheet.setFrozenRows(3);

  Logger.log('Ranking sheet configured.');
}

// ── BY Primary SOR SHEET (aggregated at Primary SOR level) ──────────────────────────────

function setupEsorRollupSheet(ss) {
  var sheet = ss.getSheetByName(ESOR_SHEET);
  if (sheet) { sheet.clear(); sheet.clearConditionalFormatRules(); }
  else sheet = ss.insertSheet(ESOR_SHEET);

  var refR = "'" + RESPONSES_SHEET + "'";

  // Title row (not merged — would conflict with frozen column 1)
  sheet.getRange(1, 1).setValue('By Primary SOR — Average score per Primary SOR per criterion (across all sub-SORs)');
  sheet.getRange(1, 1, 1, ESORS.length + 2)
       .setBackground(C_GOLD).setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(13);
  sheet.setRowHeight(1, 36);

  var row2 = ['Criterion'];
  ESORS.forEach(function(esor) { row2.push('Primary SOR ' + esor.id + ': ' + esor.title); });
  row2.push('Overall');
  sheet.getRange(2, 1, 1, row2.length).setValues([row2]);
  sheet.setRowHeight(2, 60);

  for (var qi = 0; qi < QUESTIONS.length; qi++) {
    var row = [QUESTIONS[qi].label];
    ESORS.forEach(function(esor, ei) {
      var cellRefs = [];
      esor.subSors.forEach(function(_, si) {
        var col  = responseCol(ei, si, qi);
        var colL = colLetter(col);
        cellRefs.push(refR + '!' + colL + '2:' + colL);
      });
      row.push('=IFERROR(AVERAGE(' + cellRefs.join(',') + '),"")');
    });
    var allRefs = [];
    ESORS.forEach(function(esor, ei) {
      esor.subSors.forEach(function(_, si) {
        var col  = responseCol(ei, si, qi);
        var colL = colLetter(col);
        allRefs.push(refR + '!' + colL + '2:' + colL);
      });
    });
    row.push('=IFERROR(AVERAGE(' + allRefs.join(',') + '),"")');
    sheet.getRange(3 + qi, 1, 1, row.length).setValues([row]);
  }

  var totalRow = ['Average Total /98'];
  for (var ei = 0; ei < ESORS.length; ei++) {
    var colL = colLetter(2 + ei);
    totalRow.push('=IFERROR(SUM(' + colL + '3:' + colL + (2 + QUESTIONS.length) + '),"")');
  }
  var overallColL = colLetter(2 + ESORS.length);
  totalRow.push('=IFERROR(SUM(' + overallColL + '3:' + overallColL + (2 + QUESTIONS.length) + '),"")');
  sheet.getRange(3 + QUESTIONS.length, 1, 1, totalRow.length).setValues([totalRow]);

  // Formatting
  var palette = ['#fbe9e7', '#f3e5f5', '#e3f2fd', '#e8f5e9', '#fff9c4', '#fce4ec'];
  ESORS.forEach(function(_, ei) {
    sheet.getRange(2, 2 + ei).setBackground(palette[ei]).setFontWeight('bold')
         .setFontSize(10).setWrap(true).setHorizontalAlignment('center');
  });
  sheet.getRange(2, ESORS.length + 2)
       .setBackground(C_GOLD).setFontColor('#ffffff').setFontWeight('bold')
       .setHorizontalAlignment('center');
  sheet.getRange(2, 1).setBackground(C_INK_HEADER).setFontColor('#ffffff').setFontWeight('bold');

  sheet.getRange(3, 1, QUESTIONS.length, 1).setFontWeight('bold').setBackground('#fafafa');
  sheet.getRange(3, 2, QUESTIONS.length, ESORS.length + 1)
       .setHorizontalAlignment('center').setNumberFormat('0.00');

  applySummaryColourRules(sheet, sheet.getRange(3, 2, QUESTIONS.length, ESORS.length + 1));

  var totRange = sheet.getRange(3 + QUESTIONS.length, 1, 1, ESORS.length + 2);
  totRange.setFontWeight('bold').setBackground(C_GOLD_LT).setFontColor(C_GOLD);
  sheet.getRange(3 + QUESTIONS.length, 2, 1, ESORS.length + 1)
       .setHorizontalAlignment('center').setNumberFormat('0.00');

  for (var qi2 = 0; qi2 < QUESTIONS.length; qi2++) {
    if (qi2 % 2 === 0) {
      sheet.getRange(3 + qi2, 1, 1, ESORS.length + 2).setBackground(C_GREY_ROW);
    }
  }

  sheet.setColumnWidth(1, 240);
  for (var ei2 = 0; ei2 < ESORS.length; ei2++) sheet.setColumnWidth(2 + ei2, 160);
  sheet.setColumnWidth(ESORS.length + 2, 120);

  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1);

  Logger.log('By Primary SOR sheet configured.');
}

// ── DASHBOARD SHEET (high-level overview) ────────────────────────────────────

function setupDashboardSheet(ss) {
  var sheet = ss.getSheetByName(DASHBOARD_SHEET);
  if (sheet) { sheet.clear(); sheet.clearConditionalFormatRules(); }
  else sheet = ss.insertSheet(DASHBOARD_SHEET);

  var refR = "'" + RESPONSES_SHEET + "'";
  var refK = "'" + RANKING_SHEET + "'";
  var refE = "'" + ESOR_SHEET + "'";

  // Title banner
  sheet.getRange(1, 1).setValue('SOR Prioritisation Dashboard');
  sheet.getRange(1, 1, 1, 5).merge()
       .setBackground(C_GOLD).setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(18)
       .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 56);

  sheet.getRange(2, 1).setValue('Australian HR Institute · live aggregation across all responses');
  sheet.getRange(2, 1, 1, 5).merge().setFontStyle('italic').setFontColor('#666');

  // KPI row
  var totalRow = 3 + QUESTIONS.length;
  var kpis = [
    { label: 'Respondents',          formula: '=IFERROR(COUNTA(' + refR + '!A2:A),0)' },
    { label: 'Sub-SORs assessed',    formula: '=27' },
    { label: 'Criteria per sub-SOR', formula: '=14' },
    { label: 'Top Primary SOR',            formula: '=IFERROR(INDEX(' + refE + '!B2:G2,MATCH(MAX(' + refE + '!B' + totalRow + ':G' + totalRow + '),' + refE + '!B' + totalRow + ':G' + totalRow + ',0)),"-")' },
    { label: 'Highest avg total',    formula: '=IFERROR(MAX(' + refK + '!D4:D30),"-")' }
  ];

  kpis.forEach(function(kpi, i) {
    sheet.getRange(4, 1 + i).setValue(kpi.label);
    sheet.getRange(5, 1 + i).setFormula(kpi.formula);
  });

  sheet.getRange(4, 1, 1, kpis.length)
       .setBackground(C_INK_HEADER).setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(5, 1, 1, kpis.length)
       .setBackground(C_GOLD_LT).setFontColor(C_GOLD)
       .setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  sheet.setRowHeight(4, 28);
  sheet.setRowHeight(5, 48);

  // Top 5 sub-SORs
  sheet.getRange(7, 1).setValue('TOP 5 SUB-SORs');
  sheet.getRange(7, 1, 1, 3).merge()
       .setBackground(C_INK).setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(11);

  sheet.getRange(8, 1, 1, 3).setValues([['Rank', 'Sub-SOR', 'Avg Total /98']]);
  sheet.getRange(8, 1, 1, 3).setBackground('#444').setFontColor('#fff').setFontWeight('bold');

  for (var i = 0; i < 5; i++) {
    var rank = i + 1;
    var row  = 9 + i;
    sheet.getRange(row, 1).setValue(rank);
    sheet.getRange(row, 2).setFormula(
      '=IFERROR(INDEX(' + refK + '!C$4:C$30,MATCH(LARGE(' + refK + '!D$4:D$30,' + rank + '),' + refK + '!D$4:D$30,0)),"-")'
    );
    sheet.getRange(row, 3).setFormula(
      '=IFERROR(LARGE(' + refK + '!D$4:D$30,' + rank + '),"-")'
    );
  }
  sheet.getRange(9, 1, 5, 1).setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange(9, 2, 5, 1).setWrap(true).setFontSize(10);
  sheet.getRange(9, 3, 5, 1).setHorizontalAlignment('center').setNumberFormat('0.0').setFontWeight('bold');

  // Bottom 5 sub-SORs
  sheet.getRange(7, 4).setValue('BOTTOM 5 SUB-SORs');
  sheet.getRange(7, 4, 1, 2).merge()
       .setBackground(C_INK).setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(11);

  sheet.getRange(8, 4, 1, 2).setValues([['Sub-SOR', 'Avg Total /98']]);
  sheet.getRange(8, 4, 1, 2).setBackground('#444').setFontColor('#fff').setFontWeight('bold');

  for (var j = 0; j < 5; j++) {
    var smallRank = j + 1;
    var row2      = 9 + j;
    sheet.getRange(row2, 4).setFormula(
      '=IFERROR(INDEX(' + refK + '!C$4:C$30,MATCH(SMALL(' + refK + '!D$4:D$30,' + smallRank + '),' + refK + '!D$4:D$30,0)),"-")'
    );
    sheet.getRange(row2, 5).setFormula(
      '=IFERROR(SMALL(' + refK + '!D$4:D$30,' + smallRank + '),"-")'
    );
  }
  sheet.getRange(9, 4, 5, 1).setWrap(true).setFontSize(10);
  sheet.getRange(9, 5, 5, 1).setHorizontalAlignment('center').setNumberFormat('0.0').setFontWeight('bold');

  // Primary SOR leaderboard
  sheet.getRange(15, 1).setValue('Primary SOR LEADERBOARD');
  sheet.getRange(15, 1, 1, 5).merge()
       .setBackground(C_INK).setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(11);

  sheet.getRange(16, 1, 1, 3).setValues([['Primary SOR', 'Avg Total /98', 'Avg per Criterion']]);
  sheet.getRange(16, 1, 1, 3).setBackground('#444').setFontColor('#fff').setFontWeight('bold');

  var palette = ['#fbe9e7', '#f3e5f5', '#e3f2fd', '#e8f5e9', '#fff9c4', '#fce4ec'];
  ESORS.forEach(function(esor, ei) {
    var row    = 17 + ei;
    var totRef = refE + '!' + colLetter(2 + ei) + (3 + QUESTIONS.length);
    sheet.getRange(row, 1).setValue('Primary SOR ' + esor.id + ': ' + esor.title);
    sheet.getRange(row, 1).setBackground(palette[ei]);
    sheet.getRange(row, 2).setFormula('=IFERROR(' + totRef + ',"-")');
    sheet.getRange(row, 3).setFormula('=IFERROR(B' + row + '/14,"-")');
  });

  sheet.getRange(17, 2, ESORS.length, 1).setHorizontalAlignment('center').setNumberFormat('0.0').setFontWeight('bold');
  sheet.getRange(17, 3, ESORS.length, 1).setHorizontalAlignment('center').setNumberFormat('0.00');

  var leadRange = sheet.getRange(17, 2, ESORS.length, 1);
  var rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint('#2e7d32').setGradientMinpoint('#c62828')
    .setRanges([leadRange]).build());
  sheet.setConditionalFormatRules(rules);

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 380);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 380);
  sheet.setColumnWidth(5, 130);

  sheet.setHiddenGridlines(true);

  Logger.log('Dashboard sheet configured.');
}

// ── EMAIL CONFIRMATION ───────────────────────────────────────────────────────

function sendConfirmationEmail(payload, answers) {
  if (!payload.email) return;

  try {
    var name = payload.name || 'Survey Respondent';
    var html = buildEmailHtml(name, answers);

    MailApp.sendEmail({
      to:       payload.email,
      subject:  'Your AHRI SOR Prioritisation Survey responses',
      htmlBody: html,
      name:     'AHRI SOR Prioritisation Survey'
    });
    Logger.log('Confirmation email sent to ' + payload.email);
  } catch (e) {
    Logger.log('Email send failed for ' + payload.email + ': ' + e.message);
  }
}

function buildEmailHtml(name, answers) {
  var safeName = escapeHtml(name);

  // Build sub-SOR totals
  var subSorScores = [];
  ESORS.forEach(function(esor) {
    esor.subSors.forEach(function(subSor) {
      var total = 0;
      var complete = true;
      QUESTIONS.forEach(function(q) {
        var v = answers[makeKey(esor.id, subSor.idx, q.id)];
        if (v !== undefined && v !== '') total += parseInt(v, 10);
        else complete = false;
      });
      subSorScores.push({ esor: esor, subSor: subSor, total: total, complete: complete });
    });
  });

  var top5 = subSorScores.filter(function(s) { return s.complete; })
                .sort(function(a, b) { return b.total - a.total; })
                .slice(0, 5);

  var palette = ['#fbe9e7', '#f3e5f5', '#e3f2fd', '#e8f5e9', '#fff9c4', '#fce4ec'];

  var html = '' +
    '<div style="font-family:Arial,Helvetica,sans-serif;color:#1a1a1a;line-height:1.5;max-width:680px;margin:0 auto;">' +

    // Banner
    '<div style="background:#a44f43;color:#ffffff;padding:28px 24px;text-align:center;">' +
      '<div style="font-size:11px;letter-spacing:0.18em;text-transform:uppercase;opacity:0.85;margin-bottom:8px;">Australian HR Institute</div>' +
      '<h1 style="margin:0;font-size:22px;font-weight:600;">Thank you, ' + safeName + '</h1>' +
      '<p style="margin:8px 0 0;font-size:13px;opacity:0.9;">SOR Prioritisation Survey · Response recorded</p>' +
    '</div>' +

    // Intro
    '<div style="background:#ffffff;padding:24px;border:1px solid #d4d4d4;border-top:none;font-size:14px;color:#3d3d3d;">' +
      '<p style="margin:0 0 12px;">Thank you for completing the AHRI SOR Prioritisation Survey. Your responses help inform AHRI\'s strategic direction.</p>' +
      '<p style="margin:0;">Below is a copy of your responses for your records. All individual responses are aggregated and never shared.</p>' +
    '</div>';

  // Top 5
  if (top5.length > 0) {
    html += '' +
      '<div style="background:#f5e0dd;padding:22px 24px;border:1px solid #d4d4d4;border-top:none;">' +
        '<div style="font-size:11px;letter-spacing:0.18em;text-transform:uppercase;color:#a44f43;font-weight:700;margin-bottom:14px;">Your Top 5 Sub-SORs</div>' +
        '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    top5.forEach(function(item, i) {
      html += '' +
        '<tr>' +
          '<td style="padding:8px 0;width:32px;color:#a44f43;font-weight:700;font-size:15px;vertical-align:top;">' + (i + 1) + '</td>' +
          '<td style="padding:8px 8px;color:#1a1a1a;vertical-align:top;">' + escapeHtml(item.subSor.label) +
            '<div style="font-size:11px;color:#888;margin-top:2px;">Primary SOR ' + item.esor.id + ': ' + escapeHtml(item.esor.title) + '</div>' +
          '</td>' +
          '<td style="padding:8px 0;text-align:right;font-weight:700;color:#a44f43;font-size:14px;white-space:nowrap;vertical-align:top;">' + item.total + ' / 98</td>' +
        '</tr>';
    });
    html += '</table></div>';
  }

  // Per-Primary SOR detailed breakdown
  ESORS.forEach(function(esor, ei) {
    html += '' +
      '<div style="background:#ffffff;padding:22px 24px;border:1px solid #d4d4d4;border-top:none;">' +
        '<div style="border-left:4px solid ' + palette[ei] + ';padding-left:12px;margin-bottom:18px;">' +
          '<div style="font-size:10px;letter-spacing:0.18em;text-transform:uppercase;color:#a44f43;font-weight:600;">Primary SOR ' + esor.id + '</div>' +
          '<div style="font-size:16px;font-weight:700;color:#1a1a1a;">' + escapeHtml(esor.title) + '</div>' +
        '</div>';

    esor.subSors.forEach(function(subSor) {
      var total = 0;
      var hasAny = false;
      QUESTIONS.forEach(function(q) {
        var v = answers[makeKey(esor.id, subSor.idx, q.id)];
        if (v !== undefined && v !== '') { total += parseInt(v, 10); hasAny = true; }
      });

      html += '' +
        '<div style="margin-bottom:18px;">' +
          '<table style="width:100%;border-collapse:collapse;margin-bottom:6px;">' +
            '<tr>' +
              '<td style="font-size:13px;font-weight:700;color:#1a1a1a;padding:0;">' + escapeHtml(subSor.label) + '</td>' +
              '<td style="font-size:13px;font-weight:700;color:#a44f43;text-align:right;white-space:nowrap;padding:0 0 0 12px;">' + (hasAny ? total + ' / 98' : '—') + '</td>' +
            '</tr>' +
          '</table>' +
          '<table style="width:100%;border-collapse:collapse;font-size:12px;">';

      QUESTIONS.forEach(function(q, qi) {
        var v = answers[makeKey(esor.id, subSor.idx, q.id)];
        var displayV = (v !== undefined && v !== '') ? v : '—';
        var bgColor  = qi % 2 === 0 ? '#fafafa' : '#ffffff';
        var scoreColor = '#1a1a1a';
        if (v !== undefined && v !== '') {
          var n = parseInt(v, 10);
          if (n <= 2)      scoreColor = '#c0392b';
          else if (n >= 6) scoreColor = '#27ae60';
        }
        html += '' +
          '<tr>' +
            '<td style="padding:5px 10px;background:' + bgColor + ';color:#3d3d3d;border-left:2px solid #f0f0f0;">' + escapeHtml(q.label) + '</td>' +
            '<td style="padding:5px 10px;background:' + bgColor + ';text-align:right;font-weight:700;color:' + scoreColor + ';width:40px;">' + displayV + '</td>' +
          '</tr>';
      });
      html += '</table></div>';
    });
    html += '</div>';
  });

  // Footer
  html += '' +
    '<div style="background:#fafafa;padding:20px 24px;border:1px solid #d4d4d4;border-top:none;font-size:11px;color:#666;line-height:1.6;">' +
      '<p style="margin:0 0 6px;"><strong>Scoring:</strong> Each criterion is rated 1 (low alignment) to 7 (high alignment). Maximum total per sub-SOR is 98 (14 criteria × 7).</p>' +
      '<p style="margin:0 0 6px;">All responses are recorded in aggregate and reported anonymously. Individual scores are never shared externally.</p>' +
      '<p style="margin:0;">If you have any questions, please contact AHRI.</p>' +
    '</div>' +
  '</div>';

  return html;
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

// Run this manually once to authorise MailApp and verify email rendering.
// Replace the address with your own before running.
function testEmail() {
  var fakeAnswers = {};
  ESORS.forEach(function(esor) {
    esor.subSors.forEach(function(subSor) {
      QUESTIONS.forEach(function(q) {
        fakeAnswers[makeKey(esor.id, subSor.idx, q.id)] = Math.floor(Math.random() * 7) + 1;
      });
    });
  });
  sendConfirmationEmail(
    { name: 'Test User', email: Session.getActiveUser().getEmail() },
    fakeAnswers
  );
}

// ── WEB APP HANDLERS ──────────────────────────────────────────────────────────

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    if (payload.surveyType !== 'AHRI-SOR-Prioritisation-v2') {
      return jsonResponse({ status: 'ignored', reason: 'unknown surveyType' });
    }

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(RESPONSES_SHEET);
    if (!sheet) { setupSheets(); sheet = ss.getSheetByName(RESPONSES_SHEET); }

    var answers = payload.answers || {};
    var keys    = keyOrder();

    var timestamp = Utilities.formatDate(
      payload.submittedAt ? new Date(payload.submittedAt) : new Date(),
      TIMEZONE, 'dd/MM/yyyy HH:mm:ss'
    );

    var ratingValues = keys.map(function(key) {
      var v = answers[key];
      return (v !== undefined && v !== '') ? parseInt(v, 10) : '';
    });

    var subtotals = [];
    ESORS.forEach(function(esor) {
      esor.subSors.forEach(function(subSor) {
        var ssKeys = QUESTIONS.map(function(q) {
          return makeKey(esor.id, subSor.idx, q.id);
        });
        var vals = ssKeys.map(function(k) { return answers[k]; })
                        .filter(function(v) { return v !== undefined && v !== ''; });
        subtotals.push(vals.length === QUESTIONS.length
          ? vals.reduce(function(sum, v) { return sum + parseInt(v, 10); }, 0)
          : '');
      });
    });

    var row = [timestamp, payload.name || '', payload.email || '']
                .concat(ratingValues)
                .concat(subtotals);
    sheet.appendRow(row);

    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 4, 1, ratingValues.length + subtotals.length)
         .setHorizontalAlignment('center');

    // Send confirmation email (failures are logged but don't break the submission)
    sendConfirmationEmail(payload, answers);

    return jsonResponse({ status: 'ok' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function doGet(e) {
  return jsonResponse({ status: 'ok', message: 'AHRI SOR Prioritisation endpoint active' });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── UTILITY ─────────────────────────────────────────────────────────────────
// Run manually any time you want a fresh ranking sort. Auto-sort via formulas is
// awkward because ranking values are themselves formulas referencing Responses.

function sortRankingSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RANKING_SHEET);
  if (!sheet) return;
  var n = flatSubSors().length;
  sheet.getRange(4, 1, n, 6).sort({ column: 4, ascending: false });
  Logger.log('Ranking sheet sorted by Avg Total descending.');
}
