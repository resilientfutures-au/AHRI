// AHRI SOR Prioritisation Survey — Google Apps Script
//
// SETUP:
//   1. Open your destination Google Spreadsheet.
//   2. Go to Extensions > Apps Script, paste this file.
//   3. Run setupSheets() once to create the Responses and Summary sheets.
//   4. Deploy as Web App: Execute as "Me", access "Anyone".
//   5. Copy the deployment URL into the survey index.html (replace APPS_SCRIPT_URL).
//
// MAINTENANCE:
//   • Run setupSheets() again after adding new respondent columns — it is safe to re-run.
//   • The Summary sheet auto-calculates from Responses via AVERAGE formulas.
//   • Scores are stored as plain integers (1–7); empty = not answered.

// ── CONFIGURATION ─────────────────────────────────────────────────────────────

var RESPONSES_SHEET = 'Responses';
var SUMMARY_SHEET   = 'Summary';
var TIMEZONE        = 'Australia/Melbourne';

var ESORS = [
  { id: 1, title: 'Scope of Play' },
  { id: 2, title: 'Certification and Professional Pathway' },
  { id: 3, title: 'AI and Workforce Transformation' },
  { id: 4, title: 'Advocacy and Strategic Voice' },
  { id: 5, title: 'Platform / Ecosystem Model' },
  { id: 6, title: 'Membership & Member Economics' }
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

// ── COLUMN MAP ────────────────────────────────────────────────────────────────
// Col 1          : Timestamp
// Cols 2–15      : ESOR 1 Q1–Q14
// Cols 16–29     : ESOR 2 Q1–Q14
// Cols 30–43     : ESOR 3 Q1–Q14
// Cols 44–57     : ESOR 4 Q1–Q14
// Cols 58–71     : ESOR 5 Q1–Q14
// Cols 72–85     : ESOR 6 Q1–Q14
// Cols 86–91     : ESOR 1–6 Subtotals (sum of 14 scores per ESOR)

function keyOrder() {
  var keys = [];
  ESORS.forEach(function(esor) {
    QUESTIONS.forEach(function(q) {
      keys.push('esor' + esor.id + '_q' + q.id);
    });
  });
  return keys;
}

function responseHeaders() {
  var headers = ['Timestamp'];
  ESORS.forEach(function(esor) {
    QUESTIONS.forEach(function(q) {
      headers.push('ESOR ' + esor.id + ': ' + esor.title + ' — Q' + q.id + ': ' + q.label);
    });
  });
  ESORS.forEach(function(esor) {
    headers.push('ESOR ' + esor.id + ' Total (/98)');
  });
  return headers;
}

// ── SETUP ─────────────────────────────────────────────────────────────────────

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  setupResponsesSheet(ss);
  setupSummarySheet(ss);

  Logger.log('Setup complete.');
}

function setupResponsesSheet(ss) {
  var sheet = ss.getSheetByName(RESPONSES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RESPONSES_SHEET);
    Logger.log('Created: ' + RESPONSES_SHEET);
  }

  var headers = responseHeaders();

  // Only write headers if the sheet is empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    // Re-apply headers to row 1 in case of changes
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  // Format header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1a1a1a');
  headerRange.setFontColor('#ffffff');
  headerRange.setWrap(true);
  sheet.setFrozenRows(1);

  // Column widths
  sheet.setColumnWidth(1, 160);   // Timestamp

  // ESOR question columns: narrow (they contain numeric 1–7 scores)
  for (var e = 0; e < ESORS.length; e++) {
    var startCol = 2 + (e * QUESTIONS.length);
    sheet.setColumnWidths(startCol, QUESTIONS.length, 80);
  }

  // Subtotal columns
  var totalStartCol = 2 + ESORS.length * QUESTIONS.length;
  sheet.setColumnWidths(totalStartCol, ESORS.length, 110);

  // Colour-band ESOR column groups for readability
  var bandColors = ['#fff9f8', '#f5f0ef', '#fff9f8', '#f5f0ef', '#fff9f8', '#f5f0ef'];
  for (var e = 0; e < ESORS.length; e++) {
    var startCol = 2 + (e * QUESTIONS.length);
    var headerCell = sheet.getRange(1, startCol, 1, QUESTIONS.length);
    headerCell.setBackground(e % 2 === 0 ? '#3d3d3d' : '#555555');
  }

  Logger.log('Responses sheet configured.');
}

function setupSummarySheet(ss) {
  var sheet = ss.getSheetByName(SUMMARY_SHEET);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(SUMMARY_SHEET);
    Logger.log('Created: ' + SUMMARY_SHEET);
  }

  // ── Build the summary layout ──────────────────────────────────────────────
  //
  //   Row 1 : (blank)  | ESOR 1 | ESOR 2 | … | ESOR 6 | (blank)
  //   Row 2 : Criterion| ESOR 1 | ESOR 2 | … | ESOR 6 | (blank)
  //   Row 3–16 : Q label | avg  | avg    | … | avg    |
  //   Row 17: TOTAL    | sum   | sum    | … | sum    |
  //   Row 18: (blank)
  //   Row 19: Respondents | count
  //
  // The AVERAGE formulas point at the Responses sheet data columns.
  // They use IFERROR so empty cells (unanswered) are ignored gracefully.

  var responsesRef = "'" + RESPONSES_SHEET + "'";

  // Helper: column number → A1-style letter (1 = A, 27 = AA, etc.)
  function colLetter(n) {
    var s = '';
    while (n > 0) {
      var r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  // Responses sheet column positions (1-indexed)
  // Col 1 = Timestamp, Cols 2–85 = scores, Cols 86–91 = totals
  function responseCol(esorIdx, qIdx) {
    return 1 + (esorIdx * QUESTIONS.length) + qIdx + 1;
    // esorIdx: 0-based; qIdx: 0-based
  }

  // ── Write header rows ─────────────────────────────────────────────────────

  var headerRow1 = [''];
  var headerRow2 = ['Criterion'];
  ESORS.forEach(function(esor) {
    headerRow1.push('ESOR ' + esor.id + ': ' + esor.title);
    headerRow2.push('ESOR ' + esor.id);
  });
  headerRow1.push('');
  headerRow2.push('');

  sheet.getRange(1, 1, 1, headerRow1.length).setValues([headerRow1]);
  sheet.getRange(2, 1, 1, headerRow2.length).setValues([headerRow2]);

  // Merge ESOR title cells across header row 1 (purely cosmetic — skip if it causes issues)
  // We'll skip auto-merge to keep the script simple and error-free.

  // ── Write criterion rows with AVERAGE formulas ────────────────────────────

  for (var qi = 0; qi < QUESTIONS.length; qi++) {
    var row = [QUESTIONS[qi].label];
    for (var ei = 0; ei < ESORS.length; ei++) {
      var col = responseCol(ei, qi);
      var colL = colLetter(col);
      // AVERAGEIF-style: only count rows where this column has a numeric value
      var formula = '=IFERROR(AVERAGEIF(' + responsesRef + '!' + colL + ':' + colL + ',"<>",' + responsesRef + '!' + colL + ':' + colL + '),"-")';
      row.push(formula);
    }
    row.push(''); // blank trailing column
    sheet.getRange(3 + qi, 1, 1, row.length).setValues([row]);
  }

  // ── Total row (sum of averages per ESOR) ──────────────────────────────────

  var totalRow = ['Average Total'];
  for (var ei = 0; ei < ESORS.length; ei++) {
    // SUM of the 14 average cells in this ESOR column
    var summaryCol = colLetter(2 + ei); // Summary sheet col B = ESOR1, C = ESOR2, etc.
    var firstDataRow = 3;
    var lastDataRow  = 2 + QUESTIONS.length;
    var formula = '=IFERROR(SUM(' + summaryCol + firstDataRow + ':' + summaryCol + lastDataRow + '),"-")';
    totalRow.push(formula);
  }
  totalRow.push('');
  sheet.getRange(3 + QUESTIONS.length, 1, 1, totalRow.length).setValues([totalRow]);

  // ── Respondent count ──────────────────────────────────────────────────────

  sheet.getRange(3 + QUESTIONS.length + 2, 1).setValue('Respondents');
  // Count non-empty rows in the Responses sheet (excluding header)
  sheet.getRange(3 + QUESTIONS.length + 2, 2).setFormula(
    '=IFERROR(COUNTA(' + responsesRef + '!A2:A)-1,0)'
  );

  // ── Formatting ────────────────────────────────────────────────────────────

  var totalCols = 1 + ESORS.length + 1; // criterion label + 6 ESORs + trailing blank

  // Title header row (row 1)
  var titleRange = sheet.getRange(1, 1, 1, totalCols);
  titleRange.setFontWeight('bold');
  titleRange.setBackground('#a44f43');
  titleRange.setFontColor('#ffffff');
  titleRange.setWrap(true);

  // Sub-header row (row 2)
  var subHeaderRange = sheet.getRange(2, 1, 1, totalCols);
  subHeaderRange.setFontWeight('bold');
  subHeaderRange.setBackground('#1a1a1a');
  subHeaderRange.setFontColor('#ffffff');

  // Criterion label column (col A, rows 3 onward)
  var labelRange = sheet.getRange(3, 1, QUESTIONS.length + 1, 1);
  labelRange.setFontWeight('bold');
  labelRange.setFontColor('#1a1a1a');

  // Data cells: centre-align and apply number format
  var dataRange = sheet.getRange(3, 2, QUESTIONS.length, ESORS.length);
  dataRange.setHorizontalAlignment('center');
  dataRange.setNumberFormat('0.00');

  // Total row
  var totalRowRange = sheet.getRange(3 + QUESTIONS.length, 1, 1, totalCols);
  totalRowRange.setFontWeight('bold');
  totalRowRange.setBackground('#f5f0ef');
  totalRowRange.setFontColor('#a44f43');
  var totalDataRange = sheet.getRange(3 + QUESTIONS.length, 2, 1, ESORS.length);
  totalDataRange.setHorizontalAlignment('center');
  totalDataRange.setNumberFormat('0.00');

  // Alternating row shading on criterion rows
  for (var qi = 0; qi < QUESTIONS.length; qi++) {
    if (qi % 2 === 0) {
      sheet.getRange(3 + qi, 1, 1, totalCols).setBackground('#fafafa');
    }
  }

  // Column widths
  sheet.setColumnWidth(1, 240);    // Criterion label
  for (var ei = 0; ei < ESORS.length; ei++) {
    sheet.setColumnWidth(2 + ei, 130);
  }

  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1);

  Logger.log('Summary sheet configured.');
}

// ── WEB APP HANDLERS ──────────────────────────────────────────────────────────

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    if (payload.surveyType !== 'AHRI-SOR-Prioritisation') {
      return jsonResponse({ status: 'ignored', reason: 'unknown surveyType' });
    }

    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var sheet   = ss.getSheetByName(RESPONSES_SHEET);

    // Auto-create sheet if missing (e.g. first submission before setup was run)
    if (!sheet) {
      setupSheets();
      sheet = ss.getSheetByName(RESPONSES_SHEET);
    }

    var answers = payload.answers || {};
    var keys    = keyOrder();

    var timestamp = payload.submittedAt
      ? Utilities.formatDate(new Date(payload.submittedAt), TIMEZONE, 'dd/MM/yyyy HH:mm:ss')
      : Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy HH:mm:ss');

    // Rating columns (1–7 integer, or blank if unanswered)
    var ratingValues = keys.map(function(key) {
      var v = answers[key];
      return (v !== undefined && v !== '') ? parseInt(v, 10) : '';
    });

    // Subtotals per ESOR (sum of 14 ratings; blank if any are missing)
    var subtotals = ESORS.map(function(esor) {
      var esorKeys = QUESTIONS.map(function(q) { return 'esor' + esor.id + '_q' + q.id; });
      var vals = esorKeys.map(function(k) { return answers[k]; }).filter(function(v) { return v !== undefined && v !== ''; });
      return vals.length === QUESTIONS.length
        ? vals.reduce(function(sum, v) { return sum + parseInt(v, 10); }, 0)
        : '';
    });

    var row = [timestamp].concat(ratingValues).concat(subtotals);
    sheet.appendRow(row);

    // Style the new data row: centre-align numeric columns
    var lastRow  = sheet.getLastRow();
    var numCols  = ratingValues.length + subtotals.length;
    sheet.getRange(lastRow, 2, 1, numCols).setHorizontalAlignment('center');

    return jsonResponse({ status: 'ok' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function doGet(e) {
  return jsonResponse({ status: 'ok', message: 'AHRI SOR Prioritisation endpoint active' });
}

// ── HELPERS ───────────────────────────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
