/**
 * Demo data + instructions generator.
 *
 * Run once ("Sales Tools → Set up demo data") to populate a realistic
 * Transactions tab — including deliberately messy rows (blank line, refunded
 * status, currency-formatted and unparseable amounts) so the report's
 * defensive handling is visible in the before/after, not just claimed.
 */
function setupDemo() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- Transactions (raw input) ---
    var tx = ss.getSheetByName(CONFIG.sourceSheet) || ss.insertSheet(CONFIG.sourceSheet);
    tx.clearContents();
    var rows = [
      ['Date', 'Rep', 'Category', 'Amount', 'Status'],
      ['2026-06-01', 'Alex',  'Hardware',  '£89.00',  'Paid'],
      ['2026-06-01', 'Sam',   'Software',  149,       'Completed'],
      ['2026-06-02', 'Alex',  'Hardware',  39.5,      'Paid'],
      ['2026-06-02', 'Priya', 'Services',  500,       'Pending'],     // not countable
      ['2026-06-03', 'Sam',   'Software',  '1,200.00','Paid'],        // currency-formatted
      ['', '', '', '', ''],                                            // blank row mid-sheet
      ['2026-06-03', 'Priya', 'Hardware',  219,       'Completed'],
      ['2026-06-04', 'Alex',  '',          75,        'Paid'],        // missing category
      ['2026-06-04', 'Sam',   'Services',  'n/a',     'Paid'],        // unparseable amount
      ['2026-06-05', 'Priya', 'Software',  149,       'Refunded'],    // not countable
      ['2026-06-05', 'Alex',  'Services',  320,       'Completed'],
      ['2026-06-06', 'Sam',   'Hardware',  89,        'Paid'],
    ];
    tx.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    tx.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
    tx.setFrozenRows(1);

    // --- Instructions tab (in-sheet, so the buyer never needs the editor) ---
    writeInstructions(ss);

    SpreadsheetApp.getUi().alert(
      'Demo ready',
      'A "Transactions" tab was created with sample data (including a few messy rows on purpose).\n\n'
        + 'Now run "Sales Tools → Build sales report" to generate the Summary.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    handleError('setupDemo', err);
  }
}

function writeInstructions(ss) {
  var sheet = ss.getSheetByName(CONFIG.instructionsSheet) || ss.insertSheet(CONFIG.instructionsSheet, 0);
  sheet.clearContents();
  var lines = [
    ['Sales Report Builder — how to use'],
    [''],
    ['1.  Put your data on a tab named "Transactions" (or run "Set up demo data").'],
    ['2.  It needs these columns (in any order): Date, Rep, Category, Amount, Status.'],
    ['3.  Menu: Sales Tools → Build sales report. The "Summary" tab is created/updated.'],
    [''],
    ['Notes'],
    ['•  Only Paid / Completed rows count toward revenue.'],
    ['•  Columns are matched by name, so you can reorder or recolour them freely.'],
    ['•  Blank rows and bad amounts are skipped (and counted), never silently miscounted.'],
    ['•  Big sheets continue automatically within Google’s 6-minute limit.'],
    ['•  If anything fails you get an on-screen message and an email explaining why.'],
  ];
  sheet.getRange(1, 1, lines.length, 1).setValues(lines);
  sheet.getRange('A1').setFontWeight('bold').setFontSize(14);
  sheet.setColumnWidth(1, 560);
}
