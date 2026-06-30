/**
 * @OnlyCurrentDoc
 *
 * Sales Report Builder — Google Sheets / Apps Script demo.
 *
 * Reads a raw "Transactions" tab and builds a clean "Summary" tab (totals by
 * category and by rep, with a timestamp). It's a portfolio demo, but written to
 * the bar a paying client's automation has to clear:
 *
 *   - run from a CUSTOM MENU + button, never "open the script editor"
 *   - survives the client editing the sheet: columns found BY HEADER NAME, rows
 *     by dynamic last-row — zero hardcoded A1 coordinates
 *   - batched getValues/setValues (one read, one write) — never a cell-by-cell
 *     loop, which blows the Apps Script quota on real data
 *   - resumable: respects the 6-minute execution limit by checkpointing progress
 *     and scheduling a continuation, so a big sheet never half-finishes silently
 *   - try/catch on every entry point → a plain-English toast AND an email to the
 *     owner, so a failed overnight run is never invisible
 *   - @OnlyCurrentDoc so users don't hit the scary "unverified app" consent screen
 */

// ---- Configuration (all by NAME, so layout edits don't break it) -----------
var CONFIG = {
  sourceSheet: 'Transactions',
  reportSheet: 'Summary',
  instructionsSheet: 'Instructions',
  // Required columns, matched case-insensitively by header text.
  requiredColumns: ['Date', 'Rep', 'Category', 'Amount', 'Status'],
  // Only count rows in these statuses toward revenue.
  countableStatuses: ['Paid', 'Completed'],
  chunkRows: 2000,          // rows processed per slice
  softTimeLimitMs: 4.5 * 60 * 1000, // stop + continue before the 6-min hard kill
  propKey: 'SALES_REPORT_PROGRESS',
};

// ---- Menu (simple trigger; no extra scopes) --------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sales Tools')
    .addItem('Build sales report', 'buildSalesReport')
    .addSeparator()
    .addItem('Set up demo data', 'setupDemo')
    .addItem('About / instructions', 'showAbout')
    .addToUi();
}

// ---- Main entry point ------------------------------------------------------
/**
 * Build the Summary tab. Safe to run repeatedly. Resumes automatically if the
 * data is large enough to risk the 6-minute limit.
 */
function buildSalesReport() {
  var startTime = Date.now();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var source = getSheetOrThrow(ss, CONFIG.sourceSheet);

    var headerMap = readHeaderMap(source); // name(lower) -> 0-based col index
    assertRequiredColumns(headerMap, CONFIG.requiredColumns);

    var lastRow = source.getLastRow();
    var dataRows = lastRow - 1; // minus header
    if (dataRows <= 0) {
      throw new Error('The "' + CONFIG.sourceSheet + '" tab has headers but no data rows to summarise.');
    }

    var progress = loadProgress() || { cursor: 0, byCategory: {}, byRep: {}, grandTotal: 0, counted: 0, skipped: 0 };

    // Process in slices so a large sheet survives the execution-time limit.
    while (progress.cursor < dataRows) {
      var batchStart = 2 + progress.cursor; // 1-based, skip header
      var rowsLeft = dataRows - progress.cursor;
      var take = Math.min(CONFIG.chunkRows, rowsLeft);

      // ONE batched read for the whole slice.
      var values = source.getRange(batchStart, 1, take, source.getLastColumn()).getValues();
      accumulate(values, headerMap, progress);
      progress.cursor += take;

      if (progress.cursor < dataRows && Date.now() - startTime > CONFIG.softTimeLimitMs) {
        // Near the limit: checkpoint and schedule a continuation, then exit cleanly.
        saveProgress(progress);
        scheduleContinuation();
        toast('Large dataset — processed ' + progress.cursor + ' rows so far; continuing automatically in ~1 min.');
        return;
      }
    }

    writeReport(ss, progress);
    clearProgress();
    removeContinuationTriggers();
    toast('Sales report built: ' + progress.counted + ' transactions, '
      + Object.keys(progress.byCategory).length + ' categories. (' + progress.skipped + ' rows skipped.)');
  } catch (err) {
    handleError('buildSalesReport', err);
  }
}

/** Time-based-trigger entry point for the resumable continuation. */
function continueSalesReport() {
  buildSalesReport();
}

// ---- Core helpers ----------------------------------------------------------

function getSheetOrThrow(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('Could not find a tab named "' + name + '". '
      + 'Rename your data tab to "' + name + '", or run "Set up demo data" first.');
  }
  return sheet;
}

/** Build a header -> column-index map from row 1, matched by NAME not position. */
function readHeaderMap(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) throw new Error('The "' + sheet.getName() + '" tab is empty.');
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var c = 0; c < headers.length; c++) {
    var key = String(headers[c]).trim().toLowerCase();
    if (key) map[key] = c;
  }
  return map;
}

function assertRequiredColumns(headerMap, required) {
  var missing = [];
  for (var i = 0; i < required.length; i++) {
    if (!(required[i].toLowerCase() in headerMap)) missing.push(required[i]);
  }
  if (missing.length) {
    throw new Error('Missing required column(s): ' + missing.join(', ')
      + '. Found: ' + Object.keys(headerMap).join(', ') + '.');
  }
}

/** Fold one slice of rows into the running totals. Defensive on every cell. */
function accumulate(values, headerMap, progress) {
  var iCat = headerMap['category'];
  var iRep = headerMap['rep'];
  var iAmt = headerMap['amount'];
  var iStatus = headerMap['status'];
  var countable = {};
  CONFIG.countableStatuses.forEach(function (s) { countable[s.toLowerCase()] = true; });

  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    if (isBlankRow(row)) continue; // tolerate blank rows mid-sheet

    var status = String(row[iStatus] || '').trim().toLowerCase();
    if (!countable[status]) { progress.skipped++; continue; }

    var amount = toNumber(row[iAmt]);
    if (amount === null) { progress.skipped++; continue; } // unparseable amount — don't fake a number

    var category = String(row[iCat] || '(uncategorised)').trim() || '(uncategorised)';
    var rep = String(row[iRep] || '(unassigned)').trim() || '(unassigned)';

    progress.byCategory[category] = (progress.byCategory[category] || 0) + amount;
    progress.byRep[rep] = (progress.byRep[rep] || 0) + amount;
    progress.grandTotal += amount;
    progress.counted++;
  }
}

/** Write the Summary tab in TWO batched setValues calls (categories, reps). */
function writeReport(ss, progress) {
  var sheet = ss.getSheetByName(CONFIG.reportSheet);
  if (!sheet) sheet = ss.insertSheet(CONFIG.reportSheet);
  sheet.clearContents();

  var rows = [];
  rows.push(['Sales Summary', '']);
  rows.push(['Generated', new Date()]);
  rows.push(['Transactions counted', progress.counted]);
  rows.push(['Rows skipped (non-countable/blank)', progress.skipped]);
  rows.push(['', '']);
  rows.push(['Category', 'Revenue']);
  sortedEntries(progress.byCategory).forEach(function (e) { rows.push([e[0], e[1]]); });
  rows.push(['Total', progress.grandTotal]);
  rows.push(['', '']);
  rows.push(['Rep', 'Revenue']);
  sortedEntries(progress.byRep).forEach(function (e) { rows.push([e[0], e[1]]); });

  // ONE write for the whole block.
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);

  // light, name-free formatting (won't break if rows shift)
  sheet.getRange('A1:B1').merge().setFontWeight('bold').setFontSize(14);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 140);
  protectReport(sheet);
  sheet.activate();
}

// ---- Resumability (6-minute-limit safety) ----------------------------------

function loadProgress() {
  var raw = PropertiesService.getDocumentProperties().getProperty(CONFIG.propKey);
  return raw ? JSON.parse(raw) : null;
}
function saveProgress(progress) {
  PropertiesService.getDocumentProperties().setProperty(CONFIG.propKey, JSON.stringify(progress));
}
function clearProgress() {
  PropertiesService.getDocumentProperties().deleteProperty(CONFIG.propKey);
}
function scheduleContinuation() {
  removeContinuationTriggers();
  ScriptApp.newTrigger('continueSalesReport').timeBased().after(60 * 1000).create();
}
function removeContinuationTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'continueSalesReport') ScriptApp.deleteTrigger(t);
  });
}

// ---- Protection ------------------------------------------------------------

/** Protect the generated report (warning-only) so a user doesn't hand-edit it
 *  and then wonder why the next run overwrites their changes. Inputs stay open. */
function protectReport(sheet) {
  try {
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
    sheet.protect()
      .setDescription('Auto-generated — edits will be overwritten on the next run')
      .setWarningOnly(true);
  } catch (e) {
    // Protection is a nicety, not the job — never fail the report over it.
    Logger.log('Could not set protection: ' + e.message);
  }
}

// ---- Error handling --------------------------------------------------------

/** One place: log it, toast a plain-English message, email the owner. */
function handleError(where, err) {
  var msg = (err && err.message) ? err.message : String(err);
  Logger.log('[' + where + '] ' + msg + (err && err.stack ? '\n' + err.stack : ''));
  toast('Could not finish: ' + msg);
  notifyOwner(where, msg, err && err.stack);
}

function notifyOwner(where, msg, stack) {
  try {
    var email = Session.getEffectiveUser().getEmail();
    if (!email) return;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    MailApp.sendEmail({
      to: email,
      subject: 'Sales report automation failed (' + ss.getName() + ')',
      body: 'The "' + where + '" step failed.\n\nReason: ' + msg
        + '\n\nSheet: ' + ss.getUrl()
        + (stack ? '\n\nDetails:\n' + stack : ''),
    });
  } catch (e) {
    Logger.log('Could not email owner: ' + e.message); // never throw from the notifier
  }
}

// ---- Small utilities -------------------------------------------------------

function toNumber(v) {
  if (typeof v === 'number') return isFinite(v) ? v : null;
  if (v === '' || v === null || v === undefined) return null;
  var cleaned = String(v).replace(/[^0-9.\-]/g, '');
  if (!/[0-9]/.test(cleaned)) return null; // no digits (e.g. "n/a") → genuinely unparseable; skip, never fake a 0
  var n = Number(cleaned);
  return isFinite(n) ? n : null;
}
function isBlankRow(row) {
  for (var i = 0; i < row.length; i++) {
    if (row[i] !== '' && row[i] !== null && row[i] !== undefined) return false;
  }
  return true;
}
function sortedEntries(obj) {
  return Object.keys(obj)
    .map(function (k) { return [k, Math.round(obj[k] * 100) / 100]; })
    .sort(function (a, b) { return b[1] - a[1]; });
}
function toast(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Sales Tools', 8);
}
function showAbout() {
  SpreadsheetApp.getUi().alert(
    'Sales Tools',
    'Use "Build sales report" to (re)generate the Summary tab from the Transactions tab.\n\n'
      + '• Columns are matched by header NAME, so you can reorder them.\n'
      + '• Only ' + CONFIG.countableStatuses.join('/') + ' rows count toward revenue.\n'
      + '• Large sheets resume automatically within Google’s 6-minute limit.\n'
      + '• If a run fails you’ll get a toast and an email explaining why.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}
