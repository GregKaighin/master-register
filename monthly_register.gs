/**
 * Piano Lesson Register - Monthly Sheet Creator
 * Greg Kaighin - gkpianolessons@gmail.com
 *
 * Layout (confirmed from April 2026 register):
 *   Col A        : Row numbers  (1., 2., 3., …)
 *   Col B        : Pupil names  (Surname, Forename)
 *   Col D (col 4): Day 1  —  Col AH (col 34): Day 31  (always 31 columns)
 *   Row 6        : Day numbers
 *   Row 7+       : Pupil data
 *
 * Strategy: always keep 31 day columns. For shorter months, day numbers and
 * lesson data beyond the month end are simply cleared. No columns are ever
 * inserted or deleted, so borders and merged cells are never disturbed.
 *
 * To install: open Extensions > Apps Script, paste this file,
 * then run installMonthlyTrigger() once.
 */

var FIRST_DAY_COL  = 4;   // Column D
var DAY_HEADER_ROW = 6;   // Row containing day numbers
var FIRST_PUPIL_ROW = 7;  // First pupil data row

function createNewMonthSheet(overrideDate) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var today = overrideDate || new Date();

  var newYear  = today.getFullYear();
  var newMonth = today.getMonth();
  var prevMonth = newMonth === 0 ? 11 : newMonth - 1;
  var prevYear  = newMonth === 0 ? newYear - 1 : newYear;

  var newSheetName  = formatSheetName(newMonth, newYear);
  var prevSheetName = formatSheetName(prevMonth, prevYear);

  if (ss.getSheetByName(newSheetName)) {
    Logger.log('Sheet "%s" already exists — skipping.', newSheetName);
    return;
  }

  var prevSheet = ss.getSheetByName(prevSheetName);
  if (!prevSheet) {
    SpreadsheetApp.getUi().alert(
      'Could not find the sheet "' + prevSheetName + '".\n\n' +
      'Please rename the previous month tab to exactly "' + prevSheetName + '".'
    );
    return;
  }

  // ── 1. Copy and position ───────────────────────────────────────────────────
  var newSheet = prevSheet.copyTo(ss);
  newSheet.setName(newSheetName);
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(prevSheet.getIndex()); // place to the left of previous month

  // ── 2. Update month name and year in the header area ──────────────────────
  var oldMonthName = fullMonthName(prevMonth);
  var newMonthName = fullMonthName(newMonth);
  var headerRows   = DAY_HEADER_ROW - 1;
  var searchRange  = newSheet.getRange(1, 1, headerRows, newSheet.getLastColumn());
  var values       = searchRange.getValues();

  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var v = String(values[r][c]);
      if (v === oldMonthName) {
        searchRange.getCell(r + 1, c + 1).setValue(newMonthName);
      } else if (v === String(prevYear)) {
        searchRange.getCell(r + 1, c + 1).setValue(String(newYear));
      }
    }
  }

  // ── 3. Add or remove day columns so the table width matches the new month ──
  var daysInNewMonth  = new Date(newYear, newMonth + 1, 0).getDate();
  var daysInPrevMonth = new Date(prevYear, prevMonth + 1, 0).getDate();
  var oldLastDayCol   = FIRST_DAY_COL + daysInPrevMonth - 1;

  if (daysInNewMonth > daysInPrevMonth) {
    var extra = daysInNewMonth - daysInPrevMonth;
    newSheet.insertColumnsAfter(oldLastDayCol, extra);
    // Carry formatting from the previous last day column into the new column(s)
    newSheet.getRange(1, oldLastDayCol, newSheet.getLastRow(), 1)
            .copyTo(newSheet.getRange(1, oldLastDayCol + 1, newSheet.getLastRow(), extra),
                    SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    // Remove the right border that was on the old last day column
    newSheet.getRange(1, oldLastDayCol, newSheet.getLastRow(), 1)
            .setBorder(null, null, null, false, null, null);
  } else if (daysInPrevMonth > daysInNewMonth) {
    newSheet.deleteColumns(FIRST_DAY_COL + daysInNewMonth, daysInPrevMonth - daysInNewMonth);
  }

  // ── 4. Clear any stray value in the cell immediately left of day 1 ─────────
  newSheet.getRange(DAY_HEADER_ROW, FIRST_DAY_COL - 1).clearContent();

  // ── 5. Rewrite day numbers ────────────────────────────────────────────────
  for (var d = 1; d <= daysInNewMonth; d++) {
    newSheet.getRange(DAY_HEADER_ROW, FIRST_DAY_COL + d - 1).setValue(d);
  }

  // ── 6. Clear lesson data ──────────────────────────────────────────────────
  var lastRow   = newSheet.getLastRow();
  var pupilRows = lastRow - FIRST_PUPIL_ROW + 1;

  if (pupilRows > 0) {
    newSheet.getRange(FIRST_PUPIL_ROW, FIRST_DAY_COL, pupilRows, daysInNewMonth)
            .clearContent();
  }

  Logger.log('Created "%s" (%d days).', newSheetName, daysInNewMonth);
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function formatSheetName(monthIndex, year) {
  var yy = String(year).slice(-2);
  var mm = ('0' + (monthIndex + 1)).slice(-2);
  return yy + '/' + mm;
}

function fullMonthName(monthIndex) {
  return ['January','February','March','April','May','June',
          'July','August','September','October','November','December'][monthIndex];
}

// ── Install trigger (run once) ────────────────────────────────────────────────

function installMonthlyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'createNewMonthSheet'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('createNewMonthSheet')
    .timeBased()
    .onMonthDay(1)
    .atHour(1)
    .create();

  SpreadsheetApp.getUi().alert('Trigger installed. New tabs will be created at 1 AM on the 1st of each month.');
}

// ── Manual test ───────────────────────────────────────────────────────────────

function testCreateNewMonthSheet() {
  var nextMonth = new Date();
  nextMonth.setMonth(nextMonth.getMonth() + 1);
  nextMonth.setDate(1);
  createNewMonthSheet(nextMonth);
  SpreadsheetApp.getUi().alert('Done — check your spreadsheet for the new sheet.');
}
