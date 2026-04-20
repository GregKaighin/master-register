/**
 * PIANO LESSONS WITH GREG KAIGHIN — Monthly Register Script
 * ----------------------------------------------------------
 * Automatically creates a new monthly register tab on the 1st of each month.
 *
 * What it does:
 *  - Copies the most recent YY/MM tab as the template
 *  - Names the new tab in YY/MM format (e.g. "26/05")
 *  - Inserts/hides day columns to match the new month's length
 *  - Writes day numbers 1…N for the new month
 *  - Clears all lesson attendance data (keeps names & formatting)
 *  - Fixes the right border on the last day column
 *  - Preserves "REGISTER" label in O2
 *  - Updates the month/year display (AA1 / AA2)
 *  - Inserts the new tab at position 1 (leftmost)
 *  - Deletes the partial tab automatically if anything goes wrong
 *
 * SETUP INSTRUCTIONS
 * ------------------
 * 1. Open your Google Sheet → Extensions → Apps Script.
 * 2. Paste this entire script, replacing any existing code.
 * 3. Click Save (💾).
 * 4. Select installTrigger in the dropdown → click ▶ Run → approve permissions.
 *
 * To test immediately run testCreateTab — it simulates the 1st of next month.
 */

// ── Configuration ──────────────────────────────────────────────────────────

const CFG = {
  // Row numbers (1-indexed)
  DAY_NUMBER_ROW:  6,     // Row 6 — day numbers 1…31 live here (D6:AH6)
  FIRST_DATA_ROW:  7,     // Row 7 — first pupil row

  // Column numbers (1-indexed: A=1, B=2, C=3, D=4 …)
  FIRST_DAY_COL:   4,     // Column D — always day 1

  // Header cells
  MONTH_CELL:     "AA1",  // Merged AA1:AG1 — month name e.g. "May"
  YEAR_CELL:      "AA2",  // Merged AA2:AG4 — year e.g. "2026"

  // "REGISTER" label — merged O2:T2
  REGISTER_CELL:  "O2",
  REGISTER_TEXT:  "REGISTER",
};

// ── Main function (triggered automatically on the 1st) ────────────────────

function createMonthlyTab() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const tz  = ss.getSpreadsheetTimeZone();

  const localDay = parseInt(Utilities.formatDate(now, tz, "d"), 10);
  if (localDay !== 1) {
    Logger.log("Not the 1st in spreadsheet timezone — exiting.");
    return;
  }

  _buildNewTab(ss, now);
}

// ── Core logic ────────────────────────────────────────────────────────────

function _buildNewTab(ss, targetDate) {
  const tz = ss.getSpreadsheetTimeZone();

  const newTabName  = Utilities.formatDate(targetDate, tz, "yy/MM");
  const monthName   = Utilities.formatDate(targetDate, tz, "MMMM");
  const yearStr     = Utilities.formatDate(targetDate, tz, "yyyy");
  const monthNum    = parseInt(Utilities.formatDate(targetDate, tz, "M"),    10);
  const yearNum     = parseInt(Utilities.formatDate(targetDate, tz, "yyyy"), 10);
  const daysInMonth = new Date(yearNum, monthNum, 0).getDate();

  Logger.log(`Building tab: ${newTabName} | ${monthName} ${yearStr} | ${daysInMonth} days`);

  if (ss.getSheetByName(newTabName)) {
    Logger.log(`"${newTabName}" already exists — skipping.`);
    return;
  }

  const sourceSheet = _getMostRecentTab(ss, newTabName);
  if (!sourceSheet) throw new Error("No YY/MM tab found to copy from.");
  Logger.log(`Copying from: "${sourceSheet.getName()}"`);

  const newSheet = sourceSheet.copyTo(ss);
  newSheet.setName(newTabName);
  newSheet.activate();
  ss.moveActiveSheet(1);

  try {
    // 1. Determine how many day columns the copied sheet has.
    //    Count only non-hidden columns — a shorter-month source leaves hidden surplus columns
    //    that still contain stale data, which would inflate getLastColumn().
    const prevLastCol = newSheet.getLastColumn();
    let prevDayCols = 0;
    for (let c = CFG.FIRST_DAY_COL; c <= prevLastCol; c++) {
      if (!newSheet.isColumnHiddenByUser(c)) prevDayCols++;
      else break;
    }
    if (prevDayCols === 0) prevDayCols = prevLastCol - CFG.FIRST_DAY_COL + 1;
    Logger.log(`Source has ${prevLastCol} total columns, ${prevDayCols} visible day columns.`);

    // 2. Insert columns if new month is longer than the source.
    //    Insert after the last VISIBLE day column, not prevLastCol, which may include hidden
    //    surplus columns from a previous shorter month sitting beyond the visible range.
    if (daysInMonth > prevDayCols) {
      const colsToAdd        = daysInMonth - prevDayCols;
      const lastVisibleDayCol = CFG.FIRST_DAY_COL + prevDayCols - 1;
      newSheet.insertColumnsAfter(lastVisibleDayCol, colsToAdd);
      // Copy formatting from the last visible day column to the new columns
      const numFormatRows = _lastPupilRow(newSheet);
      newSheet.getRange(1, lastVisibleDayCol, numFormatRows, 1)
              .copyTo(
                newSheet.getRange(1, lastVisibleDayCol + 1, numFormatRows, colsToAdd),
                SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false
              );
      newSheet.getRange(CFG.DAY_NUMBER_ROW, lastVisibleDayCol + 1, 1, colsToAdd).setNumberFormat('0');
      Logger.log(`Inserted ${colsToAdd} column(s) after col ${lastVisibleDayCol}.`);
    }

    // 3. Hide surplus columns if new month is shorter than the source.
    //    Also clear their content so getLastColumn() stays accurate for future copies.
    if (daysInMonth < prevDayCols) {
      const firstHideCol = CFG.FIRST_DAY_COL + daysInMonth;
      const numHide      = prevDayCols - daysInMonth;
      newSheet.hideColumns(firstHideCol, numHide);
      const clearRows = _lastPupilRow(newSheet) - CFG.DAY_NUMBER_ROW + 1;
      newSheet.getRange(CFG.DAY_NUMBER_ROW, firstHideCol, clearRows, numHide).clearContent();
      Logger.log(`Hid and cleared ${numHide} surplus column(s) from col ${firstHideCol}.`);
    }

    // 4. Unhide the day columns we need (in case source had hidden columns)
    newSheet.showColumns(CFG.FIRST_DAY_COL, daysInMonth);

    // 5. Write day numbers 1…N in a single API call
    const dayHeaderRange = newSheet.getRange(CFG.DAY_NUMBER_ROW, CFG.FIRST_DAY_COL, 1, daysInMonth);
    dayHeaderRange.clearContent();
    dayHeaderRange.setNumberFormat('0');
    dayHeaderRange.setValues([Array.from({ length: daysInMonth }, (_, i) => i + 1)]);
    Logger.log(`Wrote day numbers 1–${daysInMonth}.`);

    // 6. Clear all attendance data
    const lastPupilRow = _lastPupilRow(newSheet);
    const numRows      = lastPupilRow - CFG.FIRST_DATA_ROW + 1;
    newSheet.getRange(CFG.FIRST_DATA_ROW, CFG.FIRST_DAY_COL, numRows, daysInMonth).clearContent();
    Logger.log(`Cleared attendance: ${numRows} rows × ${daysInMonth} columns.`);

    // 7. Re-merge header ranges that span to the last day column.
    //    breakApart must reference the FULL old merged range (old last col),
    //    otherwise Sheets rejects the call with "must select all cells in merged range".
    const lastDayCol    = CFG.FIRST_DAY_COL + daysInMonth - 1;
    const lastDayColA1  = columnToLetter(lastDayCol);
    const oldLastDayCol = CFG.FIRST_DAY_COL + prevDayCols - 1;
    const oldLastColA1  = columnToLetter(oldLastDayCol);
    Logger.log(`Last day column: ${lastDayColA1} (col ${lastDayCol}), old: ${oldLastColA1}`);

    [
      [`AA1`, `1`],   // Month name
      [`AA2`, `3`],   // Year
      [`C4`,  `4`],   // "PIANO LESSONS WITH GREG KAIGHIN"
      [`D5`,  `5`],   // "Lesson Date"
    ].forEach(([start, row]) => {
      newSheet.getRange(`${start}:${oldLastColA1}${row}`).breakApart();
      newSheet.getRange(`${start}:${lastDayColA1}${row}`).merge();
      Logger.log(`Re-merged: ${start}:${lastDayColA1}${row}`);
    });

    // 8. Fix right border on the last day column
    const borderRows = lastPupilRow - CFG.DAY_NUMBER_ROW + 1;
    const SOLID      = SpreadsheetApp.BorderStyle.SOLID;

    // Remove right border from the old last day column if the month length changed
    if (daysInMonth !== prevDayCols) {
      const oldLastDayCol = CFG.FIRST_DAY_COL + prevDayCols - 1;
      newSheet.getRange(CFG.DAY_NUMBER_ROW, oldLastDayCol, borderRows, 1)
              .setBorder(null, null, null, false, null, null, null, null);
      Logger.log(`Removed right border from old last col: ${columnToLetter(oldLastDayCol)}`);

      // When inserting, add thin grey left + inner-vertical borders to ALL new columns
      if (daysInMonth > prevDayCols) {
        const firstNewCol = CFG.FIRST_DAY_COL + prevDayCols;
        const numNewCols  = daysInMonth - prevDayCols;
        newSheet.getRange(CFG.DAY_NUMBER_ROW, firstNewCol, borderRows, numNewCols)
                .setBorder(null, true, null, null, true, null, '#cccccc', SOLID);
        Logger.log(`Added thin left/inner borders to new cols ${columnToLetter(firstNewCol)}–${lastDayColA1}`);
      }
    }

    // Apply right border directly to the new last day column (data rows)
    newSheet.getRange(CFG.DAY_NUMBER_ROW, lastDayCol, borderRows, 1)
            .setBorder(null, null, null, true, null, null, '#000000', SOLID);
    Logger.log(`Applied right border to data rows at: ${lastDayColA1}`);

    // Apply right border to each merged header range individually.
    // setBorder on a column slice doesn't reach the right edge of merged cells —
    // the full merged range must be referenced so Sheets knows which cell owns the border.
    const hBorder = ['#000000', SpreadsheetApp.BorderStyle.SOLID];
    [
      `AA1:${lastDayColA1}1`,
      `AA2:${lastDayColA1}3`,
      `C4:${lastDayColA1}4`,
      `D5:${lastDayColA1}5`,
    ].forEach(addr => {
      newSheet.getRange(addr).setBorder(null, null, null, true, null, null, ...hBorder);
    });
    Logger.log(`Applied right border to header merged ranges ending at ${lastDayColA1}`);

    // 9. Restore "REGISTER" label and update month/year header
    newSheet.getRange(CFG.REGISTER_CELL).setValue(CFG.REGISTER_TEXT);
    newSheet.getRange(CFG.MONTH_CELL).clearContent();
    newSheet.getRange(CFG.YEAR_CELL).clearContent();
    SpreadsheetApp.flush();
    newSheet.getRange(CFG.MONTH_CELL).setValue(monthName);
    newSheet.getRange(CFG.YEAR_CELL).setValue(yearStr);
    Logger.log(`Header updated: ${monthName} ${yearStr}`);

    SpreadsheetApp.flush();
    Logger.log(`✅ Tab "${newTabName}" created successfully.`);

  } catch (e) {
    // Clean up the partial sheet so the spreadsheet isn't left in a broken state
    Logger.log(`❌ Error: ${e.message} — deleting partial tab "${newTabName}".`);
    ss.deleteSheet(newSheet);
    throw e;
  }
}

// ── Helper: last row containing pupil data (scans column A downward) ─────────

function _lastPupilRow(sheet) {
  const col    = sheet.getRange(CFG.FIRST_DATA_ROW, 1, sheet.getLastRow() - CFG.FIRST_DATA_ROW + 1, 1)
                      .getValues();
  let last = CFG.FIRST_DATA_ROW;
  for (let i = 0; i < col.length; i++) {
    if (col[i][0] !== '') last = CFG.FIRST_DATA_ROW + i;
  }
  return last;
}

// ── Helper: find most recent YY/MM tab ────────────────────────────────────────

function _getMostRecentTab(ss, excludeName) {
  const pattern = /^(\d{2})\/(\d{2})$/;
  let best = null, bestDate = null;

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name === excludeName) return;
    const m = name.match(pattern);
    if (m) {
      const year  = 2000 + parseInt(m[1], 10);
      const month = parseInt(m[2], 10) - 1;
      const d     = new Date(year, month, 1);
      if (!bestDate || d > bestDate) { bestDate = d; best = sheet; }
    }
  });

  return best;
}

// ── Helper: convert column number to A1 letter(s) e.g. 27→AA, 34→AH ─────────

function columnToLetter(col) {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// ── One-time trigger installer ────────────────────────────────────────────────

function installTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "createMonthlyTab")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("createMonthlyTab")
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .create();

  Logger.log("✅ Trigger installed — fires on the 1st of each month at midnight.");
}

// ── Test helpers ──────────────────────────────────────────────────────────────

// Simulates the 1st of NEXT month (quick default test).
function testCreateTab() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const tz  = ss.getSpreadsheetTimeZone();
  const now = new Date();

  const currentMonth = parseInt(Utilities.formatDate(now, tz, "M"),    10);
  const currentYear  = parseInt(Utilities.formatDate(now, tz, "yyyy"), 10);
  const nextMonth    = currentMonth === 12 ? 1               : currentMonth + 1;
  const nextYear     = currentMonth === 12 ? currentYear + 1 : currentYear;

  const testDate = new Date(`${nextYear}-${String(nextMonth).padStart(2, '0')}-01T12:00:00`);
  Logger.log(`Test: simulating 1st of ${Utilities.formatDate(testDate, tz, "MMMM yyyy")}`);
  _buildNewTab(ss, testDate);
}

// Simulates the 1st of a specific month — use this to test any month length.
// Set YEAR and MONTH below, then run testCreateTabForMonth.
//   30-day test : MONTH = 6  (June)
//   28-day test : MONTH = 2, YEAR = 2027 (February)
//   31-day test : MONTH = 7  (July)
function testCreateTabForMonth() {
  const YEAR  = 2026;
  const MONTH = 6;    // 1=Jan … 12=Dec

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const tz       = ss.getSpreadsheetTimeZone();
  const testDate = new Date(`${YEAR}-${String(MONTH).padStart(2, '0')}-01T12:00:00`);
  Logger.log(`Test: simulating 1st of ${Utilities.formatDate(testDate, tz, "MMMM yyyy")}`);
  _buildNewTab(ss, testDate);
}
