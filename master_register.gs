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
  DAY_NUMBER_ROW: 6,          // Row 6 — day numbers 1…31 (D6:AH6)
  FIRST_DATA_ROW: 7,          // Row 7 — first pupil row
  FIRST_DAY_COL:  4,          // Column D — always day 1
  MONTH_CELL:    "AA1",       // Merged AA1:AG1 — month name e.g. "May"
  YEAR_CELL:     "AA2",       // Merged AA2:AG4 — year e.g. "2026"
  REGISTER_CELL: "O2",        // Merged O2:T2 — "REGISTER" label
  REGISTER_TEXT: "REGISTER",
};

const SOLID = SpreadsheetApp.BorderStyle.SOLID;

// ── Main function (triggered automatically on the 1st) ────────────────────

function createMonthlyTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  if (parseInt(Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "d"), 10) !== 1) {
    Logger.log("Not the 1st in spreadsheet timezone — exiting.");
    return;
  }
  _buildNewTab(ss, now);
}

// ── Core logic ────────────────────────────────────────────────────────────

function _buildNewTab(ss, targetDate) {
  const tz  = ss.getSpreadsheetTimeZone();
  const fmt = (f) => Utilities.formatDate(targetDate, tz, f);

  const newTabName  = fmt("yy/MM");
  const monthName   = fmt("MMMM");
  const yearStr     = fmt("yyyy");
  const monthNum    = parseInt(fmt("M"), 10);
  const daysInMonth = new Date(+yearStr, monthNum, 0).getDate();

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
    // 1. Count visible day columns in the source.
    //    Hidden surplus columns from a shorter-month source must not be counted.
    const prevLastCol = newSheet.getLastColumn();
    let prevDayCols = 0;
    for (let c = CFG.FIRST_DAY_COL; c <= prevLastCol; c++) {
      if (!newSheet.isColumnHiddenByUser(c)) prevDayCols++;
      else break;
    }
    if (prevDayCols === 0) prevDayCols = prevLastCol - CFG.FIRST_DAY_COL + 1;
    Logger.log(`Source: ${prevLastCol} total columns, ${prevDayCols} visible day columns.`);

    const lastPupilRow  = _lastPupilRow(newSheet);
    const oldLastDayCol = CFG.FIRST_DAY_COL + prevDayCols - 1;

    // 2. Insert columns if new month is longer than the source.
    //    Insert after the last VISIBLE day column — prevLastCol may include hidden surplus.
    if (daysInMonth > prevDayCols) {
      const colsToAdd = daysInMonth - prevDayCols;
      newSheet.insertColumnsAfter(oldLastDayCol, colsToAdd);
      newSheet.getRange(1, oldLastDayCol, lastPupilRow, 1)
              .copyTo(
                newSheet.getRange(1, oldLastDayCol + 1, lastPupilRow, colsToAdd),
                SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false
              );
      newSheet.getRange(CFG.DAY_NUMBER_ROW, oldLastDayCol + 1, 1, colsToAdd).setNumberFormat('0');
      Logger.log(`Inserted ${colsToAdd} column(s) after col ${oldLastDayCol}.`);
    }

    // 3. Hide surplus columns if new month is shorter, and clear their content.
    if (daysInMonth < prevDayCols) {
      const firstHideCol = CFG.FIRST_DAY_COL + daysInMonth;
      const numHide      = prevDayCols - daysInMonth;
      newSheet.hideColumns(firstHideCol, numHide);
      newSheet.getRange(CFG.DAY_NUMBER_ROW, firstHideCol, lastPupilRow - CFG.DAY_NUMBER_ROW + 1, numHide).clearContent();
      Logger.log(`Hid and cleared ${numHide} surplus column(s) from col ${firstHideCol}.`);
    }

    // 4. Unhide the day columns we need (in case source had hidden columns).
    newSheet.showColumns(CFG.FIRST_DAY_COL, daysInMonth);

    // 5. Write day numbers 1…N.
    const dayHeaderRange = newSheet.getRange(CFG.DAY_NUMBER_ROW, CFG.FIRST_DAY_COL, 1, daysInMonth);
    dayHeaderRange.setNumberFormat('0');
    dayHeaderRange.setValues([Array.from({ length: daysInMonth }, (_, i) => i + 1)]);
    Logger.log(`Wrote day numbers 1–${daysInMonth}.`);

    // 6. Clear all attendance data.
    newSheet.getRange(CFG.FIRST_DATA_ROW, CFG.FIRST_DAY_COL, lastPupilRow - CFG.FIRST_DATA_ROW + 1, daysInMonth).clearContent();
    Logger.log(`Cleared attendance: ${lastPupilRow - CFG.FIRST_DATA_ROW + 1} rows × ${daysInMonth} columns.`);

    // 7. Re-merge header ranges that span to the last day column.
    //    breakApart must reference the FULL old merged range; otherwise Sheets rejects the call.
    const lastDayCol = CFG.FIRST_DAY_COL + daysInMonth - 1;
    const lastDayA1  = columnToLetter(lastDayCol);
    const oldLastA1  = columnToLetter(oldLastDayCol);

    [["AA1", "1"], ["AA2", "3"], ["C4", "4"], ["D5", "5"]].forEach(([start, row]) => {
      newSheet.getRange(`${start}:${oldLastA1}${row}`).breakApart();
      newSheet.getRange(`${start}:${lastDayA1}${row}`).merge();
    });
    Logger.log(`Re-merged header ranges to column ${lastDayA1}.`);

    // 8. Fix borders on the last day column.
    const borderRows = lastPupilRow - CFG.DAY_NUMBER_ROW + 1;

    if (daysInMonth !== prevDayCols) {
      // Remove right border from the old last day column.
      newSheet.getRange(CFG.DAY_NUMBER_ROW, oldLastDayCol, borderRows, 1)
              .setBorder(null, null, null, false, null, null, null, null);
      Logger.log(`Removed right border from old last col: ${oldLastA1}`);

      // Add left + inner-vertical grey borders to all newly inserted columns.
      if (daysInMonth > prevDayCols) {
        const firstNewCol = oldLastDayCol + 1;
        const numNewCols  = daysInMonth - prevDayCols;
        newSheet.getRange(CFG.DAY_NUMBER_ROW, firstNewCol, borderRows, numNewCols)
                .setBorder(null, true, null, null, true, null, '#cccccc', SOLID);
        Logger.log(`Added grey borders to new cols ${columnToLetter(firstNewCol)}–${lastDayA1}`);
      }
    }

    // Apply right border to the new last day column — data rows and each merged header range.
    // setBorder on a column slice doesn't reach the right edge of merged cells, so each
    // merged range must be referenced in full.
    newSheet.getRange(CFG.DAY_NUMBER_ROW, lastDayCol, borderRows, 1)
            .setBorder(null, null, null, true, null, null, '#000000', SOLID);
    [`AA1:${lastDayA1}1`, `AA2:${lastDayA1}3`, `C4:${lastDayA1}4`, `D5:${lastDayA1}5`]
      .forEach(addr => newSheet.getRange(addr).setBorder(null, null, null, true, null, null, '#000000', SOLID));
    Logger.log(`Applied right border at column ${lastDayA1}.`);

    // 9. Restore "REGISTER" label and update month/year header.
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
    Logger.log(`❌ Error: ${e.message} — deleting partial tab "${newTabName}".`);
    ss.deleteSheet(newSheet);
    throw e;
  }
}

// ── Helper: last row containing pupil data (scans column A downward) ─────────

function _lastPupilRow(sheet) {
  const values = sheet.getRange(CFG.FIRST_DATA_ROW, 1, sheet.getLastRow() - CFG.FIRST_DATA_ROW + 1, 1).getValues();
  let last = CFG.FIRST_DATA_ROW;
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] !== '') last = CFG.FIRST_DATA_ROW + i;
  }
  return last;
}

// ── Helper: find most recent YY/MM tab ────────────────────────────────────────

function _getMostRecentTab(ss, excludeName) {
  return ss.getSheets()
    .filter(s => s.getName() !== excludeName && /^\d{2}\/\d{2}$/.test(s.getName()))
    .reduce((best, sheet) => {
      const [yy, mm] = sheet.getName().split('/').map(Number);
      const d = new Date(2000 + yy, mm - 1, 1);
      return !best || d > best.d ? { sheet, d } : best;
    }, null)
    ?.sheet ?? null;
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
  const m   = parseInt(Utilities.formatDate(now, tz, "M"),    10);
  const y   = parseInt(Utilities.formatDate(now, tz, "yyyy"), 10);
  const nm  = m === 12 ? 1     : m + 1;
  const ny  = m === 12 ? y + 1 : y;
  const testDate = new Date(`${ny}-${String(nm).padStart(2, '0')}-01T12:00:00`);
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
