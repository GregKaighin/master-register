/**
 * PIANO LESSONS WITH GREG KAIGHIN — Individual Pupil Register Script
 * ------------------------------------------------------------------
 * Automatically creates a new monthly register tab on the 1st of each month.
 * Copies the most recent YY/MM tab, adjusts day columns, clears lesson data,
 * updates IMPORTRANGE formulas to reference the new master register tab,
 * and places the new tab at position 1. Deletes the partial tab on error.
 *
 * SETUP: Extensions → Apps Script → paste → save → run installTrigger → approve.
 * TEST:  Run testCreateTab (next month) or testCreateTabForMonth (specific month).
 */

// ── Configuration ──────────────────────────────────────────────────────────

const CFG = {
  DAY_NUMBER_ROW:    6,
  FIRST_DATA_ROW:    7,
  FIRST_DAY_COL:     4,
  PUPIL_SCAN_COL:    2,
  MONTH_CELL:       "AA1",
  YEAR_CELL:        "AA2",
  REGISTER_CELL:    "O2",
  REGISTER_TEXT:    "REGISTER",
  FORMULA_START_ROW: 7,
  FORMULA_NUM_COLS:  2,
};

const SOLID = SpreadsheetApp.BorderStyle.SOLID;

// ── Main function (triggered automatically on the 1st) ────────────────────

function createMonthlyTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  if (parseInt(Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "d"), 10) !== 1) return;
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
  const oldTabName = sourceSheet.getName();
  Logger.log(`Copying from: "${oldTabName}"`);

  const newSheet = sourceSheet.copyTo(ss);
  newSheet.setName(newTabName);
  newSheet.activate();
  ss.moveActiveSheet(1);

  try {
    // Count visible day columns — hidden surplus from a shorter source must not be counted.
    const prevLastCol = newSheet.getLastColumn();
    let prevDayCols = 0;
    for (let c = CFG.FIRST_DAY_COL; c <= prevLastCol; c++) {
      if (!newSheet.isColumnHiddenByUser(c)) prevDayCols++;
      else break;
    }
    if (prevDayCols === 0) prevDayCols = prevLastCol - CFG.FIRST_DAY_COL + 1;

    const lastPupilRow  = _lastPupilRow(newSheet);
    const oldLastDayCol = CFG.FIRST_DAY_COL + prevDayCols - 1;
    const lastDayCol    = CFG.FIRST_DAY_COL + daysInMonth - 1;
    const borderRows    = lastPupilRow - CFG.DAY_NUMBER_ROW + 1;

    Logger.log(`Source: ${prevLastCol} total cols, ${prevDayCols} visible day cols.`);

    if (daysInMonth > prevDayCols) {
      const colsToAdd = daysInMonth - prevDayCols;
      newSheet.insertColumnsAfter(oldLastDayCol, colsToAdd);
      newSheet.getRange(1, oldLastDayCol, lastPupilRow, 1)
              .copyTo(newSheet.getRange(1, oldLastDayCol + 1, lastPupilRow, colsToAdd),
                      SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      newSheet.getRange(CFG.DAY_NUMBER_ROW, oldLastDayCol + 1, 1, colsToAdd).setNumberFormat('0');
      Logger.log(`Inserted ${colsToAdd} col(s) after col ${oldLastDayCol}.`);
    }

    if (daysInMonth < prevDayCols) {
      const firstHideCol = CFG.FIRST_DAY_COL + daysInMonth;
      const numHide      = prevDayCols - daysInMonth;
      newSheet.hideColumns(firstHideCol, numHide);
      newSheet.getRange(CFG.DAY_NUMBER_ROW, firstHideCol, lastPupilRow - CFG.DAY_NUMBER_ROW + 1, numHide).clearContent();
      Logger.log(`Hid ${numHide} surplus col(s) from col ${firstHideCol}.`);
    }

    newSheet.showColumns(CFG.FIRST_DAY_COL, daysInMonth);

    const dayHeaderRange = newSheet.getRange(CFG.DAY_NUMBER_ROW, CFG.FIRST_DAY_COL, 1, daysInMonth);
    dayHeaderRange.setNumberFormat('0');
    dayHeaderRange.setValues([Array.from({ length: daysInMonth }, (_, i) => i + 1)]);

    newSheet.getRange(CFG.FIRST_DATA_ROW, CFG.FIRST_DAY_COL,
                      lastPupilRow - CFG.FIRST_DATA_ROW + 1, daysInMonth).clearContent();
    Logger.log(`Cleared lesson data: ${lastPupilRow - CFG.FIRST_DATA_ROW + 1} rows × ${daysInMonth} cols.`);

    if (daysInMonth !== prevDayCols) {
      SpreadsheetApp.flush(); // commit column ops before reading merge state

      // Re-merge header ranges ending at the old last day col, extend to new last day col,
      // and apply the right border — all in a single pass.
      newSheet.getRange(1, 1, CFG.DAY_NUMBER_ROW - 1, Math.max(oldLastDayCol, lastDayCol))
        .getMergedRanges()
        .filter(m => m.getColumn() + m.getNumColumns() - 1 === oldLastDayCol)
        .forEach(m => {
          const newMerge = newSheet.getRange(m.getRow(), m.getColumn(), m.getNumRows(),
                                             lastDayCol - m.getColumn() + 1);
          m.breakApart();
          newMerge.merge();
          newMerge.setBorder(null, null, null, true, null, null, '#000000', SOLID);
        });
      Logger.log(`Re-merged header ranges to column ${columnToLetter(lastDayCol)}.`);

      newSheet.getRange(CFG.DAY_NUMBER_ROW, oldLastDayCol, borderRows, 1)
              .setBorder(null, null, null, false, null, null, null, null);

      if (daysInMonth > prevDayCols) {
        newSheet.getRange(CFG.DAY_NUMBER_ROW, oldLastDayCol + 1, borderRows, daysInMonth - prevDayCols)
                .setBorder(null, true, null, null, true, null, '#cccccc', SOLID);
      }
    }

    newSheet.getRange(CFG.DAY_NUMBER_ROW, lastDayCol, borderRows, 1)
            .setBorder(null, null, null, true, null, null, '#000000', SOLID);
    Logger.log(`Applied right border at column ${columnToLetter(lastDayCol)}.`);

    newSheet.getRange(CFG.REGISTER_CELL).setValue(CFG.REGISTER_TEXT);
    newSheet.getRange(CFG.MONTH_CELL).clearContent();
    newSheet.getRange(CFG.YEAR_CELL).clearContent();
    SpreadsheetApp.flush();
    newSheet.getRange(CFG.MONTH_CELL).setValue(monthName);
    newSheet.getRange(CFG.YEAR_CELL).setValue(yearStr);
    Logger.log(`Header updated: ${monthName} ${yearStr}`);

    _updateFormulas(newSheet, oldTabName, newTabName);

    SpreadsheetApp.flush();
    Logger.log(`✅ Tab "${newTabName}" created successfully.`);

  } catch (e) {
    Logger.log(`❌ Error: ${e.message} — deleting partial tab "${newTabName}".`);
    ss.deleteSheet(newSheet);
    throw e;
  }
}

// ── Helper: update IMPORTRANGE formulas with the new master tab name ──────────

function _updateFormulas(sheet, oldTabName, newTabName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CFG.FORMULA_START_ROW) { Logger.log("No formula rows found."); return; }
  const range    = sheet.getRange(CFG.FORMULA_START_ROW, 2, lastRow - CFG.FORMULA_START_ROW + 1, CFG.FORMULA_NUM_COLS);
  const formulas = range.getFormulas();
  let updated    = 0;
  formulas.forEach((row, r) => {
    row.forEach((f, c) => {
      if (f.includes(`"${oldTabName}!`)) {
        range.getCell(r + 1, c + 1).setFormula(f.replaceAll(`"${oldTabName}!`, `"${newTabName}!`));
        updated++;
      }
    });
  });
  Logger.log(`Updated ${updated} IMPORTRANGE formula(s): "${oldTabName}" → "${newTabName}"`);
}

// ── Helper: last row containing pupil data ────────────────────────────────────

function _lastPupilRow(sheet) {
  const values = sheet.getRange(CFG.FIRST_DATA_ROW, CFG.PUPIL_SCAN_COL,
                                sheet.getLastRow() - CFG.FIRST_DATA_ROW + 1, 1).getValues();
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

// ── Helper: column number to A1 letter(s) e.g. 27→AA, 34→AH ─────────────────

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
  ScriptApp.newTrigger("createMonthlyTab").timeBased().onMonthDay(1).atHour(0).create();
  Logger.log("✅ Trigger installed — fires on the 1st of each month at midnight.");
}

// ── Test helpers ──────────────────────────────────────────────────────────────

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

// 30-day test: MONTH = 6 | 28-day test: MONTH = 2, YEAR = 2027 | 31-day test: MONTH = 7
function testCreateTabForMonth() {
  const YEAR  = 2026;
  const MONTH = 6;    // 1=Jan … 12=Dec

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const tz       = ss.getSpreadsheetTimeZone();
  const testDate = new Date(`${YEAR}-${String(MONTH).padStart(2, '0')}-01T12:00:00`);
  Logger.log(`Test: simulating 1st of ${Utilities.formatDate(testDate, tz, "MMMM yyyy")}`);
  _buildNewTab(ss, testDate);
}
