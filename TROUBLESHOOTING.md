# Troubleshooting

## The trigger didn't fire on the 1st

**Check the trigger exists:**
1. Open Apps Script → click the **Triggers** icon (clock) in the left sidebar.
2. You should see `createMonthlyTab` with a monthly, day-of-month trigger set to day 1.
3. If it's missing, run `installTrigger` again.

**Check the timezone:**
The trigger fires at midnight in the spreadsheet's timezone, not your local time. If you're in a different timezone the tab may appear a few hours before or after midnight local time.

**Check the execution log:**
Apps Script → **Executions** (left sidebar) shows every run, its status, and any log output. If the trigger fired but something went wrong, the error will be here.

---

## "No YY/MM tab found to copy from"

The script looks for the most recent tab whose name matches `YY/MM` (e.g. `26/04`). This error means none exist.

- Make sure at least one tab is named in `YY/MM` format.
- Tab names are case-sensitive and must use a forward slash, not a dash or dot.

---

## Tab created with the wrong number of day columns

**Symptom:** The new tab has too many or too few day columns.

**Most likely cause:** The source tab has hidden columns that confused the column count.

The script counts only *visible* day columns from the source. If a previous test run left a tab with hidden columns in an unexpected state, delete that tab and re-run from a clean source.

**Check the log:** The execution log prints the source tab name, how many visible day columns it detected, and how many it inserted or hid. This quickly shows where the count diverged.

---

## Inner grey vertical borders are missing between day columns

This can happen if the new month is more than one day longer than the source tab (e.g. copying from a 29-day February to a 31-day December adds 2 columns). The script applies left and inner-vertical grey borders to the entire range of new columns in one call, so all dividers should be present. If you see a missing border:

1. Check which columns are affected (hover to confirm column letters).
2. Run `testCreateTabForMonth` with that specific month to reproduce.
3. Review the execution log for the border steps — each step is logged.

---

## Right border on the last day column is missing or doubled

The script removes the right border from the old last day column and applies it to the new one. If the border looks wrong:

- **Missing:** The `setBorder` call for the new last column may not have run. Check the log for `Applied right border to data rows at: …`.
- **Doubled:** The old border wasn't removed. The log should show `Removed right border from old last col: …` — if that line is absent, the `daysInMonth !== prevDayCols` condition wasn't met, meaning the column count was read incorrectly.

---

## "REGISTER" label disappeared

The script restores the `REGISTER` label in cell `O2` in step 9. If it's missing, check `CFG.REGISTER_CELL` and `CFG.REGISTER_TEXT` match your spreadsheet.

---

## Month name or year in the header is wrong or blank

The script clears `AA1` and `AA2` then writes the new values using `SpreadsheetApp.flush()` between operations. If the cells are blank:

- Confirm `CFG.MONTH_CELL` is `"AA1"` and `CFG.YEAR_CELL` is `"AA2"`.
- Check the log for `Header updated: …` — if that line is missing, an earlier error interrupted execution.

---

## Partial/broken tab was left behind

The script wraps all operations in a `try/catch`. If an error occurs mid-run, it deletes the partial tab and re-throws the error, so the spreadsheet is left clean. If a partial tab does exist, it means the delete itself failed (rare). Delete it manually, fix the underlying error, and run again.

---

## Script changes aren't taking effect

After editing the script, save with Ctrl+S and re-run from the dropdown. Refreshing the spreadsheet alone does not reload script logic — you must save in the Apps Script editor.

---

## Checking the execution log

1. Open Apps Script (Extensions → Apps Script).
2. Click **Executions** in the left sidebar.
3. Click any execution to expand the log output.

Every major step is logged with `Logger.log(...)`, so you can trace exactly what the script did and where it stopped.
