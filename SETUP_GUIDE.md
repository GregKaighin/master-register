# Setup Guide

## Step 1 — Open Apps Script

1. Open your Master Register spreadsheet in Google Sheets.
2. Click **Extensions → Apps Script**.
3. A new browser tab opens with the script editor.

## Step 2 — Paste the script

1. Select all existing code in the editor (Ctrl+A).
2. Delete it.
3. Copy the entire contents of `master_register.gs` from this repository.
4. Paste it into the editor (Ctrl+V).
5. Save (Ctrl+S). The title bar should show the project name.

## Step 3 — Install the trigger

1. In the function dropdown (next to the ▶ Run button), select **`installTrigger`**.
2. Click **▶ Run**.
3. A permissions dialog appears — click **Review permissions**, select your Google account, and click **Allow**.
4. The trigger is now installed. You will see this in the log:

   ```
   ✅ Trigger installed — fires on the 1st of each month at midnight.
   ```

   You only need to do this once. To confirm it exists, click the **Triggers** icon (clock) in the left sidebar — you should see `createMonthlyTab` listed with a monthly time-based trigger.

## Step 4 — Test it

Run **`testCreateTab`** to simulate the 1st of next month:

1. Select **`testCreateTab`** in the function dropdown.
2. Click **▶ Run**.
3. Switch back to your spreadsheet — a new `YY/MM` tab should appear at the far left.
4. Check the new tab:
   - Month name and year in the header are correct.
   - Day numbers in row 6 match the month's length.
   - Attendance data is cleared.
   - Pupil names and formatting are intact.
   - Borders are correct, including the right border on the last day column.

To test a specific month (e.g. February or December), use **`testCreateTabForMonth`** instead:

1. Open the script editor and find `testCreateTabForMonth` near the bottom.
2. Set the `YEAR` and `MONTH` constants to the month you want to test.
3. Save, then run the function.

## Step 5 — Go live

No further action needed. The trigger fires automatically at midnight on the 1st of each month in the spreadsheet's timezone. Each run creates the new tab and logs its progress — you can review logs via **Execution log** in Apps Script.

## Reinstalling the trigger

If you ever need to reinstall (e.g. after copying the script to a new spreadsheet), just run `installTrigger` again. It removes any existing `createMonthlyTab` trigger before creating a new one, so there's no risk of duplicates.

## Configuration

The `CFG` object is pre-configured for this spreadsheet. Only change it if the layout of the register changes:

```javascript
const CFG = {
  DAY_NUMBER_ROW: 6,          // Row 6 — day numbers 1…31 (D6:AH6)
  FIRST_DATA_ROW: 7,          // Row 7 — first pupil row
  FIRST_DAY_COL:  4,          // Column D — always day 1
  MONTH_CELL:    "AA1",       // Merged AA1:AG1 — month name e.g. "May"
  YEAR_CELL:     "AA2",       // Merged AA2:AG4 — year e.g. "2026"
  REGISTER_CELL: "O2",        // Merged O2:T2 — "REGISTER" label
  REGISTER_TEXT: "REGISTER",
};
```

To find column numbers: A=1, B=2, C=3, D=4 … Z=26, AA=27, AB=28 … AH=34.
