# Piano Lessons with Greg Kaighin — Master Register Script

A Google Apps Script that automatically creates a new monthly register tab on the 1st of each month.

## What it does

- Copies the most recent `YY/MM` tab as the template for the new month
- Names the new tab in `YY/MM` format (e.g. `26/05`)
- Inserts or hides day columns to match the new month's exact length (28–31 days)
- Writes day numbers 1…N in row 6
- Clears all attendance data while preserving pupil names and formatting
- Re-merges header ranges that span to the last day column
- Fixes the right border on the last day column
- Updates the month name (`AA1`) and year (`AA2`) in the header
- Places the new tab at position 1 (leftmost)
- Deletes the partial tab automatically if anything goes wrong

## Setup

1. Open your Master Register spreadsheet in Google Sheets.
2. Click **Extensions → Apps Script**.
3. Delete any existing code and paste the entire contents of `master_register.gs`.
4. Click **Save** (Ctrl+S).
5. In the function dropdown, select **`installTrigger`** and click **▶ Run**.
6. Approve the permissions dialog — this is a one-time step.

The trigger is now installed. It fires automatically at midnight on the 1st of each month in your spreadsheet's timezone.

## Testing

To test without waiting for the 1st of the month:

**`testCreateTab`** — simulates the 1st of next month. Select it in the dropdown and click ▶ Run.

**`testCreateTabForMonth`** — simulates the 1st of any specific month. Edit the `YEAR` and `MONTH` constants at the top of the function, then run it.

```javascript
function testCreateTabForMonth() {
  const YEAR  = 2026;
  const MONTH = 6;    // 1=Jan … 12=Dec
  ...
}
```

## Configuration

The `CFG` object at the top of the script is pre-configured for this spreadsheet layout:

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

If your spreadsheet layout ever changes, update these values to match.

## How it handles different month lengths

| Scenario | Action |
|---|---|
| New month is longer than source | Inserts the extra columns after the last visible day column, copies formatting, adds inner grey borders |
| New month is shorter than source | Hides the surplus columns and clears their content |
| Same length | No column changes needed |

The last pupil row is detected automatically by scanning column A downward — no manual configuration needed.

## File structure

```
master-register/
├── master_register.gs    # Main script
├── monthly_register.gs   # Legacy / reference copy
├── README.md
├── SETUP_GUIDE.md
├── TROUBLESHOOTING.md
├── EXAMPLES.md
├── CHANGELOG.md
└── LICENSE
```

## Requirements

- A Google Sheets spreadsheet with at least one existing `YY/MM` tab to copy from
- Edit access to the spreadsheet and permission to run Apps Script

## License

MIT — see [LICENSE](LICENSE).
