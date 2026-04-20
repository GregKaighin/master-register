# Piano Lessons with Greg Kaighin — Monthly Register Scripts

Two Google Apps Scripts that automatically create a new monthly register tab on the 1st of each month.

## Scripts

### `master_register.gs` — Master Register

For the main attendance register covering all pupils.

**What it does:**
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

---

### `individual_register.gs` — Individual Pupil Registers

For individual pupil registers (single pupils, or siblings/parents sharing one register). Identical to the master register script with one addition:

- Updates `IMPORTRANGE` formulas in columns B–C to reference the new master register tab

The formulas pull attendance data from the master register into each individual register. They contain the master tab name twice (data range + filter column), so both occurrences are updated on each run.

---

## Setup

Both scripts are set up the same way:

1. Open the relevant Google Sheet → **Extensions → Apps Script**.
2. Delete any existing code and paste the entire contents of the script file.
3. Click **Save** (Ctrl+S).
4. In the function dropdown, select **`installTrigger`** and click **▶ Run**.
5. Approve the permissions dialog — this is a one-time step.

The trigger fires automatically at midnight on the 1st of each month in the spreadsheet's timezone.

See [SETUP_GUIDE.md](SETUP_GUIDE.md) for detailed instructions.

## Testing

**`testCreateTab`** — simulates the 1st of next month.

**`testCreateTabForMonth`** — simulates the 1st of any specific month. Edit `YEAR` and `MONTH` at the top of the function, then run it.

```javascript
function testCreateTabForMonth() {
  const YEAR  = 2026;
  const MONTH = 6;    // 1=Jan … 12=Dec
  ...
}
```

## Configuration

Both scripts use a `CFG` object pre-configured for this spreadsheet layout. Only change it if the layout changes.

**`master_register.gs`:**
```javascript
const CFG = {
  DAY_NUMBER_ROW: 6,          // Row 6 — day numbers 1…31 (D6:AH6)
  FIRST_DATA_ROW: 7,          // Row 7 — first pupil row
  FIRST_DAY_COL:  4,          // Column D — always day 1
  MONTH_CELL:    "AA1",       // Month name e.g. "May"
  YEAR_CELL:     "AA2",       // Year e.g. "2026"
  REGISTER_CELL: "O2",        // "REGISTER" label
  REGISTER_TEXT: "REGISTER",
};
```

**`individual_register.gs`** — same as above, plus:
```javascript
  PUPIL_SCAN_COL:    2,       // Column B — scanned to find the last pupil row
  FORMULA_START_ROW: 7,       // First row containing IMPORTRANGE formulas
  FORMULA_NUM_COLS:  2,       // Number of formula columns (B and C)
```

## How it handles different month lengths

| Scenario | Action |
|---|---|
| New month is longer than source | Inserts extra columns, copies formatting, adds inner grey borders |
| New month is shorter than source | Hides surplus columns and clears their content |
| Same length | No column changes needed |

## File structure

```
├── master_register.gs       # Master register script
├── individual_register.gs   # Individual pupil register script
├── README.md
├── SETUP_GUIDE.md
├── TROUBLESHOOTING.md
└── LICENSE
```

## Requirements

- A Google Sheets spreadsheet with at least one existing `YY/MM` tab to copy from
- Edit access to the spreadsheet and permission to run Apps Script

## License

MIT — see [LICENSE](LICENSE).
