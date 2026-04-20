# Setup Guide - Master Register Script

This guide walks you through installing and configuring the Master Register script for your specific spreadsheet.

## Prerequisites

- Google Sheets account
- Access to your Master Register spreadsheet
- Basic familiarity with Google Sheets
- Read and write permissions on the spreadsheet

## Installation Steps

### Step 1: Open Google Apps Script Editor

1. Open your Master Register spreadsheet in Google Sheets
2. Click **Extensions** in the menu bar
3. Click **Apps Script**
4. A new tab will open with the script editor

### Step 2: Paste the Script

1. In the Apps Script editor, you'll see code in the main panel
2. **Select all** existing code (Ctrl+A / Cmd+A)
3. **Delete** it
4. **Copy** the entire contents of `master_register.gs` from this repository
5. **Paste** it into the editor (Ctrl+V / Cmd+V)
6. **Save** the script (Ctrl+S / Cmd+S)

### Step 3: Locate Your Configuration Values

You need to find 6 values that match your spreadsheet layout. These go in the `CFG` object at the top of the script.

#### Finding monthCell

1. In your spreadsheet, find the cell containing your current month name (e.g., "April")
2. Click on that cell
3. Look at the **Name Box** in the top-left corner of your spreadsheet (it shows coordinates like "AH1")
4. **Note this address** — this is your `monthCell` value

Example: If the Name Box shows `AH1`, then `monthCell: 'AH1'`

#### Finding yearCell

1. Find the cell containing your current year (e.g., "2026")
2. Click on that cell
3. Check the Name Box for the cell address
4. **Note this address** — this is your `yearCell` value

Example: If the Name Box shows `AH2`, then `yearCell: 'AH2'`

#### Finding dayNumbersRow

1. Look for the row containing your day numbers (1, 2, 3, 4, 5... across columns)
2. Click any cell in that row that contains a day number
3. The Name Box will show something like `D6` — the number after the letter is your row number
4. **Note the row number** — this is your `dayNumbersRow` value

Example: If the Name Box shows `D6`, then `dayNumbersRow: 6`

#### Finding firstDayCol

1. Find the first cell in your lesson date grid (where day 1's lessons go)
2. Click on that cell
3. The Name Box shows something like `D7` — the letter is your column
4. Convert the letter to a number:
   - A = 1, B = 2, C = 3, D = 4, E = 5, F = 6, etc.
   - Or count columns from the left
5. **Note this number** — this is your `firstDayCol` value

Example: If the Name Box shows `D7`, then `firstDayCol: 4`

#### Finding firstPupilRow

1. Locate the first row containing a pupil name or data
2. Click on that row
3. The Name Box shows the row number (e.g., `D7` = row 7)
4. **Note the row number** — this is your `firstPupilRow` value

Example: If the Name Box shows `A7`, then `firstPupilRow: 7`

#### Finding lastPupilRow

1. Scroll down to the last row containing pupil data
2. Click that row
3. The Name Box shows the row number
4. **Note this number** — this is your `lastPupilRow` value

Example: If the Name Box shows `A48`, then `lastPupilRow: 48`

### Step 4: Update the Configuration

1. In the Apps Script editor, find the `CFG` object at the top (lines ~16-22)
2. Update each value with the numbers you found:

```javascript
const CFG = {
  monthCell:     'AH1', // ← Your month cell
  yearCell:      'AH2', // ← Your year cell
  dayNumbersRow: 6,     // ← Your day numbers row
  firstDayCol:   4,     // ← Your first day column number
  firstPupilRow: 7,     // ← Your first pupil row
  lastPupilRow:  48,    // ← Your last pupil row
};
```

### Step 5: Save and Test

1. **Save** the script (Ctrl+S)
2. **Close** the Apps Script tab
3. **Refresh** your spreadsheet (F5 or Cmd+R)
4. Look for a new **"Register"** menu in the top menu bar
5. Click **Register** → **Create next month tab** to test
6. If it works, you'll see a confirmation message with the new sheet name

## Verification Checklist

Before using the script in production, verify:

- ✅ The "Register" menu appears after refresh
- ✅ Configuration values are saved in the script
- ✅ Month and year cells contain correct values
- ✅ You can successfully create a test sheet
- ✅ New sheet has the correct month/year in the header
- ✅ Day columns match the new month's day count
- ✅ Pupil names are preserved
- ✅ Attendance data is cleared
- ✅ Formatting and formulas are maintained

## Common Setup Issues

### "Register" menu doesn't appear

1. Check that you saved the script (Ctrl+S in Apps Script editor)
2. Refresh your spreadsheet completely (F5 or Cmd+R)
3. Try closing and reopening the spreadsheet
4. Check browser console for errors (F12 → Console tab)

### Configuration values are wrong

1. Recheck each cell in your spreadsheet using the Name Box
2. Make sure you're looking at the correct sheet (if you have multiple sheets)
3. Verify month name is spelled correctly (case-sensitive)
4. Year should be a 4-digit number

### Permission denied errors

1. Make sure you have edit access to the spreadsheet
2. Try refreshing the page
3. Sign out and back into your Google account
4. Check sharing settings on the spreadsheet

### Script errors when clicking "Create next month tab"

1. Verify all configuration values are correct
2. Check that month and year cells contain valid data
3. Ensure no sheet with next month's name already exists
4. Check the browser console (F12) for detailed error messages

## Getting Help

If you encounter issues:

1. Check [TROUBLESHOOTING.md](TROUBLESHOOTING.md) for common problems
2. Review [EXAMPLES.md](EXAMPLES.md) to compare your layout
3. Verify each configuration value matches exactly
4. Check the error message for clues

## Next Steps

After successful setup:

1. Test creating a sheet for next month
2. Verify the new sheet is formatted correctly
3. Add any month-specific data before archiving the old sheet
4. Consider setting a calendar reminder to create sheets monthly
