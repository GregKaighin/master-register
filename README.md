# Master Register - Monthly Tab Creator

Automate the monthly creation of new register sheets in Google Sheets. This Google Apps Script eliminates manual copying, formatting, and adjustment work by automatically generating next month's register tab with all settings configured correctly.

## Features

✅ **One-Click Sheet Creation** — Creates next month's register sheet with a single click from a menu  
✅ **Automatic Month Calculation** — Reads current month/year and calculates the next month automatically  
✅ **Smart Day Adjustment** — Automatically adds or removes columns to match the new month's day count  
✅ **Format Preservation** — Maintains all formatting, formulas, and structure from the current sheet  
✅ **Data Clearing** — Clears attendance data while preserving pupil names and structure  
✅ **Error Handling** — Validates configuration and provides helpful error messages  
✅ **Duplicate Prevention** — Won't create a tab if one for that month already exists  

## Installation

### Quick Start (2 minutes)

1. **Open your Master Register spreadsheet** in Google Sheets
2. **Click** `Extensions` → `Apps Script`
3. **Delete** any existing code and paste the entire contents of `master_register.gs`
4. **Update** the configuration (see [Setup Guide](#setup-guide))
5. **Save** (Ctrl+S) and close the script editor
6. **Refresh** your spreadsheet — a new "Register" menu will appear
7. **Click** "Register" → "Create next month tab" to test

For detailed setup instructions, see [SETUP_GUIDE.md](SETUP_GUIDE.md).

## Usage

1. Navigate to your current month's register sheet
2. Click the **Register** menu in the top-right corner
3. Select **Create next month tab**
4. A confirmation dialog appears with the new sheet name and day count
5. The new sheet is created immediately after the current sheet

## Configuration

The script requires 6 configuration values that match your spreadsheet layout:

```javascript
const CFG = {
  monthCell:     'AH1', // Cell containing the month name
  yearCell:      'AH2', // Cell containing the year
  dayNumbersRow: 6,     // Row containing day numbers (1, 2, 3...)
  firstDayCol:   4,     // First lesson date column (D=4)
  firstPupilRow: 7,     // First row with pupil data
  lastPupilRow:  48,    // Last row with pupil data
};
```

**Important:** These values must match your actual spreadsheet layout. See [SETUP_GUIDE.md](SETUP_GUIDE.md) for step-by-step instructions to find your values.

### How to Find Your Configuration Values

1. **monthCell & yearCell:**
   - Click the cell containing your month name
   - Look at the **Name Box** (top-left) — it shows the cell address (e.g., AH1)
   - Repeat for the year cell

2. **dayNumbersRow:**
   - Find the row containing your day numbers (1, 2, 3, 4...)
   - Click any day number cell
   - The row number appears in the Name Box

3. **firstDayCol & firstPupilRow:**
   - Click the first lesson date cell in your grid
   - The Name Box shows the column letter and row number

4. **lastPupilRow:**
   - Scroll to the bottom of your pupil data
   - Click the last pupil's row
   - Note the row number

See [EXAMPLES.md](EXAMPLES.md) for example spreadsheet layouts.

## Troubleshooting

### "Could not read month/year" error
- Verify `monthCell` and `yearCell` in configuration match your actual cells
- Ensure those cells contain the month name (e.g., "April") and year (e.g., 2026)

### "Could not parse month/year as a date" error
- Month name must be a full English month name (e.g., "April", not "Apr" or "04")
- Check spelling and capitalization

### New tab not created, no error message
- Check if a tab with that month/year already exists (duplicate prevention)
- Verify you have permission to edit the spreadsheet
- Try creating the sheet manually first to test permissions

### Sheet created but day columns are wrong
- Verify `firstDayCol` and `dayNumbersRow` match your layout
- Ensure day numbers are in a single row in sequential order (1, 2, 3...)

See [TROUBLESHOOTING.md](TROUBLESHOOTING.md) for more help.

## What Gets Copied?

✅ **Copied to new sheet:**
- All formatting (colors, fonts, borders, sizing)
- Pupil names and structure
- Formulas and calculations (except attendance data)
- Headers and labels

❌ **NOT copied:**
- Attendance data (lesson numbers in day columns)
- These are automatically cleared for the new month

## File Structure

```
master-register/
├── master_register.gs      # Main script file
├── README.md               # This file
├── SETUP_GUIDE.md         # Detailed setup instructions
├── EXAMPLES.md            # Configuration examples
├── TROUBLESHOOTING.md     # Common issues and solutions
├── CHANGELOG.md           # Version history
├── LICENSE                # MIT License
└── .gitignore            # Git ignore file
```

## Requirements

- Google Sheets account
- Access to edit the spreadsheet
- Spreadsheet must have at least one "current month" sheet set up
- Month name must be in full English format (January, February, etc.)

## Compatibility

- Works with any Google Sheets spreadsheet
- Compatible with all modern browsers
- No external dependencies required

## Version History

See [CHANGELOG.md](CHANGELOG.md) for version history and updates.

## License

This project is licensed under the MIT License — see [LICENSE](LICENSE) file for details.

## Contributing

Found a bug or have a feature request? Please create an issue or pull request.

## Support

For help with:
- **Setup issues** → See [SETUP_GUIDE.md](SETUP_GUIDE.md)
- **Common problems** → See [TROUBLESHOOTING.md](TROUBLESHOOTING.md)
- **Configuration examples** → See [EXAMPLES.md](EXAMPLES.md)

---

Made for Piano Lessons with Greg Kaighin
