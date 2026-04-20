# Troubleshooting Guide

This guide helps you resolve common issues with the Master Register script.

## Installation Issues

### Issue: "Register" menu doesn't appear

**Symptoms:** After pasting the script and saving, you don't see the "Register" menu in your spreadsheet.

**Solutions:**

1. **Save the script properly:**
   - In the Apps Script editor, press Ctrl+S (or Cmd+S)
   - Wait for the "Saved" indicator to appear
   - Close the Apps Script tab

2. **Refresh the spreadsheet:**
   - Press F5 (or Cmd+R on Mac)
   - Wait for the page to fully reload
   - The menu should now appear

3. **Check if code was pasted correctly:**
   - Open Apps Script again (Extensions → Apps Script)
   - Verify the `onOpen()` function exists at the bottom
   - Look for: `function onOpen() { ... }`

4. **Check browser permissions:**
   - Verify you have edit access to the spreadsheet
   - Try in a different browser
   - Clear browser cache and try again

### Issue: Authorization dialog appears on every refresh

**Symptoms:** When you open the spreadsheet, Google asks permission to use the script repeatedly.

**Solutions:**

1. Click **Review permissions** and grant the necessary permissions
2. Select your Google account
3. Click **Allow** on the permissions dialog
4. The script should remember your choice thereafter

---

## Configuration Issues

### Issue: "Could not read month/year" error

**Error message:** *"Could not read month/year. Check that monthCell and yearCell in the configuration point to the correct cells."*

**Causes:**
- `monthCell` or `yearCell` addresses are incorrect
- The cells are empty or contain invalid data
- Configuration wasn't saved after editing

**Solutions:**

1. **Verify your configuration values:**
   - Open the spreadsheet
   - Click the cell containing your month name
   - Check the Name Box (top-left) for the address
   - Update `monthCell` in the script to match exactly

2. **Check the cells contain data:**
   - Verify the month cell contains a month name (e.g., "April")
   - Verify the year cell contains a number (e.g., "2026")
   - Ensure there are no extra spaces before/after the text

3. **Resave the configuration:**
   - Update the values in the script
   - Press Ctrl+S to save
   - Close and reopen the spreadsheet

4. **Check for typos:**
   - Cell addresses are case-sensitive: `AH1` ≠ `ah1`
   - Example: `monthCell: 'AH1'` (with quotes)

### Issue: "Could not parse 'April 2026' as a date" error

**Error message:** *"Could not parse \"April 2026\" as a date. Check the monthCell setting."*

**Causes:**
- Month name is not spelled correctly
- Month name is abbreviated (e.g., "Apr" instead of "April")
- Month name has unusual capitalization
- Year is not a 4-digit number

**Solutions:**

1. **Verify month name format:**
   - Month must be the full English name
   - ✅ Correct: "April", "January", "December"
   - ❌ Wrong: "Apr", "april", "4", "04"
   - Capitalization matters (first letter uppercase)

2. **Check year format:**
   - Year must be a 4-digit number
   - ✅ Correct: "2026", "2025", "1999"
   - ❌ Wrong: "'26", "26", "2,026"

3. **Update your spreadsheet:**
   - Edit the month/year cells to use correct format
   - Month: full name, first letter capitalized
   - Year: 4-digit number with no extra characters

---

## Runtime Issues

### Issue: "Tab 'May 2026' already exists" error

**Symptoms:** You get an error saying the month tab already exists when trying to create it.

**Cause:** A sheet for next month already exists (duplicate prevention feature).

**Solutions:**

1. **Check existing sheets:**
   - Look at the sheet tabs at the bottom
   - See if a tab for next month already exists
   - Delete the old one if it's not needed, or skip creation if it is

2. **To delete an existing sheet:**
   - Right-click on the sheet tab
   - Select "Delete sheet"
   - Confirm the deletion

3. **Try again:**
   - Click Register → Create next month tab
   - If the tab exists, you'll get the same error (this prevents accidental duplicates)

### Issue: New sheet is created but day columns are wrong

**Symptoms:** The new sheet is created, but the number of day columns doesn't match the month (too many or too few).

**Causes:**
- `dayNumbersRow` is incorrect
- `firstDayCol` is incorrect
- Day numbers in the source sheet are not sequential or not in a single row

**Solutions:**

1. **Verify dayNumbersRow:**
   - In your current month sheet, find the row containing day numbers
   - Click any day number cell
   - Check the Name Box shows the correct row
   - Update `dayNumbersRow` if needed

2. **Verify firstDayCol:**
   - Click the first lesson date cell (day 1)
   - Check the Name Box for the column letter
   - Convert to number: A=1, B=2, C=3, D=4, etc.
   - Update `firstDayCol` if needed

3. **Check source sheet day numbers:**
   - Day numbers must be sequential (1, 2, 3...)
   - All in the same row
   - Starting from `firstDayCol`

4. **Resave configuration:**
   - Update the script with correct values
   - Save (Ctrl+S)
   - Close the Apps Script editor
   - Try again

### Issue: New sheet loses data or formatting

**Symptoms:** The new sheet is missing pupil data, formulas, or formatting from the original.

**Causes:**
- `firstPupilRow` or `lastPupilRow` incorrect
- Source sheet has unusual structure
- Copy operation was interrupted

**Solutions:**

1. **Verify pupil row configuration:**
   - Find your first pupil row → update `firstPupilRow`
   - Find your last pupil row → update `lastPupilRow`
   - Make sure range includes all pupil data

2. **Check source sheet before creating:**
   - Verify current month sheet has all data intact
   - Ensure no rows are hidden
   - Ensure no columns are hidden

3. **Recreate the sheet:**
   - Delete the incorrectly created sheet
   - Delete the old month's sheet if needed
   - Update configuration values
   - Try creating again

4. **Manual verification:**
   - After creating a new sheet, manually check:
     - All pupil names present
     - All formulas intact
     - All formatting preserved
     - Correct day count

---

## Permission and Access Issues

### Issue: "You do not have permission to edit this sheet" error

**Causes:**
- You don't have edit access to the spreadsheet
- The script is trying to write to a read-only sheet
- Your Google account permissions changed

**Solutions:**

1. **Check spreadsheet sharing:**
   - Ask the spreadsheet owner to give you edit access
   - Or request a copy with edit permissions

2. **Check your account:**
   - Sign out and back into your Google account
   - Verify you're using the correct account

3. **Contact spreadsheet owner:**
   - If shared with you, request edit permissions
   - If it's your sheet, check account status

### Issue: Apps Script won't save changes

**Symptoms:** You update the configuration but changes don't stick after saving.

**Solutions:**

1. **Save properly:**
   - Use Ctrl+S (or Cmd+S) in the Apps Script editor
   - Wait for "Saved" confirmation message

2. **Check for syntax errors:**
   - Make sure you didn't accidentally delete or break code
   - Look for red error indicators in the editor
   - Fix any syntax errors (missing commas, brackets, etc.)

3. **Refresh and try again:**
   - Close the Apps Script tab
   - Refresh the spreadsheet (F5)
   - Test by creating a sheet

---

## Advanced Troubleshooting

### Checking browser console for errors

1. Open your spreadsheet
2. Press F12 to open developer tools
3. Click the **Console** tab
4. Try running "Create next month tab"
5. Look for red error messages
6. Error details may help diagnose issues

### Resetting the script

If everything is broken:

1. Open Apps Script (Extensions → Apps Script)
2. Delete all code (Ctrl+A, Delete)
3. Paste the complete script fresh
4. Update configuration values again
5. Save and test

### Getting help

If you've tried these solutions:

1. Review [SETUP_GUIDE.md](SETUP_GUIDE.md) for setup steps
2. Check [EXAMPLES.md](EXAMPLES.md) to compare your layout
3. Verify each configuration value matches exactly
4. Try with a fresh test spreadsheet if available

---

## Common Success Indicators

✅ You'll know it's working when:
- "Register" menu appears in your spreadsheet
- Clicking "Create next month tab" doesn't show errors
- New sheet is created immediately after current sheet
- Month/year in new sheet matches the expected month
- Day columns match the new month's day count (28/29/30/31)
- Pupil names and formatting are preserved
- Attendance data is cleared for the new month

---

## Questions?

If you encounter an issue not listed here:

1. Check [README.md](README.md) for feature overview
2. Review [SETUP_GUIDE.md](SETUP_GUIDE.md) for installation details
3. Compare your setup with [EXAMPLES.md](EXAMPLES.md)
4. Check configuration values match your spreadsheet exactly
