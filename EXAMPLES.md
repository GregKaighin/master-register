# Configuration Examples

This file shows example configurations for different spreadsheet layouts.

## Example 1: Default Piano Lesson Register (Music Studio)

**Spreadsheet Structure:**
- Columns A-C: Student names and information
- Column D onwards: Daily lesson slots (columns D-AG for 31-day months)
- Row 6: Day numbers (1, 2, 3... 31)
- Rows 7-48: Student records

**Configuration:**
```javascript
const CFG = {
  monthCell:     'AH1',
  yearCell:      'AH2',
  dayNumbersRow: 6,
  firstDayCol:   4,     // Column D
  firstPupilRow: 7,
  lastPupilRow:  48,
};
```

**How it works:**
- Month name in cell AH1, year in AH2
- Days 1-31 appear in row 6, starting from column D
- Students are in rows 7-48, with lesson attendance recorded in columns D-AG

---

## Example 2: Compact Layout (Small Studio)

**Spreadsheet Structure:**
- Columns A-B: Student names
- Columns C onwards: Daily slots
- Row 3: Day numbers (1-31)
- Rows 4-20: Student records

**Configuration:**
```javascript
const CFG = {
  monthCell:     'B1',
  yearCell:      'B2',
  dayNumbersRow: 3,
  firstDayCol:   3,     // Column C
  firstPupilRow: 4,
  lastPupilRow:  20,
};
```

---

## Example 3: Expanded Layout (Large Studio)

**Spreadsheet Structure:**
- Columns A-E: Student details (name, level, contact, notes, rate)
- Columns F onwards: Daily slots
- Row 1: Month/year info, Row 5: Day numbers
- Rows 6-100: Student records

**Configuration:**
```javascript
const CFG = {
  monthCell:     'A1',
  yearCell:      'A2',
  dayNumbersRow: 5,
  firstDayCol:   6,     // Column F
  firstPupilRow: 6,
  lastPupilRow:  100,
};
```

---

## Example 4: Multiple Month Tracking

**Spreadsheet Structure:**
- Month/year in cells B2/C2
- Day numbers in row 4, starting from column D
- Multiple student groups
- Students in rows 6-80

**Configuration:**
```javascript
const CFG = {
  monthCell:     'B2',
  yearCell:      'C2',
  dayNumbersRow: 4,
  firstDayCol:   4,     // Column D
  firstPupilRow: 6,
  lastPupilRow:  80,
};
```

---

## How to Determine Your Values

### monthCell and yearCell
- These are usually in a header area of your spreadsheet
- Example cells: A1, B1, Z1 (top area)
- Don't need to be adjacent
- Must contain month name (e.g., "April") and year (e.g., 2026)

### dayNumbersRow
- The row that contains 1, 2, 3... through the end of the month
- All day numbers must be in the same row
- Must be in sequential order starting from 1

### firstDayCol
- The column containing the lesson slots for day 1
- Usually column C, D, E, or F (depending on how many info columns you have)
- Convert letters to numbers: A=1, B=2, C=3, D=4, etc.

### firstPupilRow and lastPupilRow
- firstPupilRow: The row containing the first student's data
- lastPupilRow: The row containing the last student's data
- Include header rows in the row count if they exist

---

## Finding Your Values Visually

### Step 1: Open your spreadsheet and click a cell

When you click any cell, the **Name Box** (top-left) shows the cell address:

```
Name Box: AH1
↓
[AH1] [fx] ← This is where you see the address
```

### Step 2: Match your layout to an example

1. Find your month and year cells → note their addresses for `monthCell` and `yearCell`
2. Find the row with day numbers → note the row number for `dayNumbersRow`
3. Find the first day column → convert the letter to a number for `firstDayCol`
4. Find the first and last student rows → note the numbers for `firstPupilRow` and `lastPupilRow`

### Step 3: Update your configuration

Replace the values in the `CFG` object with your values.

---

## Common Cell Address Patterns

| Cell | Name Box Shows | Usage |
|------|---|---|
| First cell in sheet | A1 | Often used for month/year |
| Further right | AH1 | Common for month name storage |
| Below month | AH2 | Year cell (one row below month) |
| Far right column | AZ5 | Could be any configuration value |

---

## Column Number Reference

| Column | Number | Column | Number |
|--------|--------|--------|--------|
| A | 1 | N | 14 |
| B | 2 | O | 15 |
| C | 3 | P | 16 |
| D | 4 | ... | ... |
| E | 5 | Z | 26 |
| F | 6 | AA | 27 |
| G | 7 | AB | 28 |
| H | 8 | AC | 29 |
| I | 9 | AD | 30 |
| J | 10 | AE | 31 |
| K | 11 | AF | 32 |
| L | 12 | AG | 33 |
| M | 13 | AH | 34 |

---

## Testing Your Configuration

Once configured, test with these steps:

1. Make sure the current sheet has valid month/year values
2. Click **Register** → **Create next month tab**
3. Verify:
   - ✅ New sheet created with correct month/year
   - ✅ Day columns match the new month's length
   - ✅ Student names preserved
   - ✅ Attendance data cleared
   - ✅ Formatting maintained

If anything looks wrong, compare with the examples above and adjust configuration values.

---

## Questions or Issues?

See [TROUBLESHOOTING.md](TROUBLESHOOTING.md) for common issues and solutions.
