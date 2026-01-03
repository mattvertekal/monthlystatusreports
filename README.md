# Monthly Status Report (MSR) Automation

Automated updating of Vertekal Monthly Status Reports (MSRs) with timesheet data from TSheets.

## Overview

This tool automates the monthly process of updating three MSRs:
- **Athena TO1** - Extension Period MSR
- **Athena TO4 PIVOT** - Development (CLIN 0001AD) and O&M (CLIN 0002AD)
- **Athena TO6** - Runway Option 4 MSR

Each month, simply export timesheet data from TSheets and tell Claude Code which month to update. The tool handles:
- Parsing timesheet CSV and aggregating hours by employee/charge code
- Finding the correct month column in each MSR
- Updating employee hours with proper formatting
- Matching cell colors to previous month for consistency
- Calculating totals

## Quick Start (Claude Code Integration)

### One-Time Setup

1. **Clone this repository:**
   ```bash
   cd ~/path/to/your/skills
   git clone https://github.com/mattvertekal/monthlystatusreports.git
   ```

2. **Install dependencies:**
   ```bash
   pip install openpyxl
   ```

### Monthly Usage

1. **Export timesheet from TSheets** → save as CSV
2. **Download current MSR files** from OneDrive
3. **Tell Claude Code:**
   - Drop the timesheet CSV
   - Drop the three MSR files (TO1, TO4, TO6)
   - Say: *"Update MSRs for Jan-26"* (or whatever month)

4. **Claude Code will:**
   - Parse the timesheet
   - Find the correct column in each MSR
   - Update all employee hours
   - Return updated MSR files

5. **Upload updated MSRs** back to OneDrive

## File Structure

```
monthlystatusreports/
├── README.md                      # This file
├── skill.json                     # Claude Code skill definition
├── update_msrs.py                 # Main orchestrator script
├── agents/
│   ├── to1_updater.py            # TO1 MSR update logic
│   ├── to4_updater.py            # TO4 MSR update logic (both CLINs)
│   └── to6_updater.py            # TO6 MSR update logic
├── config/
│   ├── employee_mappings.json    # Employee → charge code → MSR row mappings
│   └── msr_settings.json         # MSR-specific settings (fill ranges, etc.)
├── utils/
│   ├── timesheet_parser.py       # CSV parsing utilities
│   └── date_finder.py            # Month column finding utilities
└── examples/
    └── (example timesheet and MSR files)
```

## Configuration Files

### `config/employee_mappings.json`

Maps each employee to their charge codes and MSR rows. Example:

```json
{
  "employees": {
    "Samuel Aldrich": {
      "msrs": ["TO1"],
      "charge_codes": {
        "Athena TO1 Ext Telework": {
          "msr": "TO1",
          "sheet": "Extension Period MSR",
          "row": 11
        }
      }
    }
  }
}
```

**When to update:** When employees change projects or new employees are added.

### `config/msr_settings.json`

Contains MSR-specific settings like status rows, fill ranges, etc.

**When to update:** When MSR formats change or new MSRs are added.

## Command-Line Usage (Optional)

You can also run the tool directly from command line:

```bash
python update_msrs.py <timesheet.csv> <to1.xlsx> <to4.xlsx> <to6.xlsx> <month>
```

Example:
```bash
python update_msrs.py \
  timesheet_2025-12-01_thru_2025-12-31.csv \
  "Athena TO1 MSR Nov 2025.xlsx" \
  "Athena TO4 MSR 2025.11.xlsx" \
  "Athena TO6 MSR Nov 2025.xlsx" \
  "Dec-25"
```

## Month Format Examples

The tool accepts various month formats:
- `Jan-26`, `Feb-26`, `Dec-25` (short form)
- `January 2026`, `February 2026` (long form)
- `2026-01`, `2026-02` (ISO format)

## How It Works

### 1. Timesheet Parsing
- Reads TSheets CSV export
- Aggregates hours by employee and charge code
- Excludes PTO and Holiday time
- Uses detailed charge codes (jobcode_2) when available

### 2. Month Column Finding
- Parses your month input (e.g., "Jan-26")
- Searches each MSR for the matching date column
- Handles various date formats in MSR headers

### 3. MSR Updating
- Updates status row to "Actual"
- Fills each employee's hours in correct rows
- Copies cell formatting from previous month
- Calculates and updates totals

### 4. Formatting Rules
- **TO1**: Each row's color matches its previous month color
- **TO4 CLIN 0001AD**: Full column (rows 3-54) filled with status color
- **TO4 CLIN 0002AD**: Rows 3-15 filled with status color
- **TO6**: Each row's color matches its previous month color

## Troubleshooting

### "Could not find column for [month]"
- Check that the MSR actually has that month
- Verify month format matches MSR headers
- MSR may need to be extended if month doesn't exist yet

### "Employee not found in mappings"
- Update `config/employee_mappings.json` with new employee
- Add their charge codes and MSR row numbers

### "Hours appear in wrong row"
- Check `config/employee_mappings.json` for correct row numbers
- Verify charge code matches exactly (including charge code number)

### Colors don't match
- Tool copies colors from previous month column
- If previous month has wrong colors, they'll propagate
- Fix manually in Excel, then colors will be correct next month

## Adding New Employees

1. Get employee's name (as it appears in TSheets)
2. Get their charge codes from TSheets
3. Find their row number in the MSR
4. Add to `config/employee_mappings.json`:

```json
"New Employee": {
  "msrs": ["TO1"],
  "charge_codes": {
    "Charge Code Name": {
      "msr": "TO1",
      "sheet": "Extension Period MSR",
      "row": 18,
      "description": "Description"
    }
  }
}
```

5. Commit changes to GitHub

## Maintenance

### Monthly
- No maintenance required if employees/MSRs haven't changed
- Just run the tool each month

### When Employees Change
- Update `config/employee_mappings.json`
- Commit and push to GitHub

### When MSR Format Changes
- Update `config/msr_settings.json`
- Test with sample data
- Commit and push to GitHub

## Backup Strategy

All code lives in GitHub for disaster recovery. Your local laptop just clones the repo. If your Mac dies:
1. Clone repo on new machine
2. Install dependencies
3. Continue where you left off

## Support

Questions or issues? Check the GitHub issues or create a new one:
https://github.com/mattvertekal/monthlystatusreports/issues

## Version History

**v1.0.0** - Initial release
- TO1, TO4, TO6 MSR automation
- Claude Code skill integration
- Configurable employee mappings
