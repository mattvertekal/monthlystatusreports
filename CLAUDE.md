# Monthly Status Reports Agent

Automates Monthly Status Report (MSR) updates with timesheet data from TSheets.

## Quick Start

When the user says "update MSRs", "run MSR for [month]", or similar:

1. Ask which month to report
2. Query TSheets API for employee hours
3. Show hours summary for confirmation
4. Update each MSR Excel file with hours
5. Save updated files to output directory

## Project Structure

```
monthlystatusreports/
├── CLAUDE.md              # This file
├── update_msrs.py         # Main orchestrator
├── agents/
│   ├── to1_updater.py     # Athena TO1 MSR updater
│   ├── to4_updater.py     # Athena TO4 MSR updater
│   ├── to6_updater.py     # Athena TO6 MSR updater
│   └── wsr_updater.py     # Weekly Status Report updater (handles Emmett)
├── config/
│   ├── employee_mappings.json  # Employee -> charge code -> row mappings
│   ├── msr_settings.json       # MSR sheet configurations
│   └── tsheets_config.json     # TSheets API credentials and user mappings
└── utils/
    ├── timesheet_parser.py    # TSheets API client + CSV parser
    └── date_finder.py         # Date/column finding utilities
```

## Usage

### Update MSRs using TSheets API (recommended)

```bash
python update_msrs.py --api to1.xlsx to4.xlsx to6.xlsx "Jan-26"
```

### Update MSRs using CSV export (legacy)

```bash
python update_msrs.py timesheet.csv to1.xlsx to4.xlsx to6.xlsx "Jan-26"
```

### Update Emmett/WSR

```bash
# Monthly (Emmett MSR)
python agents/emmett_updater.py wsr_file.xlsx "Jan-26"

# Weekly (WSR)
python agents/wsr_updater.py wsr_file.xlsb --week 2026-01-05
```

### Fetch hours only (no MSR update)

```python
from utils.timesheet_parser import get_timesheets_for_month, print_timesheet_summary

data = get_timesheets_for_month(2026, 1)  # January 2026
print_timesheet_summary(data)
```

## TSheets API

**Base URL:** `https://rest.tsheets.com/api/v1`
**Token:** See `config/tsheets_config.json`
**Token Expiry:** 2026-03-09

### Tracked Employees

| Name | TSheets ID | MSRs |
|------|------------|------|
| Samuel Aldrich | 7326536 | TO1 |
| Keith Mosley | 7326538 | TO1 |
| Samuel Martin | 3285612 | TO1 |
| Matthew Nicely | 3285614 | TO4 |
| Greg Mihokovich | 3377124 | TO4 |
| Neil Franklin | 7326540 | TO4 |
| Ryan Robertson | 3362664 | TO4 |
| Rachel Palmer | 7326498 | TO6 |
| Daniel Quillen | 8162260 | TO6 |
| David Thompson | 8499572 | Emmett |
| Nathan Ruf | 8131040 | Emmett |

## Contracts

### Athena TO1
- **Sheet:** Extension Period MSR
- **Employees:** Samuel Aldrich, Keith Mosley, Samuel Martin
- **Charge Codes:** Athena TO1 Ext Telework, Athena TO1 Ext

### Athena TO4 PIVOT
- **Sheets:** CLIN 0001AD (Development), CLIN 0002AD (O&M)
- **Employees:** Matthew Nicely, Greg Mihokovich, Neil Franklin, Ryan Robertson
- **Charge Codes:** AB11662.004.03.* codes

### Athena TO6
- **Sheet:** Option 4 MSR
- **Employees:** Rachel Palmer, Daniel Quillen
- **Charge Codes:** Athena TO6 CLIN 0005

### Emmett (Vertekal -> Booz Allen)
- **Sheet:** CLIN Level Detail
- **Employees:** David Thompson, Nathan Ruf, Philip Yang
- **Charge Code:** Emmett – Magni HA – R2026Q1
- **Hourly Rates:** $211.15 (Thompson, Yang), $187.41 (Ruf)

## Adding New Employees

1. Update `config/employee_mappings.json`:
```json
{
  "New Employee": {
    "msrs": ["TO4"],
    "tsheets_id": "12345",
    "charge_codes": {
      "Charge Code Name": {
        "msr": "TO4",
        "sheet": "CLIN 0001AD",
        "row": 15,
        "description": "Description"
      }
    }
  }
}
```

2. Update `config/tsheets_config.json` users mapping:
```json
{
  "users": {
    "12345": "New Employee"
  }
}
```

## Dependencies

```bash
pip3 install requests openpyxl xlwings
```

- **requests:** TSheets API calls
- **openpyxl:** .xlsx file manipulation
- **xlwings:** .xlsb file manipulation (requires Excel on Mac)
