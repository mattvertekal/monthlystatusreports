# Monthly Status Report (MSR) Automation

Automated updating of Vertekal Monthly Status Reports (MSRs) and Weekly Status Reports (WSRs) with timesheet data from TSheets API.

## Overview

This tool automates the monthly and weekly reporting process:

**MSR Agents (Monthly):**
- **Athena TO1** - Extension Period MSR
- **Athena TO4 PIVOT** - Development (CLIN 0001AD) and O&M (CLIN 0002AD)
- **Athena TO6** - Runway Option 4 MSR

**WSR Agent (Weekly/Monthly):**
- **Weekly** - Updates CLIN Level Detail with hours, changes Estimate→Actual
- **Monthly** - Rolls up weeks to Data tab for invoicing

## Quick Start

### One-Time Setup

1. **Clone this repository:**
   ```bash
   cd ~/Documents
   git clone https://github.com/mattvertekal/monthlystatusreports.git
   ```

2. **Install dependencies:**
   ```bash
   pip3 install requests openpyxl xlwings
   ```

3. **Configure TSheets API:**
   ```bash
   cp config/tsheets_config.template.json config/tsheets_config.json
   # Edit with your API token and user mappings
   ```

4. **Set up folder structure:**
   ```bash
   mkdir -p ~/Documents/MSRs/templates ~/Documents/MSRs/completed
   mkdir -p ~/Documents/WSR/templates ~/Documents/WSR/completed
   ```

### Monthly MSR Usage

Simply tell Claude Code:
> "Run MSRs for January"

The tool will:
1. Query TSheets API for the month's hours
2. Find the latest MSR files (carry-forward from previous month)
3. Update all employee hours with proper formatting
4. Save to `~/Documents/MSRs/completed/2026/01-Jan/`

### Weekly WSR Usage

Tell Claude Code:
> "Run WSR for Jan 12-16"

The tool will:
1. Query TSheets for the week's hours
2. Find the latest WSR file
3. Update CLIN Level Detail with hours
4. Change status from Estimate → Actual
5. Apply highlighting
6. Save to `~/Documents/WSR/completed/2026/Q1/`

### Monthly WSR Roll-up

At month end, tell Claude Code:
> "Roll up WSR for January"

This sums all weekly hours to the Data tab for invoicing.

## File Structure

```
monthlystatusreports/
├── README.md                      # This file
├── CLAUDE.md                      # Claude Code instructions
├── ARCHITECTURE.md                # Business process diagrams
├── update_msrs.py                 # Main MSR orchestrator
├── agents/
│   ├── to1_updater.py            # TO1 MSR update logic
│   ├── to4_updater.py            # TO4 MSR update logic
│   ├── to6_updater.py            # TO6 MSR update logic
│   └── wsr_updater.py            # WSR weekly + monthly logic
├── config/
│   ├── employee_mappings.json    # Employee → charge code → row mappings
│   ├── msr_settings.json         # MSR-specific settings
│   ├── tsheets_config.json       # TSheets API credentials (gitignored)
│   └── tsheets_config.template.json  # Template for credentials
└── utils/
    ├── timesheet_parser.py       # TSheets API client
    └── date_finder.py            # Date/column utilities
```

### Output Folders

```
~/Documents/MSRs/
├── templates/                    # Fallback templates
└── completed/
    └── 2026/
        ├── 01-Jan/
        │   ├── Athena TO1 Vertekal MSR Jan 2026.xlsx
        │   ├── Athena TO4_PIVOT_OP3_Vertekal MSR_2026.01.xlsx
        │   └── Athena TO6 Vertekal MSR Opt3 January 2026.xlsx
        └── 02-Feb/
            └── ...

~/Documents/WSR/
├── templates/
│   └── Vertekal- Draft WSR.xlsb
└── completed/
    └── 2026/
        └── Q1/
            ├── Vertekal_WSR_2026-01-05_to_2026-01-09.xlsb
            └── ...
```

## Carry-Forward Workflow

Each period's output becomes the next period's input:

```
MSR Chain (Monthly)
Dec 2025 MSR ──→ Jan 2026 MSR ──→ Feb 2026 MSR ──→ ...

WSR Chain (Weekly)
Jan 5-9 WSR ──→ Jan 12-16 WSR ──→ Jan 19-23 WSR ──→ ...
```

No need to manually specify input files - the tool auto-finds the latest.

## Tracked Employees

### Athena Contracts (MSR)

| Name | TSheets ID | Contract |
|------|------------|----------|
| Samuel Aldrich | 7326536 | TO1 |
| Keith Mosley | 7326538 | TO1 |
| Samuel Martin | 3285612 | TO1 |
| Matthew Nicely | 3285614 | TO4 |
| Greg Mihokovich | 3377124 | TO4 |
| Neil Franklin | 7326540 | TO4 |
| Ryan Robertson | 3362664 | TO4 |
| Rachel Palmer | 7326498 | TO6 |
| Daniel Quillen | 8162260 | TO6 |

### Vertekal Subcontract (WSR)

| Name | TSheets ID | Hourly Rate |
|------|------------|-------------|
| David Thompson | 8499572 | $211.15 |
| Nathan Ruf | 8131040 | $187.41 |
| Philip Yang | - | $211.15 |

## Command-Line Usage

### MSR Update (API mode)
```bash
python update_msrs.py --api "Jan-26"
```

### WSR Weekly Update
```bash
python agents/wsr_updater.py --weekly 2026-01-12
```

### WSR Monthly Roll-up
```bash
python agents/wsr_updater.py --monthly 2026-01
```

### Fetch Hours Only (no update)
```python
from utils.timesheet_parser import get_timesheets_for_month, print_timesheet_summary

data = get_timesheets_for_month(2026, 1)  # January 2026
print_timesheet_summary(data)
```

## TSheets API

- **Base URL:** `https://rest.tsheets.com/api/v1`
- **Token:** Stored in `config/tsheets_config.json`
- **Endpoints used:** `/timesheets`, `/users`, `/jobcodes`

## Adding New Employees

1. Update `config/employee_mappings.json`:
```json
"New Employee": {
  "msrs": ["TO4"],
  "charge_codes": {
    "Charge Code Name": {
      "msr": "TO4",
      "sheet": "CLIN 0001AD",
      "row": 15,
      "description": "Description"
    }
  }
}
```

2. Update `config/tsheets_config.json` users mapping:
```json
"users": {
  "12345": "New Employee"
}
```

3. Commit and push to GitHub

## Troubleshooting

### "Could not find column for [month]"
- MSR may need month column added manually first
- Check date format matches MSR headers

### "Employee not found in mappings"
- Add to `config/employee_mappings.json`
- Add TSheets ID to `config/tsheets_config.json`

### TSheets API errors
- Check token hasn't expired (see `config/tsheets_config.json`)
- Verify network connectivity

### WSR picks up temp file
- Close Excel before running WSR agent
- Tool filters `~$` files but Excel must be closed

## Dependencies

| Package | Purpose |
|---------|---------|
| `requests` | TSheets API calls |
| `openpyxl` | .xlsx file manipulation |
| `xlwings` | .xlsb file manipulation (requires Excel on Mac) |

## Documentation

- **CLAUDE.md** - Quick reference for Claude Code
- **ARCHITECTURE.md** - Visual business process diagrams

## Version History

- **v2.0.0** - TSheets API integration, auto-find, WSR agent, carry-forward
- **v1.0.0** - Initial release (CSV-based)
