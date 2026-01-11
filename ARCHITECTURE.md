# Business Process Architecture

## System Overview

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                              TSHEETS                                         │
│                         (Time Tracking)                                      │
│                                                                              │
│   Employees log hours daily → Jobcodes (charge codes) → Duration            │
└─────────────────────────────────────┬───────────────────────────────────────┘
                                      │
                                      │ API Query
                                      ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                        TIMESHEET PARSER                                      │
│                   (utils/timesheet_parser.py)                                │
│                                                                              │
│   • Fetches hours by date range                                              │
│   • Maps TSheets user IDs → Employee names                                   │
│   • Returns: {Employee: {Charge Code: Hours}}                                │
└───────────┬─────────────────────────┬─────────────────────────┬─────────────┘
            │                         │                         │
            ▼                         ▼                         ▼
┌───────────────────┐   ┌───────────────────┐   ┌───────────────────────────┐
│   MSR AGENTS      │   │   WSR AGENT       │   │   INVOICING AGENT         │
│   (Monthly)       │   │   (Weekly/Monthly)│   │   (Monthly)               │
├───────────────────┤   ├───────────────────┤   ├───────────────────────────┤
│                   │   │                   │   │                           │
│ • TO1 Updater     │   │ WEEKLY MODE:      │   │ • Pull hours from TSheets │
│ • TO4 Updater     │   │ • Update CLIN     │   │ • Calculate: hrs × rate   │
│ • TO6 Updater     │   │   Level Detail    │   │ • Preview invoice lines   │
│                   │   │ • Estimate→Actual │   │ • Update QBO invoice      │
│                   │   │ • Apply highlight │   │                           │
│ Input: Previous   │   │                   │   │ Connects to:              │
│   month's MSR     │   │ MONTHLY MODE:     │   │ QuickBooks Online API     │
│                   │   │ • Sum all weeks   │   │                           │
│ Output: Updated   │   │ • Roll up to      │   │                           │
│   MSR with hours  │   │   Data tab        │   │                           │
│                   │   │ • Ready for       │   │                           │
│                   │   │   invoicing       │   │                           │
└────────┬──────────┘   └────────┬──────────┘   └─────────────┬─────────────┘
         │                       │                             │
         ▼                       ▼                             ▼
┌───────────────────┐   ┌───────────────────┐   ┌───────────────────────────┐
│ ~/Documents/MSRs/ │   │ ~/Documents/WSR/  │   │ QuickBooks Online         │
│                   │   │                   │   │                           │
│ completed/        │   │ completed/        │   │ • Invoices updated        │
│  └─2026/          │   │  └─2026/          │   │ • Ready to send           │
│     └─01-Jan/     │   │     └─Q1/         │   │                           │
│        └─TO1_MSR  │   │        └─WSR_     │   │                           │
│        └─TO4_MSR  │   │          Jan-12   │   │                           │
│        └─TO6_MSR  │   │                   │   │                           │
└───────────────────┘   └───────────────────┘   └───────────────────────────┘
         │                       │                             │
         ▼                       ▼                             ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                           MANUAL ACTIONS                                     │
│                                                                              │
│   MSRs:  Email to primes → Drop in Vertekal OneDrive                        │
│   WSR:   Review Data tab → Feeds into invoicing                             │
│   Invoice: Send from QBO                                                     │
└─────────────────────────────────────────────────────────────────────────────┘
```

## Carry-Forward Flow

### MSR Chain (Monthly)
```
Dec 2025 MSR ──→ Jan 2026 MSR ──→ Feb 2026 MSR ──→ ...
     │                │                │
     └── Input ──────→└── Input ──────→└── Input
```

### WSR Chain (Weekly → Monthly)
```
Jan 5-9 WSR ──→ Jan 12-16 WSR ──→ Jan 19-23 WSR ──→ ... ──→ Monthly Roll-up
     │                │                 │                        │
     └── Input ──────→└── Input ───────→└── Input               ▼
                                                           Data Tab
                                                               │
                                                               ▼
                                                          Invoicing
```

## Trigger Summary

| Trigger Phrase | Agent | Action |
|----------------|-------|--------|
| "Run MSRs for January" | MSR | Update TO1, TO4, TO6 with monthly hours |
| "Run WSR for Jan 12-16" | WSR Weekly | Update CLIN Level Detail |
| "Roll up WSR for January" | WSR Monthly | Sum weeks → Data tab |
| "Preview invoice for Emmett" | Invoicing | Calculate invoice from TSheets |

## Folder Structure

### MSRs
```
~/Documents/MSRs/
├── templates/                    # Fallback templates
│   ├── TO1_MSR.xlsx
│   ├── TO4_MSR.xlsx
│   └── TO6_MSR.xlsx
└── completed/                    # Chain of updated MSRs
    └── 2026/
        ├── 01-Jan/
        │   ├── Athena TO1 Vertekal MSR Jan 2026.xlsx
        │   ├── Athena TO4_PIVOT_OP3_Vertekal MSR_2026.01.xlsx
        │   └── Athena TO6 Vertekal MSR Opt3 January 2026.xlsx
        └── 02-Feb/
            └── ...
```

### WSR
```
~/Documents/WSR/
├── templates/
│   └── Vertekal- Draft WSR.xlsb
└── completed/
    └── 2026/
        └── Q1/
            ├── Vertekal_WSR_2026-01-05_to_2026-01-09.xlsb
            ├── Vertekal_WSR_2026-01-12_to_2026-01-16.xlsb
            └── ...
```

### Invoicing
```
~/Documents/Invoicing/
├── qbo_auth.py                   # OAuth authentication
├── qbo_tokens.json               # Access/refresh tokens
├── agents/
│   └── invoice_updater.py
├── config/
│   └── invoice_mappings.json
└── utils/
    └── qbo_client.py
```

## Contracts & Employees

### Athena Contracts (MSR Agents)

| Contract | Employees | Charge Codes |
|----------|-----------|--------------|
| TO1 | Samuel Aldrich, Keith Mosley, Samuel Martin | Athena TO1 Ext Telework, Athena TO1 Ext |
| TO4 | Matthew Nicely, Greg Mihokovich, Neil Franklin, Ryan Robertson | AB11662.004.03.* codes |
| TO6 | Rachel Palmer, Daniel Quillen | Athena TO6 CLIN 0005 |

### Vertekal Subcontract (WSR + Invoicing)

| Employee | TSheets ID | Hourly Rate | Project |
|----------|------------|-------------|---------|
| David Thompson | 8499572 | $211.15 | Emmett (Magni HA) |
| Nathan Ruf | 8131040 | $187.41 | Emmett (Magni HA) |
| Philip Yang | - | $211.15 | Emmett (Magni HA) |

## Data Flow Details

### 1. TSheets → Timesheet Parser
- API Token stored in `config/tsheets_config.json`
- User ID mapping: TSheets ID → Employee Name
- Jobcode mapping: TSheets Jobcode → Charge Code
- Skip list: PTO, Holiday, etc.

### 2. Timesheet Parser → MSR Agents
- Query by month (first day to last day)
- Returns hours grouped by employee and charge code
- MSR agents look up row numbers from `employee_mappings.json`

### 3. Timesheet Parser → WSR Agent
- Query by week (Monday to Friday)
- Returns hours grouped by employee
- WSR agent updates specific column based on week label

### 4. WSR Data Tab → Invoicing
- Monthly roll-up populates Data tab
- Data tab contains: Employee, Hours, Rate, Cost, Month
- Invoice agent reads this for invoice line items

### 5. Invoicing → QuickBooks Online
- OAuth 2.0 authentication (auto-refresh tokens)
- Find/create invoice for customer
- Update line items with hours × rate
- Currently using Sandbox (switch to Production when ready)
