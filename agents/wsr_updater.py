"""
WSR (Weekly Status Report) Updater Agent
Handles both weekly updates and monthly roll-ups for the Vertekal subcontract.

Two modes:
1. Weekly: Update CLIN Level Detail with hours from TSheets
2. Monthly: Roll up weekly hours to Data tab for invoicing
"""

import json
import copy
from datetime import datetime, timedelta
from calendar import monthrange
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import sys
sys.path.append(str(Path(__file__).parent.parent))

# Employee rates (for cost calculation)
EMPLOYEE_RATES = {
    "David Thompson": 211.15,
    "Nathan Ruf": 187.41,
    "Philip Yang": 211.15
}

# WSR folder paths
WSR_BASE_DIR = Path.home() / "Documents" / "WSR"
TEMPLATES_DIR = WSR_BASE_DIR / "templates"
COMPLETED_DIR = WSR_BASE_DIR / "completed"


def get_week_dates(date_str: Optional[str] = None) -> Tuple[str, str]:
    """
    Get the Monday-Friday date range for a given week.

    Args:
        date_str: Date string (YYYY-MM-DD) or None for most recent complete week

    Returns:
        Tuple of (start_date, end_date) in YYYY-MM-DD format
    """
    if date_str:
        target_date = datetime.strptime(date_str, "%Y-%m-%d")
    else:
        # Find most recent completed Friday
        today = datetime.now()
        days_since_friday = (today.weekday() - 4) % 7
        if days_since_friday == 0 and today.hour < 18:
            days_since_friday = 7
        last_friday = today - timedelta(days=days_since_friday)
        target_date = last_friday

    # Find Monday of that week
    days_to_monday = target_date.weekday()
    monday = target_date - timedelta(days=days_to_monday)
    friday = monday + timedelta(days=4)

    return monday.strftime("%Y-%m-%d"), friday.strftime("%Y-%m-%d")


def format_week_label(start_date: str, end_date: str) -> str:
    """Format week as label like 'Jan 5-9'."""
    start = datetime.strptime(start_date, "%Y-%m-%d")
    end = datetime.strptime(end_date, "%Y-%m-%d")

    if start.month == end.month:
        return f"{start.strftime('%b')} {start.day}-{end.day}"
    else:
        return f"{start.strftime('%b')} {start.day}-{end.strftime('%b')} {end.day}"


def get_tsheets_hours_for_week(start_date: str, end_date: str) -> Dict[str, float]:
    """Fetch hours from TSheets for a specific week."""
    import requests

    config_dir = Path(__file__).parent.parent / 'config'
    with open(config_dir / 'tsheets_config.json') as f:
        config = json.load(f)

    headers = {"Authorization": f"Bearer {config['api_token']}"}
    base_url = config['base_url']

    params = {"start_date": start_date, "end_date": end_date}
    response = requests.get(f"{base_url}/timesheets", headers=headers, params=params)
    response.raise_for_status()
    timesheets = response.json().get('results', {}).get('timesheets', {})

    user_map = config['users']
    hours = {}

    for ts_id, ts in timesheets.items():
        user_id = str(ts['user_id'])
        duration = ts.get('duration', 0) / 3600.0

        if user_id not in user_map:
            continue

        emp_name = user_map[user_id]
        hours[emp_name] = hours.get(emp_name, 0) + duration

    return hours


def find_latest_wsr() -> Optional[Path]:
    """Find the most recent WSR file."""
    # Search completed folders in reverse chronological order
    now = datetime.now()
    year = now.year
    quarter = (now.month - 1) // 3 + 1

    for _ in range(8):  # Search up to 8 quarters back
        folder = COMPLETED_DIR / str(year) / f"Q{quarter}"
        if folder.exists():
            # Filter out temp files (starting with ~$)
            files = sorted([f for f in folder.glob("*.xlsb") if not f.name.startswith("~$")], reverse=True)
            if files:
                return files[0]

        quarter -= 1
        if quarter < 1:
            quarter = 4
            year -= 1

    # Fall back to template
    template = TEMPLATES_DIR / "Vertekal- Draft WSR.xlsb"
    if template.exists():
        return template

    return None


def find_week_column(sheet, week_label: str, year: int) -> Optional[int]:
    """Find the column for a specific week."""
    for col in range(100, 200):
        cell_value = sheet.cells(3, col).value
        if cell_value and week_label in str(cell_value):
            # Verify year by checking nearby Total column
            for check_col in range(col, col + 10):
                check_val = sheet.cells(3, check_col).value
                if check_val and "Total" in str(check_val) and str(year) in str(check_val):
                    return col
            return col
    return None


def find_month_total_column(sheet, year: int, month: int) -> Optional[int]:
    """Find the monthly total column (e.g., 'January 2026 Total')."""
    month_name = datetime(year, month, 1).strftime("%B")
    search_term = f"{month_name} {year} Total"

    for col in range(100, 200):
        cell_value = sheet.cells(3, col).value
        if cell_value and search_term in str(cell_value):
            return col
    return None


def get_weeks_in_month(year: int, month: int) -> List[Tuple[str, str]]:
    """Get all work weeks (Mon-Fri) that fall within a month."""
    weeks = []
    first_day = datetime(year, month, 1)
    last_day = datetime(year, month, monthrange(year, month)[1])

    # Find first Monday on or before the 1st
    current = first_day
    while current.weekday() != 0:  # Monday
        current -= timedelta(days=1)

    while current <= last_day:
        week_start = current
        week_end = current + timedelta(days=4)  # Friday

        # Include week if any part falls in the target month
        if week_start.month == month or week_end.month == month:
            # Adjust dates to stay within month boundaries for partial weeks
            if week_start.month != month:
                week_start = first_day
            if week_end.month != month:
                week_end = last_day

            weeks.append((week_start.strftime("%Y-%m-%d"), week_end.strftime("%Y-%m-%d")))

        current += timedelta(days=7)

    return weeks


# =============================================================================
# WEEKLY UPDATE
# =============================================================================

def update_weekly(wsr_path: str, week_start: str, week_end: str, output_path: Optional[str] = None) -> Dict:
    """
    Update WSR with weekly hours from TSheets.

    Args:
        wsr_path: Path to the WSR file
        week_start: Week start date (YYYY-MM-DD)
        week_end: Week end date (YYYY-MM-DD)
        output_path: Output path (default: completed/YYYY/Q#/)

    Returns:
        Result dictionary
    """
    try:
        import xlwings as xw
    except ImportError:
        return {'error': 'xlwings not installed. Run: pip3 install xlwings'}

    # Fetch hours from TSheets
    print(f"Fetching TSheets data for {week_start} to {week_end}...")
    hours = get_tsheets_hours_for_week(week_start, week_end)

    week_label = format_week_label(week_start, week_end)
    year = int(week_start[:4])

    print(f"\nWeek: {week_label} {year}")
    print(f"Hours retrieved:")
    for emp, hrs in sorted(hours.items()):
        print(f"  {emp}: {hrs:.2f} hrs")

    # Open Excel
    app = xw.App(visible=False)
    wb = app.books.open(wsr_path)

    try:
        sheet = wb.sheets["CLIN Level Detail"]

        # Find week column
        target_col = find_week_column(sheet, week_label, year)
        if not target_col:
            return {'error': f"Could not find column for week {week_label} {year}"}

        print(f"\nTarget column: {target_col}")

        # Employee rows
        employee_rows = {
            "David Thompson": 4,
            "Nathan Ruf": 5,
            "Philip Yang": 6
        }

        # Highlight color
        highlight_color = (217, 226, 243)

        # Copy formatting from previous column
        source_col = target_col - 1

        # Update status row (Estimate -> Actual)
        sheet.cells(2, target_col).value = "Actual"

        # Update employee rows
        updates = []
        total_hours = 0.0

        for emp_name, row in employee_rows.items():
            emp_hours = hours.get(emp_name, 0)

            # Copy format from source column
            source_range = sheet.range((row, source_col))
            target_range = sheet.range((row, target_col))
            source_range.copy()
            target_range.paste(paste='formats')

            # Write hours and apply highlight
            sheet.cells(row, target_col).value = emp_hours
            sheet.cells(row, target_col).color = highlight_color

            total_hours += emp_hours
            updates.append({'employee': emp_name, 'row': row, 'hours': emp_hours})

        # Set output path
        if output_path is None:
            quarter = (int(week_start[5:7]) - 1) // 3 + 1
            output_dir = COMPLETED_DIR / str(year) / f"Q{quarter}"
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = str(output_dir / f"Vertekal_WSR_{week_start}_to_{week_end}.xlsb")

        wb.save(output_path)
        print(f"\nSaved to: {output_path}")

        return {
            'success': True,
            'mode': 'weekly',
            'week': week_label,
            'year': year,
            'total_hours': total_hours,
            'updates': updates,
            'output_path': output_path
        }

    finally:
        wb.close()
        app.quit()


# =============================================================================
# MONTHLY ROLL-UP
# =============================================================================

def rollup_monthly(wsr_path: str, year: int, month: int, output_path: Optional[str] = None) -> Dict:
    """
    Roll up weekly hours to Data tab for monthly invoicing.

    Args:
        wsr_path: Path to the WSR file
        year: Year
        month: Month (1-12)
        output_path: Output path (default: same file)

    Returns:
        Result dictionary
    """
    try:
        import xlwings as xw
    except ImportError:
        return {'error': 'xlwings not installed. Run: pip3 install xlwings'}

    month_name = datetime(year, month, 1).strftime("%B")
    print(f"Rolling up {month_name} {year} to Data tab...")

    # Open Excel
    app = xw.App(visible=False)
    wb = app.books.open(wsr_path)

    try:
        clin_sheet = wb.sheets["CLIN Level Detail"]
        data_sheet = wb.sheets["Data"]

        # Employee rows and their data from CLIN Level Detail
        employees = {}
        for row in [4, 5, 6]:
            emp_name = clin_sheet.cells(row, 2).value  # Column B
            if emp_name and emp_name != "Employee":
                employees[emp_name] = {
                    'row': row,
                    'position': clin_sheet.cells(row, 1).value,      # A - Position (PLC)
                    'clin': clin_sheet.cells(row, 3).value,          # C - CLIN
                    'detail': clin_sheet.cells(row, 4).value,        # D - Code/Detail
                    'rate': clin_sheet.cells(row, 5).value,          # E - Rate
                    'wbs': clin_sheet.cells(row, 6).value,           # F - WBS
                    'charge_no': clin_sheet.cells(row, 14).value,    # N - Booz Allen Charge No
                }

        # Get weeks in this month and sum hours
        weeks = get_weeks_in_month(year, month)
        print(f"Weeks in {month_name}: {len(weeks)}")

        monthly_hours = {emp: 0.0 for emp in employees}

        for week_start, week_end in weeks:
            week_label = format_week_label(week_start, week_end)
            col = find_week_column(clin_sheet, week_label, year)

            if col:
                status = clin_sheet.cells(2, col).value
                if status == "Actual":
                    for emp_name, emp_data in employees.items():
                        hrs = clin_sheet.cells(emp_data['row'], col).value or 0
                        monthly_hours[emp_name] += float(hrs)
                    print(f"  {week_label}: Found (Actual)")
                else:
                    print(f"  {week_label}: Found but status is '{status}' - skipping")
            else:
                print(f"  {week_label}: Not found")

        print(f"\nMonthly hours:")
        for emp, hrs in monthly_hours.items():
            print(f"  {emp}: {hrs:.2f} hrs")

        # Find next empty row in Data tab
        next_row = 2
        while data_sheet.cells(next_row, 1).value:
            next_row += 1

        print(f"\nAdding to Data tab starting at row {next_row}...")

        # Add rows to Data tab
        month_str = f"{month_name} {year}"
        rollup_data = []

        for emp_name, emp_data in employees.items():
            hrs = monthly_hours[emp_name]
            if hrs > 0:
                rate = emp_data['rate'] or EMPLOYEE_RATES.get(emp_name, 0)
                cost = hrs * rate

                # Write to Data tab
                data_sheet.cells(next_row, 1).value = "Vertekal"                    # Company
                data_sheet.cells(next_row, 2).value = emp_data['charge_no']         # Booz Allen Charge No
                data_sheet.cells(next_row, 3).value = emp_name                      # Employee
                data_sheet.cells(next_row, 4).value = emp_data['clin']              # CLIN
                data_sheet.cells(next_row, 5).value = emp_data['position']          # PLC
                data_sheet.cells(next_row, 6).value = rate                          # BY Rate
                data_sheet.cells(next_row, 7).value = emp_data['wbs']               # WBS
                data_sheet.cells(next_row, 8).value = emp_data['detail']            # Detail
                data_sheet.cells(next_row, 9).value = hrs                           # Hours
                data_sheet.cells(next_row, 10).value = cost                         # Cost
                data_sheet.cells(next_row, 11).value = month_str                    # Month

                rollup_data.append({
                    'employee': emp_name,
                    'hours': hrs,
                    'rate': rate,
                    'cost': cost,
                    'row': next_row
                })

                print(f"  Row {next_row}: {emp_name} - {hrs:.2f} hrs x ${rate:.2f} = ${cost:.2f}")
                next_row += 1

        # Save
        if output_path is None:
            output_path = wsr_path

        wb.save(output_path)
        print(f"\nSaved to: {output_path}")

        total_hours = sum(monthly_hours.values())
        total_cost = sum(d['cost'] for d in rollup_data)

        return {
            'success': True,
            'mode': 'monthly',
            'month': month_str,
            'total_hours': total_hours,
            'total_cost': total_cost,
            'rollup_data': rollup_data,
            'output_path': output_path
        }

    finally:
        wb.close()
        app.quit()


# =============================================================================
# MAIN
# =============================================================================

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description='WSR Updater - Weekly updates and Monthly roll-ups',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Weekly update (most recent week):
  python wsr_updater.py --weekly

  # Weekly update (specific week):
  python wsr_updater.py --weekly --week 2026-01-05

  # Monthly roll-up:
  python wsr_updater.py --monthly --month "Jan-26"
        """
    )

    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument('--weekly', action='store_true', help='Update weekly hours')
    mode.add_argument('--monthly', action='store_true', help='Roll up monthly to Data tab')

    parser.add_argument('--week', help='Week date (YYYY-MM-DD) for weekly mode')
    parser.add_argument('--month', help='Month (e.g., "Jan-26") for monthly mode')
    parser.add_argument('--wsr', help='Path to WSR file (default: find latest)')
    parser.add_argument('--output', '-o', help='Output file path')
    parser.add_argument('--preview', action='store_true', help='Preview only, do not update')

    args = parser.parse_args()

    # Find WSR file
    if args.wsr:
        wsr_path = args.wsr
    else:
        wsr_path = find_latest_wsr()
        if not wsr_path:
            print("Error: Could not find WSR file")
            sys.exit(1)
        wsr_path = str(wsr_path)

    print("=" * 60)
    print("WSR Updater Agent")
    print("=" * 60)
    print(f"WSR file: {wsr_path}")

    if args.weekly:
        week_start, week_end = get_week_dates(args.week)
        print(f"Mode: Weekly update for {week_start} to {week_end}")

        if args.preview:
            hours = get_tsheets_hours_for_week(week_start, week_end)
            print("\nPreview - Hours from TSheets:")
            for emp, hrs in sorted(hours.items()):
                print(f"  {emp}: {hrs:.2f} hrs")
        else:
            result = update_weekly(wsr_path, week_start, week_end, args.output)
            if 'error' in result:
                print(f"\nError: {result['error']}")
                sys.exit(1)

    elif args.monthly:
        if not args.month:
            print("Error: --month required for monthly mode (e.g., --month 'Jan-26')")
            sys.exit(1)

        # Parse month
        from utils.date_finder import parse_month_input
        year, month = parse_month_input(args.month)
        print(f"Mode: Monthly roll-up for {datetime(year, month, 1).strftime('%B %Y')}")

        if args.preview:
            print("\nPreview mode - would roll up weekly hours to Data tab")
        else:
            result = rollup_monthly(wsr_path, year, month, args.output)
            if 'error' in result:
                print(f"\nError: {result['error']}")
                sys.exit(1)

    print("\nDone!")
