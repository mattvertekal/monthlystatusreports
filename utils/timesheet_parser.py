"""
Timesheet Parser
Fetches timesheet data from TSheets API or parses CSV export.
"""

import csv
import json
import requests
from collections import defaultdict
from datetime import datetime, date
from pathlib import Path
from typing import Dict, List, Tuple, Optional


def load_tsheets_config() -> dict:
    """Load TSheets configuration."""
    config_path = Path(__file__).parent.parent / 'config' / 'tsheets_config.json'
    with open(config_path) as f:
        return json.load(f)


def get_tsheets_timesheets(start_date: str, end_date: str) -> Dict[str, Dict[str, float]]:
    """
    Fetch timesheets from TSheets API for a date range.

    Args:
        start_date: Start date in YYYY-MM-DD format
        end_date: End date in YYYY-MM-DD format

    Returns:
        Dictionary mapping employee names to their charge code hours
        {
            "Employee Name": {
                "Charge Code": hours,
                ...
            },
            ...
        }
    """
    config = load_tsheets_config()

    headers = {"Authorization": f"Bearer {config['api_token']}"}
    base_url = config['base_url']

    # First, get jobcodes to map IDs to names
    jobcodes_response = requests.get(f"{base_url}/jobcodes", headers=headers)
    jobcodes_response.raise_for_status()
    jobcodes_data = jobcodes_response.json()

    jobcode_map = {}
    for jc_id, jc in jobcodes_data.get('results', {}).get('jobcodes', {}).items():
        jobcode_map[int(jc_id)] = jc['name']

    # Fetch timesheets
    params = {
        "start_date": start_date,
        "end_date": end_date
    }

    timesheets_response = requests.get(f"{base_url}/timesheets", headers=headers, params=params)
    timesheets_response.raise_for_status()
    timesheets_data = timesheets_response.json()

    # Aggregate hours by user and jobcode
    user_map = config['users']
    skip_jobcodes = set(config.get('skip_jobcodes', ['PTO', 'Holiday']))

    employees = defaultdict(lambda: defaultdict(float))

    for ts_id, ts in timesheets_data.get('results', {}).get('timesheets', {}).items():
        user_id = str(ts['user_id'])
        jobcode_id = ts['jobcode_id']
        duration_seconds = ts.get('duration', 0)

        # Skip if user not in our mapping
        if user_id not in user_map:
            continue

        employee_name = user_map[user_id]

        # Get jobcode name
        jobcode_name = jobcode_map.get(jobcode_id, '')

        # Skip PTO, Holiday, etc.
        if jobcode_name in skip_jobcodes:
            continue

        # Convert seconds to hours
        hours = duration_seconds / 3600.0

        if jobcode_name and hours > 0:
            employees[employee_name][jobcode_name] += hours

    # Convert defaultdict to regular dict
    return {name: dict(codes) for name, codes in employees.items()}


def get_timesheets_for_month(year: int, month: int) -> Dict[str, Dict[str, float]]:
    """
    Fetch timesheets for an entire month.

    Args:
        year: Year (e.g., 2026)
        month: Month (1-12)

    Returns:
        Dictionary mapping employee names to their charge code hours
    """
    from calendar import monthrange

    start_date = f"{year}-{month:02d}-01"
    last_day = monthrange(year, month)[1]
    end_date = f"{year}-{month:02d}-{last_day:02d}"

    return get_tsheets_timesheets(start_date, end_date)


def parse_timesheet_csv(csv_path: str) -> Dict[str, Dict[str, float]]:
    """
    Parse timesheet CSV and return hours by employee and charge code.
    (Legacy function - kept for backward compatibility)

    Args:
        csv_path: Path to the timesheet CSV file

    Returns:
        Dictionary mapping employee names to their charge code hours
        {
            "Employee Name": {
                "Charge Code": hours,
                ...
            },
            ...
        }
    """
    employees = defaultdict(lambda: defaultdict(float))

    with open(csv_path, 'r') as f:
        reader = csv.DictReader(f)
        for row in reader:
            # Extract employee name
            fname = row.get('fname', '')
            lname = row.get('lname', '')
            name = f"{fname} {lname}"

            # Extract hours
            hours = float(row.get('hours', 0)) if row.get('hours') else 0.0

            # Extract charge code - use jobcode_2 if available, otherwise jobcode_1
            jobcode_2 = row.get('jobcode_2', '').strip()
            jobcode_1 = row.get('jobcode_1', '').strip()

            # Skip PTO and Holiday
            if jobcode_1 in ['PTO', 'Holiday']:
                continue

            # Use jobcode_2 if available and not empty, otherwise use jobcode_1
            charge_code = jobcode_2 if jobcode_2 else jobcode_1

            if charge_code and hours > 0:
                employees[name][charge_code] += hours

    # Convert defaultdict to regular dict
    return {name: dict(codes) for name, codes in employees.items()}


def get_employee_hours_summary(timesheet_data: Dict[str, Dict[str, float]]) -> List[Tuple[str, float]]:
    """
    Get total hours per employee.

    Args:
        timesheet_data: Output from parse_timesheet_csv or get_tsheets_timesheets

    Returns:
        List of (employee_name, total_hours) tuples
    """
    summary = []
    for name, codes in timesheet_data.items():
        total = sum(codes.values())
        summary.append((name, total))
    return sorted(summary, key=lambda x: x[0])


def print_timesheet_summary(timesheet_data: Dict[str, Dict[str, float]]):
    """Print a formatted summary of timesheet data."""
    print("="*80)
    print("TIMESHEET SUMMARY")
    print("="*80)

    for name in sorted(timesheet_data.keys()):
        codes = timesheet_data[name]
        total = sum(codes.values())
        print(f"\n{name} ({total:.2f} hrs total)")
        print("-"*80)
        for code, hours in sorted(codes.items()):
            print(f"  {code:60s} {hours:>8.2f} hrs")

    print("\n" + "="*80)


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        arg = sys.argv[1]

        # Check if it's a month format (Jan-26, 2026-01, etc.)
        if '-' in arg and not arg.endswith('.csv'):
            from date_finder import parse_month_input
            year, month = parse_month_input(arg)
            print(f"Fetching timesheets for {year}-{month:02d} from TSheets API...")
            data = get_timesheets_for_month(year, month)
        else:
            # Assume it's a CSV path
            print(f"Parsing CSV file: {arg}")
            data = parse_timesheet_csv(arg)

        print_timesheet_summary(data)
    else:
        print("Usage:")
        print("  python timesheet_parser.py <csv_file>     # Parse CSV export")
        print("  python timesheet_parser.py <month>        # Fetch from TSheets API")
        print("  Examples: Jan-26, February 2026, 2026-01")
