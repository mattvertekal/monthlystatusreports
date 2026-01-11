"""
WSR (Weekly Status Report) Updater Agent
Updates Vertekal subcontract WSR with weekly timesheet data from TSheets.

This handles the weekly status report for the Vertekal -> Booz Allen subcontract,
which tracks hours by employee for each work week.
"""

import json
import copy
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, Optional, Tuple

import sys
sys.path.append(str(Path(__file__).parent.parent))


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
            days_since_friday = 7  # Use previous week if still Friday
        last_friday = today - timedelta(days=days_since_friday)
        target_date = last_friday

    # Find Monday of that week
    days_to_monday = target_date.weekday()
    monday = target_date - timedelta(days=days_to_monday)

    # Find Friday of that week
    friday = monday + timedelta(days=4)

    return monday.strftime("%Y-%m-%d"), friday.strftime("%Y-%m-%d")


def format_week_label(start_date: str, end_date: str) -> str:
    """Format week as label like 'Jan 5-9' or 'Dec 29-31'."""
    start = datetime.strptime(start_date, "%Y-%m-%d")
    end = datetime.strptime(end_date, "%Y-%m-%d")

    if start.month == end.month:
        return f"{start.strftime('%b')} {start.day}-{end.day}"
    else:
        return f"{start.strftime('%b')} {start.day}-{end.strftime('%b')} {end.day}"


def get_tsheets_hours_for_week(
    start_date: str,
    end_date: str,
    user_ids: Optional[list] = None
) -> Dict[str, float]:
    """
    Fetch hours from TSheets for a specific week.

    Args:
        start_date: Week start date (YYYY-MM-DD)
        end_date: Week end date (YYYY-MM-DD)
        user_ids: Optional list of user IDs to filter

    Returns:
        Dictionary mapping employee names to total hours for the week
    """
    import requests

    config_dir = Path(__file__).parent.parent / 'config'
    with open(config_dir / 'tsheets_config.json') as f:
        config = json.load(f)

    headers = {"Authorization": f"Bearer {config['api_token']}"}
    base_url = config['base_url']

    # Fetch timesheets
    params = {
        "start_date": start_date,
        "end_date": end_date
    }
    if user_ids:
        params['user_ids'] = ','.join(user_ids)

    response = requests.get(f"{base_url}/timesheets", headers=headers, params=params)
    response.raise_for_status()
    timesheets = response.json().get('results', {}).get('timesheets', {})

    # Aggregate by user
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


def find_week_column_xlwings(wb, sheet_name: str, week_label: str, year: int) -> Optional[int]:
    """
    Find the column for a specific week in the WSR using xlwings.

    Args:
        wb: xlwings workbook
        sheet_name: Sheet name to search
        week_label: Week label like 'Jan 5-9'
        year: Year to disambiguate weeks

    Returns:
        Column number (1-indexed) or None if not found
    """
    sheet = wb.sheets[sheet_name]

    # Scan row 3 for week labels (columns 100+)
    for col in range(100, 200):
        cell_value = sheet.cells(3, col).value
        if cell_value and week_label in str(cell_value):
            # Verify year by checking nearby Total column
            for check_col in range(col, col + 10):
                check_val = sheet.cells(3, check_col).value
                if check_val and "Total" in str(check_val) and str(year) in str(check_val):
                    return col
            # If no year verification, still return the column
            return col

    return None


def update_wsr(
    wsr_path: str,
    week_start: str,
    week_end: str,
    output_path: Optional[str] = None,
    preview_only: bool = False
) -> Dict[str, any]:
    """
    Update WSR with weekly hours from TSheets.

    Args:
        wsr_path: Path to the WSR file (.xlsb)
        week_start: Week start date (YYYY-MM-DD)
        week_end: Week end date (YYYY-MM-DD)
        output_path: Path to save (default: completed/YYYY/Q#/)
        preview_only: If True, just show what would be updated

    Returns:
        Update result dictionary
    """
    try:
        import xlwings as xw
    except ImportError:
        return {
            'error': 'xlwings not installed. Run: pip3 install xlwings',
            'requires_excel': True
        }

    # Load employee config
    config_dir = Path(__file__).parent.parent / 'config'
    with open(config_dir / 'msr_settings.json') as f:
        settings = json.load(f)
    with open(config_dir / 'employee_mappings.json') as f:
        mappings = json.load(f)

    emmett_settings = settings['Emmett']

    # Fetch hours from TSheets
    print(f"Fetching TSheets data for {week_start} to {week_end}...")
    hours = get_tsheets_hours_for_week(week_start, week_end)

    # Get employee rows
    employee_rows = {}
    for emp_name, emp_info in mappings['employees'].items():
        if 'Emmett' in emp_info.get('msrs', []):
            for charge_code, mapping in emp_info['charge_codes'].items():
                if mapping['msr'] == 'Emmett':
                    employee_rows[emp_name] = mapping['row']

    week_label = format_week_label(week_start, week_end)
    year = int(week_start[:4])

    print(f"\nWeek: {week_label} {year}")
    print(f"Hours retrieved:")
    for emp, hrs in hours.items():
        row = employee_rows.get(emp, '?')
        print(f"  {emp} (row {row}): {hrs:.2f} hrs")

    if preview_only:
        return {
            'preview': True,
            'week': week_label,
            'year': year,
            'hours': hours,
            'employee_rows': employee_rows
        }

    # Open Excel with xlwings
    app = xw.App(visible=False)
    wb = app.books.open(wsr_path)

    try:
        sheet = wb.sheets[emmett_settings['sheet_name']]

        # Find week column
        target_col = find_week_column_xlwings(wb, emmett_settings['sheet_name'], week_label, year)

        if not target_col:
            return {
                'error': f"Could not find column for week {week_label} {year}",
                'hours': hours
            }

        print(f"\nTarget column: {target_col}")

        # Copy formatting from existing Actual column
        source_col = 50  # A column with correct formatting
        highlight_color = tuple(emmett_settings.get('highlight_color', [217, 226, 243]))

        # Update status row (Estimate -> Actual)
        sheet.cells(emmett_settings['status_row'], target_col).value = "Actual"

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
            updates.append({
                'employee': emp_name,
                'row': row,
                'hours': emp_hours
            })

        # Set output path
        if output_path is None:
            quarter = (int(week_start[5:7]) - 1) // 3 + 1
            output_dir = Path(wsr_path).parent / "completed" / str(year) / f"Q{quarter}"
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = str(output_dir / f"Vertekal_WSR_{week_start}_to_{week_end}.xlsb")

        # Save
        wb.save(output_path)
        print(f"\nSaved to: {output_path}")

        return {
            'success': True,
            'week': week_label,
            'year': year,
            'total_hours': total_hours,
            'updates': updates,
            'output_path': output_path
        }

    finally:
        wb.close()
        app.quit()


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description='Update Vertekal WSR with TSheets hours')
    parser.add_argument('wsr_path', help='Path to WSR file (.xlsb)')
    parser.add_argument('--week', help='Week date (YYYY-MM-DD) or leave blank for most recent')
    parser.add_argument('--preview', action='store_true', help='Preview only, do not update')
    parser.add_argument('--output', '-o', help='Output file path')

    args = parser.parse_args()

    # Get week dates
    week_start, week_end = get_week_dates(args.week)

    print("=" * 60)
    print("WSR Update Agent")
    print("=" * 60)

    result = update_wsr(
        wsr_path=args.wsr_path,
        week_start=week_start,
        week_end=week_end,
        output_path=args.output,
        preview_only=args.preview
    )

    if 'error' in result:
        print(f"\nError: {result['error']}")
    elif result.get('preview'):
        print("\nPreview complete. Run without --preview to update.")
    else:
        print("\nUpdate complete!")
