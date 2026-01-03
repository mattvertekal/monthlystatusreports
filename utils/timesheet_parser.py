"""
Timesheet CSV Parser
Parses TSheets export CSV and aggregates hours by employee and charge code.
"""

import csv
from collections import defaultdict
from typing import Dict, List, Tuple


def parse_timesheet_csv(csv_path: str) -> Dict[str, Dict[str, float]]:
    """
    Parse timesheet CSV and return hours by employee and charge code.

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
        timesheet_data: Output from parse_timesheet_csv

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
        data = parse_timesheet_csv(sys.argv[1])
        print_timesheet_summary(data)
    else:
        print("Usage: python timesheet_parser.py <csv_file>")
