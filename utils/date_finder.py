"""
Date Column Finder
Finds the column in an MSR that matches a given month/year.
"""

from datetime import datetime
from typing import Optional, Tuple
import re


def parse_month_input(month_str: str) -> Tuple[int, int]:
    """
    Parse various month input formats to (year, month).

    Supports:
    - "Jan-26", "Feb-26", etc.
    - "January 2026", "February 2026", etc.
    - "2026-01", "2026-02", etc.

    Args:
        month_str: Month string in various formats

    Returns:
        Tuple of (year, month) as integers

    Raises:
        ValueError: If format cannot be parsed
    """
    month_str = month_str.strip()

    # Try "Jan-26" format
    match = re.match(r'([A-Za-z]+)-(\d{2})', month_str)
    if match:
        month_name = match.group(1)
        year_short = int(match.group(2))
        year = 2000 + year_short
        month = datetime.strptime(month_name, '%b').month
        return (year, month)

    # Try "January 2026" format
    try:
        dt = datetime.strptime(month_str, '%B %Y')
        return (dt.year, dt.month)
    except ValueError:
        pass

    # Try "2026-01" format
    try:
        dt = datetime.strptime(month_str, '%Y-%m')
        return (dt.year, dt.month)
    except ValueError:
        pass

    raise ValueError(f"Cannot parse month format: {month_str}")


def find_month_column(worksheet, date_header_row: int, target_year: int, target_month: int) -> Optional[int]:
    """
    Find the column number that contains the target month/year.

    Args:
        worksheet: openpyxl worksheet object
        date_header_row: Row number containing date headers
        target_year: Year to find (e.g., 2026)
        target_month: Month to find (e.g., 1 for January)

    Returns:
        Column number (1-indexed) or None if not found
    """
    max_col = worksheet.max_column

    for col in range(1, max_col + 1):
        cell_value = worksheet.cell(row=date_header_row, column=col).value

        if isinstance(cell_value, datetime):
            if cell_value.year == target_year and cell_value.month == target_month:
                return col

    return None


def format_month_display(year: int, month: int) -> str:
    """Format year/month for display."""
    dt = datetime(year, month, 1)
    return dt.strftime('%b-%y')  # e.g., "Jan-26"


if __name__ == "__main__":
    # Test parsing
    test_cases = ["Jan-26", "February 2026", "2026-03"]
    for test in test_cases:
        try:
            year, month = parse_month_input(test)
            display = format_month_display(year, month)
            print(f"{test:20s} -> {year}-{month:02d} ({display})")
        except ValueError as e:
            print(f"{test:20s} -> ERROR: {e}")
