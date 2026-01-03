#!/usr/bin/env python3
"""
MSR Update Orchestrator
Coordinates updating all three MSRs (TO1, TO4, TO6) with timesheet data.
"""

import json
import sys
from pathlib import Path
from typing import Dict, List

from agents.to1_updater import update_to1_msr
from agents.to4_updater import update_to4_msr
from agents.to6_updater import update_to6_msr
from utils.timesheet_parser import parse_timesheet_csv
from utils.date_finder import parse_month_input, find_month_column, format_month_display
from openpyxl import load_workbook


def update_all_msrs(
    timesheet_csv: str,
    to1_msr: str,
    to4_msr: str,
    to6_msr: str,
    month_str: str,
    output_dir: str = None
) -> Dict[str, any]:
    """
    Update all MSRs with timesheet data for the specified month.

    Args:
        timesheet_csv: Path to timesheet CSV file
        to1_msr: Path to TO1 MSR file
        to4_msr: Path to TO4 MSR file
        to6_msr: Path to TO6 MSR file
        month_str: Month to update (e.g., "Jan-26", "February 2026")
        output_dir: Directory to save updated MSRs (default: same as input)

    Returns:
        Dictionary with results from all updates
    """
    print("="*80)
    print("MSR Update Process Starting")
    print("="*80)

    # Parse month
    try:
        target_year, target_month = parse_month_input(month_str)
        month_display = format_month_display(target_year, target_month)
        print(f"\n Target Month: {month_display} ({target_year}-{target_month:02d})")
    except ValueError as e:
        print(f"\nERROR: {e}")
        return None

    # Parse timesheet
    print(f"\n Parsing timesheet: {timesheet_csv}")
    timesheet_data = parse_timesheet_csv(timesheet_csv)
    total_employees = len(timesheet_data)
    print(f"   Found {total_employees} employees with billable hours")

    # Set output directory
    if output_dir is None:
        output_dir = Path(timesheet_csv).parent
    else:
        output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Load config for settings
    config_dir = Path(__file__).parent / 'config'
    with open(config_dir / 'msr_settings.json') as f:
        settings = json.load(f)

    results = {}

    # Update TO1
    print("\n" + "-"*80)
    print("Updating TO1 MSR...")
    print("-"*80)
    try:
        wb1 = load_workbook(to1_msr, data_only=True)
        ws1 = wb1[settings['TO1']['sheet_name']]
        date_row = settings['TO1']['date_header_row']
        col1 = find_month_column(ws1, date_row, target_year, target_month)
        wb1.close()

        if col1:
            print(f"   Found column: {col1}")
            output1 = output_dir / f"TO1_MSR_{month_display}_UPDATED.xlsx"
            result1 = update_to1_msr(to1_msr, timesheet_data, col1, str(output1))
            results['TO1'] = result1
            print(f"   ✓ Updated {result1['total_hours']:.2f} hours")
            print(f"   ✓ Saved: {output1}")
        else:
            print(f"   ✗ Could not find column for {month_display}")
            results['TO1'] = {'error': 'Column not found'}
    except Exception as e:
        print(f"   ✗ Error: {e}")
        results['TO1'] = {'error': str(e)}

    # Update TO4
    print("\n" + "-"*80)
    print("Updating TO4 MSR...")
    print("-"*80)
    try:
        wb4 = load_workbook(to4_msr, data_only=True)
        ws4 = wb4['CLIN 0001AD']
        date_row = settings['TO4']['sheets']['CLIN 0001AD']['date_header_row']
        col4 = find_month_column(ws4, date_row, target_year, target_month)
        wb4.close()

        if col4:
            print(f"   Found column: {col4}")
            output4 = output_dir / f"TO4_MSR_{month_display}_UPDATED.xlsx"
            result4 = update_to4_msr(to4_msr, timesheet_data, col4, str(output4))
            results['TO4'] = result4
            print(f"   ✓ CLIN 0001AD: {result4['clin_0001ad_hours']:.2f} hours")
            print(f"   ✓ CLIN 0002AD: {result4['clin_0002ad_hours']:.2f} hours")
            print(f"   ✓ Total: {result4['total_hours']:.2f} hours")
            print(f"   ✓ Saved: {output4}")
        else:
            print(f"   ✗ Could not find column for {month_display}")
            results['TO4'] = {'error': 'Column not found'}
    except Exception as e:
        print(f"   ✗ Error: {e}")
        results['TO4'] = {'error': str(e)}

    # Update TO6
    print("\n" + "-"*80)
    print("Updating TO6 MSR...")
    print("-"*80)
    try:
        wb6 = load_workbook(to6_msr, data_only=True)
        ws6 = wb6[settings['TO6']['sheet_name']]
        date_row = settings['TO6']['date_header_row']
        col6 = find_month_column(ws6, date_row, target_year, target_month)
        wb6.close()

        if col6:
            print(f"   Found column: {col6}")
            output6 = output_dir / f"TO6_MSR_{month_display}_UPDATED.xlsx"
            result6 = update_to6_msr(to6_msr, timesheet_data, col6, str(output6))
            results['TO6'] = result6
            print(f"   ✓ Updated {result6['total_hours']:.2f} hours")
            print(f"   ✓ Saved: {output6}")
        else:
            print(f"   ✗ Could not find column for {month_display}")
            results['TO6'] = {'error': 'Column not found'}
    except Exception as e:
        print(f"   ✗ Error: {e}")
        results['TO6'] = {'error': str(e)}

    # Summary
    print("\n" + "="*80)
    print("MSR Update Complete")
    print("="*80)

    successful = sum(1 for r in results.values() if 'error' not in r)
    print(f"\nSuccessfully updated: {successful}/3 MSRs")

    return results


if __name__ == "__main__":
    if len(sys.argv) < 6:
        print("Usage: python update_msrs.py <timesheet.csv> <to1.xlsx> <to4.xlsx> <to6.xlsx> <month>")
        print("Example: python update_msrs.py timesheet.csv to1.xlsx to4.xlsx to6.xlsx 'Jan-26'")
        sys.exit(1)

    update_all_msrs(
        timesheet_csv=sys.argv[1],
        to1_msr=sys.argv[2],
        to4_msr=sys.argv[3],
        to6_msr=sys.argv[4],
        month_str=sys.argv[5]
    )
