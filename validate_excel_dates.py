#!/usr/bin/env python3
"""
Excel Data Validator - Comprehensive Data Validation Tool
==========================================================
This script validates that all dates from individual store Excel files
are present in their respective merge files:
- Sales lines individual files against sales lines merge file
- Payment lines individual files against payment lines merge file

It generates a detailed HTML and text report showing:
- Summary statistics
- Missing dates per file
- Extra dates not in merge file
- Amount validation (individual file totals vs merge file totals)
- Grand total validation (all files combined)
- Complete validation results

CONFIGURATION:
-------------
You can customize the merge file names and patterns below to match your files:
- SALES_MERGE_FILE: Name of the sales lines merge file
- PAYMENT_MERGE_FILE: Name of the payment lines merge file
- SALES_FILE_PATTERN: Regex pattern to identify sales line files
- PAYMENT_FILE_PATTERN: Regex pattern to identify payment line files
- Column names for both merge and individual files
"""

import pandas as pd
import os
from datetime import datetime
from pathlib import Path
import warnings
import re
warnings.filterwarnings('ignore')

# Configuration - Sales Lines
SALES_MERGE_FILE = '49_stores_sales_lines.xlsx'
SALES_FILE_PATTERN = r'sale.*line'  # Pattern to identify sales line files (matches "sale line", "sales line", etc.)
SALES_MERGE_DATE_COLUMN = 'Order Lines/Order Ref/Date'
SALES_INDIVIDUAL_DATE_COLUMN = 'Order Lines/Order Ref/Date'  # Sales files use same structure as merge
SALES_MERGE_ORDER_REF_COLUMN = 'Order Lines/Order Ref'
SALES_MERGE_AMOUNT_COLUMN = 'Order Lines/Subtotal'
SALES_INDIVIDUAL_ORDER_REF_COLUMN = 'Order Lines/Order Ref'  # Sales files use same structure as merge
SALES_INDIVIDUAL_AMOUNT_COLUMN = 'Order Lines/Subtotal'  # Sales files use same structure as merge
SALES_INDIVIDUAL_BRANCH_COLUMN = 'Branch'  # Note: Sales files may not have this column

# Configuration - Payment Lines
PAYMENT_MERGE_FILE = '49_stores_payment_lines.xlsx'
PAYMENT_FILE_PATTERN = r'payment.*line'  # Pattern to identify payment line files (matches "payment line", "payment lines", etc.)
PAYMENT_MERGE_DATE_COLUMN = 'Date'  # Payment merge uses same structure as individual files
PAYMENT_INDIVIDUAL_DATE_COLUMN = 'Date'
PAYMENT_MERGE_ORDER_REF_COLUMN = 'Order Ref'  # Payment merge uses same structure as individual files
PAYMENT_MERGE_AMOUNT_COLUMN = 'Payments/Amount'  # Payment merge uses same structure as individual files
PAYMENT_INDIVIDUAL_ORDER_REF_COLUMN = 'Order Ref'
PAYMENT_INDIVIDUAL_AMOUNT_COLUMN = 'Payments/Amount'
PAYMENT_INDIVIDUAL_BRANCH_COLUMN = 'Branch'

# General Configuration
REPORT_HTML = 'validation_report.html'
REPORT_TXT = 'validation_report.txt'
AMOUNT_TOLERANCE = 0.01  # Allow small rounding differences


class ExcelDateValidator:
    """Validates dates and amounts across multiple Excel files against their respective merge files."""

    def __init__(self, base_dir, merge_file, file_type='payment',
                 date_column='Date', merge_date_column='Order Lines/Order Ref/Date',
                 merge_order_ref_column='Order Lines/Order Ref', merge_amount_column='Order Lines/Subtotal',
                 individual_order_ref_column='Order Ref', individual_amount_column='Payments/Amount',
                 individual_branch_column='Branch'):
        self.base_dir = Path(base_dir)
        self.merge_file = self.base_dir / merge_file
        self.file_type = file_type  # 'sales' or 'payment'
        self.results = []
        self.merge_dates = set()
        self.all_individual_dates = set()
        self.merge_df = None
        self.merge_amounts_by_branch = {}
        self.grand_total_individual = 0
        self.grand_total_merge = 0

        # Column configurations
        self.merge_date_column = merge_date_column
        self.individual_date_column = date_column
        self.merge_order_ref_column = merge_order_ref_column
        self.merge_amount_column = merge_amount_column
        self.individual_order_ref_column = individual_order_ref_column
        self.individual_amount_column = individual_amount_column
        self.individual_branch_column = individual_branch_column

    def load_merge_file(self):
        """Load and extract dates and amounts from the merge file."""
        print(f"Loading {self.file_type} merge file: {self.merge_file}")
        try:
            self.merge_df = pd.read_excel(self.merge_file)
            if self.merge_date_column not in self.merge_df.columns:
                raise ValueError(f"Column '{self.merge_date_column}' not found in merge file")
            if self.merge_order_ref_column not in self.merge_df.columns:
                raise ValueError(f"Column '{self.merge_order_ref_column}' not found in merge file")
            if self.merge_amount_column not in self.merge_df.columns:
                raise ValueError(f"Column '{self.merge_amount_column}' not found in merge file")

            # Convert to datetime and extract date only
            # Handle Excel serial dates (stored as floats) by checking the dtype
            if self.merge_df[self.merge_date_column].dtype in ['float64', 'float32', 'object']:
                # Check if values look like Excel serial dates (numbers > 1000)
                sample_value = self.merge_df[self.merge_date_column].dropna().iloc[0] if len(self.merge_df[self.merge_date_column].dropna()) > 0 else None
                if sample_value is not None and isinstance(sample_value, (int, float)) and sample_value > 1000:
                    # Convert Excel serial dates (days since 1899-12-30)
                    self.merge_df[self.merge_date_column] = pd.to_datetime(self.merge_df[self.merge_date_column], unit='D', origin='1899-12-30')
                else:
                    # Regular datetime conversion
                    self.merge_df[self.merge_date_column] = pd.to_datetime(self.merge_df[self.merge_date_column])
            else:
                # Already datetime or can be converted normally
                self.merge_df[self.merge_date_column] = pd.to_datetime(self.merge_df[self.merge_date_column])

            self.merge_dates = set(self.merge_df[self.merge_date_column].dt.date)

            # Pre-compute amounts by branch for faster lookup
            # Extract branch name from Order Ref (e.g., "SALAMRYD/12345" -> "SALAMRYD")
            self.merge_df['branch'] = self.merge_df[self.merge_order_ref_column].str.split('/').str[0]

            # Group by branch and order ref to get order-level amounts
            for branch in self.merge_df['branch'].unique():
                branch_df = self.merge_df[self.merge_df['branch'] == branch]
                # Sum amounts by order ref to get order total (since merge file has line items)
                order_amounts = branch_df.groupby(self.merge_order_ref_column)[self.merge_amount_column].sum()
                total_amount = order_amounts.sum()
                self.merge_amounts_by_branch[branch] = {
                    'total': total_amount,
                    'order_count': len(order_amounts),
                    'order_amounts': order_amounts.to_dict()
                }

            self.grand_total_merge = self.merge_df[self.merge_amount_column].sum()

            print(f"✓ Loaded {len(self.merge_df)} rows with {len(self.merge_dates)} unique dates")
            print(f"✓ Computed amounts for {len(self.merge_amounts_by_branch)} branches")
            return True
        except Exception as e:
            print(f"✗ Error loading merge file: {e}")
            return False

    def get_individual_files(self):
        """Get all individual Excel files (excluding the merge file)."""
        all_files = list(self.base_dir.glob('*.xlsx')) + list(self.base_dir.glob('*.xls'))
        individual_files = [f for f in all_files if f.name != self.merge_file.name]
        return sorted(individual_files)

    def validate_file(self, file_path):
        """Validate a single Excel file against the merge file (dates and amounts)."""
        file_name = file_path.name
        print(f"\nValidating: {file_name}")

        result = {
            'file_name': file_name,
            'status': 'Success',
            'total_rows': 0,
            'unique_dates': 0,
            'missing_dates': [],
            'missing_count': 0,
            'individual_amount': 0,
            'merge_amount': 0,
            'amount_difference': 0,
            'amount_match': True,
            'branch_name': None,
            'unique_orders_individual': 0,
            'unique_orders_merge': 0,
            'orders_in_individual_not_merge': [],
            'orders_in_merge_not_individual': [],
            'error': None
        }

        try:
            df = pd.read_excel(file_path)

            # Validate required columns exist
            if self.individual_date_column not in df.columns:
                result['status'] = 'Error'
                result['error'] = f"Column '{self.individual_date_column}' not found"
                print(f"  ✗ {result['error']}")
                return result

            # Date validation
            # Handle Excel serial dates (stored as floats) by checking the dtype
            if df[self.individual_date_column].dtype in ['float64', 'float32', 'object']:
                # Check if values look like Excel serial dates (numbers > 1000)
                sample_value = df[self.individual_date_column].dropna().iloc[0] if len(df[self.individual_date_column].dropna()) > 0 else None
                if sample_value is not None and isinstance(sample_value, (int, float)) and sample_value > 1000:
                    # Convert Excel serial dates (days since 1899-12-30)
                    df[self.individual_date_column] = pd.to_datetime(df[self.individual_date_column], unit='D', origin='1899-12-30')
                else:
                    # Regular datetime conversion
                    df[self.individual_date_column] = pd.to_datetime(df[self.individual_date_column])
            else:
                # Already datetime or can be converted normally
                df[self.individual_date_column] = pd.to_datetime(df[self.individual_date_column])

            file_dates = set(df[self.individual_date_column].dt.date)

            result['total_rows'] = len(df)
            result['unique_dates'] = len(file_dates)

            # Store all individual dates for extra date detection
            self.all_individual_dates.update(file_dates)

            # Find missing dates
            missing = file_dates - self.merge_dates
            result['missing_dates'] = sorted(missing)
            result['missing_count'] = len(missing)

            if missing:
                result['status'] = 'Missing Dates'
                print(f"  ✗ Found {len(missing)} missing date(s)")
                for date in sorted(missing)[:5]:
                    print(f"    - {date}")
                if len(missing) > 5:
                    print(f"    ... and {len(missing) - 5} more")
            else:
                print(f"  ✓ All {result['unique_dates']} unique dates are present in merge file")

            # Amount validation (if columns exist)
            if (self.individual_amount_column in df.columns and
                self.individual_branch_column in df.columns and
                self.individual_order_ref_column in df.columns):

                # Get branch name
                branch_names = df[self.individual_branch_column].unique()
                if len(branch_names) > 0:
                    result['branch_name'] = branch_names[0]

                    # Calculate total amount from individual file
                    result['individual_amount'] = df[self.individual_amount_column].sum()
                    self.grand_total_individual += result['individual_amount']

                    # Get unique orders from individual file
                    individual_orders = set(df[self.individual_order_ref_column].unique())
                    result['unique_orders_individual'] = len(individual_orders)

                    # Get merge file amount for this branch
                    if result['branch_name'] in self.merge_amounts_by_branch:
                        merge_data = self.merge_amounts_by_branch[result['branch_name']]
                        result['merge_amount'] = merge_data['total']
                        result['unique_orders_merge'] = merge_data['order_count']

                        # Get merge orders
                        merge_orders = set(merge_data['order_amounts'].keys())

                        # Find order mismatches
                        result['orders_in_individual_not_merge'] = sorted(list(individual_orders - merge_orders))[:10]
                        result['orders_in_merge_not_individual'] = sorted(list(merge_orders - individual_orders))[:10]

                        # Calculate difference
                        result['amount_difference'] = result['individual_amount'] - result['merge_amount']

                        # Check if amounts match (within tolerance)
                        if abs(result['amount_difference']) > AMOUNT_TOLERANCE:
                            result['amount_match'] = False
                            if result['status'] == 'Success':
                                result['status'] = 'Amount Mismatch'
                            elif result['status'] == 'Missing Dates':
                                result['status'] = 'Missing Dates & Amount Mismatch'
                            print(f"  ✗ Amount mismatch: Individual={result['individual_amount']:.2f}, Merge={result['merge_amount']:.2f}, Diff={result['amount_difference']:.2f}")
                        else:
                            print(f"  ✓ Amount matches: {result['individual_amount']:.2f}")
                    else:
                        result['error'] = f"Branch '{result['branch_name']}' not found in merge file"
                        result['status'] = 'Error'
                        print(f"  ✗ {result['error']}")

        except Exception as e:
            result['status'] = 'Error'
            result['error'] = str(e)
            print(f"  ✗ Error: {e}")

        return result

    def validate_all(self):
        """Validate all individual files."""
        if not self.load_merge_file():
            return False

        individual_files = self.get_individual_files()
        print(f"\nFound {len(individual_files)} individual files to validate")
        print("=" * 80)

        for file_path in individual_files:
            result = self.validate_file(file_path)
            self.results.append(result)

        return True

    def generate_summary(self):
        """Generate summary statistics."""
        total_files = len(self.results)
        success_files = len([r for r in self.results if r['status'] == 'Success'])
        files_with_missing = len([r for r in self.results if 'Missing Dates' in r['status']])
        files_with_amount_mismatch = len([r for r in self.results if not r['amount_match'] and r['individual_amount'] > 0])
        error_files = len([r for r in self.results if r['status'] == 'Error'])
        total_missing_dates = sum(r['missing_count'] for r in self.results)

        # Find dates in individual files but not in merge
        extra_dates = self.all_individual_dates - self.merge_dates

        # Calculate grand total difference
        grand_total_difference = self.grand_total_individual - self.grand_total_merge
        grand_total_match = abs(grand_total_difference) <= AMOUNT_TOLERANCE

        summary = {
            'total_files': total_files,
            'success_files': success_files,
            'files_with_missing': files_with_missing,
            'files_with_amount_mismatch': files_with_amount_mismatch,
            'error_files': error_files,
            'total_missing_dates': total_missing_dates,
            'extra_dates': sorted(extra_dates),
            'extra_dates_count': len(extra_dates),
            'merge_unique_dates': len(self.merge_dates),
            'individual_unique_dates': len(self.all_individual_dates),
            'grand_total_individual': self.grand_total_individual,
            'grand_total_merge': self.grand_total_merge,
            'grand_total_difference': grand_total_difference,
            'grand_total_match': grand_total_match
        }

        return summary

    def generate_text_report(self):
        """Generate a detailed text report."""
        summary = self.generate_summary()

        report_lines = []
        report_lines.append("=" * 80)
        report_lines.append("EXCEL DATA VALIDATION REPORT")
        report_lines.append("=" * 80)
        report_lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report_lines.append(f"Merge File: {self.merge_file.name}")
        report_lines.append("")

        # Summary
        report_lines.append("SUMMARY")
        report_lines.append("-" * 80)
        report_lines.append(f"Total Files Validated:        {summary['total_files']}")
        report_lines.append(f"Files with All Dates Present: {summary['success_files']} ✓")
        report_lines.append(f"Files with Missing Dates:     {summary['files_with_missing']} ✗")
        report_lines.append(f"Files with Amount Mismatch:   {summary['files_with_amount_mismatch']} ✗")
        report_lines.append(f"Files with Errors:            {summary['error_files']}")
        report_lines.append(f"Total Missing Date Entries:   {summary['total_missing_dates']}")
        report_lines.append("")
        report_lines.append(f"Unique Dates in Merge File:   {summary['merge_unique_dates']}")
        report_lines.append(f"Unique Dates in All Individual Files: {summary['individual_unique_dates']}")
        report_lines.append("")
        report_lines.append("AMOUNT VALIDATION")
        report_lines.append("-" * 80)
        report_lines.append(f"Grand Total (All Individual Files): {summary['grand_total_individual']:,.2f}")
        report_lines.append(f"Grand Total (Merge File):           {summary['grand_total_merge']:,.2f}")
        report_lines.append(f"Grand Total Difference:             {summary['grand_total_difference']:,.2f}")
        if summary['grand_total_match']:
            report_lines.append("Grand Total Status:                 ✓ MATCH")
        else:
            report_lines.append("Grand Total Status:                 ✗ MISMATCH")

        if summary['extra_dates']:
            report_lines.append(f"\n⚠ WARNING: Found {summary['extra_dates_count']} date(s) in individual files NOT in merge file:")
            for date in summary['extra_dates'][:10]:
                report_lines.append(f"  - {date}")
            if summary['extra_dates_count'] > 10:
                report_lines.append(f"  ... and {summary['extra_dates_count'] - 10} more")

        report_lines.append("")
        report_lines.append("")

        # Detailed Results
        report_lines.append("DETAILED RESULTS")
        report_lines.append("=" * 80)

        # Group by status - include all statuses
        for status in ['Amount Mismatch', 'Missing Dates & Amount Mismatch', 'Missing Dates', 'Success', 'Error']:
            files = [r for r in self.results if r['status'] == status]
            if not files:
                continue

            report_lines.append("")
            report_lines.append(f"{status.upper()} ({len(files)} files)")
            report_lines.append("-" * 80)

            for result in files:
                report_lines.append(f"\nFile: {result['file_name']}")
                if result['branch_name']:
                    report_lines.append(f"  Branch: {result['branch_name']}")
                report_lines.append(f"  Total Rows: {result['total_rows']}")
                report_lines.append(f"  Unique Dates: {result['unique_dates']}")

                # Amount information
                if result['individual_amount'] > 0:
                    report_lines.append(f"  Individual File Amount: {result['individual_amount']:,.2f}")
                    report_lines.append(f"  Merge File Amount: {result['merge_amount']:,.2f}")
                    report_lines.append(f"  Amount Difference: {result['amount_difference']:,.2f}")
                    if result['amount_match']:
                        report_lines.append(f"  Amount Status: ✓ MATCH")
                    else:
                        report_lines.append(f"  Amount Status: ✗ MISMATCH")

                # Orders information
                if result['unique_orders_individual'] > 0:
                    report_lines.append(f"  Unique Orders (Individual): {result['unique_orders_individual']}")
                    report_lines.append(f"  Unique Orders (Merge): {result['unique_orders_merge']}")
                    if result['orders_in_individual_not_merge']:
                        report_lines.append(f"  Orders in Individual but not Merge: {len(result['orders_in_individual_not_merge'])}")
                    if result['orders_in_merge_not_individual']:
                        report_lines.append(f"  Orders in Merge but not Individual: {len(result['orders_in_merge_not_individual'])}")

                if result['error']:
                    report_lines.append(f"  Error: {result['error']}")
                elif result['missing_dates']:
                    report_lines.append(f"  Missing Dates ({result['missing_count']}):")
                    for date in result['missing_dates']:
                        report_lines.append(f"    - {date}")

        return "\n".join(report_lines)

    def generate_html_report(self):
        """Generate a detailed HTML report."""
        summary = self.generate_summary()

        # Determine overall status
        if (summary['files_with_missing'] == 0 and
            summary['error_files'] == 0 and
            summary['files_with_amount_mismatch'] == 0 and
            summary['grand_total_match']):
            overall_status = "✓ PASS"
            status_color = "#28a745"
        else:
            overall_status = "✗ FAIL"
            status_color = "#dc3545"

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Data Validation Report</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        h1 {{
            color: #333;
            border-bottom: 3px solid #007bff;
            padding-bottom: 10px;
        }}
        .header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 30px;
        }}
        .status {{
            font-size: 24px;
            font-weight: bold;
            color: {status_color};
        }}
        .summary {{
            background-color: #f8f9fa;
            padding: 20px;
            border-radius: 5px;
            margin-bottom: 30px;
        }}
        .summary-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
            margin-top: 15px;
        }}
        .summary-item {{
            background-color: white;
            padding: 15px;
            border-radius: 5px;
            border-left: 4px solid #007bff;
        }}
        .summary-label {{
            font-size: 14px;
            color: #666;
            margin-bottom: 5px;
        }}
        .summary-value {{
            font-size: 24px;
            font-weight: bold;
            color: #333;
        }}
        .warning {{
            background-color: #fff3cd;
            border: 1px solid #ffc107;
            border-radius: 5px;
            padding: 15px;
            margin: 20px 0;
        }}
        .warning-title {{
            font-weight: bold;
            color: #856404;
            margin-bottom: 10px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }}
        th, td {{
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }}
        th {{
            background-color: #007bff;
            color: white;
            font-weight: 600;
        }}
        tr:hover {{
            background-color: #f5f5f5;
        }}
        .status-success {{
            color: #28a745;
            font-weight: bold;
        }}
        .status-error {{
            color: #dc3545;
            font-weight: bold;
        }}
        .status-missing {{
            color: #ffc107;
            font-weight: bold;
        }}
        .missing-dates {{
            background-color: #f8f9fa;
            padding: 10px;
            border-radius: 3px;
            font-size: 12px;
            max-height: 100px;
            overflow-y: auto;
        }}
        .timestamp {{
            color: #666;
            font-size: 14px;
        }}
        .section {{
            margin-top: 40px;
        }}
        .section-title {{
            font-size: 20px;
            color: #333;
            border-bottom: 2px solid #dee2e6;
            padding-bottom: 10px;
            margin-bottom: 20px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div>
                <h1>📊 Excel Data Validation Report</h1>
                <p class="timestamp">Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
                <p><strong>Merge File:</strong> {self.merge_file.name}</p>
            </div>
            <div class="status">{overall_status}</div>
        </div>

        <div class="summary">
            <h2>Summary</h2>
            <div class="summary-grid">
                <div class="summary-item">
                    <div class="summary-label">Total Files Validated</div>
                    <div class="summary-value">{summary['total_files']}</div>
                </div>
                <div class="summary-item" style="border-left-color: #28a745;">
                    <div class="summary-label">Files with All Dates ✓</div>
                    <div class="summary-value" style="color: #28a745;">{summary['success_files']}</div>
                </div>
                <div class="summary-item" style="border-left-color: #ffc107;">
                    <div class="summary-label">Files with Missing Dates</div>
                    <div class="summary-value" style="color: #ffc107;">{summary['files_with_missing']}</div>
                </div>
                <div class="summary-item" style="border-left-color: #ff6b6b;">
                    <div class="summary-label">Files with Amount Mismatch</div>
                    <div class="summary-value" style="color: #ff6b6b;">{summary['files_with_amount_mismatch']}</div>
                </div>
                <div class="summary-item" style="border-left-color: #dc3545;">
                    <div class="summary-label">Files with Errors</div>
                    <div class="summary-value" style="color: #dc3545;">{summary['error_files']}</div>
                </div>
                <div class="summary-item">
                    <div class="summary-label">Merge File Unique Dates</div>
                    <div class="summary-value">{summary['merge_unique_dates']}</div>
                </div>
                <div class="summary-item">
                    <div class="summary-label">Total Missing Date Entries</div>
                    <div class="summary-value">{summary['total_missing_dates']}</div>
                </div>
                <div class="summary-item" style="border-left-color: #17a2b8;">
                    <div class="summary-label">Grand Total (Individual Files)</div>
                    <div class="summary-value" style="font-size: 18px;">{summary['grand_total_individual']:,.2f}</div>
                </div>
                <div class="summary-item" style="border-left-color: #17a2b8;">
                    <div class="summary-label">Grand Total (Merge File)</div>
                    <div class="summary-value" style="font-size: 18px;">{summary['grand_total_merge']:,.2f}</div>
                </div>
                <div class="summary-item" style="border-left-color: {'#28a745' if summary['grand_total_match'] else '#dc3545'};">
                    <div class="summary-label">Grand Total Status</div>
                    <div class="summary-value" style="color: {'#28a745' if summary['grand_total_match'] else '#dc3545'}; font-size: 18px;">{'✓ MATCH' if summary['grand_total_match'] else '✗ MISMATCH'}</div>
                </div>
            </div>
        </div>
"""

        # Warning for extra dates
        if summary['extra_dates']:
            html += f"""
        <div class="warning">
            <div class="warning-title">⚠ Warning: Dates in Individual Files NOT in Merge File</div>
            <p>Found {summary['extra_dates_count']} unique date(s) present in individual files but missing from the merge file:</p>
            <div class="missing-dates">
"""
            for date in summary['extra_dates'][:20]:
                html += f"                {date}<br>\n"
            if summary['extra_dates_count'] > 20:
                html += f"                <em>... and {summary['extra_dates_count'] - 20} more</em>\n"
            html += """            </div>
        </div>
"""

        # Detailed results table
        html += """
        <div class="section">
            <h2 class="section-title">Detailed Results</h2>
            <table>
                <thead>
                    <tr>
                        <th>File Name</th>
                        <th>Branch</th>
                        <th>Status</th>
                        <th>Total Rows</th>
                        <th>Unique Dates</th>
                        <th>Missing Dates</th>
                        <th>Individual Amount</th>
                        <th>Merge Amount</th>
                        <th>Amount Diff</th>
                        <th>Details</th>
                    </tr>
                </thead>
                <tbody>
"""

        for result in sorted(self.results, key=lambda x: (x['status'] != 'Success', x['missing_count']), reverse=True):
            if result['status'] == 'Success':
                status_class = 'status-success'
                status_icon = '✓'
            elif result['status'] == 'Error':
                status_class = 'status-error'
                status_icon = '✗'
            else:
                status_class = 'status-missing'
                status_icon = '⚠'

            details = ''
            if result['error']:
                details = f"Error: {result['error']}"
            elif result['missing_dates']:
                dates_str = '<br>'.join([str(d) for d in result['missing_dates'][:10]])
                if result['missing_count'] > 10:
                    dates_str += f"<br><em>... and {result['missing_count'] - 10} more</em>"
                details = f'<div class="missing-dates">{dates_str}</div>'
            else:
                details = 'All dates present'

            # Amount display
            if result['individual_amount'] > 0:
                individual_amt = f"{result['individual_amount']:,.2f}"
                merge_amt = f"{result['merge_amount']:,.2f}"
                diff_amt = f"{result['amount_difference']:,.2f}"
                if not result['amount_match']:
                    diff_amt = f'<span style="color: #dc3545; font-weight: bold;">{diff_amt}</span>'
            else:
                individual_amt = '-'
                merge_amt = '-'
                diff_amt = '-'

            html += f"""
                    <tr>
                        <td>{result['file_name']}</td>
                        <td>{result['branch_name'] if result['branch_name'] else '-'}</td>
                        <td class="{status_class}">{status_icon} {result['status']}</td>
                        <td>{result['total_rows']:,}</td>
                        <td>{result['unique_dates']}</td>
                        <td>{result['missing_count']}</td>
                        <td>{individual_amt}</td>
                        <td>{merge_amt}</td>
                        <td>{diff_amt}</td>
                        <td>{details}</td>
                    </tr>
"""

        html += """
                </tbody>
            </table>
        </div>
    </div>
</body>
</html>
"""
        return html

    def save_reports(self):
        """Save both HTML and text reports."""
        # Text report
        text_report = self.generate_text_report()
        text_path = self.base_dir / REPORT_TXT
        with open(text_path, 'w', encoding='utf-8') as f:
            f.write(text_report)
        print(f"\n✓ Text report saved: {text_path}")

        # HTML report
        html_report = self.generate_html_report()
        html_path = self.base_dir / REPORT_HTML
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_report)
        print(f"✓ HTML report saved: {html_path}")

        return text_path, html_path


def run_dual_validation(base_dir, sales_files, payment_files):
    """Run validation for both sales and payment lines separately."""
    all_results = {'sales': None, 'payment': None}

    # Validate sales lines if merge file exists and there are sales files
    sales_merge_path = Path(base_dir) / SALES_MERGE_FILE
    if sales_merge_path.exists() and sales_files:
        print("\n" + "=" * 80)
        print(f"VALIDATING SALES LINES ({len(sales_files)} files)")
        print("=" * 80)
        sales_validator = ExcelDateValidator(
            base_dir, SALES_MERGE_FILE, file_type='sales',
            date_column=SALES_INDIVIDUAL_DATE_COLUMN,
            merge_date_column=SALES_MERGE_DATE_COLUMN,
            merge_order_ref_column=SALES_MERGE_ORDER_REF_COLUMN,
            merge_amount_column=SALES_MERGE_AMOUNT_COLUMN,
            individual_order_ref_column=SALES_INDIVIDUAL_ORDER_REF_COLUMN,
            individual_amount_column=SALES_INDIVIDUAL_AMOUNT_COLUMN,
            individual_branch_column=SALES_INDIVIDUAL_BRANCH_COLUMN
        )

        # Temporarily replace individual files with only sales files
        original_get_files = sales_validator.get_individual_files
        sales_validator.get_individual_files = lambda: sales_files

        if sales_validator.validate_all():
            all_results['sales'] = sales_validator
    elif not sales_merge_path.exists():
        print(f"\n⚠ Sales merge file not found: {SALES_MERGE_FILE}")
        print(f"  Skipping sales lines validation")
    elif not sales_files:
        print(f"\n⚠ No sales line files found")
        print(f"  Skipping sales lines validation")

    # Validate payment lines if merge file exists and there are payment files
    payment_merge_path = Path(base_dir) / PAYMENT_MERGE_FILE
    if payment_merge_path.exists() and payment_files:
        print("\n" + "=" * 80)
        print(f"VALIDATING PAYMENT LINES ({len(payment_files)} files)")
        print("=" * 80)
        payment_validator = ExcelDateValidator(
            base_dir, PAYMENT_MERGE_FILE, file_type='payment',
            date_column=PAYMENT_INDIVIDUAL_DATE_COLUMN,
            merge_date_column=PAYMENT_MERGE_DATE_COLUMN,
            merge_order_ref_column=PAYMENT_MERGE_ORDER_REF_COLUMN,
            merge_amount_column=PAYMENT_MERGE_AMOUNT_COLUMN,
            individual_order_ref_column=PAYMENT_INDIVIDUAL_ORDER_REF_COLUMN,
            individual_amount_column=PAYMENT_INDIVIDUAL_AMOUNT_COLUMN,
            individual_branch_column=PAYMENT_INDIVIDUAL_BRANCH_COLUMN
        )

        # Temporarily replace individual files with only payment files
        original_get_files = payment_validator.get_individual_files
        payment_validator.get_individual_files = lambda: payment_files

        if payment_validator.validate_all():
            all_results['payment'] = payment_validator
    elif not payment_merge_path.exists():
        print(f"\n⚠ Payment merge file not found: {PAYMENT_MERGE_FILE}")
        print(f"  Skipping payment lines validation")
    elif not payment_files:
        print(f"\n⚠ No payment line files found")
        print(f"  Skipping payment lines validation")

    return all_results


def categorize_files(base_dir):
    """Categorize individual files into sales and payment lines."""
    base_path = Path(base_dir)
    all_files = list(base_path.glob('*.xlsx')) + list(base_path.glob('*.xls'))

    sales_files = []
    payment_files = []
    uncategorized_files = []

    # Remove merge files from the list
    sales_merge = base_path / SALES_MERGE_FILE
    payment_merge = base_path / PAYMENT_MERGE_FILE

    for file_path in all_files:
        # Skip merge files
        if file_path == sales_merge or file_path == payment_merge:
            continue

        file_name_lower = file_path.name.lower()

        # Check if file matches payment pattern first (to avoid misclassifying as sales)
        # Since "sale" could match in words like "sales", we check payment first
        if re.search(PAYMENT_FILE_PATTERN, file_name_lower, re.IGNORECASE):
            payment_files.append(file_path)
        # Check if file matches sales pattern
        elif re.search(SALES_FILE_PATTERN, file_name_lower, re.IGNORECASE):
            sales_files.append(file_path)
        else:
            # If no pattern matches, add to uncategorized list
            uncategorized_files.append(file_path)
            print(f"⚠ Warning: Could not categorize file: {file_path.name}")

    return sorted(sales_files), sorted(payment_files)


def main():
    """Main execution function."""
    print("=" * 80)
    print("Excel Data Validator - Sales & Payment Lines")
    print("=" * 80)

    # Get the current directory
    base_dir = os.path.dirname(os.path.abspath(__file__))
    if not base_dir:
        base_dir = '.'

    # Categorize files into sales and payment
    print("\nCategorizing files...")
    sales_files, payment_files = categorize_files(base_dir)
    print(f"✓ Found {len(sales_files)} sales line file(s)")
    print(f"✓ Found {len(payment_files)} payment line file(s)")

    # Run dual validation
    results = run_dual_validation(base_dir, sales_files, payment_files)

    # Generate and save reports for each type
    print("\n" + "=" * 80)
    print("Generating reports...")
    print("=" * 80)

    validation_success = True

    # Sales report
    if results['sales']:
        sales_validator = results['sales']
        sales_text_path = Path(base_dir) / 'validation_report_sales.txt'
        sales_html_path = Path(base_dir) / 'validation_report_sales.html'

        text_report = sales_validator.generate_text_report()
        with open(sales_text_path, 'w', encoding='utf-8') as f:
            f.write(text_report)
        print(f"✓ Sales text report saved: {sales_text_path}")

        html_report = sales_validator.generate_html_report()
        with open(sales_html_path, 'w', encoding='utf-8') as f:
            f.write(html_report)
        print(f"✓ Sales HTML report saved: {sales_html_path}")

        # Check sales validation status
        summary = sales_validator.generate_summary()
        if not (summary['files_with_missing'] == 0 and
                summary['error_files'] == 0 and
                summary['files_with_amount_mismatch'] == 0 and
                summary['grand_total_match']):
            validation_success = False

    # Payment report
    if results['payment']:
        payment_validator = results['payment']
        payment_text_path = Path(base_dir) / 'validation_report_payment.txt'
        payment_html_path = Path(base_dir) / 'validation_report_payment.html'

        text_report = payment_validator.generate_text_report()
        with open(payment_text_path, 'w', encoding='utf-8') as f:
            f.write(text_report)
        print(f"✓ Payment text report saved: {payment_text_path}")

        html_report = payment_validator.generate_html_report()
        with open(payment_html_path, 'w', encoding='utf-8') as f:
            f.write(html_report)
        print(f"✓ Payment HTML report saved: {payment_html_path}")

        # Check payment validation status
        summary = payment_validator.generate_summary()
        if not (summary['files_with_missing'] == 0 and
                summary['error_files'] == 0 and
                summary['files_with_amount_mismatch'] == 0 and
                summary['grand_total_match']):
            validation_success = False

    # Print final summary
    print("\n" + "=" * 80)
    print("VALIDATION COMPLETE")
    print("=" * 80)

    # Sales summary
    if results['sales']:
        print("\nSALES LINES VALIDATION:")
        summary = results['sales'].generate_summary()
        if (summary['files_with_missing'] == 0 and
            summary['error_files'] == 0 and
            summary['files_with_amount_mismatch'] == 0 and
            summary['grand_total_match']):
            print("✓ SUCCESS: All sales validations passed!")
        else:
            print("✗ ISSUES FOUND:")
            if summary['files_with_missing'] > 0:
                print(f"  - {summary['files_with_missing']} file(s) have missing dates")
            if summary['files_with_amount_mismatch'] > 0:
                print(f"  - {summary['files_with_amount_mismatch']} file(s) have amount mismatches")
            if not summary['grand_total_match']:
                print(f"  - Grand total mismatch: Diff={summary['grand_total_difference']:,.2f}")
            if summary['error_files'] > 0:
                print(f"  - {summary['error_files']} file(s) had errors")

    # Payment summary
    if results['payment']:
        print("\nPAYMENT LINES VALIDATION:")
        summary = results['payment'].generate_summary()
        if (summary['files_with_missing'] == 0 and
            summary['error_files'] == 0 and
            summary['files_with_amount_mismatch'] == 0 and
            summary['grand_total_match']):
            print("✓ SUCCESS: All payment validations passed!")
        else:
            print("✗ ISSUES FOUND:")
            if summary['files_with_missing'] > 0:
                print(f"  - {summary['files_with_missing']} file(s) have missing dates")
            if summary['files_with_amount_mismatch'] > 0:
                print(f"  - {summary['files_with_amount_mismatch']} file(s) have amount mismatches")
            if not summary['grand_total_match']:
                print(f"  - Grand total mismatch: Diff={summary['grand_total_difference']:,.2f}")
            if summary['error_files'] > 0:
                print(f"  - {summary['error_files']} file(s) had errors")

    if not results['sales'] and not results['payment']:
        print("\n✗ No validations were performed. Please check:")
        print("  - Merge files exist and are named correctly")
        print("  - Individual files exist and match the expected patterns")
        return 1

    return 0 if validation_success else 1


if __name__ == '__main__':
    exit(main())
