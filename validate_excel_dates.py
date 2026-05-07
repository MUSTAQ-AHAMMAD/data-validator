#!/usr/bin/env python3
"""
Excel Date Validator - Comprehensive Data Validation Tool
==========================================================
This script validates that all dates from individual store Excel files
are present in the merge file (49_stores_sales_lines.xlsx).

It generates a detailed HTML and text report showing:
- Summary statistics
- Missing dates per file
- Extra dates not in merge file
- Complete validation results
"""

import pandas as pd
import os
from datetime import datetime
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

# Configuration
MERGE_FILE = '49_stores_sales_lines.xlsx'
REPORT_HTML = 'validation_report.html'
REPORT_TXT = 'validation_report.txt'
MERGE_DATE_COLUMN = 'Order Lines/Order Ref/Date'
INDIVIDUAL_DATE_COLUMN = 'Date'


class ExcelDateValidator:
    """Validates dates across multiple Excel files against a merge file."""

    def __init__(self, base_dir, merge_file):
        self.base_dir = Path(base_dir)
        self.merge_file = self.base_dir / merge_file
        self.results = []
        self.merge_dates = set()
        self.all_individual_dates = set()

    def load_merge_file(self):
        """Load and extract dates from the merge file."""
        print(f"Loading merge file: {self.merge_file}")
        try:
            df = pd.read_excel(self.merge_file)
            if MERGE_DATE_COLUMN not in df.columns:
                raise ValueError(f"Column '{MERGE_DATE_COLUMN}' not found in merge file")

            # Convert to datetime and extract date only
            df[MERGE_DATE_COLUMN] = pd.to_datetime(df[MERGE_DATE_COLUMN])
            self.merge_dates = set(df[MERGE_DATE_COLUMN].dt.date)

            print(f"✓ Loaded {len(df)} rows with {len(self.merge_dates)} unique dates")
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
        """Validate a single Excel file against the merge file."""
        file_name = file_path.name
        print(f"\nValidating: {file_name}")

        result = {
            'file_name': file_name,
            'status': 'Success',
            'total_rows': 0,
            'unique_dates': 0,
            'missing_dates': [],
            'missing_count': 0,
            'error': None
        }

        try:
            df = pd.read_excel(file_path)

            if INDIVIDUAL_DATE_COLUMN not in df.columns:
                result['status'] = 'Error'
                result['error'] = f"Column '{INDIVIDUAL_DATE_COLUMN}' not found"
                print(f"  ✗ {result['error']}")
                return result

            # Convert to datetime and extract date only
            df[INDIVIDUAL_DATE_COLUMN] = pd.to_datetime(df[INDIVIDUAL_DATE_COLUMN])
            file_dates = set(df[INDIVIDUAL_DATE_COLUMN].dt.date)

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
                for date in sorted(missing)[:5]:  # Show first 5
                    print(f"    - {date}")
                if len(missing) > 5:
                    print(f"    ... and {len(missing) - 5} more")
            else:
                print(f"  ✓ All {result['unique_dates']} unique dates are present in merge file")

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
        files_with_missing = len([r for r in self.results if r['status'] == 'Missing Dates'])
        error_files = len([r for r in self.results if r['status'] == 'Error'])
        total_missing_dates = sum(r['missing_count'] for r in self.results)

        # Find dates in individual files but not in merge
        extra_dates = self.all_individual_dates - self.merge_dates

        summary = {
            'total_files': total_files,
            'success_files': success_files,
            'files_with_missing': files_with_missing,
            'error_files': error_files,
            'total_missing_dates': total_missing_dates,
            'extra_dates': sorted(extra_dates),
            'extra_dates_count': len(extra_dates),
            'merge_unique_dates': len(self.merge_dates),
            'individual_unique_dates': len(self.all_individual_dates)
        }

        return summary

    def generate_text_report(self):
        """Generate a detailed text report."""
        summary = self.generate_summary()

        report_lines = []
        report_lines.append("=" * 80)
        report_lines.append("EXCEL DATE VALIDATION REPORT")
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
        report_lines.append(f"Files with Errors:            {summary['error_files']}")
        report_lines.append(f"Total Missing Date Entries:   {summary['total_missing_dates']}")
        report_lines.append("")
        report_lines.append(f"Unique Dates in Merge File:   {summary['merge_unique_dates']}")
        report_lines.append(f"Unique Dates in All Individual Files: {summary['individual_unique_dates']}")

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

        # Group by status
        for status in ['Missing Dates', 'Success', 'Error']:
            files = [r for r in self.results if r['status'] == status]
            if not files:
                continue

            report_lines.append("")
            report_lines.append(f"{status.upper()} ({len(files)} files)")
            report_lines.append("-" * 80)

            for result in files:
                report_lines.append(f"\nFile: {result['file_name']}")
                report_lines.append(f"  Total Rows: {result['total_rows']}")
                report_lines.append(f"  Unique Dates: {result['unique_dates']}")

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
        if summary['files_with_missing'] == 0 and summary['error_files'] == 0:
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
    <title>Excel Date Validation Report</title>
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
                <h1>📊 Excel Date Validation Report</h1>
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
                        <th>Status</th>
                        <th>Total Rows</th>
                        <th>Unique Dates</th>
                        <th>Missing Dates</th>
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

            html += f"""
                    <tr>
                        <td>{result['file_name']}</td>
                        <td class="{status_class}">{status_icon} {result['status']}</td>
                        <td>{result['total_rows']:,}</td>
                        <td>{result['unique_dates']}</td>
                        <td>{result['missing_count']}</td>
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


def main():
    """Main execution function."""
    print("=" * 80)
    print("Excel Date Validator")
    print("=" * 80)

    # Get the current directory
    base_dir = os.path.dirname(os.path.abspath(__file__))
    if not base_dir:
        base_dir = '.'

    # Create validator
    validator = ExcelDateValidator(base_dir, MERGE_FILE)

    # Run validation
    if validator.validate_all():
        # Generate and save reports
        print("\n" + "=" * 80)
        print("Generating reports...")
        print("=" * 80)

        text_path, html_path = validator.save_reports()

        # Print summary to console
        summary = validator.generate_summary()
        print("\n" + "=" * 80)
        print("VALIDATION COMPLETE")
        print("=" * 80)

        if summary['files_with_missing'] == 0 and summary['error_files'] == 0:
            print("✓ SUCCESS: All dates from individual files are present in the merge file!")
        else:
            print("✗ ISSUES FOUND:")
            if summary['files_with_missing'] > 0:
                print(f"  - {summary['files_with_missing']} file(s) have missing dates")
            if summary['error_files'] > 0:
                print(f"  - {summary['error_files']} file(s) had errors")

        print(f"\nDetailed reports generated:")
        print(f"  - HTML: {html_path}")
        print(f"  - Text: {text_path}")
    else:
        print("\n✗ Validation failed. Please check the merge file.")
        return 1

    return 0


if __name__ == '__main__':
    exit(main())
