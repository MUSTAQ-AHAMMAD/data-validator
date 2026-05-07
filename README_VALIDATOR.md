# Excel Data Validator

A comprehensive Python script that validates whether all dates from individual store Excel files are present in the merge file `49_stores_sales_lines.xlsx`, and validates that payment amounts match between individual files and the merge file.

## Features

✨ **Comprehensive Validation**
- Validates all Excel files in the directory against the merge file
- Identifies missing dates in the merge file
- Detects extra dates in individual files not present in merge
- **NEW:** Validates payment amounts for each individual file against the merge file
- **NEW:** Validates grand total (sum of all individual files vs merge file total)
- **NEW:** Identifies order-level mismatches between individual and merge files

📊 **Detailed Reporting**
- Generates both HTML and text reports
- Summary statistics with visual indicators
- File-by-file detailed results including amount validation
- Amount discrepancy reporting
- Order count comparison
- Beautiful, easy-to-read HTML dashboard

🚀 **Easy to Use**
- Simple one-command execution
- Automatic file discovery
- Clear console output with progress indicators

## Requirements

- Python 3.6 or higher
- Required packages (auto-installable):
  - pandas
  - openpyxl
  - xlrd

## Installation

1. Install required Python packages:
```bash
pip3 install pandas openpyxl xlrd
```

Or:
```bash
pip install pandas openpyxl xlrd
```

## Usage

### Basic Usage

Simply run the script in the directory containing your Excel files:

```bash
python3 validate_excel_dates.py
```

Or:
```bash
python validate_excel_dates.py
```

### What It Does

1. **Loads the merge file**: Reads `49_stores_sales_lines.xlsx` and extracts all unique dates and amounts by branch
2. **Scans for individual files**: Automatically finds all other `.xlsx` and `.xls` files
3. **Validates each file**:
   - Checks if all dates in each individual file exist in the merge file
   - Validates that payment amounts match between individual file and merge file for that branch
   - Compares order counts between individual and merge files
4. **Validates grand totals**: Ensures the sum of all individual file amounts matches the merge file total
5. **Generates reports**: Creates detailed HTML and text reports with all validation results

## Output

The script generates two report files:

### 1. HTML Report (`validation_report.html`)
- Interactive, visually appealing report
- Color-coded status indicators
- Summary dashboard with key metrics
- Sortable detailed results table
- Open in any web browser

### 2. Text Report (`validation_report.txt`)
- Plain text format for easy viewing in terminal or text editor
- Same comprehensive information as HTML report
- Perfect for automated processing or quick terminal review

## Report Contents

Both reports include:

- **Summary Statistics**:
  - Total files validated
  - Files with all dates present
  - Files with missing dates
  - **NEW:** Files with amount mismatches
  - Files with errors
  - Total missing date entries
  - Unique dates comparison
  - **NEW:** Grand total amounts (individual files vs merge file)
  - **NEW:** Grand total validation status

- **Detailed Results per File**:
  - File name
  - Branch name
  - Validation status (Success/Missing Dates/Amount Mismatch/Error)
  - Total rows in file
  - Unique dates count
  - List of missing dates (if any)
  - **NEW:** Individual file total amount
  - **NEW:** Merge file total amount for that branch
  - **NEW:** Amount difference
  - **NEW:** Amount validation status (Match/Mismatch)
  - **NEW:** Order count comparison
  - Error messages (if any)

- **Warnings**:
  - Dates present in individual files but not in merge file
  - **NEW:** Amount mismatches with detailed discrepancy information

## Understanding the Results

### Success Status ✓
All dates from this file are present in the merge file AND amounts match (within tolerance).

### Missing Dates Status ⚠
Some dates from this file are NOT present in the merge file. The report will list all missing dates.

### Amount Mismatch Status ⚠
The total payment amount in the individual file does not match the total in the merge file for that branch. The report shows the exact difference.

### Error Status ✗
An error occurred while processing this file (e.g., missing date column, file format issue, branch not found in merge file).

## Example Output

```
================================================================================
Excel Data Validator
================================================================================
Loading merge file: 49_stores_sales_lines.xlsx
✓ Loaded 417424 rows with 27 unique dates
✓ Computed amounts for 91 branches

Found 49 individual files to validate
================================================================================

Validating: SALAMRYD payment line 5 to 31 March.xlsx
  ✓ All 27 unique dates are present in merge file
  ✗ Amount mismatch: Individual=465431.00, Merge=465339.34, Diff=91.66

Validating: NAKHDMM payment line 5 to 31 March.xlsx
  ✓ All 27 unique dates are present in merge file
  ✗ Amount mismatch: Individual=844299.00, Merge=845854.60, Diff=-1555.60

...

================================================================================
VALIDATION COMPLETE
================================================================================
✗ ISSUES FOUND:
  - 46 file(s) have amount mismatches
  - Grand total mismatch: Individual=23,977,423.81, Merge=23,963,714.69, Diff=13,709.12
  - 2 file(s) had errors

Detailed reports generated:
  - HTML: validation_report.html
  - Text: validation_report.txt
```

## File Structure Expected

The script expects the following structure:

```
your-directory/
├── 49_stores_sales_lines.xlsx          # Merge file (required)
├── STORE1 payment lines.xlsx           # Individual store files
├── STORE2 payment lines.xlsx
├── STORE3 payment lines.xlsx
├── ...
└── validate_excel_dates.py             # This script
```

## Column Names

### Merge File
The script looks for these columns:
- **Date column**: `Order Lines/Order Ref/Date`
- **Order reference column**: `Order Lines/Order Ref`
- **Amount column**: `Order Lines/Subtotal`

### Individual Files
The script looks for these columns:
- **Date column**: `Date`
- **Order reference column**: `Order Ref`
- **Amount column**: `Payments/Amount`
- **Branch column**: `Branch`

If your files use different column names, you can modify the constants at the top of the script:
```python
MERGE_DATE_COLUMN = 'Order Lines/Order Ref/Date'
INDIVIDUAL_DATE_COLUMN = 'Date'
MERGE_ORDER_REF_COLUMN = 'Order Lines/Order Ref'
MERGE_AMOUNT_COLUMN = 'Order Lines/Subtotal'
INDIVIDUAL_ORDER_REF_COLUMN = 'Order Ref'
INDIVIDUAL_AMOUNT_COLUMN = 'Payments/Amount'
INDIVIDUAL_BRANCH_COLUMN = 'Branch'
AMOUNT_TOLERANCE = 0.01  # Allow small rounding differences
```

## Troubleshooting

### "Module not found" Error
Install the required packages:
```bash
pip3 install pandas openpyxl xlrd
```

### "Column not found" Error
Check that your Excel files have the correct column names. Update the column name constants in the script if needed.

### Memory Issues
If processing very large files, ensure you have sufficient RAM available. The script loads files into memory for processing.

## Customization

You can customize the script by modifying these constants at the top:

```python
MERGE_FILE = '49_stores_sales_lines.xlsx'      # Name of merge file
REPORT_HTML = 'validation_report.html'          # HTML report filename
REPORT_TXT = 'validation_report.txt'            # Text report filename
MERGE_DATE_COLUMN = 'Order Lines/Order Ref/Date'    # Date column in merge file
INDIVIDUAL_DATE_COLUMN = 'Date'                     # Date column in individual files
MERGE_ORDER_REF_COLUMN = 'Order Lines/Order Ref'    # Order ref column in merge file
MERGE_AMOUNT_COLUMN = 'Order Lines/Subtotal'        # Amount column in merge file
INDIVIDUAL_ORDER_REF_COLUMN = 'Order Ref'           # Order ref column in individual files
INDIVIDUAL_AMOUNT_COLUMN = 'Payments/Amount'        # Amount column in individual files
INDIVIDUAL_BRANCH_COLUMN = 'Branch'                 # Branch column in individual files
AMOUNT_TOLERANCE = 0.01  # Allow small rounding differences
```

## How Amount Validation Works

The script validates amounts by:

1. **For each individual file**:
   - Sums all payment amounts from the `Payments/Amount` column
   - Extracts the branch name from the `Branch` column
   - Compares with the merge file's total for that branch

2. **For the merge file**:
   - Groups order lines by `Order Lines/Order Ref` (order reference)
   - Sums the `Order Lines/Subtotal` for each order to get order totals
   - Groups orders by branch (extracted from order reference prefix)
   - Calculates total amount per branch

3. **Comparison**:
   - Individual file total amount vs. Merge file total for that branch
   - Allows a small tolerance (0.01) for rounding differences
   - Reports any discrepancies greater than the tolerance

4. **Grand total validation**:
   - Sums all individual file amounts
   - Compares with the total of all merge file amounts
   - Reports overall match or mismatch

## License

Free to use and modify as needed.

## Support

For issues or questions, please check:
1. Excel files are in the correct format
2. Required Python packages are installed
3. Column names match the expected format
4. Files are not corrupted or password-protected

---

**Created for validating store sales data across multiple Excel files.**
