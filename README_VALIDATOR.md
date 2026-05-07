# Excel Date Validator

A comprehensive Python script that validates whether all dates from individual store Excel files are present in the merge file `49_stores_sales_lines.xlsx`.

## Features

✨ **Comprehensive Validation**
- Validates all Excel files in the directory against the merge file
- Identifies missing dates in the merge file
- Detects extra dates in individual files not present in merge

📊 **Detailed Reporting**
- Generates both HTML and text reports
- Summary statistics with visual indicators
- File-by-file detailed results
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

1. **Loads the merge file**: Reads `49_stores_sales_lines.xlsx` and extracts all unique dates
2. **Scans for individual files**: Automatically finds all other `.xlsx` and `.xls` files
3. **Validates each file**: Checks if all dates in each individual file exist in the merge file
4. **Generates reports**: Creates detailed HTML and text reports

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
  - Files with errors
  - Total missing date entries
  - Unique dates comparison

- **Detailed Results per File**:
  - File name
  - Validation status (Success/Missing Dates/Error)
  - Total rows in file
  - Unique dates count
  - List of missing dates (if any)
  - Error messages (if any)

- **Warnings**:
  - Dates present in individual files but not in merge file

## Understanding the Results

### Success Status ✓
All dates from this file are present in the merge file.

### Missing Dates Status ⚠
Some dates from this file are NOT present in the merge file. The report will list all missing dates.

### Error Status ✗
An error occurred while processing this file (e.g., missing date column, file format issue).

## Example Output

```
================================================================================
Excel Date Validator
================================================================================
Loading merge file: 49_stores_sales_lines.xlsx
✓ Loaded 417424 rows with 27 unique dates

Found 49 individual files to validate
================================================================================

Validating: SALAMRYD payment line 5 to 31 March.xlsx
  ✓ All 27 unique dates are present in merge file

Validating: NAKHDMM payment line 5 to 31 March.xlsx
  ✓ All 27 unique dates are present in merge file

...

================================================================================
VALIDATION COMPLETE
================================================================================
✓ SUCCESS: All dates from individual files are present in the merge file!

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
The script looks for the date column: `Order Lines/Order Ref/Date`

### Individual Files
The script looks for the date column: `Date`

If your files use different column names, you can modify the constants at the top of the script:
```python
MERGE_DATE_COLUMN = 'Order Lines/Order Ref/Date'
INDIVIDUAL_DATE_COLUMN = 'Date'
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
```

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
