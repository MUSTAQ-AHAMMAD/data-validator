# Data Validator

## Excel Date Validator Tool

This repository contains a comprehensive Python script that validates whether all dates from individual store Excel files are present in the merge file `49_stores_sales_lines.xlsx`.

### Quick Start

**For Linux/Mac:**
```bash
./run_validator.sh
```

**For Windows:**
```
run_validator.bat
```

**Or run directly:**
```bash
python3 validate_excel_dates.py
```

### What You Get

✅ **Validation of 49 store Excel files** against the master merge file
✅ **Beautiful HTML report** with visual dashboard and detailed results
✅ **Text report** for easy terminal viewing
✅ **Missing date detection** - identifies any dates not in merge file
✅ **Extra date detection** - finds dates in individual files but not in merge
✅ **Comprehensive statistics** and file-by-file analysis

### Documentation

See [README_VALIDATOR.md](README_VALIDATOR.md) for complete documentation including:
- Installation instructions
- Usage guide
- Report interpretation
- Customization options
- Troubleshooting

### Output Files

After running the validator, you'll get:
- `validation_report.html` - Interactive HTML report (open in browser)
- `validation_report.txt` - Plain text report

### Requirements

- Python 3.6 or higher
- pandas, openpyxl, xlrd (auto-installed by run scripts)