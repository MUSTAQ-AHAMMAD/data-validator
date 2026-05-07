# Data Validator

## Excel Date Validator Tool - Sales & Payment Lines

This repository contains a comprehensive Python script that validates dates and amounts from individual store Excel files against their respective merge files:

- **Sales line files** → validated against sales merge file
- **Payment line files** → validated against payment merge file
- **No cross-validation** between sales and payments

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

✅ **Separate validation** for sales lines and payment lines
✅ **Automatic file categorization** based on filename patterns
✅ **Beautiful HTML reports** with visual dashboard (separate for sales and payment)
✅ **Text reports** for easy terminal viewing
✅ **Date validation** - ensures all dates are present in merge files
✅ **Amount validation** - compares totals between individual and merge files
✅ **Grand total validation** - verifies sum of all individual files
✅ **Comprehensive statistics** and file-by-file analysis

### Documentation

See [DUAL_VALIDATION_GUIDE.md](DUAL_VALIDATION_GUIDE.md) for detailed setup and configuration guide.

See [README_VALIDATOR.md](README_VALIDATOR.md) for complete documentation including:
- Installation instructions
- Usage guide
- Report interpretation
- Customization options
- Troubleshooting

### Output Files

After running the validator, you'll get separate reports for each validation type:

**Sales Validation:**
- `validation_report_sales.html` - Interactive HTML report (open in browser)
- `validation_report_sales.txt` - Plain text report

**Payment Validation:**
- `validation_report_payment.html` - Interactive HTML report (open in browser)
- `validation_report_payment.txt` - Plain text report

### Requirements

- Python 3.6 or higher
- pandas, openpyxl, xlrd (auto-installed by run scripts)