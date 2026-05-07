# Data Validator

## Excel Date Validator Tool - Sales & Payment Lines

This repository contains a comprehensive Python script that validates dates and amounts from individual store Excel files against their respective merge files:

- **Sales line files** → validated against `49_stores_sales_lines.xlsx`
- **Payment line files** → validated against `49_stores_payment_lines.xlsx`
- **Automatic file categorization** based on filename patterns
- **No cross-validation** between sales and payments

### ✅ Key Features

✅ **Proper file categorization** - Sales files match with sales merge, payment files match with payment merge
✅ **Flexible configuration** - Easy to customize for different file names and patterns
✅ **Separate validation** for sales lines and payment lines
✅ **Beautiful HTML reports** with visual dashboard (separate for sales and payment)
✅ **Text reports** for easy terminal viewing
✅ **Date validation** - ensures all dates are present in merge files
✅ **Amount validation** - compares totals between individual and merge files
✅ **Grand total validation** - verifies sum of all individual files
✅ **Comprehensive statistics** and file-by-file analysis

### Quick Start

**The script is already configured for your current files!** Just run:

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

### Changing File Names or Patterns

If your merge files or naming patterns change, see **[CONFIGURATION_GUIDE.md](CONFIGURATION_GUIDE.md)** for detailed instructions on how to customize the validator.

### What You Get

After running the validator, you'll get separate reports for each validation type:

**Sales Validation:**
- `validation_report_sales.html` - Interactive HTML report (open in browser)
- `validation_report_sales.txt` - Plain text report

**Payment Validation:**
- `validation_report_payment.html` - Interactive HTML report (open in browser)
- `validation_report_payment.txt` - Plain text report

### Documentation

See **[CONFIGURATION_GUIDE.md](CONFIGURATION_GUIDE.md)** - **NEW!** Step-by-step guide for customizing file names and patterns

See [DUAL_VALIDATION_GUIDE.md](DUAL_VALIDATION_GUIDE.md) for detailed setup and configuration guide

See [README_VALIDATOR.md](README_VALIDATOR.md) for complete documentation including:
- Installation instructions
- Usage guide
- Report interpretation
- Troubleshooting