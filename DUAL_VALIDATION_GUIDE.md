# Dual Validation Guide - Sales & Payment Lines

## Overview

The validator now supports **separate validation** for sales lines and payment lines:
- **Sales line files** are validated against the **sales merge file**
- **Payment line files** are validated against the **payment merge file**
- **No cross-validation** between sales and payments

## Current File Structure

Based on the repository contents:

### Merge Files
- ✅ **Sales Merge File**: `49_stores_sales_lines.xlsx` (18MB, exists)
- ❌ **Payment Merge File**: `49_stores_payment_lines.xlsx` (needs to be added)

### Individual Files
- **Sales Files**: 0 files (pattern: files containing "sales" and "line" in name)
- **Payment Files**: 49 files (pattern: files containing "payment" and "line" in name)
  - Example: `SALAMRYD payment line 5 to 31 March.xlsx`

## Configuration

Edit the top of `validate_excel_dates.py` to configure:

```python
# Configuration - Sales Lines
SALES_MERGE_FILE = '49_stores_sales_lines.xlsx'
SALES_FILE_PATTERN = r'.*sales.*line.*\.xlsx$'  # Pattern to identify sales line files
SALES_MERGE_DATE_COLUMN = 'Order Lines/Order Ref/Date'
SALES_MERGE_AMOUNT_COLUMN = 'Order Lines/Subtotal'
# ... more sales config ...

# Configuration - Payment Lines
PAYMENT_MERGE_FILE = '49_stores_payment_lines.xlsx'  # UPDATE THIS!
PAYMENT_FILE_PATTERN = r'.*payment.*line.*\.xlsx$'  # Pattern to identify payment line files
PAYMENT_MERGE_DATE_COLUMN = 'Payment Lines/Order Ref/Date'
PAYMENT_MERGE_AMOUNT_COLUMN = 'Payment Lines/Amount'
# ... more payment config ...
```

## What You Need to Do

### Option 1: Add Payment Merge File (Recommended)
1. Add your payment lines merge file to the repository
2. Update `PAYMENT_MERGE_FILE` in the script to match your file name
3. Update column names if different from defaults
4. Run the validator: `python3 validate_excel_dates.py`

### Option 2: Rename Existing Merge File
If `49_stores_sales_lines.xlsx` is actually a payment merge file:
1. Rename it to indicate it's for payments (e.g., `49_stores_payment_lines.xlsx`)
2. Update `PAYMENT_MERGE_FILE` configuration
3. Run the validator

### Option 3: Create Sales Individual Files
If you have sales data in separate files:
1. Add your individual sales line files to the repository
2. Ensure filenames contain "sales" and "line" (e.g., `STORE_sales_lines.xlsx`)
3. The validator will automatically categorize them
4. Run the validator

## File Categorization

The validator automatically categorizes files based on their names:

- **Sales files**: Match pattern `.*sales.*line.*\.xlsx$` (case-insensitive)
  - Example: `STORE_sales_lines.xlsx`, `Store Sales Line Data.xlsx`

- **Payment files**: Match pattern `.*payment.*line.*\.xlsx$` (case-insensitive)
  - Example: `STORE payment line 5 to 31 March.xlsx`

- **Merge files**: Files matching `SALES_MERGE_FILE` or `PAYMENT_MERGE_FILE` are excluded from validation

## Running the Validator

```bash
# Run the validator
python3 validate_excel_dates.py

# Or use the convenience scripts
./run_validator.sh    # Linux/Mac
run_validator.bat     # Windows
```

## Output Reports

The validator generates **separate reports** for each validation type:

### Sales Validation Reports
- `validation_report_sales.html` - Interactive HTML report
- `validation_report_sales.txt` - Plain text report

### Payment Validation Reports
- `validation_report_payment.html` - Interactive HTML report
- `validation_report_payment.txt` - Plain text report

## Console Output

The validator shows separate sections for each validation:

```
================================================================================
Excel Data Validator - Sales & Payment Lines
================================================================================

Categorizing files...
✓ Found 2 sales line file(s)
✓ Found 49 payment line file(s)

================================================================================
VALIDATING SALES LINES (2 files)
================================================================================
Loading sales merge file: 49_stores_sales_lines.xlsx
✓ Loaded 417424 rows with 27 unique dates
...

================================================================================
VALIDATING PAYMENT LINES (49 files)
================================================================================
Loading payment merge file: 49_stores_payment_lines.xlsx
✓ Loaded 125000 rows with 27 unique dates
...

================================================================================
VALIDATION COMPLETE
================================================================================

SALES LINES VALIDATION:
✓ SUCCESS: All sales validations passed!

PAYMENT LINES VALIDATION:
✓ SUCCESS: All payment validations passed!
```

## Column Configuration

### Sales Lines Columns

**Merge File:**
- Date: `Order Lines/Order Ref/Date`
- Order Reference: `Order Lines/Order Ref`
- Amount: `Order Lines/Subtotal`

**Individual Files:**
- Date: `Date`
- Order Reference: `Order Ref`
- Amount: `Subtotal`
- Branch: `Branch`

### Payment Lines Columns

**Merge File:**
- Date: `Payment Lines/Order Ref/Date`
- Order Reference: `Payment Lines/Order Ref`
- Amount: `Payment Lines/Amount`

**Individual Files:**
- Date: `Date`
- Order Reference: `Order Ref`
- Amount: `Payments/Amount`
- Branch: `Branch`

Update these in the configuration if your files use different column names.

## Validation Logic

For each file type (sales or payment):

1. **Date Validation**: Checks if all dates in individual files exist in the corresponding merge file
2. **Amount Validation**: Compares total amounts per branch between individual files and merge file
3. **Grand Total**: Validates sum of all individual files matches merge file total
4. **Order Matching**: Identifies orders present in one file but not the other

## Troubleshooting

### "No validations were performed"
- Check that merge files exist with correct names
- Verify individual files match the expected patterns
- Check file permissions

### "Merge file not found"
- Verify the merge file name in the configuration
- Ensure the file is in the same directory as the script

### "Column not found"
- Check that your Excel files have the expected column names
- Update column configuration if your files use different names

### Files not categorized correctly
- Update `SALES_FILE_PATTERN` or `PAYMENT_FILE_PATTERN` to match your naming convention
- Use Python regex syntax for patterns

## Benefits of Dual Validation

✅ **Separate validation**: Sales and payment data are validated independently
✅ **No cross-contamination**: Prevents validating payment files against sales merge file
✅ **Clear reporting**: Separate reports for each validation type
✅ **Flexible configuration**: Customize patterns and columns for each file type
✅ **Backward compatible**: If only one type exists, it validates just that type

## Questions?

If you encounter issues:
1. Check the configuration matches your file structure
2. Verify column names in your Excel files
3. Ensure merge files exist and are in the correct location
4. Review the console output for specific errors
