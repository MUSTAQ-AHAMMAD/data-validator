# Configuration Guide - Excel Data Validator

This guide explains how to configure the validator script to work with different file names and patterns.

## Quick Start

The validator script is **already configured** to work with your current files:
- `49_stores_sales_lines.xlsx` (sales merge file)
- `49_stores_payment_lines.xlsx` (payment merge file)
- Individual sales files (matching pattern: "sale...line" or "sales...line")
- Individual payment files (matching pattern: "payment...line" or "payment...lines")

## When to Reconfigure

You need to update the configuration if:
1. Your merge file names change (e.g., `50_stores_sales_lines.xlsx`)
2. Your individual files have different naming patterns
3. Your Excel columns have different names
4. You want to adjust the amount tolerance for mismatches

## Configuration Location

All configuration is at the top of `validate_excel_dates.py` (lines 36-61):

```python
# Configuration - Sales Lines
SALES_MERGE_FILE = '49_stores_sales_lines.xlsx'
SALES_FILE_PATTERN = r'sale.*line'
# ... (more settings)

# Configuration - Payment Lines
PAYMENT_MERGE_FILE = '49_stores_payment_lines.xlsx'
PAYMENT_FILE_PATTERN = r'payment.*line'
# ... (more settings)
```

## Configuration Options

### 1. Merge File Names

```python
SALES_MERGE_FILE = '49_stores_sales_lines.xlsx'
PAYMENT_MERGE_FILE = '49_stores_payment_lines.xlsx'
```

**When to change:** When your merge files have different names.

**Examples:**
- `SALES_MERGE_FILE = '50_stores_sales_lines.xlsx'`
- `PAYMENT_MERGE_FILE = 'all_payments.xlsx'`
- `SALES_MERGE_FILE = 'master_sales_data.xlsx'`

### 2. File Patterns

```python
SALES_FILE_PATTERN = r'sale.*line'
PAYMENT_FILE_PATTERN = r'payment.*line'
```

**When to change:** When your individual files follow a different naming convention.

**How it works:**
- These are **regex patterns** (regular expressions)
- They match file names in a case-insensitive manner
- `.*` means "any characters"

**Examples:**

If your sales files are named like:
- `STORE_ABC_sales_march.xlsx`
- `STORE_XYZ_sales_march.xlsx`

Use: `SALES_FILE_PATTERN = r'sales.*march'`

If your payment files are named like:
- `payments_STORE1.xlsx`
- `payments_STORE2.xlsx`

Use: `PAYMENT_FILE_PATTERN = r'^payments_'`

If your files have a prefix:
- `2026_STORE_payment_lines.xlsx`

Use: `PAYMENT_FILE_PATTERN = r'\d{4}_.*payment.*line'`

### 3. Column Names

#### Sales Files Configuration

```python
SALES_INDIVIDUAL_DATE_COLUMN = 'Order Lines/Order Ref/Date'
SALES_INDIVIDUAL_ORDER_REF_COLUMN = 'Order Lines/Order Ref'
SALES_INDIVIDUAL_AMOUNT_COLUMN = 'Order Lines/Subtotal'
SALES_INDIVIDUAL_BRANCH_COLUMN = 'Branch'
```

**When to change:** When your sales files have different column names.

#### Payment Files Configuration

```python
PAYMENT_INDIVIDUAL_DATE_COLUMN = 'Date'
PAYMENT_INDIVIDUAL_ORDER_REF_COLUMN = 'Order Ref'
PAYMENT_INDIVIDUAL_AMOUNT_COLUMN = 'Payments/Amount'
PAYMENT_INDIVIDUAL_BRANCH_COLUMN = 'Branch'
```

**When to change:** When your payment files have different column names.

**How to find column names:**
```bash
# Check your file columns
python3 -c "
import pandas as pd
df = pd.read_excel('YOUR_FILE.xlsx', nrows=1)
print('Columns:', list(df.columns))
"
```

### 4. Merge File Column Names

These follow the same pattern as individual files but for merge files:

```python
SALES_MERGE_DATE_COLUMN = 'Order Lines/Order Ref/Date'
SALES_MERGE_ORDER_REF_COLUMN = 'Order Lines/Order Ref'
SALES_MERGE_AMOUNT_COLUMN = 'Order Lines/Subtotal'

PAYMENT_MERGE_DATE_COLUMN = 'Date'
PAYMENT_MERGE_ORDER_REF_COLUMN = 'Order Ref'
PAYMENT_MERGE_AMOUNT_COLUMN = 'Payments/Amount'
```

### 5. Amount Tolerance

```python
AMOUNT_TOLERANCE = 0.01  # Allow small rounding differences
```

**When to change:** To adjust sensitivity to amount mismatches.

**Examples:**
- `AMOUNT_TOLERANCE = 0.001` - Very strict (0.001 difference allowed)
- `AMOUNT_TOLERANCE = 1.00` - Lenient (1.00 difference allowed)
- `AMOUNT_TOLERANCE = 0` - Exact match required

## Common Scenarios

### Scenario 1: New Month, Same Structure

Your files for April are named:
- `49_stores_sales_lines_april.xlsx`
- `49_stores_payment_lines_april.xlsx`
- Individual files remain the same pattern

**Changes needed:**
```python
SALES_MERGE_FILE = '49_stores_sales_lines_april.xlsx'
PAYMENT_MERGE_FILE = '49_stores_payment_lines_april.xlsx'
# No other changes needed!
```

### Scenario 2: Different Store Count

You now have 50 stores:
- `50_stores_sales_lines.xlsx`
- `50_stores_payment_lines.xlsx`

**Changes needed:**
```python
SALES_MERGE_FILE = '50_stores_sales_lines.xlsx'
PAYMENT_MERGE_FILE = '50_stores_payment_lines.xlsx'
# No other changes needed!
```

### Scenario 3: Completely Different File Names

Your new structure:
- Merge files: `master_sales.xlsx`, `master_payments.xlsx`
- Individual files: `branch_ABC_sales.xlsx`, `branch_ABC_payments.xlsx`

**Changes needed:**
```python
SALES_MERGE_FILE = 'master_sales.xlsx'
PAYMENT_MERGE_FILE = 'master_payments.xlsx'
SALES_FILE_PATTERN = r'branch_.*_sales'
PAYMENT_FILE_PATTERN = r'branch_.*_payments'
```

### Scenario 4: Different Column Names

Your files use different column names:
- Date column: `Transaction Date`
- Amount column: `Total Amount`
- Order reference: `Invoice Number`

**Changes needed:**
```python
PAYMENT_INDIVIDUAL_DATE_COLUMN = 'Transaction Date'
PAYMENT_INDIVIDUAL_AMOUNT_COLUMN = 'Total Amount'
PAYMENT_INDIVIDUAL_ORDER_REF_COLUMN = 'Invoice Number'
# Update merge columns similarly if they differ
```

## Testing Your Configuration

After making changes:

1. **Test file categorization:**
```bash
python3 -c "
from validate_excel_dates import categorize_files
sales, payments = categorize_files('.')
print(f'Sales files: {len(sales)}')
print(f'Payment files: {len(payments)}')
"
```

2. **Run the full validation:**
```bash
python3 validate_excel_dates.py
```

3. **Check the output:**
- Look for "Found X sales line file(s)"
- Look for "Found X payment line file(s)"
- Verify the counts match your expectations

## Troubleshooting

### Files Not Being Categorized

**Problem:** The script says "Found 0 sales line file(s)" or "Found 0 payment line file(s)"

**Solution:**
1. Check your file names match the pattern
2. Test your pattern:
   ```python
   import re
   pattern = r'sale.*line'
   filename = 'STORE_sales_line.xlsx'
   if re.search(pattern, filename.lower()):
       print("Match!")
   else:
       print("No match - adjust pattern")
   ```

### Column Not Found Error

**Problem:** Script says "Column 'XYZ' not found"

**Solution:**
1. Check the column names in your file:
   ```bash
   python3 -c "
   import pandas as pd
   df = pd.read_excel('YOUR_FILE.xlsx', nrows=1)
   for col in df.columns:
       print(f'  - {col}')
   "
   ```
2. Update the configuration with the exact column name (case-sensitive)

### Files Categorized Incorrectly

**Problem:** Sales files being treated as payment files or vice versa

**Solution:**
1. Make patterns more specific
2. Check the order - payment pattern is checked first
3. Example fix:
   ```python
   # Too broad (matches both)
   SALES_FILE_PATTERN = r'sale'

   # Better (more specific)
   SALES_FILE_PATTERN = r'sale.*line'
   ```

## Advanced: Regex Pattern Guide

Quick reference for common regex patterns:

| Pattern | Matches | Example |
|---------|---------|---------|
| `r'sales'` | Contains "sales" | `ABC_sales_123.xlsx` |
| `r'^sales'` | Starts with "sales" | `sales_ABC.xlsx` |
| `r'sales$'` | Ends with "sales" | `ABC_sales.xlsx` |
| `r'sales.*line'` | "sales" followed by "line" | `sales_march_line.xlsx` |
| `r'\d{4}_sales'` | 4 digits + "_sales" | `2026_sales.xlsx` |
| `r'(sale\|sales).*line'` | "sale" OR "sales" + "line" | `sale_line.xlsx`, `sales_line.xlsx` |

## Getting Help

If you're unsure about a configuration:
1. Check this guide
2. Look at the examples in this document
3. Test with a small subset of files first
4. Check the generated reports to verify correct categorization

## Summary

**For most users:** You don't need to change anything! The script is already configured for your current file structure.

**If files change:** Update the file names and patterns at the top of `validate_excel_dates.py`.

**Key principle:** The script is designed to be flexible. All important settings are clearly marked and documented at the top of the file.
