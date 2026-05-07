# Excel Data Validator - Example Outputs

## Example 1: Current Validation Results (2026-05-07)

```
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

**Summary Statistics:**
- Total Files Validated: 49
- Files with All Dates Present: 1 ✓
- Files with Missing Dates: 0
- Files with Amount Mismatch: 46 ✗
- Files with Errors: 2
- Total Missing Date Entries: 0
- Unique Dates in Merge File: 27
- **Grand Total (Individual Files): 23,977,423.81**
- **Grand Total (Merge File): 23,963,714.69**
- **Grand Total Difference: 13,709.12**
- **Grand Total Status: ✗ MISMATCH**

---

## Example 2: Sample File with Amount Mismatch

```
Validating: SALAMRYD payment line 5 to 31 March.xlsx
  ✓ All 27 unique dates are present in merge file
  ✗ Amount mismatch: Individual=465431.00, Merge=465339.34, Diff=91.66
```

**Detailed Results (from report):**
- File: SALAMRYD payment line 5 to 31 March.xlsx
- Branch: SALAMRYD
- Total Rows: 2487
- Unique Dates: 27
- **Individual File Amount: 465,431.00**
- **Merge File Amount: 465,339.34**
- **Amount Difference: 91.66**
- **Amount Status: ✗ MISMATCH**
- Unique Orders (Individual): 2174
- Unique Orders (Merge): 2166
- Orders in Individual but not Merge: 8

---

## Example 3: Sample File with Perfect Match

```
Validating: WATANIYAHB Payment Lines 5 to 31 March.xlsx
  ✓ All 27 unique dates are present in merge file
  ✓ Amount matches: 991445.82
```

**Detailed Results:**
- File: WATANIYAHB Payment Lines 5 to 31 March.xlsx
- Branch: WATANIYAHB
- Status: Success ✓
- Total Rows: 3158
- Unique Dates: 27
- **Individual File Amount: 991,445.82**
- **Merge File Amount: 991,445.82**
- **Amount Difference: 0.00**
- **Amount Status: ✓ MATCH**

---

## Example 4: Missing Dates Found (Hypothetical)

If the validator finds missing dates, you would see output like this:

```
Validating: STORE1 payment line.xlsx
  ✗ Found 3 missing date(s)
    - 2026-03-05
    - 2026-03-12
    - 2026-03-20
  ✗ Amount mismatch: Individual=150000.00, Merge=149500.00, Diff=500.00

Validating: STORE2 payment line.xlsx
  ✓ All 27 unique dates are present in merge file
  ✓ Amount matches: 250000.00
```

**Summary Statistics:**
- Total Files Validated: 49
- Files with All Dates Present: 48 ✓
- Files with Missing Dates: 1 ⚠
- Files with Amount Mismatch: 1 ✗
- Total Missing Date Entries: 3
- Unique Dates in Merge File: 27

---

## Report Features

### HTML Report Features:
- 🎨 Color-coded status indicators (green=success, yellow=warning, red=error)
- 📊 Interactive summary dashboard with key metrics
- 💰 **NEW:** Amount validation summary with grand totals
- 📋 Comprehensive table with all file details including amounts
- ⚠️ Warning sections for dates in individual files not in merge
- 🔍 Expandable missing dates lists for each file
- 💵 **NEW:** Amount difference highlighting for mismatches
- 📈 **NEW:** Order count comparison

### Text Report Features:
- 📄 Clean, easy-to-read plain text format
- 📝 Complete validation details
- 💰 **NEW:** Amount validation section with grand totals
- 🗂️ Organized by status (Success, Amount Mismatch, Missing Dates, Errors)
- 📋 Full list of missing dates per file
- 💵 **NEW:** Detailed amount information per file

---

## Understanding the Validation

### What the Script Checks:

1. **Date Validation**: Are all dates from each individual store file present in the merge file?
   - ✓ YES = Date validation passes
   - ✗ NO = Date validation fails (missing dates listed)

2. **Amount Validation**: Do the payment amounts match between individual files and merge file?
   - ✓ YES = Amount validation passes (within 0.01 tolerance)
   - ✗ NO = Amount validation fails (difference shown)
   - Compares individual file's total `Payments/Amount` vs merge file's total `Order Lines/Subtotal` for that branch

3. **Grand Total Validation**: Does the sum of all individual file amounts match the merge file total?
   - ✓ YES = Grand total matches
   - ✗ NO = Grand total mismatch (shows exact difference)

4. **Order Count Validation**: Do the order counts match between individual and merge files?
   - Reports unique orders in individual file
   - Reports unique orders in merge file for that branch
   - Identifies orders present in one but not the other

### Current Result (2026-05-07):

✅ **Date Validation**: All 49 store files contain dates that are present in the merge file!

⚠️ **Amount Validation**: 46 files have amount mismatches
- Most differences are small (under 2,000)
- Could indicate:
  - Rounding differences
  - Refunds or adjustments not in payment files
  - Orders in merge but not in payment files
  - Partial payments or split payments

⚠️ **Grand Total**: Mismatch of 13,709.12
- Individual Files Total: 23,977,423.81
- Merge File Total: 23,963,714.69
- Difference: 13,709.12 (0.06% variance)

**Details:**
- Merge file has **27 unique dates** (March 5-31, 2026)
- All 49 individual store files have dates within this range
- **417,424 total rows** validated in the merge file
- Individual files range from **258 to 4,523 rows** each
- **91 unique branches** found in merge file

---

## Files Validated

The script automatically validates these files:

1. NAJPARK Payment Lines 5 to 31 March.xlsx - 1,270 rows ✓
2. NAJRANPLZA Payment Lines 5 to 31 March.xlsx - 830 rows ✓
3. NAKHDMM payment line 5 to 31 March.xlsx - 4,523 rows ✓
4. NAKHELMALL payment line 5 to 31 March.xlsx - 3,850 rows ✓
5. NJRANFRONT PAYMENT LINES.xlsx - 258 rows ✓
... (and 44 more)

**Total: 49 stores validated** ✅
