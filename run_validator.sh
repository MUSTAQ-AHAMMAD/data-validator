#!/bin/bash
# Quick Start Script for Excel Date Validator
# This script helps you quickly set up and run the validator

echo "=========================================="
echo "Excel Date Validator - Quick Start"
echo "=========================================="
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null
then
    echo "❌ Python 3 is not installed. Please install Python 3.6 or higher."
    exit 1
fi

echo "✓ Python 3 is installed"
echo ""

# Check if required packages are installed
echo "Checking required packages..."
python3 -c "import pandas" 2>/dev/null
PANDAS_INSTALLED=$?

python3 -c "import openpyxl" 2>/dev/null
OPENPYXL_INSTALLED=$?

if [ $PANDAS_INSTALLED -ne 0 ] || [ $OPENPYXL_INSTALLED -ne 0 ]; then
    echo "⚠ Some required packages are missing."
    echo "Installing required packages..."
    pip3 install pandas openpyxl xlrd --user
    echo ""
fi

echo "✓ All required packages are installed"
echo ""

# Check if merge file exists
if [ ! -f "49_stores_sales_lines.xlsx" ]; then
    echo "❌ Merge file '49_stores_sales_lines.xlsx' not found in current directory."
    echo "Please make sure you're running this script in the correct directory."
    exit 1
fi

echo "✓ Merge file found"
echo ""

# Run the validator
echo "Running validation..."
echo "=========================================="
echo ""
python3 validate_excel_dates.py

echo ""
echo "=========================================="
echo "✓ Done!"
echo ""
echo "Reports have been generated:"
echo "  📄 validation_report.html - Open in your web browser"
echo "  📄 validation_report.txt  - View in text editor"
echo ""
