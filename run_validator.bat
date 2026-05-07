@echo off
REM Quick Start Script for Excel Date Validator (Windows)
REM This script helps you quickly set up and run the validator

echo ==========================================
echo Excel Date Validator - Quick Start
echo ==========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo X Python is not installed. Please install Python 3.6 or higher.
    pause
    exit /b 1
)

echo [OK] Python is installed
echo.

REM Check if merge file exists
if not exist "49_stores_sales_lines.xlsx" (
    echo X Merge file '49_stores_sales_lines.xlsx' not found in current directory.
    echo Please make sure you're running this script in the correct directory.
    pause
    exit /b 1
)

echo [OK] Merge file found
echo.

REM Install required packages
echo Installing/checking required packages...
python -m pip install pandas openpyxl xlrd --user --quiet
echo.

REM Run the validator
echo Running validation...
echo ==========================================
echo.
python validate_excel_dates.py

echo.
echo ==========================================
echo [OK] Done!
echo.
echo Reports have been generated:
echo   validation_report.html - Open in your web browser
echo   validation_report.txt  - View in text editor
echo.
pause
