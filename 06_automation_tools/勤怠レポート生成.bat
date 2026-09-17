@echo off
cd /d "%~dp0"

echo Checking and installing required packages...
pip install numpy pandas scikit-learn openpyxl >nul 2>&1

python generate_report.py %*
if errorlevel 1 pause