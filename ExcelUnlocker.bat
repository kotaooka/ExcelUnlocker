@echo off
where /q python
if %ERRORLEVEL% neq 0 (
    echo Python is not installed. Please install Python and try again.
    exit /b
)
if not exist ExcelUnlocker.py (
    echo ExcelUnlocker.py does not exist. Please check the file path and try again.
    exit /b
)
python ExcelUnlocker.py
