@echo off
echo Checking Python installation...

where python >nul 2>nul
if errorlevel 1 (
    echo Python is not installed or not in PATH.
    echo Please install Python from https://www.python.org/downloads/
    pause
    exit /b
)

echo Installing required packages...
pip install -r requirements.txt

echo Launching the tool...
python report_builder_gui.pyw
