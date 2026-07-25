@echo off
REM Double-click this to turn a Squarespace orders CSV into a formatted Excel file.
REM Keep it in the same folder as format_orders.py.
cd /d "%~dp0"

REM 1. Find Python (python or py launcher).
set "PY="
where python >nul 2>nul && set "PY=python"
if not defined PY ( where py >nul 2>nul && set "PY=py" )
if not defined PY (
  echo Python is not installed yet.
  echo Install it from https://www.python.org/downloads/  ^(tick "Add python.exe to PATH"^),
  echo then double-click this file again.
  echo.
  pause
  exit /b
)

REM 2. Make sure the one dependency is present.
%PY% -c "import openpyxl" 2>nul || %PY% -m pip install --quiet openpyxl

REM 3. Choose the CSV -- dropped onto this icon, or via a file picker.
set "CSV=%~1"
if "%CSV%"=="" (
  for /f "delims=" %%F in ('powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; $d=New-Object System.Windows.Forms.OpenFileDialog; $d.Title='Choose your Squarespace orders export (.csv)'; $d.Filter='CSV files (*.csv)|*.csv'; if($d.ShowDialog() -eq 'OK'){$d.FileName}"') do set "CSV=%%F"
)
if "%CSV%"=="" exit /b

REM 4. Run it and open the result.
set "OUT=%CSV:.csv=_formatted.xlsx%"
%PY% format_orders.py "%CSV%" "%OUT%"
if exist "%OUT%" (
  start "" "%OUT%"
) else (
  echo Something went wrong. Make sure you picked a Squarespace orders CSV.
  pause
)
