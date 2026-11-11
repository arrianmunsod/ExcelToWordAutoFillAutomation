@echo off
cd /d "%~dp0"
py -3 -c "import PySide6, docx, openpyxl" >nul 2>&1
if errorlevel 1 (
  echo Installing required Python packages...
  py -3 -m pip install -r requirements.txt
  if errorlevel 1 (
    echo.
    echo Could not install dependencies. Make sure Python 3.11 or newer is installed and available as py.
    pause
    exit /b 1
  )
)
py -3 app.py
if errorlevel 1 pause
