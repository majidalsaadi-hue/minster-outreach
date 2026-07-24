@echo off
title MISA Investor Relations CRM
cd /d "%~dp0"

echo =============================================
echo   MISA Investor Relations CRM
echo =============================================
echo.

:: Pull latest updates from GitHub
echo Checking for updates...
git fetch origin claude/determined-sagan-YNQZB >nul 2>&1
git pull origin claude/determined-sagan-YNQZB
echo.

:: Activate virtual environment if it exists
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
) else if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
) else (
    echo No venv found - using system Python.
)

:: Install / update dependencies silently (pillow needed for donut charts)
python -m pip install -q -r requirements.txt pillow 2>nul

:: Clear pycache so new code is always loaded
rd /s /q modules\__pycache__ 2>nul
rd /s /q __pycache__ 2>nul

echo Starting CRM... browser will open automatically.
echo Press Ctrl+C to stop the server.
echo.

python -m streamlit run app.py --server.headless false --browser.gatherUsageStats false

pause
