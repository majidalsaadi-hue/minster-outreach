@echo off
title MISA Investor Relations CRM
cd /d "%~dp0"

echo =============================================
echo   MISA Investor Relations CRM
echo =============================================
echo.

:: Activate virtual environment if it exists
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
) else if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
) else (
    echo No venv found - using system Python.
)

:: Install / update dependencies silently
python -m pip install -q -r requirements.txt 2>nul

echo Starting CRM... browser will open automatically.
echo Press Ctrl+C to stop the server.
echo.

python -m streamlit run app.py --server.headless false --browser.gatherUsageStats false

pause
