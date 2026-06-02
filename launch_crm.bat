@echo off
cd /d "%~dp0"

:: Activate virtual environment (works with venv or .venv)
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
) else if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
)

:: Open browser after a short delay
start "" timeout /t 2 /nobreak >nul & start "" http://localhost:8501

:: Launch the CRM
streamlit run app.py --server.headless false --browser.gatherUsageStats false
