@echo off
title MISA CRM — First-time Setup
cd /d "%~dp0"
color 0A

echo.
echo  =========================================================
echo    MISA CRM — New PC Setup
echo  =========================================================
echo.

:: ── 1. Check Python ───────────────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found.
    echo  Download from https://python.org and re-run this script.
    echo  Make sure to tick "Add Python to PATH" during install.
    pause & exit /b 1
)
echo  [OK] Python found.

:: ── 2. Check Git ──────────────────────────────────────────────────────────────
git --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Git not found.
    echo  Download from https://git-scm.com and re-run this script.
    pause & exit /b 1
)
echo  [OK] Git found.

:: ── 3. Create virtual environment ─────────────────────────────────────────────
if not exist "venv\Scripts\activate.bat" (
    echo.
    echo  Creating virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo  [ERROR] Failed to create venv.
        pause & exit /b 1
    )
    echo  [OK] Virtual environment created.
) else (
    echo  [OK] Virtual environment already exists.
)

:: ── 4. Activate venv and install dependencies ─────────────────────────────────
echo.
echo  Installing / updating packages (this may take a minute)...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip -q
python -m pip install -r requirements.txt pillow -q
if errorlevel 1 (
    echo  [ERROR] Package install failed.
    pause & exit /b 1
)
echo  [OK] All packages installed.

:: ── 5. Create desktop shortcut ────────────────────────────────────────────────
echo.
echo  Creating desktop shortcut...
set SCRIPT_DIR=%~dp0
set VBS_PATH=%SCRIPT_DIR%launch_crm.vbs
set SHORTCUT_NAME=MISA CRM

:: Write VBS launcher (silent, no CMD flash)
(
echo Set oWS = WScript.CreateObject^("WScript.Shell"^)
echo oWS.Run Chr^(34^) ^& "%SCRIPT_DIR%launch_crm.bat" ^& Chr^(34^), 1, False
) > "%VBS_PATH%"

:: Create the .lnk on Desktop
powershell -NoProfile -Command ^
  "$ws = New-Object -ComObject WScript.Shell; ^
   $s = $ws.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\%SHORTCUT_NAME%.lnk'); ^
   $s.TargetPath = '%VBS_PATH:\=\\%'; ^
   $s.WorkingDirectory = '%SCRIPT_DIR:\=\\%'; ^
   $s.Description = 'MISA Investor Relations CRM'; ^
   $s.Save()"

if exist "%USERPROFILE%\Desktop\%SHORTCUT_NAME%.lnk" (
    echo  [OK] Shortcut "MISA CRM" created on Desktop.
) else (
    echo  [WARN] Could not auto-create shortcut.
    echo  Manually right-click launch_crm.bat ^> Send to ^> Desktop.
)

echo.
echo  =========================================================
echo    Setup complete! Double-click "MISA CRM" on your Desktop
echo    to launch. It will auto-update every time you open it.
echo  =========================================================
echo.
pause
