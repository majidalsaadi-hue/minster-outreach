@echo off
:: Run this ONCE to create the desktop icon.

set "PROJECT=%~dp0"
set "LAUNCHER=%PROJECT%launch_crm.bat"
set "ICON=%PROJECT%assets\crm.ico"
set "SHORTCUT=%USERPROFILE%\Desktop\MISA CRM.lnk"

:: Generate the ICO file
python "%PROJECT%assets\make_icon.py"

:: Create the desktop shortcut (single-line — avoids ^ continuation issues)
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -ComObject WScript.Shell; $s=$ws.CreateShortcut('%SHORTCUT%'); $s.TargetPath='%LAUNCHER%'; $s.WorkingDirectory='%PROJECT%'; $s.IconLocation='%ICON%'; $s.Description='Open MISA Investor CRM'; $s.WindowStyle=7; $s.Save()"

if %errorlevel%==0 (
    echo.
    echo  Desktop icon created: MISA CRM
    echo  Double-click it to launch the CRM.
) else (
    echo.
    echo  ERROR: shortcut creation failed. Try running as Administrator.
)
echo.
pause
