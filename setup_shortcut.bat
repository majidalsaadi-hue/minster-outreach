@echo off
:: Run this ONCE to create the desktop icon.
:: After that, double-click the desktop icon to open the CRM.

set "PROJECT=%~dp0"
set "LAUNCHER=%PROJECT%launch_crm.bat"
set "ICON=%PROJECT%assets\crm.ico"
set "SHORTCUT=%USERPROFILE%\Desktop\MISA CRM.lnk"

:: Generate the ICO file (needs Python in PATH — same Python used for the CRM)
python "%PROJECT%assets\make_icon.py"

:: Create the desktop shortcut via PowerShell
powershell -NoProfile -Command ^
  "$ws = New-Object -ComObject WScript.Shell; ^
   $s = $ws.CreateShortcut('%SHORTCUT%'); ^
   $s.TargetPath = '%LAUNCHER%'; ^
   $s.WorkingDirectory = '%PROJECT%'; ^
   $s.IconLocation = '%ICON%'; ^
   $s.Description = 'Open MISA Investor CRM'; ^
   $s.WindowStyle = 7; ^
   $s.Save()"

echo.
echo  Desktop icon created: MISA CRM
echo  Double-click it to launch the CRM.
echo.
pause
