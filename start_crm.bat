@echo off
title MISA CRM
cd /d C:\Users\malsaadi\minster-outreach
call venv\Scripts\activate.bat
echo.
echo  Starting MISA CRM...
echo  Open your browser at http://localhost:8501
echo.
streamlit run app.py
pause
