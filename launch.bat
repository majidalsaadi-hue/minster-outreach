@echo off
title MISA CRM Dashboard
cd /d "%~dp0"

echo Updating to latest version...
git pull origin claude/determined-sagan-YNQZB

echo Starting MISA CRM Dashboard...
streamlit run app.py --server.port 8501

pause
