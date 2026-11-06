@echo off
title ARMMC Pharmacy - RIS ^& RSMI Automation App
cd /d "%~dp0"
echo ============================================================
echo   Starting ARMMC Pharmacy RIS ^& RSMI Automation App...
echo   Opening http://127.0.0.1:5050 in your default browser...
echo ============================================================
start "" "http://127.0.0.1:5050"
python "app\server.py" 5050
pause

