@echo off
title ARMMC Pharmacy - RIS ^& RSMI Automation App
cd /d "%~dp0"
echo ============================================================
echo   Starting ARMMC Pharmacy RIS ^& RSMI Automation App...
echo   Stopping any previous instance on port 5050...
echo ============================================================
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :5050 ^| findstr LISTENING') do taskkill /F /PID %%a >nul 2>&1
echo   Opening http://127.0.0.1:5050 in your default browser...
start "" "http://127.0.0.1:5050"
python "app\server.py" 5050
pause

