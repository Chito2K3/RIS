@echo off
title ARMMC Pharmacy - RIS ^& RSMI Automation App
cd /d "%~dp0"
if exist "%~dp0RIS\app\server.py" cd /d "%~dp0RIS"

echo ============================================================
echo   Starting ARMMC Pharmacy RIS ^& RSMI Automation App...
echo   Stopping any previous instance on port 5050...
echo ============================================================
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :5050 ^| findstr LISTENING') do taskkill /F /PID %%a >nul 2>&1

set "PYTHONW_EXE="
set "PYTHON_EXE="
for %%P in ("%~dp0.python\pythonw.exe" "%~dp0RIS\.python\pythonw.exe" "%CD%\.python\pythonw.exe") do (
    if not defined PYTHONW_EXE (
        if exist "%%~fP" set "PYTHONW_EXE=%%~fP"
    )
)
for %%P in ("%~dp0.python\python.exe" "%~dp0RIS\.python\python.exe" "%CD%\.python\python.exe" "%LOCALAPPDATA%\Programs\Python\Python311\python.exe") do (
    if not defined PYTHON_EXE (
        if exist "%%~fP" set "PYTHON_EXE=%%~fP"
    )
)

if defined PYTHONW_EXE (
    echo   Starting background server: "%PYTHONW_EXE%"
    start "" "%PYTHONW_EXE%" "app\server.py" 5050 --open
    echo   Server is now running in the background at http://127.0.0.1:5050 !
    timeout /t 3 >nul
    exit /b 0
)

if not defined PYTHON_EXE set "PYTHON_EXE=python"
echo   Using Python: "%PYTHON_EXE%"
echo   Launching server and opening http://127.0.0.1:5050 ...
"%PYTHON_EXE%" -u "app\server.py" 5050 --open
pause
