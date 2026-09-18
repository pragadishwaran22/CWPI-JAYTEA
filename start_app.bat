@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title MJIPL Printing Issue Report

set "PYTHON_COMMAND="
where py >nul 2>nul
if not errorlevel 1 set "PYTHON_COMMAND=py -3"
if not defined PYTHON_COMMAND (
    where python >nul 2>nul
    if not errorlevel 1 set "PYTHON_COMMAND=python"
)

if not defined PYTHON_COMMAND goto :no_python

if not exist ".venv\Scripts\python.exe" (
    echo [1/3] Creating the application environment...
    %PYTHON_COMMAND% -m venv ".venv"
    if errorlevel 1 goto :failed
) else (
    echo [1/3] Existing application environment found.
)

echo [2/3] Installing or checking required packages...
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto :failed

echo [3/3] Starting the application...
echo The browser will open at http://127.0.0.1:8501
echo Keep this window open while using the application.
start "" powershell.exe -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 4; Start-Process 'http://127.0.0.1:8501'"
".venv\Scripts\python.exe" -m streamlit run app.py --server.headless true --server.address 127.0.0.1 --server.port 8501 --browser.gatherUsageStats false
if errorlevel 1 goto :failed
exit /b 0

:no_python
echo.
echo Python 3 was not found.
echo Install Python 3.11 or newer and select "Add Python to PATH" during installation.
goto :stop

:failed
echo.
echo The application could not start. Review the error message above.
echo You can copy the message or send a screenshot for diagnosis.

:stop
echo.
pause
exit /b 1
