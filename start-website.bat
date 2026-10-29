@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python 3.10 or newer, then run this file again.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Setting up the website environment...
  py -m venv .venv
  if errorlevel 1 goto failed
)

echo Checking website packages...
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto failed

echo Starting the website at http://127.0.0.1:5000
start "Multiple VIN Search" http://127.0.0.1:5000
.venv\Scripts\python.exe app.py
goto end

:failed
echo Website setup failed. Check the error above and confirm you have an internet connection for the first setup.
pause

:end
endlocal
