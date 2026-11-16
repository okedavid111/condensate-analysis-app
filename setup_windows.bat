@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not defined LOCALAPPDATA set "LOCALAPPDATA=%USERPROFILE%\AppData\Local"
set "VENV_DIR=%LOCALAPPDATA%\CondensateAnalysis\app_v1"

where py >nul 2>nul
if %errorlevel%==0 (
  set "PYTHON_CMD=py -3"
) else (
  where python >nul 2>nul
  if errorlevel 1 (
    echo Python 3.11 or newer is required.
    echo Install it from https://www.python.org/downloads/windows/
    pause
    exit /b 1
  )
  set "PYTHON_CMD=python"
)

%PYTHON_CMD% -c "import sys; raise SystemExit(0 if sys.version_info >= (3,11) else 1)"
if errorlevel 1 (
  echo Python 3.11 or newer is required.
  pause
  exit /b 1
)

echo Creating a private Windows Python environment at:
echo %VENV_DIR%
%PYTHON_CMD% -m venv "%VENV_DIR%"
if errorlevel 1 goto :failed

"%VENV_DIR%\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :failed
"%VENV_DIR%\Scripts\python.exe" -m pip install -r "requirements.txt"
if errorlevel 1 goto :failed
"%VENV_DIR%\Scripts\python.exe" -c "import numpy, pandas, tifffile, imagecodecs, PIL, scipy, skimage, openpyxl"
if errorlevel 1 goto :failed

echo.
echo Setup completed successfully.
echo Double-click run_app_windows.bat to open the app.
pause
exit /b 0

:failed
echo.
echo Setup failed. Read the error above, correct it, and try again.
pause
exit /b 1
