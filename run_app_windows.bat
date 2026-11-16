@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PORT=18818"
set "APP_SCRIPT=%~dp0condensate_analysis_app.py"
if not defined LOCALAPPDATA set "LOCALAPPDATA=%USERPROFILE%\AppData\Local"
set "VENV_DIR=%LOCALAPPDATA%\CondensateAnalysis\app_v1"

if exist "%VENV_DIR%\Scripts\python.exe" (
  set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"
  set "PYTHON_ARGS="
  goto :python_found
)

where py >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_EXE=py"
  set "PYTHON_ARGS=-3"
  goto :python_found
)

where python >nul 2>nul
if errorlevel 1 (
  echo Python 3 could not be found.
  pause
  exit /b 1
)
set "PYTHON_EXE=python"
set "PYTHON_ARGS="

:python_found

if not exist "%APP_SCRIPT%" (
  echo The app script is missing:
  echo %APP_SCRIPT%
  pause
  exit /b 1
)

if not defined FIJI_EXECUTABLE if exist "%~dp0..\FIJI FOR WINDOWS\Fiji.app\ImageJ-win64.exe" set "FIJI_EXECUTABLE=%~dp0..\FIJI FOR WINDOWS\Fiji.app\ImageJ-win64.exe"
if not defined FIJI_EXECUTABLE if exist "%~dp0..\FIJI FOR WINDOWS\Fiji.app\fiji-windows-x64.exe" set "FIJI_EXECUTABLE=%~dp0..\FIJI FOR WINDOWS\Fiji.app\fiji-windows-x64.exe"
if not defined FIJI_EXECUTABLE if exist "C:\Program Files\Fiji.app\ImageJ-win64.exe" set "FIJI_EXECUTABLE=C:\Program Files\Fiji.app\ImageJ-win64.exe"
if not defined FIJI_PCC_MANDERS_CODE_ROOT set "FIJI_PCC_MANDERS_CODE_ROOT=%~dp0CODE FOR ANALYSIS"

echo Checking Python dependencies...
"%PYTHON_EXE%" %PYTHON_ARGS% -c "import numpy, pandas, tifffile, imagecodecs, PIL, scipy, skimage, openpyxl"
if errorlevel 1 (
  echo.
  echo Required Python packages are missing.
  echo Run setup_windows.bat first, then reopen this launcher.
  pause
  exit /b 1
)

echo.
echo Starting CONDENSATE ANALYSIS APP...
echo Keep this window open while using the app.
echo.

set "CONDENSATE_OPEN_BROWSER=1"
"%PYTHON_EXE%" %PYTHON_ARGS% "%APP_SCRIPT%" %PORT%
set "APP_STATUS=%errorlevel%"
echo.
echo The app stopped with exit code %APP_STATUS%.
pause
exit /b %APP_STATUS%
