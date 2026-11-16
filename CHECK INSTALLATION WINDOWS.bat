@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not defined LOCALAPPDATA set "LOCALAPPDATA=%USERPROFILE%\AppData\Local"
set "VENV_DIR=%LOCALAPPDATA%\CondensateAnalysis\app_v1"

if not exist "%VENV_DIR%\Scripts\python.exe" (
  echo FAIL: The private Python environment is missing.
  echo Run setup_windows.bat first.
  pause
  exit /b 1
)

"%VENV_DIR%\Scripts\python.exe" "check_installation.py"
set "CHECK_STATUS=%errorlevel%"
echo.
if not "%CHECK_STATUS%"=="0" echo Correct the FAIL message above, then run this check again.
pause
exit /b %CHECK_STATUS%
