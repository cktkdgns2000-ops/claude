@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
  py -3 hwp_batch_save.py %*
  goto :eof
)
where python >nul 2>nul
if %errorlevel%==0 (
  python hwp_batch_save.py %*
  goto :eof
)
echo.
echo Python is not installed on this PC.
echo Install it from https://www.python.org/downloads/  (check "Add python.exe to PATH" during setup)
echo Then double-click run.bat again.
pause
