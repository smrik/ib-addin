@echo off
setlocal

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Uninstall-FASAddins.ps1" %*
set "exitCode=%ERRORLEVEL%"

echo.
if not "%exitCode%"=="0" echo Uninstall failed. Review the message above.
if "%exitCode%"=="0" echo Uninstall completed.
pause
exit /b %exitCode%
