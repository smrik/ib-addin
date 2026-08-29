@echo off
setlocal

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-FASAddins.ps1" %*
set "exitCode=%ERRORLEVEL%"

echo.
if not "%exitCode%"=="0" echo Installation failed. Review the message above.
if "%exitCode%"=="0" echo Installation completed. Reopen Excel and/or PowerPoint.
pause
exit /b %exitCode%
