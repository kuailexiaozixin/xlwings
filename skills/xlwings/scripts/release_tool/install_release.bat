@echo off
rem xlwings release tool - double-click wrapper for install_release.ps1
rem ASCII only. Forwards all arguments to the PowerShell installer.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_release.ps1" %*
echo.
pause
