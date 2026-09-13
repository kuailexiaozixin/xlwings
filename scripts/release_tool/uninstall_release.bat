@echo off
rem xlwings release tool - double-click wrapper for uninstall_release.ps1
rem ASCII only. Forwards all arguments to the PowerShell uninstaller.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall_release.ps1" %*
echo.
pause
