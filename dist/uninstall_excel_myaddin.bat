@echo off
setlocal
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\uninstall_excel_myaddin.ps1"
if errorlevel 1 (
  echo Uninstallation failed. Review PowerShell output above.
  pause
  exit /b 1
)
echo Microsoft Excel add-in uninstalled successfully.
pause
