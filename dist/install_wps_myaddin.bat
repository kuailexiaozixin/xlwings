@echo off
setlocal
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\install_wps_myaddin.ps1"
if errorlevel 1 (
  echo Installation failed. Review PowerShell output above.
  pause
  exit /b 1
)
echo WPS Spreadsheets add-in installed successfully.
pause
