@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install-wps.ps1"
if errorlevel 1 (
  echo Installation failed.
  pause
  exit /b 1
)
echo Installation completed. Restart WPS Office before using the plugin.
pause
