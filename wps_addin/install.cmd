@echo off
setlocal
set AUTOUP=0
if exist "%TEMP%\picture_replace_auto_restart.flag" set AUTOUP=1
if exist "%APPDATA%\kingsoft\wps\jsaddons\picture_replace_auto_restart.flag" set AUTOUP=1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install-wps.ps1"
if errorlevel 1 (
  echo Installation failed.
  if not "%AUTOUP%"=="1" pause
  exit /b 1
)
echo Installation completed.
if not "%AUTOUP%"=="1" (
  echo Restart WPS Office before using the plugin.
  pause
)
exit /b 0
