@echo off
chcp 65001 >nul
pushd "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0WB_Ads_7days_EXPERIMENTAL.ps1"
echo.
pause
popd
