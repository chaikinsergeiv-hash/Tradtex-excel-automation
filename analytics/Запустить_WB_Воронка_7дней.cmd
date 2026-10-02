@echo off
chcp 65001 >nul
pushd "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\WB_Sales_Funnel_7days.ps1"
popd
echo.
pause
