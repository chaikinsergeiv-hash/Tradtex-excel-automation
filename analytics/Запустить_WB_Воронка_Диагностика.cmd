@echo off
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0WB_Sales_Funnel_Diagnostic.ps1"
echo.
pause
