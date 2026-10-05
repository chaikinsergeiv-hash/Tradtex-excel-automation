@echo off
chcp 65001 >nul
pushd "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\WB_Warehouse_Stocks_Diagnostic.ps1"
popd
echo.
pause
