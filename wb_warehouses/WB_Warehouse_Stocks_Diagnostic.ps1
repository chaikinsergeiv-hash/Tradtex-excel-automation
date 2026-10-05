# WB warehouse stocks diagnostic
$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TokenPath = Join-Path $ScriptDir "wb_analytics_token.txt"
$RawPath = Join-Path $ScriptDir "01_WB_WAREHOUSES_RAW.json"
$CsvPath = Join-Path $ScriptDir "01_WB_WAREHOUSES_SAMPLE.csv"

if (-not (Test-Path -LiteralPath $TokenPath)) {
    throw "Не найден wb_analytics_token.txt"
}

$Token = (Get-Content -LiteralPath $TokenPath -Raw).Trim()

$Body = @{
    nmIds = @()
    chrtIds = @()
    limit = 1000
    offset = 0
} | ConvertTo-Json -Depth 5

Write-Host "Диагностика остатков WB. Будет выполнен 1 запрос." -ForegroundColor Cyan

$Client = New-Object System.Net.WebClient
$Client.Encoding = [System.Text.Encoding]::UTF8
$Client.Headers.Add("Authorization", ("Bearer " + $Token))
$Client.Headers.Add("Content-Type", "application/json; charset=utf-8")

try {
    $ResponseText = $Client.UploadString(
        "https://seller-analytics-api.wildberries.ru/api/analytics/v1/stocks-report/wb-warehouses",
        "POST",
        $Body
    )
}
finally {
    $Client.Dispose()
}

$Response = $ResponseText | ConvertFrom-Json

$Utf8Bom = New-Object System.Text.UTF8Encoding($true)
$RawJson = $Response | ConvertTo-Json -Depth 30
[System.IO.File]::WriteAllText($RawPath, $RawJson, $Utf8Bom)

$Items = @($Response.data.items)

if ($Items.Count -gt 0) {
    $CsvLines = $Items | ConvertTo-Csv -NoTypeInformation -Delimiter ";"
    $CsvEncoding = [System.Text.Encoding]::GetEncoding(1251)
    [System.IO.File]::WriteAllLines($CsvPath, $CsvLines, $CsvEncoding)
}

Write-Host "Готово. Строк: $($Items.Count)" -ForegroundColor Green
Write-Host "RAW: $RawPath"
Write-Host "CSV: $CsvPath"
