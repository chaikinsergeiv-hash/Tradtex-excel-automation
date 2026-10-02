# ================================================================
# WB ANALYTICS — ДИАГНОСТИКА ИСТОРИИ ОСТАТКОВ
#
# Цель:
#   выяснить фактические колонки STOCK_HISTORY_DAILY_CSV
#   и затем посчитать дни в наличии за текущий/прошлый периоды.
#
# Использует первые 20 nmId из 01_SALES_FUNNEL_RAW.json.
# Период: 14 полных дней (прошлый 7д + текущий 7д).
# stockType = "" -> все склады: WB + продавец/FBS.
#
# ВАЖНО:
#   создание отчёта расходует 1 задачу из суточного лимита CSV-отчётов WB.
# ================================================================

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TokenPath = Join-Path $ScriptDir "wb_analytics_token.txt"
$FunnelRawPath = Join-Path $ScriptDir "01_SALES_FUNNEL_RAW.json"

$ZipPath = Join-Path $ScriptDir "02_STOCK_HISTORY_RAW.zip"
$ExtractDir = Join-Path $ScriptDir "02_STOCK_HISTORY_EXTRACTED"
$SamplePath = Join-Path $ScriptDir "02_STOCK_HISTORY_SAMPLE.csv"

function Invoke-WBPostUtf8 {
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [Parameter(Mandatory=$true)][string]$JsonBody,
        [Parameter(Mandatory=$true)][string]$AuthorizationHeader,
        [int]$TimeoutMilliseconds = 120000
    )

    $request = [System.Net.HttpWebRequest]::Create($Uri)
    $request.Method = "POST"
    $request.Timeout = $TimeoutMilliseconds
    $request.ReadWriteTimeout = $TimeoutMilliseconds
    $request.Accept = "application/json"
    $request.ContentType = "application/json; charset=utf-8"
    $request.Headers["Authorization"] = $AuthorizationHeader

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($JsonBody)
    $request.ContentLength = $bytes.Length

    $requestStream = $request.GetRequestStream()
    try {
        $requestStream.Write($bytes, 0, $bytes.Length)
    }
    finally {
        $requestStream.Dispose()
    }

    $response = $null
    try {
        $response = [System.Net.HttpWebResponse]$request.GetResponse()
        $stream = $response.GetResponseStream()
        $memory = New-Object System.IO.MemoryStream

        try {
            $stream.CopyTo($memory)
            $rawBytes = $memory.ToArray()
        }
        finally {
            if ($null -ne $stream) { $stream.Dispose() }
            $memory.Dispose()
        }

        $text = [System.Text.Encoding]::UTF8.GetString($rawBytes)
        if ([string]::IsNullOrWhiteSpace($text)) { return $null }
        return ($text | ConvertFrom-Json)
    }
    finally {
        if ($null -ne $response) { $response.Dispose() }
    }
}

function Invoke-WBGetUtf8 {
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [Parameter(Mandatory=$true)][string]$AuthorizationHeader,
        [int]$TimeoutMilliseconds = 120000
    )

    $request = [System.Net.HttpWebRequest]::Create($Uri)
    $request.Method = "GET"
    $request.Timeout = $TimeoutMilliseconds
    $request.ReadWriteTimeout = $TimeoutMilliseconds
    $request.Accept = "application/json"
    $request.Headers["Authorization"] = $AuthorizationHeader

    $response = $null
    try {
        $response = [System.Net.HttpWebResponse]$request.GetResponse()
        $stream = $response.GetResponseStream()
        $memory = New-Object System.IO.MemoryStream

        try {
            $stream.CopyTo($memory)
            $rawBytes = $memory.ToArray()
        }
        finally {
            if ($null -ne $stream) { $stream.Dispose() }
            $memory.Dispose()
        }

        $text = [System.Text.Encoding]::UTF8.GetString($rawBytes)
        if ([string]::IsNullOrWhiteSpace($text)) { return $null }
        return ($text | ConvertFrom-Json)
    }
    finally {
        if ($null -ne $response) { $response.Dispose() }
    }
}

function Download-WBFile {
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [Parameter(Mandatory=$true)][string]$AuthorizationHeader,
        [Parameter(Mandatory=$true)][string]$Path,
        [int]$TimeoutMilliseconds = 180000
    )

    $request = [System.Net.HttpWebRequest]::Create($Uri)
    $request.Method = "GET"
    $request.Timeout = $TimeoutMilliseconds
    $request.ReadWriteTimeout = $TimeoutMilliseconds
    $request.Headers["Authorization"] = $AuthorizationHeader

    $response = $null
    try {
        $response = [System.Net.HttpWebResponse]$request.GetResponse()
        $stream = $response.GetResponseStream()
        $file = [System.IO.File]::Create($Path)

        try {
            $stream.CopyTo($file)
        }
        finally {
            if ($null -ne $stream) { $stream.Dispose() }
            $file.Dispose()
        }
    }
    finally {
        if ($null -ne $response) { $response.Dispose() }
    }
}

if (-not (Test-Path -LiteralPath $TokenPath)) {
    throw "Не найден wb_analytics_token.txt."
}

if (-not (Test-Path -LiteralPath $FunnelRawPath)) {
    throw "Не найден 01_SALES_FUNNEL_RAW.json. Сначала запустите диагностику воронки."
}

$Token = (Get-Content -LiteralPath $TokenPath -Raw).Trim()
$Auth = "Bearer $Token"

# Берём nmId из уже успешной диагностики воронки.
$funnelText = [System.IO.File]::ReadAllText($FunnelRawPath, [System.Text.Encoding]::UTF8)
$funnel = $funnelText | ConvertFrom-Json

$nmIds = @(
    @($funnel.data.products) |
        ForEach-Object { [Int64]$_.product.nmId } |
        Select-Object -First 20
)

if ($nmIds.Count -eq 0) {
    throw "В 01_SALES_FUNNEL_RAW.json не найдены nmId."
}

$DateTo = (Get-Date).Date.AddDays(-1)
$DateFrom = $DateTo.AddDays(-13)

$BeginDate = $DateFrom.ToString("yyyy-MM-dd")
$EndDate = $DateTo.ToString("yyyy-MM-dd")

$ReportId = [Guid]::NewGuid().ToString()

$bodyObject = [ordered]@{
    id = $ReportId
    reportType = "STOCK_HISTORY_DAILY_CSV"
    userReportName = "Tradtex stock history diagnostic"
    params = [ordered]@{
        nmIds = $nmIds
        subjectIds = @()
        brandNames = @()
        tagIds = @()
        currentPeriod = [ordered]@{
            start = $BeginDate
            end = $EndDate
        }
        stockType = ""
        skipDeletedNm = $true
    }
}

$bodyJson = $bodyObject | ConvertTo-Json -Depth 10

Write-Host "WB Analytics — диагностика истории остатков" -ForegroundColor Cyan
Write-Host "Период: $BeginDate - $EndDate"
Write-Host "Артикулов в тесте: $($nmIds.Count)"
Write-Host "Склады: все (WB + FBS)"
Write-Host "Report ID: $ReportId"
Write-Host ""
Write-Host "Создаю 1 CSV-отчёт WB..."

$create = Invoke-WBPostUtf8 `
    -Uri "https://seller-analytics-api.wildberries.ru/api/v2/nm-report/downloads" `
    -JsonBody $bodyJson `
    -AuthorizationHeader $Auth

Write-Host "Ответ создания: $($create.data)"

$statusUri = "https://seller-analytics-api.wildberries.ru/api/v2/nm-report/downloads?filter%5BdownloadIds%5D=$ReportId"
$status = ""
$maxChecks = 10

for ($i = 1; $i -le $maxChecks; $i++) {
    Write-Host "Жду 22 сек. перед проверкой статуса ($i/$maxChecks)..."
    Start-Sleep -Seconds 22

    $statusResponse = Invoke-WBGetUtf8 -Uri $statusUri -AuthorizationHeader $Auth
    $items = @($statusResponse.data)

    $report = $items | Where-Object { $_.id -eq $ReportId } | Select-Object -First 1

    if ($null -eq $report) {
        Write-Host "Отчёт пока не найден в списке."
        continue
    }

    $status = [string]$report.status
    Write-Host "Статус: $status"

    if ($status -eq "SUCCESS") {
        break
    }

    if ($status -eq "FAILED") {
        throw "WB вернул статус FAILED. Report ID: $ReportId"
    }
}

if ($status -ne "SUCCESS") {
    throw "Отчёт не успел подготовиться за отведённое время. Report ID: $ReportId"
}

# Держим интервал API перед скачиванием.
Write-Host "Отчёт готов. Жду 22 сек. перед скачиванием..."
Start-Sleep -Seconds 22

$downloadUri = "https://seller-analytics-api.wildberries.ru/api/v2/nm-report/downloads/file/$ReportId"
Download-WBFile -Uri $downloadUri -AuthorizationHeader $Auth -Path $ZipPath

if (Test-Path -LiteralPath $ExtractDir) {
    Remove-Item -LiteralPath $ExtractDir -Recurse -Force
}

New-Item -ItemType Directory -Path $ExtractDir | Out-Null

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $ExtractDir)

$csvFiles = @(Get-ChildItem -LiteralPath $ExtractDir -Filter "*.csv" -File -Recurse)

if ($csvFiles.Count -eq 0) {
    throw "В ZIP не найден CSV."
}

Copy-Item -LiteralPath $csvFiles[0].FullName -Destination $SamplePath -Force

Write-Host ""
Write-Host "ГОТОВО" -ForegroundColor Green
Write-Host "ZIP: $ZipPath"
Write-Host "CSV для проверки: $SamplePath"
Write-Host ""
Write-Host "Пришлите 02_STOCK_HISTORY_SAMPLE.csv для проверки колонок и расчёта дней в наличии."
