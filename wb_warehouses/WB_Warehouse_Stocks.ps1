# ================================================================
# WB — ОСТАТКИ ПО СКЛАДАМ WILDBERRIES — TEST
#
# Итог:
#   1 строка = 1 артикул WB (nmId)
#
# Текущая особенность WB:
#   API может возвращать только агрегированный склад "Склад WB".
#   Скрипт сделан динамически: если WB снова вернёт реальные warehouseName,
#   они автоматически появятся отдельными колонками.
#
# Поля:
#   Артикул WB
#   <динамические колонки складов>
#   Итого WB
#   В пути к клиенту
#   В пути от клиента
#
# Если один nmId имеет несколько chrtId/размеров, они суммируются.
# ================================================================

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ================================================================
# 1. НАСТРОЙКИ
# ================================================================

$PageSize = 1000
$ApiPauseSeconds = 22

# ================================================================
# 2. ПУТИ
# ================================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$TokenPath = Join-Path $ScriptDir "wb_analytics_token.txt"

$OutputPath = Join-Path $ScriptDir "WB_Остатки_Склады_WB_TEST.csv"
$TempPath = Join-Path $ScriptDir "WB_Остатки_Склады_WB_TEST.tmp.csv"
$LogPath = Join-Path $ScriptDir "WB_Остатки_Склады_WB_TEST.log"

# ================================================================
# 3. ФУНКЦИИ
# ================================================================

function Write-Log {
    param([string]$Message)

    $line = "{0}  {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
}

function Convert-ToInt64OrZero {
    param($Value)

    if ($null -eq $Value) {
        return [Int64]0
    }

    $text = ([string]$Value).Trim()

    if ([string]::IsNullOrWhiteSpace($text)) {
        return [Int64]0
    }

    $number = [Int64]0

    if ([Int64]::TryParse($text, [ref]$number)) {
        return $number
    }

    return [Int64]0
}

function Invoke-WBPostUtf8 {
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [Parameter(Mandatory=$true)][string]$JsonBody,
        [Parameter(Mandatory=$true)][string]$AuthorizationHeader,
        [int]$TimeoutMilliseconds = 180000
    )

    $request = [System.Net.HttpWebRequest]::Create($Uri)
    $request.Method = "POST"
    $request.Timeout = $TimeoutMilliseconds
    $request.ReadWriteTimeout = $TimeoutMilliseconds
    $request.Accept = "application/json"
    $request.ContentType = "application/json; charset=utf-8"
    $request.Headers["Authorization"] = $AuthorizationHeader
    $request.AutomaticDecompression =
        [System.Net.DecompressionMethods]::GZip -bor
        [System.Net.DecompressionMethods]::Deflate

    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($JsonBody)
    $request.ContentLength = $bodyBytes.Length

    $requestStream = $request.GetRequestStream()

    try {
        $requestStream.Write($bodyBytes, 0, $bodyBytes.Length)
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
            if ($null -ne $stream) {
                $stream.Dispose()
            }

            $memory.Dispose()
        }

        $jsonText = [System.Text.Encoding]::UTF8.GetString($rawBytes)

        if ([string]::IsNullOrWhiteSpace($jsonText)) {
            return $null
        }

        return ($jsonText | ConvertFrom-Json)
    }
    catch [System.Net.WebException] {
        $errorBody = ""
        $statusCode = $null

        if ($null -ne $_.Exception.Response) {
            try {
                $errorResponse = [System.Net.HttpWebResponse]$_.Exception.Response
                $statusCode = [int]$errorResponse.StatusCode
                $errorStream = $errorResponse.GetResponseStream()
                $errorMemory = New-Object System.IO.MemoryStream

                try {
                    $errorStream.CopyTo($errorMemory)
                    $errorBody = [System.Text.Encoding]::UTF8.GetString(
                        $errorMemory.ToArray()
                    )
                }
                finally {
                    if ($null -ne $errorStream) {
                        $errorStream.Dispose()
                    }

                    $errorMemory.Dispose()
                }
            }
            catch {
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($errorBody)) {
            throw "Ошибка WB API HTTP ${statusCode}: $($_.Exception.Message).`nОтвет WB: $errorBody"
        }

        throw
    }
    finally {
        if ($null -ne $response) {
            $response.Dispose()
        }
    }
}

# ================================================================
# 4. ТОКЕН
# ================================================================

if (-not (Test-Path -LiteralPath $TokenPath)) {
    throw "Не найден wb_analytics_token.txt рядом со скриптом."
}

$Token = (Get-Content -LiteralPath $TokenPath -Raw).Trim()

if (
    [string]::IsNullOrWhiteSpace($Token) -or
    ($Token -match "ВСТАВ|TOKEN|ТОКЕН")
) {
    throw "wb_analytics_token.txt пустой или содержит шаблон."
}

$Auth = "Bearer $Token"

# ================================================================
# 5. ПОЛУЧАЕМ ВСЕ СТРАНИЦЫ API
# ================================================================

$Uri = "https://seller-analytics-api.wildberries.ru/api/analytics/v1/stocks-report/wb-warehouses"

$AllItems = @()
$offset = 0
$page = 0

Write-Log "============================================================"
Write-Log "Старт выгрузки остатков по складам WB."

while ($true) {
    $page++

    if ($page -gt 1) {
        Write-Log "Пауза $ApiPauseSeconds сек. перед страницей $page."
        Start-Sleep -Seconds $ApiPauseSeconds
    }

    $BodyObject = [ordered]@{
        nmIds = @()
        chrtIds = @()
        limit = $PageSize
        offset = $offset
    }

    $BodyJson = $BodyObject | ConvertTo-Json -Depth 10

    Write-Log "API: страница $page, offset=$offset, limit=$PageSize."

    $Response = Invoke-WBPostUtf8 `
        -Uri $Uri `
        -JsonBody $BodyJson `
        -AuthorizationHeader $Auth

    $PageItems = @()

    if (
        $null -ne $Response -and
        $null -ne $Response.data -and
        $null -ne $Response.data.items
    ) {
        $PageItems = @($Response.data.items)
    }

    $AllItems += $PageItems

    Write-Log "Получено на странице: $($PageItems.Count). Всего: $($AllItems.Count)."

    if ($PageItems.Count -lt $PageSize) {
        break
    }

    $offset += $PageSize
}

if ($AllItems.Count -eq 0) {
    throw "API не вернул данные по остаткам WB."
}

# ================================================================
# 6. АГРЕГАЦИЯ: 1 NMID = 1 СТРОКА
# ================================================================

$WarehouseNamesMap = @{}
$ArticleMap = @{}

foreach ($Item in $AllItems) {
    $nmId = Convert-ToInt64OrZero $Item.nmId

    if ($nmId -le 0) {
        continue
    }

    $nmKey = [string]$nmId

    $warehouseName = [string]$Item.warehouseName

    if ([string]::IsNullOrWhiteSpace($warehouseName)) {
        $warehouseName = "Склад WB"
    }

    $WarehouseNamesMap[$warehouseName] = $true

    if (-not $ArticleMap.ContainsKey($nmKey)) {
        $ArticleMap[$nmKey] = [ordered]@{
            NmId = $nmId
            Warehouses = @{}
            TotalQuantity = [Int64]0
            InWayToClient = [Int64]0
            InWayFromClient = [Int64]0
            ChrtIds = @{}
        }
    }

    $article = $ArticleMap[$nmKey]

    $quantity = Convert-ToInt64OrZero $Item.quantity
    $toClient = Convert-ToInt64OrZero $Item.inWayToClient
    $fromClient = Convert-ToInt64OrZero $Item.inWayFromClient
    $chrtId = Convert-ToInt64OrZero $Item.chrtId

    if (-not $article.Warehouses.ContainsKey($warehouseName)) {
        $article.Warehouses[$warehouseName] = [Int64]0
    }

    $article.Warehouses[$warehouseName] += $quantity
    $article.TotalQuantity += $quantity
    $article.InWayToClient += $toClient
    $article.InWayFromClient += $fromClient

    if ($chrtId -gt 0) {
        $article.ChrtIds[[string]$chrtId] = $true
    }
}

$WarehouseNames = @(
    $WarehouseNamesMap.Keys |
        Sort-Object
)

Write-Log "Уникальных складов/агрегатов WB: $($WarehouseNames.Count)."
Write-Log "Артикулов WB: $($ArticleMap.Count)."

foreach ($warehouseName in $WarehouseNames) {
    Write-Log "Колонка склада: $warehouseName."
}

# ================================================================
# 7. ФОРМИРУЕМ ДИНАМИЧЕСКИЙ CSV
# ================================================================

$Rows = @()

foreach ($article in ($ArticleMap.Values | Sort-Object NmId)) {
    $row = [ordered]@{
        "Артикул WB" = [Int64]$article.NmId
    }

    foreach ($warehouseName in $WarehouseNames) {
        $value = [Int64]0

        if ($article.Warehouses.ContainsKey($warehouseName)) {
            $value = [Int64]$article.Warehouses[$warehouseName]
        }

        $row[$warehouseName] = $value
    }

    $row["Итого WB"] = [Int64]$article.TotalQuantity
    $row["В пути к клиенту"] = [Int64]$article.InWayToClient
    $row["В пути от клиента"] = [Int64]$article.InWayFromClient
    $row["Количество размеров"] = $article.ChrtIds.Count

    $Rows += [pscustomobject]$row
}

$CsvLines = $Rows |
    ConvertTo-Csv -NoTypeInformation -Delimiter ";"

# Такой же формат, как у утверждённой воронки:
# стандартный CSV без sep=;, Windows-1251.
$CsvEncoding = [System.Text.Encoding]::GetEncoding(1251)

[System.IO.File]::WriteAllLines(
    $TempPath,
    $CsvLines,
    $CsvEncoding
)

# ================================================================
# 8. БЕЗОПАСНАЯ ЗАМЕНА ФАЙЛА
# ================================================================

$replaceAttempts = 12
$replaceDelaySec = 5
$replaceSucceeded = $false

for ($attempt = 1; $attempt -le $replaceAttempts; $attempt++) {
    try {
        if (Test-Path -LiteralPath $OutputPath) {
            Remove-Item -LiteralPath $OutputPath -Force -ErrorAction Stop
        }

        Move-Item -LiteralPath $TempPath -Destination $OutputPath -Force -ErrorAction Stop
        $replaceSucceeded = $true
        break
    }
    catch [System.UnauthorizedAccessException] {
        if ($attempt -lt $replaceAttempts) {
            Write-Log "CSV занят. Повтор через $replaceDelaySec сек."
            Start-Sleep -Seconds $replaceDelaySec
        }
    }
    catch [System.IO.IOException] {
        if ($attempt -lt $replaceAttempts) {
            Write-Log "CSV занят. Повтор через $replaceDelaySec сек."
            Start-Sleep -Seconds $replaceDelaySec
        }
    }
}

if (-not $replaceSucceeded) {
    $fallback = Join-Path $ScriptDir (
        "WB_Остатки_Склады_WB_TEST_NEW_{0}.csv" -f
        (Get-Date -Format "yyyyMMdd_HHmmss")
    )

    if (Test-Path -LiteralPath $TempPath) {
        Move-Item -LiteralPath $TempPath -Destination $fallback -Force
    }

    throw "Основной CSV заблокирован. Новые данные сохранены: $fallback"
}

Write-Log "ГОТОВО."
Write-Log "CSV: $OutputPath."

Write-Host ""
Write-Host "ВЫГРУЗКА ОСТАТКОВ WB ЗАВЕРШЕНА" -ForegroundColor Green
Write-Host "Строк API: $($AllItems.Count)"
Write-Host "Артикулов: $($Rows.Count)"
Write-Host "Колонок складов: $($WarehouseNames.Count)"
Write-Host "Файл: $OutputPath"
