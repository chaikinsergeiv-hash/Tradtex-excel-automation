# ================================================================
# WB ANALYTICS — ВОРОНКА ПРОДАЖ ЗА 7 ДНЕЙ — TEST
#
# Итог:
#   1 строка = 1 артикул WB
#
# Периоды:
#   текущий = последние 7 полных дней до вчера
#   прошлый = предыдущие 7 полных дней
#
# Источники:
#   1) POST /api/analytics/v3/sales-funnel/products
#      - текущая и прошлая воронка
#      - средние заказы в день
#      - конверсии
#
#   2) STOCK_HISTORY_DAILY_CSV
#      - 14 дней истории остатков
#      - stockType = "" => WB + FBS
#      - день в наличии = суммарный остаток по всем строкам NmID > 0
#
# ВАЖНО:
#   Создаётся 1 CSV-отчёт истории остатков за запуск.
# ================================================================

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ================================================================
# 1. НАСТРОЙКИ
# ================================================================

$FunnelPageSize = 1000
$ApiPauseSeconds = 22
$ReportPollSeconds = 22
$ReportMaxChecks = 20

# ================================================================
# 2. ПУТИ
# ================================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$TokenPath = Join-Path $ScriptDir "wb_analytics_token.txt"

$OutputPath = Join-Path $ScriptDir "WB_Воронка_7дней_TEST.csv"
$TempPath = Join-Path $ScriptDir "WB_Воронка_7дней_TEST.tmp.csv"
$LogPath = Join-Path $ScriptDir "WB_Воронка_7дней_TEST.log"

$StockZipPath = Join-Path $ScriptDir "WB_Остатки_14д_TEST.zip"
$StockExtractDir = Join-Path $ScriptDir "WB_Остатки_14д_TEST_EXTRACTED"
$StockCsvCopyPath = Join-Path $ScriptDir "WB_Остатки_14д_TEST.csv"

# ================================================================
# 3. ПЕРИОДЫ
# ================================================================

$DateTo = (Get-Date).Date.AddDays(-1)
$DateFrom = $DateTo.AddDays(-6)

$PastTo = $DateFrom.AddDays(-1)
$PastFrom = $PastTo.AddDays(-6)

$BeginDate = $DateFrom.ToString("yyyy-MM-dd")
$EndDate = $DateTo.ToString("yyyy-MM-dd")

$PastBeginDate = $PastFrom.ToString("yyyy-MM-dd")
$PastEndDate = $PastTo.ToString("yyyy-MM-dd")

$StockBeginDate = $PastBeginDate
$StockEndDate = $EndDate

$CurrentDateColumns = @()
$PastDateColumns = @()

for ($i = 0; $i -lt 7; $i++) {
    $PastDateColumns += $PastFrom.AddDays($i).ToString("dd.MM.yyyy")
    $CurrentDateColumns += $DateFrom.AddDays($i).ToString("dd.MM.yyyy")
}

# ================================================================
# 4. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ================================================================

function Write-Log {
    param([string]$Message)

    $line = "{0}  {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
}

function Get-Field {
    param(
        $Object,
        [string]$Name,
        $Default = $null
    )

    if ($null -eq $Object) {
        return $Default
    }

    $p = $Object.PSObject.Properties[$Name]

    if ($null -eq $p) {
        return $Default
    }

    return $p.Value
}

function Convert-ToDoubleOrZero {
    param($Value)

    if ($null -eq $Value) {
        return [double]0
    }

    $text = ([string]$Value).Trim()

    if ([string]::IsNullOrWhiteSpace($text)) {
        return [double]0
    }

    # В выгрузке WB пустой остаток может приходить как дефис/тире.
    if ($text -eq "-" -or $text -eq "—" -or $text -eq "–") {
        return [double]0
    }

    # Убираем обычные и неразрывные пробелы-разделители.
    $text = $text.Replace(([string][char]0x00A0), "").Replace(" ", "")

    $number = [double]0
    $styles = [System.Globalization.NumberStyles]::Any

    $cultures = @(
        [System.Globalization.CultureInfo]::CurrentCulture,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.CultureInfo]::GetCultureInfo("ru-RU")
    )

    foreach ($culture in $cultures) {
        if ([double]::TryParse($text, $styles, $culture, [ref]$number)) {
            return $number
        }
    }

    # Остатки должны быть числом. Неизвестное текстовое значение
    # безопасно трактуем как 0, чтобы одна ячейка не роняла весь отчёт.
    return [double]0
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
    $request.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate

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
        $status = $null

        if ($null -ne $_.Exception.Response) {
            try {
                $errorResponse = [System.Net.HttpWebResponse]$_.Exception.Response
                $status = [int]$errorResponse.StatusCode
                $errorStream = $errorResponse.GetResponseStream()
                $errorMemory = New-Object System.IO.MemoryStream

                try {
                    $errorStream.CopyTo($errorMemory)
                    $errorBody = [System.Text.Encoding]::UTF8.GetString($errorMemory.ToArray())
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
            throw "Ошибка WB API HTTP ${status}: $($_.Exception.Message).`nОтвет WB: $errorBody"
        }

        throw
    }
    finally {
        if ($null -ne $response) {
            $response.Dispose()
        }
    }
}

function Invoke-WBGetUtf8 {
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [Parameter(Mandatory=$true)][string]$AuthorizationHeader,
        [int]$TimeoutMilliseconds = 180000
    )

    $request = [System.Net.HttpWebRequest]::Create($Uri)
    $request.Method = "GET"
    $request.Timeout = $TimeoutMilliseconds
    $request.ReadWriteTimeout = $TimeoutMilliseconds
    $request.Accept = "application/json"
    $request.Headers["Authorization"] = $AuthorizationHeader
    $request.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate

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
    finally {
        if ($null -ne $response) {
            $response.Dispose()
        }
    }
}

function Download-WBFile {
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [Parameter(Mandatory=$true)][string]$AuthorizationHeader,
        [Parameter(Mandatory=$true)][string]$Path,
        [int]$TimeoutMilliseconds = 300000
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
            if ($null -ne $stream) {
                $stream.Dispose()
            }

            $file.Dispose()
        }
    }
    finally {
        if ($null -ne $response) {
            $response.Dispose()
        }
    }
}

# ================================================================
# 5. ТОКЕН
# ================================================================

if (-not (Test-Path -LiteralPath $TokenPath)) {
    throw "Не найден wb_analytics_token.txt рядом со скриптом."
}

$Token = (Get-Content -LiteralPath $TokenPath -Raw).Trim()

if ([string]::IsNullOrWhiteSpace($Token) -or ($Token -match "ВСТАВ|TOKEN|ТОКЕН")) {
    throw "wb_analytics_token.txt пустой или содержит шаблон."
}

$Auth = "Bearer $Token"

Write-Log "============================================================"
Write-Log "Старт выгрузки воронки продаж WB."
Write-Log "Текущий период: $BeginDate - $EndDate."
Write-Log "Прошлый период: $PastBeginDate - $PastEndDate."

# ================================================================
# 6. ВОРОНКА ПРОДАЖ — ВСЕ ТОВАРЫ
# ================================================================

$FunnelUri = "https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products"

$Products = @()
$offset = 0
$pageNumber = 0

while ($true) {
    $pageNumber++

    if ($pageNumber -gt 1) {
        Write-Log "Пауза $ApiPauseSeconds сек. перед следующей страницей воронки."
        Start-Sleep -Seconds $ApiPauseSeconds
    }

    $bodyObject = [ordered]@{
        selectedPeriod = [ordered]@{
            start = $BeginDate
            end = $EndDate
        }
        pastPeriod = [ordered]@{
            start = $PastBeginDate
            end = $PastEndDate
        }
        nmIds = @()
        brandNames = @()
        subjectIds = @()
        tagIds = @()
        skipDeletedNm = $true
        orderBy = [ordered]@{
            field = "orderSum"
            mode = "desc"
        }
        limit = $FunnelPageSize
        offset = $offset
    }

    $bodyJson = $bodyObject | ConvertTo-Json -Depth 10

    Write-Log "Воронка: страница $pageNumber, offset=$offset."

    $response = Invoke-WBPostUtf8 `
        -Uri $FunnelUri `
        -JsonBody $bodyJson `
        -AuthorizationHeader $Auth

    $pageProducts = @()

    if ($null -ne $response -and $null -ne $response.data -and $null -ne $response.data.products) {
        $pageProducts = @($response.data.products)
    }

    $Products += $pageProducts

    Write-Log "Воронка: получено на странице $($pageProducts.Count). Всего: $($Products.Count)."

    if ($pageProducts.Count -lt $FunnelPageSize) {
        break
    }

    $offset += $FunnelPageSize
}

if ($Products.Count -eq 0) {
    throw "API воронки не вернул товары."
}

# ================================================================
# 7. ИСТОРИЯ ОСТАТКОВ: ПОВТОРНО ИСПОЛЬЗУЕМ СВЕЖИЙ CSV, ЕСЛИ ОН УЖЕ ЕСТЬ
# ================================================================
#
# Если предыдущий запуск уже успел скачать историю остатков, но упал
# на локальной обработке, повторно НЕ создаём CSV-отчёт WB.
# Кэш используется только если в нём есть обе граничные даты
# текущего 14-дневного окна.
# ================================================================

$UseCachedStockCsv = $false
$stockSourceCsv = $null

if (Test-Path -LiteralPath $StockCsvCopyPath) {
    try {
        $cacheProbe = Import-Csv -LiteralPath $StockCsvCopyPath |
            Select-Object -First 1

        if ($null -ne $cacheProbe) {
            $firstDateColumn = $PastDateColumns[0]
            $lastDateColumn = $CurrentDateColumns[$CurrentDateColumns.Count - 1]

            $hasFirstDate = $null -ne $cacheProbe.PSObject.Properties[$firstDateColumn]
            $hasLastDate = $null -ne $cacheProbe.PSObject.Properties[$lastDateColumn]

            if ($hasFirstDate -and $hasLastDate) {
                $UseCachedStockCsv = $true
                $stockSourceCsv = $StockCsvCopyPath
                Write-Log "Использую уже скачанную историю остатков за нужный 14-дневный период."
            }
        }
    }
    catch {
        Write-Log "Локальный CSV остатков не подходит для повторного использования."
    }
}

# ================================================================
# 8. ЕСЛИ КЭША НЕТ — СОЗДАЁМ ОДИН STOCK_HISTORY_DAILY_CSV
# ================================================================

if (-not $UseCachedStockCsv) {
    Write-Log "Пауза $ApiPauseSeconds сек. перед созданием истории остатков."
    Start-Sleep -Seconds $ApiPauseSeconds

    $ReportId = [Guid]::NewGuid().ToString()

    $stockBodyObject = [ordered]@{
        id = $ReportId
        reportType = "STOCK_HISTORY_DAILY_CSV"
        userReportName = "Tradtex funnel stock history 14d"
        params = [ordered]@{
            nmIds = @()
            subjectIds = @()
            brandNames = @()
            tagIds = @()
            currentPeriod = [ordered]@{
                start = $StockBeginDate
                end = $StockEndDate
            }
            stockType = ""
            skipDeletedNm = $true
        }
    }

    $stockBodyJson = $stockBodyObject | ConvertTo-Json -Depth 10

    Write-Log "Создаю STOCK_HISTORY_DAILY_CSV. Report ID: $ReportId."

    $createResponse = Invoke-WBPostUtf8 `
        -Uri "https://seller-analytics-api.wildberries.ru/api/v2/nm-report/downloads" `
        -JsonBody $stockBodyJson `
        -AuthorizationHeader $Auth

    $statusUri = "https://seller-analytics-api.wildberries.ru/api/v2/nm-report/downloads?filter%5BdownloadIds%5D=$ReportId"
    $reportStatus = ""

    for ($check = 1; $check -le $ReportMaxChecks; $check++) {
        Write-Log "Жду $ReportPollSeconds сек. перед проверкой статуса остатков ($check/$ReportMaxChecks)."
        Start-Sleep -Seconds $ReportPollSeconds

        $statusResponse = Invoke-WBGetUtf8 `
            -Uri $statusUri `
            -AuthorizationHeader $Auth

        $statusItems = @($statusResponse.data)

        $report = $statusItems |
            Where-Object { $_.id -eq $ReportId } |
            Select-Object -First 1

        if ($null -eq $report) {
            Write-Log "Отчёт истории остатков пока не найден в списке."
            continue
        }

        $reportStatus = [string]$report.status
        Write-Log "Статус истории остатков: $reportStatus."

        if ($reportStatus -eq "SUCCESS") {
            break
        }

        if ($reportStatus -eq "FAILED") {
            throw "WB вернул FAILED для истории остатков. Report ID: $ReportId."
        }
    }

    if ($reportStatus -ne "SUCCESS") {
        throw "История остатков не успела подготовиться. Report ID: $ReportId."
    }

    Write-Log "Отчёт готов. Жду $ReportPollSeconds сек. перед скачиванием."
    Start-Sleep -Seconds $ReportPollSeconds

    $downloadUri = "https://seller-analytics-api.wildberries.ru/api/v2/nm-report/downloads/file/$ReportId"

    Download-WBFile `
        -Uri $downloadUri `
        -AuthorizationHeader $Auth `
        -Path $StockZipPath

    if (Test-Path -LiteralPath $StockExtractDir) {
        Remove-Item -LiteralPath $StockExtractDir -Recurse -Force
    }

    New-Item -ItemType Directory -Path $StockExtractDir | Out-Null

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($StockZipPath, $StockExtractDir)

    $stockCsvFiles = @(
        Get-ChildItem -LiteralPath $StockExtractDir -Filter "*.csv" -File -Recurse
    )

    if ($stockCsvFiles.Count -eq 0) {
        throw "В ZIP истории остатков не найден CSV."
    }

    $stockSourceCsv = $stockCsvFiles[0].FullName
    Copy-Item -LiteralPath $stockSourceCsv -Destination $StockCsvCopyPath -Force

    # Дальше читаем именно копию: она останется доступна при повторном запуске.
    $stockSourceCsv = $StockCsvCopyPath
}

# ================================================================
# 9. ЧИТАЕМ CSV ОСТАТКОВ
# ================================================================

$StockRows = @(Import-Csv -LiteralPath $stockSourceCsv)

if ($StockRows.Count -eq 0) {
    throw "CSV истории остатков пуст."
}

Write-Log "Строк истории остатков: $($StockRows.Count)."

# ================================================================
# 10. СУММИРУЕМ WB + FBS ПО NmID И ДНЯМ
# ================================================================
#
# В реальном CSV на один NmID есть отдельные строки:
#   OfficeName = "Свой склад"
#   OfficeName = "Склад WB"
#
# Если у товара несколько размеров/ChrtID, они также суммируются.
#
# День в наличии:
#   сумма всех строк этого NmID за дату > 0
# ================================================================

$StockMap = @{}

foreach ($stockRow in $StockRows) {
    $nmRaw = Get-Field $stockRow "NmID" $null

    if ($null -eq $nmRaw -or [string]::IsNullOrWhiteSpace([string]$nmRaw)) {
        continue
    }

    $nmId = [Int64]$nmRaw
    $nmKey = [string]$nmId

    if (-not $StockMap.ContainsKey($nmKey)) {
        $StockMap[$nmKey] = @{}
    }

    $dayMap = $StockMap[$nmKey]

    foreach ($dateColumn in ($PastDateColumns + $CurrentDateColumns)) {
        $valueRaw = Get-Field $stockRow $dateColumn 0
        $value = Convert-ToDoubleOrZero $valueRaw

        if (-not $dayMap.ContainsKey($dateColumn)) {
            $dayMap[$dateColumn] = [double]0
        }

        $dayMap[$dateColumn] += $value
    }
}

# ================================================================
# 11. ФОРМИРУЕМ ИТОГ: 1 СТРОКА = 1 АРТИКУЛ
# ================================================================

$Rows = @()
$MissingStockHistoryCount = 0

foreach ($item in $Products) {
    $product = Get-Field $item "product" $null
    $statistic = Get-Field $item "statistic" $null

    if ($null -eq $product -or $null -eq $statistic) {
        continue
    }

    $nmId = [Int64](Get-Field $product "nmId" 0)
    $nmKey = [string]$nmId

    $selected = Get-Field $statistic "selected" $null
    $past = Get-Field $statistic "past" $null
    $comparison = Get-Field $statistic "comparison" $null

    $selectedConversions = Get-Field $selected "conversions" $null
    $pastConversions = Get-Field $past "conversions" $null

    $stocks = Get-Field $product "stocks" $null

    $daysCurrent = 0
    $daysPast = 0
    $stockCheck = "OK"

    if ($StockMap.ContainsKey($nmKey)) {
        $dayMap = $StockMap[$nmKey]

        foreach ($dateColumn in $CurrentDateColumns) {
            if ($dayMap.ContainsKey($dateColumn) -and [double]$dayMap[$dateColumn] -gt 0) {
                $daysCurrent++
            }
        }

        foreach ($dateColumn in $PastDateColumns) {
            if ($dayMap.ContainsKey($dateColumn) -and [double]$dayMap[$dateColumn] -gt 0) {
                $daysPast++
            }
        }
    }
    else {
        $stockCheck = "Нет данных в STOCK_HISTORY"
        $MissingStockHistoryCount++
    }

    $ordersCurrent = [double](Get-Field $selected "orderCount" 0)
    $ordersPast = [double](Get-Field $past "orderCount" 0)

    $avgOrdersCurrentAvailable = if ($daysCurrent -gt 0) {
        [Math]::Round($ordersCurrent / $daysCurrent, 1)
    }
    else {
        0
    }

    $avgOrdersPastAvailable = if ($daysPast -gt 0) {
        [Math]::Round($ordersPast / $daysPast, 1)
    }
    else {
        0
    }

    $Rows += [pscustomobject][ordered]@{
        "Артикул WB" = $nmId
        "Артикул продавца" = [string](Get-Field $product "vendorCode" "")
        "Наименование" = [string](Get-Field $product "title" "")
        "Бренд" = [string](Get-Field $product "brandName" "")
        "Предмет" = [string](Get-Field $product "subjectName" "")

        "Текущий период с" = $BeginDate
        "Текущий период по" = $EndDate
        "Прошлый период с" = $PastBeginDate
        "Прошлый период по" = $PastEndDate

        "Переходы текущий" = [Int64](Get-Field $selected "openCount" 0)
        "Переходы прошлый" = [Int64](Get-Field $past "openCount" 0)

        "Корзины текущий" = [Int64](Get-Field $selected "cartCount" 0)
        "Корзины прошлый" = [Int64](Get-Field $past "cartCount" 0)

        "Заказы текущий" = [Int64]$ordersCurrent
        "Заказы прошлый" = [Int64]$ordersPast

        "Средние заказы WB текущий" = [double](Get-Field $selected "avgOrdersCountPerDay" 0)
        "Средние заказы WB прошлый" = [double](Get-Field $past "avgOrdersCountPerDay" 0)
        "Динамика средних заказов %" = [double](Get-Field $comparison "avgOrdersCountPerDayDynamic" 0)

        "Дней в наличии текущий" = $daysCurrent
        "Дней в наличии прошлый" = $daysPast

        # Дополнительный аналитический показатель.
        # Это НЕ поле WB: заказы делим только на дни, когда товар был в наличии.
        "Заказы/день в наличии текущий" = $avgOrdersCurrentAvailable
        "Заказы/день в наличии прошлый" = $avgOrdersPastAvailable
        "Проверка истории наличия" = $stockCheck

        "Сумма заказов текущий" = [double](Get-Field $selected "orderSum" 0)
        "Сумма заказов прошлый" = [double](Get-Field $past "orderSum" 0)

        "Выкупы текущий" = [Int64](Get-Field $selected "buyoutCount" 0)
        "Выкупы прошлый" = [Int64](Get-Field $past "buyoutCount" 0)
        "Сумма выкупов текущий" = [double](Get-Field $selected "buyoutSum" 0)
        "Сумма выкупов прошлый" = [double](Get-Field $past "buyoutSum" 0)

        "Отмены текущий" = [Int64](Get-Field $selected "cancelCount" 0)
        "Отмены прошлый" = [Int64](Get-Field $past "cancelCount" 0)
        "Сумма отмен текущий" = [double](Get-Field $selected "cancelSum" 0)
        "Сумма отмен прошлый" = [double](Get-Field $past "cancelSum" 0)

        "Средняя цена текущий" = [double](Get-Field $selected "avgPrice" 0)
        "Средняя цена прошлый" = [double](Get-Field $past "avgPrice" 0)

        "Конверсия карточка-корзина текущий" = [double](Get-Field $selectedConversions "addToCartPercent" 0)
        "Конверсия карточка-корзина прошлый" = [double](Get-Field $pastConversions "addToCartPercent" 0)

        "Конверсия корзина-заказ текущий" = [double](Get-Field $selectedConversions "cartToOrderPercent" 0)
        "Конверсия корзина-заказ прошлый" = [double](Get-Field $pastConversions "cartToOrderPercent" 0)

        "Процент выкупа текущий" = [double](Get-Field $selectedConversions "buyoutPercent" 0)
        "Процент выкупа прошлый" = [double](Get-Field $pastConversions "buyoutPercent" 0)

        "Добавили в избранное текущий" = [Int64](Get-Field $selected "addToWishlist" 0)
        "Добавили в избранное прошлый" = [Int64](Get-Field $past "addToWishlist" 0)

        "Остаток WB сейчас" = [Int64](Get-Field $stocks "wb" 0)
        "Остаток FBS сейчас" = [Int64](Get-Field $stocks "mp" 0)
    }
}

Write-Log "Итоговых строк: $($Rows.Count)."
Write-Log "Без строки в истории остатков: $MissingStockHistoryCount."

# ================================================================
# 12. CSV
# ================================================================

$csvLines = $Rows |
    ConvertTo-Csv -NoTypeInformation -Delimiter ";"

# CSV оставляем стандартным: без служебной строки sep=;.
# Power Query должен читать его явно как ; + Windows-1251.
# Это надёжнее для автоматического импорта и последующего обновления.
$excelEncoding = [System.Text.Encoding]::GetEncoding(1251)

[System.IO.File]::WriteAllLines(
    $TempPath,
    $csvLines,
    $excelEncoding
)

# ================================================================
# 13. БЕЗОПАСНАЯ ЗАМЕНА CSV
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
    $fallback = Join-Path $ScriptDir ("WB_Воронка_7дней_TEST_NEW_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

    if (Test-Path -LiteralPath $TempPath) {
        Move-Item -LiteralPath $TempPath -Destination $fallback -Force
    }

    throw "WB_Воронка_7дней_TEST.csv заблокирован. Новые данные сохранены: $fallback"
}

Write-Log "ГОТОВО."
Write-Log "CSV: $OutputPath."
Write-Log "История остатков: $StockCsvCopyPath."

Write-Host ""
Write-Host "ВЫГРУЗКА ВОРОНКИ ЗАВЕРШЕНА" -ForegroundColor Green
Write-Host "Текущий период: $BeginDate - $EndDate"
Write-Host "Прошлый период: $PastBeginDate - $PastEndDate"
Write-Host "Товаров: $($Rows.Count)"
Write-Host "Файл: $OutputPath"
