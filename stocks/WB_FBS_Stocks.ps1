# WB FBS stocks -> CSV
# Requires a WB API token with categories: Content + Marketplace.
# Compatible with Windows PowerShell 5.1 and PowerShell 7+.

$ErrorActionPreference = "Stop"

# Force TLS 1.2 for older Windows PowerShell without changing the registry.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TokenPath = Join-Path $ScriptDir "wb_token.txt"
$OutputPath = Join-Path $ScriptDir "WB_FBS_Остатки.csv"
$TempPath = Join-Path $ScriptDir "WB_FBS_Остатки.tmp.csv"
$LogPath = Join-Path $ScriptDir "WB_FBS_Остатки.log"

function Write-Log {
    param([string]$Message)
    $line = "{0}  {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
}

function Invoke-WBRequest {
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [Parameter(Mandatory=$true)][ValidateSet("GET","POST")][string]$Method,
        $Body = $null,
        [int]$MaxAttempts = 5
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $response = $null

        try {
            # Используем HttpWebRequest и читаем ответ как RAW bytes.
            # Это важно для Windows PowerShell 5.1:
            # Invoke-RestMethod иногда неверно определяет кодировку ответа WB,
            # из-за чего русские строки превращаются в "Ð.../Ñ...".
            $request = [System.Net.HttpWebRequest]::Create($Uri)
            $request.Method = $Method
            $request.Timeout = 120000
            $request.ReadWriteTimeout = 120000
            $request.Accept = "application/json"
            $request.Headers["Authorization"] = $script:Headers["Authorization"]
            $request.AutomaticDecompression = `
                [System.Net.DecompressionMethods]::GZip -bor `
                [System.Net.DecompressionMethods]::Deflate

            if ($Method -eq "POST") {
                $json = $Body | ConvertTo-Json -Depth 12 -Compress
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)

                $request.ContentType = "application/json; charset=utf-8"
                $request.ContentLength = $bytes.Length

                $requestStream = $request.GetRequestStream()
                try {
                    $requestStream.Write($bytes, 0, $bytes.Length)
                }
                finally {
                    $requestStream.Dispose()
                }
            }

            $response = [System.Net.HttpWebResponse]$request.GetResponse()

            $memory = New-Object System.IO.MemoryStream
            $stream = $response.GetResponseStream()

            try {
                $stream.CopyTo($memory)
                $rawBytes = $memory.ToArray()
            }
            finally {
                if ($null -ne $stream) { $stream.Dispose() }
                $memory.Dispose()
            }

            # WB API отдаёт JSON в UTF-8. Декодируем принудительно,
            # не доверяя старому движку Windows PowerShell 5.1.
            $jsonText = [System.Text.Encoding]::UTF8.GetString($rawBytes)

            if ([string]::IsNullOrWhiteSpace($jsonText)) {
                return $null
            }

            return ($jsonText | ConvertFrom-Json)
        }
        catch [System.Net.WebException] {
            $status = $null
            $errorBody = ""

            try {
                if ($null -ne $_.Exception.Response) {
                    $webResponse = [System.Net.HttpWebResponse]$_.Exception.Response
                    $status = [int]$webResponse.StatusCode

                    try {
                        $errorStream = $webResponse.GetResponseStream()
                        $errorMemory = New-Object System.IO.MemoryStream
                        try {
                            $errorStream.CopyTo($errorMemory)
                            $errorBytes = $errorMemory.ToArray()
                            $errorBody = [System.Text.Encoding]::UTF8.GetString($errorBytes)
                        }
                        finally {
                            if ($null -ne $errorStream) { $errorStream.Dispose() }
                            $errorMemory.Dispose()
                        }
                    } catch {}
                }
            } catch {}

            if ($status -eq 429 -and $attempt -lt $MaxAttempts) {
                $wait = [Math]::Min(30, 3 * $attempt)
                Write-Log "WB API вернул 429 (слишком много запросов). Повтор через $wait сек."
                Start-Sleep -Seconds $wait
                continue
            }

            $hint = ""
            switch ($status) {
                401 { $hint = " Проверьте корректность токена." }
                403 { $hint = " Проверьте, что токен имеет категории Контент и Маркетплейс." }
                404 { $hint = " Проверьте адрес API и ID склада." }
                429 { $hint = " Превышен лимит запросов WB API." }
            }

            $statusText = if ($null -eq $status) { "без HTTP-кода" } else { "HTTP $status" }

            if (-not [string]::IsNullOrWhiteSpace($errorBody)) {
                throw "Ошибка WB API ($statusText): $($_.Exception.Message).$hint`nОтвет WB: $errorBody`nURL: $Uri"
            } else {
                throw "Ошибка WB API ($statusText): $($_.Exception.Message).$hint`nURL: $Uri"
            }
        }
        catch {
            throw
        }
        finally {
            if ($null -ne $response) {
                $response.Dispose()
            }
        }
    }
}

# ----------------------------
# 1. Read token
# ----------------------------
if (-not (Test-Path -LiteralPath $TokenPath)) {
    throw "Не найден файл токена: $TokenPath`nСоздайте рядом со скриптом файл wb_token.txt и вставьте в него токен WB одной строкой."
}

$Token = (Get-Content -LiteralPath $TokenPath -Raw).Trim()

if ([string]::IsNullOrWhiteSpace($Token) -or $Token -match "ВСТАВ|TOKEN|ТОКЕН") {
    throw "Файл wb_token.txt пустой или содержит шаблон. Вставьте в него действующий WB API токен одной строкой."
}

$script:Headers = @{
    Authorization = "Bearer $Token"
}

Write-Log "Старт обновления FBS-остатков."

# ----------------------------
# 2. Get all active product cards (Content API)
# ----------------------------
$contentUrl = "https://content-api.wildberries.ru/content/v2/get/cards/list"
$pageSize = 100

$productsByChrt = @{}   # chrtID -> product
$productsByNm = @{}     # nmID -> product

$cursorUpdatedAt = $null
$cursorNmID = $null
$page = 0
$previousCursor = ""

while ($true) {
    $page++

    $cursor = @{
        limit = $pageSize
    }

    if ($null -ne $cursorUpdatedAt -and $null -ne $cursorNmID) {
        $cursor.updatedAt = $cursorUpdatedAt
        $cursor.nmID = [Int64]$cursorNmID
    }

    $body = @{
        settings = @{
            sort = @{
                ascending = $true
            }
            filter = @{
                withPhoto = -1
            }
            cursor = $cursor
        }
    }

    $response = Invoke-WBRequest -Uri $contentUrl -Method POST -Body $body
    $cards = @($response.cards)

    foreach ($card in $cards) {
        if ($null -eq $card.nmID) { continue }

        $nmID = [Int64]$card.nmID
        $product = [pscustomobject]@{
            NmID       = $nmID
            VendorCode = [string]$card.vendorCode
            Title      = [string]$card.title
        }

        $productsByNm[[string]$nmID] = $product

        foreach ($size in @($card.sizes)) {
            if ($null -eq $size -or $null -eq $size.chrtID) { continue }
            $productsByChrt[[string]([Int64]$size.chrtID)] = $product
        }
    }

    $count = $cards.Count
    Write-Log "Карточки: страница $page, получено $count, всего товаров $($productsByNm.Count), размеров $($productsByChrt.Count)."

    if ($count -lt $pageSize) { break }
    if ($null -eq $response.cursor) { break }
    if ($null -eq $response.cursor.updatedAt -or $null -eq $response.cursor.nmID) { break }

    $cursorUpdatedAt = [string]$response.cursor.updatedAt
    $cursorNmID = [Int64]$response.cursor.nmID

    $cursorSignature = "$cursorUpdatedAt|$cursorNmID"
    if ($cursorSignature -eq $previousCursor) {
        throw "WB вернул тот же cursor повторно. Пагинация остановлена, чтобы избежать бесконечного цикла."
    }
    $previousCursor = $cursorSignature

    # Content API limit: 100 req/min, interval about 600 ms.
    Start-Sleep -Milliseconds 650
}

if ($productsByChrt.Count -eq 0) {
    throw "Не получено ни одного chrtID из Content API."
}

# ----------------------------
# 3. Get seller warehouses (Marketplace API)
# ----------------------------
$warehouseUrl = "https://marketplace-api.wildberries.ru/api/v3/warehouses"
$allWarehouses = @(Invoke-WBRequest -Uri $warehouseUrl -Method GET)

$fbsWarehouses = @(
    foreach ($warehouse in $allWarehouses) {
        # Сначала вычисляем два отдельных логических условия.
        # Так мы избегаем ошибки преобразования Int32 -> Boolean
        # в Windows PowerShell 5.1.
        $deliveryTypeIsFbs = ($warehouse.deliveryType -eq 1)
        $warehouseIsActive = ($warehouse.isDeleting -ne $true)

        if ($deliveryTypeIsFbs -and $warehouseIsActive) {
            $warehouse
        }
    }
)

if ($fbsWarehouses.Count -eq 0) {
    throw "Не найдено активных FBS-складов (deliveryType = 1)."
}

Write-Log "Найдено активных FBS-складов: $($fbsWarehouses.Count)."

# ----------------------------
# 4. Prepare chrtID list and aggregate stocks by nmID + warehouse
# ----------------------------
$chrtIds = @(
    $productsByChrt.Keys |
        ForEach-Object { [Int64]$_ } |
        Sort-Object
)

$totals = @{}   # "warehouseId|nmID" -> total amount

foreach ($warehouse in $fbsWarehouses) {
    $warehouseId = [Int64]$warehouse.id
    $warehouseName = [string]$warehouse.name

    Write-Log "Склад '$warehouseName' (ID $warehouseId): начинаю загрузку остатков."

    for ($i = 0; $i -lt $chrtIds.Count; $i += 1000) {
        $end = [Math]::Min($i + 999, $chrtIds.Count - 1)
        $batch = @($chrtIds[$i..$end])

        $stockUrl = "https://marketplace-api.wildberries.ru/api/v3/stocks/$warehouseId"
        $stockBody = @{
            chrtIds = $batch
        }

        $stockResponse = Invoke-WBRequest -Uri $stockUrl -Method POST -Body $stockBody

        foreach ($stock in @($stockResponse.stocks)) {
            if ($null -eq $stock.chrtId) { continue }

            $chrtKey = [string]([Int64]$stock.chrtId)
            if (-not $productsByChrt.ContainsKey($chrtKey)) { continue }

            $product = $productsByChrt[$chrtKey]
            $aggKey = "$warehouseId|$($product.NmID)"
            $amount = if ($null -eq $stock.amount) { 0 } else { [Int64]$stock.amount }

            if ($totals.ContainsKey($aggKey)) {
                $totals[$aggKey] = [Int64]$totals[$aggKey] + $amount
            } else {
                $totals[$aggKey] = $amount
            }
        }

        # Marketplace stock methods: up to 300 req/min, interval about 200 ms.
        Start-Sleep -Milliseconds 250
    }

    Write-Log "Склад '$warehouseName': остатки загружены."
}

# ----------------------------
# 5. Build complete matrix Product x FBS warehouse
#    Missing API rows are written as zero.
# ----------------------------
$rows = New-Object System.Collections.Generic.List[object]

$products = @(
    $productsByNm.Values |
        Sort-Object NmID
)

foreach ($warehouse in ($fbsWarehouses | Sort-Object name)) {
    $warehouseId = [Int64]$warehouse.id
    $warehouseName = [string]$warehouse.name

    foreach ($product in $products) {
        $aggKey = "$warehouseId|$($product.NmID)"
        $amount = if ($totals.ContainsKey($aggKey)) { [Int64]$totals[$aggKey] } else { 0 }

        $rows.Add(
            [pscustomobject][ordered]@{
                "Артикул WB"        = $product.NmID
                "Артикул продавца"  = $product.VendorCode
                "Наименование"      = $product.Title
                "ID склада"         = $warehouseId
                "Склад"             = $warehouseName
                "Остаток"           = $amount
            }
        )
    }
}

# ----------------------------
# 6. Export CSV: semicolon + UTF-8 BOM for Russian Excel
# ----------------------------
$csvLines = $rows | ConvertTo-Csv -NoTypeInformation -Delimiter ";"
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllLines($TempPath, $csvLines, $utf8Bom)

# Пытаемся заменить основной CSV.
# Если файл открыт в Excel или временно занят Power Query/сетевой папкой,
# Windows может запретить замену. В таком случае ждём и повторяем.
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
            Write-Log "Файл WB_FBS_Остатки.csv сейчас занят или открыт. Попытка $attempt из $replaceAttempts. Повтор через $replaceDelaySec сек."
            Start-Sleep -Seconds $replaceDelaySec
        }
        else {
            break
        }
    }
    catch [System.IO.IOException] {
        if ($attempt -lt $replaceAttempts) {
            Write-Log "Файл WB_FBS_Остатки.csv сейчас занят или открыт. Попытка $attempt из $replaceAttempts. Повтор через $replaceDelaySec сек."
            Start-Sleep -Seconds $replaceDelaySec
        }
        else {
            break
        }
    }
}

if (-not $replaceSucceeded) {
    $fallbackPath = Join-Path $ScriptDir ("WB_FBS_Остатки_NEW_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

    if (Test-Path -LiteralPath $TempPath) {
        Move-Item -LiteralPath $TempPath -Destination $fallbackPath -Force
    }

    throw "Основной файл WB_FBS_Остатки.csv не удалось заменить, потому что он открыт или заблокирован другой программой.`nЗакройте WB_FBS_Остатки.csv в Excel и запустите скрипт ещё раз.`nНовые данные сохранены отдельно: $fallbackPath"
}

Write-Log "Готово. Строк: $($rows.Count)."
Write-Log "Файл: $OutputPath"
Write-Host ""
Write-Host "ОБНОВЛЕНИЕ ЗАВЕРШЕНО" -ForegroundColor Green
Write-Host $OutputPath
