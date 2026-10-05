# ================================================================
# WB — ЦЕНЫ И СКИДКИ — TEST
#
# Только чтение данных. Ничего в кабинете WB не изменяет.
#
# Итог:
#   1 строка = 1 артикул WB (nmID)
#
# Включает:
#   - артикул продавца;
#   - базовые цены по размерам;
#   - цены после обычной скидки;
#   - цены WB Клуба;
#   - обычную скидку;
#   - скидку WB Клуба;
#   - B2B / оптовые скидки для юрлиц;
#   - динамические уровни B2B.
# ================================================================

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ================================================================
# 1. НАСТРОЙКИ
# ================================================================

$PageSize = 1000
$PagePauseSeconds = 2
$MaxRetryCount = 4

# ================================================================
# 2. ПУТИ
# ================================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TokenPath = Join-Path $ScriptDir "wb_prices_token.txt"

$OutputPath = Join-Path $ScriptDir "WB_Цены_Скидки_TEST.csv"
$TempPath = Join-Path $ScriptDir "WB_Цены_Скидки_TEST.tmp.csv"
$LogPath = Join-Path $ScriptDir "WB_Цены_Скидки_TEST.log"

# ================================================================
# 3. ФУНКЦИИ
# ================================================================

function Write-Log {
    param([string]$Message)

    $line = "{0}  {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
}

function Convert-ToDecimalOrBlank {
    param($Value)

    if ($null -eq $Value) {
        return ""
    }

    try {
        return [decimal]$Value
    }
    catch {
        return ""
    }
}

function Get-MinValue {
    param([object[]]$Values)

    $clean = @(
        $Values |
            Where-Object { $null -ne $_ -and ([string]$_).Trim() -ne "" } |
            ForEach-Object { [decimal]$_ }
    )

    if ($clean.Count -eq 0) {
        return ""
    }

    return ($clean | Measure-Object -Minimum).Minimum
}

function Get-MaxValue {
    param([object[]]$Values)

    $clean = @(
        $Values |
            Where-Object { $null -ne $_ -and ([string]$_).Trim() -ne "" } |
            ForEach-Object { [decimal]$_ }
    )

    if ($clean.Count -eq 0) {
        return ""
    }

    return ($clean | Measure-Object -Maximum).Maximum
}

function Invoke-WBPricesGet {
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [Parameter(Mandatory=$true)][string]$TokenValue
    )

    for ($attempt = 1; $attempt -le $MaxRetryCount; $attempt++) {
        $client = New-Object System.Net.WebClient
        $client.Encoding = [System.Text.Encoding]::UTF8
        $client.Headers.Add("Authorization", $TokenValue)
        $client.Headers.Add("Accept", "application/json")

        try {
            $responseText = $client.DownloadString($Uri)

            if ([string]::IsNullOrWhiteSpace($responseText)) {
                throw "WB API вернул пустой ответ."
            }

            return ($responseText | ConvertFrom-Json)
        }
        catch [System.Net.WebException] {
            $statusCode = $null

            if ($null -ne $_.Exception.Response) {
                try {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                }
                catch {
                }
            }

            if ($statusCode -eq 429 -and $attempt -lt $MaxRetryCount) {
                $waitSeconds = 10 * $attempt
                Write-Log "Получен HTTP 429. Повтор через $waitSeconds сек."
                Start-Sleep -Seconds $waitSeconds
                continue
            }

            throw
        }
        finally {
            $client.Dispose()
        }
    }

    throw "Не удалось получить данные WB после $MaxRetryCount попыток."
}

# ================================================================
# 4. ТОКЕН
# ================================================================

if (-not (Test-Path -LiteralPath $TokenPath)) {
    throw "Не найден wb_prices_token.txt рядом со скриптом."
}

$Token = (Get-Content -LiteralPath $TokenPath -Raw).Trim()

if (
    [string]::IsNullOrWhiteSpace($Token) -or
    ($Token -match "ВСТАВ|TOKEN|ТОКЕН")
) {
    throw "wb_prices_token.txt пустой или содержит шаблон."
}

# ================================================================
# 5. ПОЛУЧАЕМ ВСЕ ТОВАРЫ
# ================================================================

$AllItems = @()
$offset = 0
$page = 0

Write-Log "============================================================"
Write-Log "Старт выгрузки цен и скидок WB. Только чтение."

while ($true) {
    $page++

    if ($page -gt 1) {
        Write-Log "Пауза $PagePauseSeconds сек. перед страницей $page."
        Start-Sleep -Seconds $PagePauseSeconds
    }

    $Uri = (
        "https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter" +
        "?limit=$PageSize&offset=$offset"
    )

    Write-Log "API: страница $page, offset=$offset, limit=$PageSize."

    $Response = Invoke-WBPricesGet -Uri $Uri -TokenValue $Token

    if ($Response.error -eq $true) {
        throw "WB API вернул ошибку: $($Response.errorText)"
    }

    $PageItems = @()

    if (
        $null -ne $Response.data -and
        $null -ne $Response.data.listGoods
    ) {
        $PageItems = @($Response.data.listGoods)
    }

    $AllItems += $PageItems

    Write-Log "Получено на странице: $($PageItems.Count). Всего: $($AllItems.Count)."

    if ($PageItems.Count -lt $PageSize) {
        break
    }

    $offset += $PageSize
}

if ($AllItems.Count -eq 0) {
    throw "WB API не вернул товары."
}

# ================================================================
# 6. ОПРЕДЕЛЯЕМ МАКСИМАЛЬНЫЙ B2B-УРОВЕНЬ
# ================================================================

$MaxB2BLevel = 0

foreach ($Item in $AllItems) {
    if ($null -eq $Item.wholesaleDiscountThreshold) {
        continue
    }

    foreach ($Level in @($Item.wholesaleDiscountThreshold)) {
        $levelNumber = 0

        try {
            $levelNumber = [int]$Level.level
        }
        catch {
            $levelNumber = 0
        }

        if ($levelNumber -gt $MaxB2BLevel) {
            $MaxB2BLevel = $levelNumber
        }
    }
}

Write-Log "Максимальный B2B уровень в выгрузке: $MaxB2BLevel."

# ================================================================
# 7. ФОРМИРУЕМ 1 СТРОКУ НА 1 NMID
# ================================================================

$Rows = @()

foreach ($Item in ($AllItems | Sort-Object nmID)) {
    $Sizes = @()

    if ($null -ne $Item.sizes) {
        $Sizes = @($Item.sizes)
    }

    $BasePrices = @()
    $DiscountedPrices = @()
    $ClubPrices = @()
    $SizeTextParts = @()

    foreach ($Size in $Sizes) {
        $BasePrices += Convert-ToDecimalOrBlank $Size.price
        $DiscountedPrices += Convert-ToDecimalOrBlank $Size.discountedPrice
        $ClubPrices += Convert-ToDecimalOrBlank $Size.clubDiscountedPrice

        $techSize = [string]$Size.techSizeName

        if ([string]::IsNullOrWhiteSpace($techSize)) {
            $techSize = [string]$Size.sizeID
        }

        $SizeTextParts += (
            "{0}: {1} -> {2}; клуб {3}" -f
            $techSize,
            $Size.price,
            $Size.discountedPrice,
            $Size.clubDiscountedPrice
        )
    }

    $BaseMin = Get-MinValue $BasePrices
    $BaseMax = Get-MaxValue $BasePrices
    $DiscountedMin = Get-MinValue $DiscountedPrices
    $DiscountedMax = Get-MaxValue $DiscountedPrices
    $ClubMin = Get-MinValue $ClubPrices
    $ClubMax = Get-MaxValue $ClubPrices

    $PricesDiffer = "Нет"

    if (
        ($BaseMin -ne "" -and $BaseMax -ne "" -and $BaseMin -ne $BaseMax) -or
        ($DiscountedMin -ne "" -and $DiscountedMax -ne "" -and $DiscountedMin -ne $DiscountedMax) -or
        ($ClubMin -ne "" -and $ClubMax -ne "" -and $ClubMin -ne $ClubMax)
    ) {
        $PricesDiffer = "Да"
    }

    $Wholesale = @()

    if ($null -ne $Item.wholesaleDiscountThreshold) {
        $Wholesale = @(
            $Item.wholesaleDiscountThreshold |
                Sort-Object level
        )
    }

    $WholesaleByLevel = @{}
    $WholesaleTextParts = @()

    foreach ($Level in $Wholesale) {
        $levelNumber = [int]$Level.level
        $WholesaleByLevel[[string]$levelNumber] = $Level

        $WholesaleTextParts += (
            "уровень {0}: от {1} шт = {2}%" -f
            $Level.level,
            $Level.minQuantity,
            $Level.wholesaleDiscount
        )
    }

    $row = [ordered]@{
        "Артикул WB" = $Item.nmID
        "Артикул продавца" = $Item.vendorCode
        "Валюта" = $Item.currencyIsoCode4217
        "Скидка, %" = $Item.discount
        "Скидка WB Клуба, %" = $Item.clubDiscount
        "Цена базовая от" = $BaseMin
        "Цена базовая до" = $BaseMax
        "Цена после скидки от" = $DiscountedMin
        "Цена после скидки до" = $DiscountedMax
        "Цена WB Клуб от" = $ClubMin
        "Цена WB Клуб до" = $ClubMax
        "Размеров" = $Sizes.Count
        "Цены по размерам различаются" = $PricesDiffer
        "Цена по размерам разрешена" = $Item.editableSizePrice
        "B2B скидка есть" = $(if ($Wholesale.Count -gt 0) { "Да" } else { "Нет" })
        "B2B уровней" = $Wholesale.Count
    }

    for ($levelNumber = 1; $levelNumber -le $MaxB2BLevel; $levelNumber++) {
        $minQty = ""
        $discount = ""
        $key = [string]$levelNumber

        if ($WholesaleByLevel.ContainsKey($key)) {
            $minQty = $WholesaleByLevel[$key].minQuantity
            $discount = $WholesaleByLevel[$key].wholesaleDiscount
        }

        $row["B2B уровень $levelNumber от, шт"] = $minQty
        $row["B2B уровень $levelNumber скидка, %"] = $discount
    }

    $row["B2B условия"] = ($WholesaleTextParts -join " | ")
    $row["Цены по размерам"] = ($SizeTextParts -join " | ")

    $Rows += [pscustomobject]$row
}

# ================================================================
# 8. CSV
# ================================================================

$CsvLines = $Rows |
    ConvertTo-Csv -NoTypeInformation -Delimiter ";"

# Такой же формат, как в утверждённых блоках:
# стандартный CSV без sep=;, Windows-1251.
$CsvEncoding = [System.Text.Encoding]::GetEncoding(1251)

[System.IO.File]::WriteAllLines(
    $TempPath,
    $CsvLines,
    $CsvEncoding
)

# ================================================================
# 9. БЕЗОПАСНАЯ ЗАМЕНА ФАЙЛА
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
        "WB_Цены_Скидки_TEST_NEW_{0}.csv" -f
        (Get-Date -Format "yyyyMMdd_HHmmss")
    )

    if (Test-Path -LiteralPath $TempPath) {
        Move-Item -LiteralPath $TempPath -Destination $fallback -Force
    }

    throw "Основной CSV заблокирован. Новые данные сохранены: $fallback"
}

$B2BCount = @(
    $Rows |
        Where-Object { $_."B2B скидка есть" -eq "Да" }
).Count

Write-Log "ГОТОВО."
Write-Log "Товаров: $($Rows.Count)."
Write-Log "Товаров с B2B скидкой: $B2BCount."
Write-Log "CSV: $OutputPath."

Write-Host ""
Write-Host "ВЫГРУЗКА ЦЕН И СКИДОК WB ЗАВЕРШЕНА" -ForegroundColor Green
Write-Host "Товаров: $($Rows.Count)"
Write-Host "С B2B скидкой: $B2BCount"
Write-Host "Максимум B2B уровней: $MaxB2BLevel"
Write-Host "Файл: $OutputPath"
