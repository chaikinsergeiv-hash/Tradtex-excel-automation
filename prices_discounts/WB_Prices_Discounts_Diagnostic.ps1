# ================================================================
# WB — ДИАГНОСТИКА ЦЕН И СКИДОК
#
# Безопасный тест:
#   только чтение данных
#   ровно 1 GET-запрос
#   limit=10, offset=0
#
# Endpoint:
#   GET /api/v2/list/goods/filter
#
# Результаты:
#   01_WB_PRICES_RAW.json
#   01_WB_PRICES_SAMPLE.csv
#
# Нужен токен категории "Цены и скидки":
#   wb_prices_token.txt
# ================================================================

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TokenPath = Join-Path $ScriptDir "wb_prices_token.txt"
$RawPath = Join-Path $ScriptDir "01_WB_PRICES_RAW.json"
$CsvPath = Join-Path $ScriptDir "01_WB_PRICES_SAMPLE.csv"

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

$Limit = 10
$Offset = 0
$Uri = "https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter?limit=$Limit&offset=$Offset"

Write-Host "WB — диагностика цен и скидок" -ForegroundColor Cyan
Write-Host "Будет выполнен ровно 1 GET-запрос."
Write-Host "Товаров в тесте: максимум $Limit."
Write-Host "Изменений цен/скидок не будет."
Write-Host ""

$Client = New-Object System.Net.WebClient
$Client.Encoding = [System.Text.Encoding]::UTF8
$Client.Headers.Add("Authorization", $Token)
$Client.Headers.Add("Accept", "application/json")

try {
    $ResponseText = $Client.DownloadString($Uri)
}
finally {
    $Client.Dispose()
}

if ([string]::IsNullOrWhiteSpace($ResponseText)) {
    throw "WB API вернул пустой ответ."
}

$Response = $ResponseText | ConvertFrom-Json

$Utf8Bom = New-Object System.Text.UTF8Encoding($true)
$RawJson = $Response | ConvertTo-Json -Depth 30
[System.IO.File]::WriteAllText($RawPath, $RawJson, $Utf8Bom)

$Items = @()

if (
    $null -ne $Response -and
    $null -ne $Response.data -and
    $null -ne $Response.data.listGoods
) {
    $Items = @($Response.data.listGoods)
}

$Rows = @()

foreach ($Item in $Items) {
    $Wholesale = @()

    if ($null -ne $Item.wholesaleDiscountThreshold) {
        $Wholesale = @($Item.wholesaleDiscountThreshold)
    }

    $WholesaleText = ""

    if ($Wholesale.Count -gt 0) {
        $parts = @()

        foreach ($Level in ($Wholesale | Sort-Object level)) {
            $parts += (
                "уровень {0}: от {1} шт = {2}%" -f
                $Level.level,
                $Level.minQuantity,
                $Level.wholesaleDiscount
            )
        }

        $WholesaleText = $parts -join " | "
    }

    $Sizes = @()

    if ($null -ne $Item.sizes) {
        $Sizes = @($Item.sizes)
    }

    $SizeText = ""

    if ($Sizes.Count -gt 0) {
        $sizeParts = @()

        foreach ($Size in $Sizes) {
            $sizeParts += (
                "{0}: цена {1}; со скидкой {2}; клуб {3}" -f
                $Size.techSizeName,
                $Size.price,
                $Size.discountedPrice,
                $Size.clubDiscountedPrice
            )
        }

        $SizeText = $sizeParts -join " | "
    }

    $Rows += [pscustomobject][ordered]@{
        "Артикул WB" = $Item.nmID
        "Артикул продавца" = $Item.vendorCode
        "Валюта" = $Item.currencyIsoCode4217
        "Скидка, %" = $Item.discount
        "Скидка WB Клуба, %" = $Item.clubDiscount
        "Цена по размерам разрешена" = $Item.editableSizePrice
        "Низкая оборачиваемость" = $Item.isBadTurnover
        "B2B скидка есть" = $(if ($Wholesale.Count -gt 0) { "Да" } else { "Нет" })
        "B2B уровней" = $Wholesale.Count
        "B2B условия" = $WholesaleText
        "Размеров" = $Sizes.Count
        "Цены по размерам" = $SizeText
    }
}

if ($Rows.Count -gt 0) {
    $CsvLines = $Rows | ConvertTo-Csv -NoTypeInformation -Delimiter ";"
    $CsvEncoding = [System.Text.Encoding]::GetEncoding(1251)
    [System.IO.File]::WriteAllLines($CsvPath, $CsvLines, $CsvEncoding)
}

Write-Host ""
Write-Host "ГОТОВО" -ForegroundColor Green
Write-Host "Получено товаров: $($Items.Count)"
Write-Host "RAW: $RawPath"
Write-Host "CSV: $CsvPath"

$B2BCount = @($Rows | Where-Object { $_."B2B скидка есть" -eq "Да" }).Count
Write-Host "Товаров с B2B скидкой в тестовой выборке: $B2BCount"
