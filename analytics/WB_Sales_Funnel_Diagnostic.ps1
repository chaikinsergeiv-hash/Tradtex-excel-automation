# ================================================================
# WB ANALYTICS — ДИАГНОСТИКА ВОРОНКИ ПРОДАЖ
#
# Безопасный тест: ровно 1 API-запрос к sales-funnel/products.
# Периоды:
#   текущий = последние 7 полных дней до вчера
#   прошлый = предыдущие 7 полных дней
#
# Результаты:
#   01_SALES_FUNNEL_RAW.json
#   01_SALES_FUNNEL_SAMPLE.csv
#
# Нужен токен WB API категории "Аналитика" в wb_analytics_token.txt
# Совместимо с Windows PowerShell 5.1.
# ================================================================

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TokenPath = Join-Path $ScriptDir "wb_analytics_token.txt"
$RawPath   = Join-Path $ScriptDir "01_SALES_FUNNEL_RAW.json"
$CsvPath   = Join-Path $ScriptDir "01_SALES_FUNNEL_SAMPLE.csv"

function Get-Field {
    param($Object, [string]$Name, $Default = $null)

    if ($null -eq $Object) { return $Default }

    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p) { return $Default }

    return $p.Value
}

if (-not (Test-Path -LiteralPath $TokenPath)) {
    throw "Не найден wb_analytics_token.txt рядом со скриптом."
}

$Token = (Get-Content -LiteralPath $TokenPath -Raw).Trim()

if ([string]::IsNullOrWhiteSpace($Token) -or ($Token -match "ВСТАВ|TOKEN|ТОКЕН")) {
    throw "wb_analytics_token.txt пустой или содержит шаблон."
}

$DateTo = (Get-Date).Date.AddDays(-1)
$DateFrom = $DateTo.AddDays(-6)

$PastTo = $DateFrom.AddDays(-1)
$PastFrom = $PastTo.AddDays(-6)

$BeginDate = $DateFrom.ToString("yyyy-MM-dd")
$EndDate = $DateTo.ToString("yyyy-MM-dd")
$PastBeginDate = $PastFrom.ToString("yyyy-MM-dd")
$PastEndDate = $PastTo.ToString("yyyy-MM-dd")

Write-Host "WB Analytics — диагностика воронки" -ForegroundColor Cyan
Write-Host "Текущий период: $BeginDate - $EndDate"
Write-Host "Прошлый период: $PastBeginDate - $PastEndDate"
Write-Host "Будет выполнен 1 API-запрос."

$bodyObject = [ordered]@{
    selectedPeriod = [ordered]@{
        start = $BeginDate
        end   = $EndDate
    }
    pastPeriod = [ordered]@{
        start = $PastBeginDate
        end   = $PastEndDate
    }
    nmIds = @()
    brandNames = @()
    subjectIds = @()
    tagIds = @()
    skipDeletedNm = $true
    orderBy = [ordered]@{
        field = "orderSum"
        mode  = "desc"
    }
    limit = 20
    offset = 0
}

$bodyJson = $bodyObject | ConvertTo-Json -Depth 10

$uri = "https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products"

$headers = @{
    Authorization = "Bearer $Token"
}

try {
    $response = Invoke-RestMethod `
        -Method Post `
        -Uri $uri `
        -Headers $headers `
        -ContentType "application/json; charset=utf-8" `
        -Body ([System.Text.Encoding]::UTF8.GetBytes($bodyJson)) `
        -TimeoutSec 120
}
catch {
    Write-Host ""
    Write-Host "ОШИБКА API" -ForegroundColor Red
    Write-Host $_.Exception.Message
    throw
}

$rawJson = $response | ConvertTo-Json -Depth 30
[System.IO.File]::WriteAllText(
    $RawPath,
    $rawJson,
    (New-Object System.Text.UTF8Encoding($true))
)

$products = @()

if ($null -ne $response -and $null -ne $response.data -and $null -ne $response.data.products) {
    $products = @($response.data.products)
}

$rows = @()

foreach ($item in $products) {
    $product = Get-Field $item "product" $null
    $statistic = Get-Field $item "statistic" $null

    $selected = Get-Field $statistic "selected" $null
    $past = Get-Field $statistic "past" $null
    $comparison = Get-Field $statistic "comparison" $null

    $selConversions = Get-Field $selected "conversions" $null
    $pastConversions = Get-Field $past "conversions" $null

    $selectedOrders = [double](Get-Field $selected "orderCount" 0)
    $pastOrders = [double](Get-Field $past "orderCount" 0)

    $selectedAvgCalculated = [Math]::Round($selectedOrders / 7, 2)
    $pastAvgCalculated = [Math]::Round($pastOrders / 7, 2)

    $rows += [pscustomobject][ordered]@{
        "Артикул WB" = Get-Field $product "nmId" ""
        "Артикул продавца" = Get-Field $product "vendorCode" ""
        "Наименование" = Get-Field $product "title" ""
        "Бренд" = Get-Field $product "brandName" ""
        "Предмет" = Get-Field $product "subjectName" ""

        "Текущий период с" = $BeginDate
        "Текущий период по" = $EndDate
        "Прошлый период с" = $PastBeginDate
        "Прошлый период по" = $PastEndDate

        "Переходы текущий" = Get-Field $selected "openCount" 0
        "Корзины текущий" = Get-Field $selected "cartCount" 0
        "Заказы текущий" = $selectedOrders
        "Сумма заказов текущий" = Get-Field $selected "orderSum" 0
        "Средние заказы API текущий" = Get-Field $selected "avgOrdersCountPerDay" ""
        "Средние заказы расчёт текущий" = $selectedAvgCalculated

        "Переходы прошлый" = Get-Field $past "openCount" 0
        "Корзины прошлый" = Get-Field $past "cartCount" 0
        "Заказы прошлый" = $pastOrders
        "Сумма заказов прошлый" = Get-Field $past "orderSum" 0
        "Средние заказы API прошлый" = Get-Field $past "avgOrdersCountPerDay" ""
        "Средние заказы расчёт прошлый" = $pastAvgCalculated

        "Выкупы текущий" = Get-Field $selected "buyoutCount" 0
        "Сумма выкупов текущий" = Get-Field $selected "buyoutSum" 0
        "Отмены текущий" = Get-Field $selected "cancelCount" 0
        "Сумма отмен текущий" = Get-Field $selected "cancelSum" 0
        "Средняя цена текущий" = Get-Field $selected "avgPrice" 0

        "Конверсия карточка-корзина текущий" = Get-Field $selConversions "addToCartPercent" 0
        "Конверсия корзина-заказ текущий" = Get-Field $selConversions "cartToOrderPercent" 0
        "Процент выкупа текущий" = Get-Field $selConversions "buyoutPercent" 0

        "Конверсия карточка-корзина прошлый" = Get-Field $pastConversions "addToCartPercent" 0
        "Конверсия корзина-заказ прошлый" = Get-Field $pastConversions "cartToOrderPercent" 0
        "Процент выкупа прошлый" = Get-Field $pastConversions "buyoutPercent" 0

        "Динамика заказов %" = Get-Field $comparison "orderCountDynamic" ""
        "Остаток WB сейчас" = Get-Field (Get-Field $product "stocks" $null) "wb" 0
        "Остаток FBS сейчас" = Get-Field (Get-Field $product "stocks" $null) "mp" 0
    }
}

if ($rows.Count -gt 0) {
    $rows | Export-Csv -LiteralPath $CsvPath -Delimiter ";" -NoTypeInformation -Encoding UTF8
}

Write-Host ""
Write-Host "ГОТОВО" -ForegroundColor Green
Write-Host "Получено товаров в тестовой странице: $($products.Count)"
Write-Host "RAW: $RawPath"
Write-Host "CSV: $CsvPath"
Write-Host ""
Write-Host "Важно: дни в наличии этим тестом пока не считаются."
Write-Host "Для них следующим шагом подключим STOCK_HISTORY_DAILY_CSV со stockType = все склады (WB + продавец)."
