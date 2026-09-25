# ================================================================
# WB РЕКЛАМА v10
#
# Архитектура:
#   1) /adv/v1/upd -> история фактических затрат за 7 полных дней
#   2) суммируем updSum по campaign ID
#   3) оставляем только кампании с расходом >= 100 руб.
#   4) только для них получаем сведения о кампании
#   5) только для них вызываем /adv/v3/fullstats
#   6) формируем CSV по товарам
#
# Нужен токен WB API категории "Продвижение".
# Совместимо с Windows PowerShell 5.1.
# ================================================================

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


# ================================================================
# 1. НАСТРОЙКИ
# ================================================================

# Кампании с расходом меньше этого значения не запрашиваем в fullstats.
$MinCampaignSpend = 100

# fullstats: 3 запроса в минуту, официальный интервал 20 секунд.
$FullStatsPauseSeconds = 22


# ================================================================
# 2. ПУТИ
# ================================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$TokenPath  = Join-Path $ScriptDir "wb_adv_token.txt"
$OutputPath = Join-Path $ScriptDir "WB_Реклама_7дней.csv"
$TempPath   = Join-Path $ScriptDir "WB_Реклама_7дней.tmp.csv"
$LogPath    = Join-Path $ScriptDir "WB_Реклама_7дней.log"


# ================================================================
# 3. ПЕРИОД
# ================================================================
# Вчера = последний полный день.
# От вчера ещё минус 6 дней = 7 дней включительно.

$DateTo   = (Get-Date).Date.AddDays(-1)
$DateFrom = $DateTo.AddDays(-6)

$BeginDate = $DateFrom.ToString("yyyy-MM-dd")
$EndDate   = $DateTo.ToString("yyyy-MM-dd")


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

    $property = $Object.PSObject.Properties[$Name]

    if ($null -eq $property) {
        return $Default
    }

    return $property.Value
}


function Has-Field {
    param(
        $Object,
        [string]$Name
    )

    if ($null -eq $Object) {
        return $false
    }

    return ($null -ne $Object.PSObject.Properties[$Name])
}


# Разбивает массив на пачки, например по 50 ID.
function Split-IntoBatches {
    param(
        [object[]]$Items,
        [int]$Size
    )

    $result = @()

    if ($null -eq $Items) {
        return $result
    }

    if ($Items.Count -eq 0) {
        return $result
    }

    for ($i = 0; $i -lt $Items.Count; $i += $Size) {
        $end = [Math]::Min($i + $Size - 1, $Items.Count - 1)
        $batch = @($Items[$i..$end])

        # Запятая означает: добавить batch как один элемент.
        $result += ,$batch
    }

    return $result
}


function Get-StatusName {
    param($Status)

    try {
        $number = [int]$Status
    }
    catch {
        return [string]$Status
    }

    switch ($number) {
        7  { return "Завершена" }
        9  { return "Активна" }
        11 { return "Приостановлена" }
        default { return [string]$number }
    }
}


function Get-BidTypeName {
    param([string]$BidType)

    if ([string]::IsNullOrWhiteSpace($BidType)) {
        return "Не определено"
    }

    switch ($BidType.ToLowerInvariant()) {
        "manual"  { return "Ручная ставка" }
        "unified" { return "Единая ставка" }
        "auto"    { return "Единая ставка" }
        default   { return "Не определено" }
    }
}


# ================================================================
# 5. GET-ЗАПРОС К WB API
# ================================================================
# Одна функция для всех GET-запросов:
#   - добавляет Bearer token;
#   - сама декодирует UTF-8;
#   - автоматически повторяет 429.

function Invoke-WBGet {
    param(
        [Parameter(Mandatory=$true)][string]$BaseUri,
        [hashtable]$Query = @{},
        [int]$MaxAttempts = 8
    )

    $parts = @()

    foreach ($key in $Query.Keys) {
        $encodedKey = [System.Uri]::EscapeDataString([string]$key)
        $encodedValue = [System.Uri]::EscapeDataString([string]$Query[$key])
        $parts += "$encodedKey=$encodedValue"
    }

    $Uri = $BaseUri

    if ($parts.Count -gt 0) {
        $Uri = $BaseUri + "?" + ($parts -join "&")
    }

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $response = $null

        try {
            $request = [System.Net.HttpWebRequest]::Create($Uri)
            $request.Method = "GET"
            $request.Timeout = 180000
            $request.ReadWriteTimeout = 180000
            $request.Accept = "application/json"
            $request.Headers["Authorization"] = $script:AuthorizationHeader
            $request.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate

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
            $status = $null
            $errorBody = ""
            $retryHeader = $null

            try {
                if ($null -ne $_.Exception.Response) {
                    $webResponse = [System.Net.HttpWebResponse]$_.Exception.Response
                    $status = [int]$webResponse.StatusCode

                    $retryHeader = $webResponse.Headers["X-Ratelimit-Retry"]

                    if ([string]::IsNullOrWhiteSpace($retryHeader)) {
                        $retryHeader = $webResponse.Headers["Retry-After"]
                    }

                    try {
                        $errorStream = $webResponse.GetResponseStream()
                        $errorMemory = New-Object System.IO.MemoryStream

                        try {
                            $errorStream.CopyTo($errorMemory)
                            $errorBytes = $errorMemory.ToArray()
                            $errorBody = [System.Text.Encoding]::UTF8.GetString($errorBytes)
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
            }
            catch {
            }

            if ($status -eq 429 -and $attempt -lt $MaxAttempts) {
                $wait = 62

                if (-not [string]::IsNullOrWhiteSpace($retryHeader)) {
                    $retryNumber = 0.0

                    if ([double]::TryParse(
                        $retryHeader,
                        [System.Globalization.NumberStyles]::Any,
                        [System.Globalization.CultureInfo]::InvariantCulture,
                        [ref]$retryNumber
                    )) {
                        if ($retryNumber -gt 0 -and $retryNumber -lt 600) {
                            $wait = [Math]::Ceiling($retryNumber) + 3
                        }
                    }
                }

                Write-Log "WB API вернул 429. Попытка $attempt из $MaxAttempts. Ждём $wait сек."
                Start-Sleep -Seconds $wait
                continue
            }

            $hint = ""

            switch ($status) {
                400 { $hint = " Проверьте параметры запроса." }
                401 { $hint = " Проверьте токен." }
                403 { $hint = " Проверьте категорию Продвижение у токена." }
                429 { $hint = " Превышен лимит запросов WB API." }
            }

            $statusText = if ($null -eq $status) {
                "без HTTP-кода"
            }
            else {
                "HTTP $status"
            }

            if (-not [string]::IsNullOrWhiteSpace($errorBody)) {
                throw "Ошибка WB API ($statusText): $($_.Exception.Message).$hint`nОтвет WB: $errorBody`nURL: $Uri"
            }

            throw "Ошибка WB API ($statusText): $($_.Exception.Message).$hint`nURL: $Uri"
        }
        finally {
            if ($null -ne $response) {
                $response.Dispose()
            }
        }
    }
}


# ================================================================
# 6. ТОКЕН
# ================================================================

if (-not (Test-Path -LiteralPath $TokenPath)) {
    throw "Не найден wb_adv_token.txt рядом со скриптом."
}

$Token = (Get-Content -LiteralPath $TokenPath -Raw).Trim()

if ([string]::IsNullOrWhiteSpace($Token) -or ($Token -match "ВСТАВ|TOKEN|ТОКЕН")) {
    throw "wb_adv_token.txt пустой или содержит шаблон."
}

$script:AuthorizationHeader = "Bearer $Token"


Write-Log "============================================================"
Write-Log "Старт выгрузки рекламной статистики WB v10."
Write-Log "Период: $BeginDate - $EndDate."
Write-Log "Шаг 1: история затрат -> фильтр расход >= $MinCampaignSpend руб."


# ================================================================
# 7. ОДИН ЗАПРОС: ИСТОРИЯ ФАКТИЧЕСКИХ ЗАТРАТ
# ================================================================
#
# GET /adv/v1/upd?from=YYYY-MM-DD&to=YYYY-MM-DD
#
# В ответе может быть несколько строк одной кампании.
# Поэтому updSum суммируем по advertId.

$UpdUrl = "https://advert-api.wildberries.ru/adv/v1/upd"

$updResponse = Invoke-WBGet `
    -BaseUri $UpdUrl `
    -Query @{
        from = $BeginDate
        to   = $EndDate
    }

$updRows = @($updResponse)

Write-Log "Строк истории затрат получено: $($updRows.Count)."


# ================================================================
# 8. СУММИРУЕМ РАСХОД ПО КАМПАНИЯМ
# ================================================================

$SpendMap = @{}

foreach ($item in $updRows) {
    $advertIdRaw = Get-Field $item "advertId" $null

    if ($null -eq $advertIdRaw) {
        continue
    }

    $advertId = [Int64]$advertIdRaw
    $updSum = [double](Get-Field $item "updSum" 0)

    $key = [string]$advertId

    if (-not $SpendMap.ContainsKey($key)) {
        $SpendMap[$key] = [pscustomobject]@{
            AdvertId    = $advertId
            Spend       = [double]0
            CampaignName = ""
            AdvertType  = ""
            PaymentType = ""
            Status      = $null
        }
    }

    $row = $SpendMap[$key]
    $row.Spend += $updSum

    $campName = [string](Get-Field $item "campName" "")
    if (-not [string]::IsNullOrWhiteSpace($campName)) {
        $row.CampaignName = $campName
    }

    $advertType = [string](Get-Field $item "advertType" "")
    if (-not [string]::IsNullOrWhiteSpace($advertType)) {
        $row.AdvertType = $advertType
    }

    $paymentType = [string](Get-Field $item "paymentType" "")
    if (-not [string]::IsNullOrWhiteSpace($paymentType)) {
        $row.PaymentType = $paymentType
    }

    $statusRaw = Get-Field $item "advertStatus" $null
    if ($null -ne $statusRaw) {
        $row.Status = $statusRaw
    }
}


$QualifiedBySpend = @(
    $SpendMap.Values |
        Where-Object { $_.Spend -ge $MinCampaignSpend } |
        Sort-Object AdvertId
)

$QualifiedIds = @(
    $QualifiedBySpend |
        ForEach-Object { [Int64]$_.AdvertId }
)

Write-Log "Кампаний с любыми затратами за период: $($SpendMap.Count)."
Write-Log "Кампаний с расходом >= $MinCampaignSpend руб.: $($QualifiedIds.Count)."


if ($QualifiedIds.Count -eq 0) {
    Write-Log "Нет кампаний, прошедших фильтр расхода."
}


# ================================================================
# 9. ОДНИМ ЗАПРОСОМ ПОЛУЧАЕМ СВЕДЕНИЯ О КАМПАНИЯХ
# ================================================================
#
# Главное изменение v10:
#
# В v9 мы делили 255 нужных кампаний на пачки по 50 и делали
# несколько запросов /api/advert/v2/adverts?ids=...
#
# Теперь делаем ОДИН запрос:
#   /api/advert/v2/adverts?statuses=7,9,11
#
# Он возвращает сведения о кампаниях поддерживаемых fullstats статусов.
# После этого PowerShell ЛОКАЛЬНО оставляет только те campaign ID,
# которые уже прошли фильтр расхода >= 100 руб.
#
# Таким образом этап сведений о кампаниях = 1 API-запрос.

Write-Log "Шаг 2: одним запросом получаем сведения о кампаниях статусов 7,9,11."

# Небольшая пауза после /upd, чтобы не делать два запроса подряд.
Start-Sleep -Seconds 2

$infoResponse = Invoke-WBGet `
    -BaseUri "https://advert-api.wildberries.ru/api/advert/v2/adverts" `
    -Query @{
        statuses = "7,9,11"
    }

$items = @()

if (Has-Field $infoResponse "adverts") {
    $items = @($infoResponse.adverts)
}
else {
    $items = @($infoResponse)
}

Write-Log "Кампаний получено из справочника статусов 7,9,11: $($items.Count)."

# Делаем быстрый набор ID, прошедших фильтр расхода.
# Ключ = текстовое представление campaign ID.
$QualifiedIdSet = @{}

foreach ($id in $QualifiedIds) {
    $QualifiedIdSet[[string]$id] = $true
}

$CampaignInfoMap = @{}

foreach ($campaign in $items) {
    $idRaw = Get-Field $campaign "id" $null

    if ($null -eq $idRaw) {
        $idRaw = Get-Field $campaign "advertId" $null
    }

    if ($null -eq $idRaw) {
        continue
    }

    $advertId = [Int64]$idRaw
    $advertKey = [string]$advertId

    # Нас интересуют только кампании, уже прошедшие spend >= 100.
    if (-not $QualifiedIdSet.ContainsKey($advertKey)) {
        continue
    }

    $settings = Get-Field $campaign "settings" $null
    $timestamps = Get-Field $campaign "timestamps" $null

    $name = [string](Get-Field $campaign "name" "")
    if ([string]::IsNullOrWhiteSpace($name) -and ($null -ne $settings)) {
        $name = [string](Get-Field $settings "name" "")
    }

    $bidType = [string](Get-Field $campaign "bid_type" "")
    if ([string]::IsNullOrWhiteSpace($bidType)) {
        $bidType = [string](Get-Field $campaign "bidType" "")
    }

    $paymentType = [string](Get-Field $campaign "payment_type" "")
    if ([string]::IsNullOrWhiteSpace($paymentType) -and ($null -ne $settings)) {
        $paymentType = [string](Get-Field $settings "payment_type" "")
    }

    $status = Get-Field $campaign "status" $null

    $created = ""
    if ($null -ne $timestamps) {
        $created = [string](Get-Field $timestamps "created" "")
    }

    if ([string]::IsNullOrWhiteSpace($created)) {
        $created = [string](Get-Field $campaign "createTime" "")
    }

    $nmSettings = @((Get-Field $campaign "nm_settings" @()))

    $CampaignInfoMap[$advertKey] = [pscustomobject]@{
        AdvertId    = $advertId
        Name        = $name
        BidType     = $bidType
        PaymentType = $paymentType
        Status      = $status
        Created     = $created
        NmSettings  = $nmSettings
    }
}

Write-Log "Из расход-фильтра найдены в справочнике 7/9/11: $($CampaignInfoMap.Count) из $($QualifiedIds.Count)."

# ================================================================
# 10. ID ДЛЯ FULLSTATS
# ================================================================
#
# Так как предыдущий запрос уже был ограничен statuses=7,9,11,
# в fullstats отправляем только пересечение:
#
#   расход >= 100
#   И
#   кампания найдена среди статусов 7/9/11

$FullStatsIds = @(
    $CampaignInfoMap.Keys |
        ForEach-Object { [Int64]$_ } |
        Sort-Object -Unique
)

$MissingInfoCount = $QualifiedIds.Count - $FullStatsIds.Count

Write-Log "Кампаний для fullstats: $($FullStatsIds.Count)."

if ($MissingInfoCount -gt 0) {
    Write-Log "Не отправлено в fullstats из-за текущего статуса/отсутствия в справочнике: $MissingInfoCount."
}


# ================================================================
# 11. FULLSTATS ТОЛЬКО ПО КАМПАНИЯМ С РАСХОДОМ >= 100
# ================================================================

$FullStatsBatches = Split-IntoBatches -Items $FullStatsIds -Size 50

Write-Log "Запросов fullstats потребуется: $($FullStatsBatches.Count)."

if ($FullStatsBatches.Count -gt 1) {
    $estimatedSeconds = ($FullStatsBatches.Count - 1) * $FullStatsPauseSeconds
    Write-Log "Минимальное время ожиданий между fullstats: около $estimatedSeconds сек."
}


$StatsMap = @{}
$batchNumber = 0

foreach ($batch in $FullStatsBatches) {
    $batchNumber++

    if ($batchNumber -gt 1) {
        Write-Log "Пауза $FullStatsPauseSeconds сек. перед следующим fullstats."
        Start-Sleep -Seconds $FullStatsPauseSeconds
    }

    Write-Log "Fullstats: пакет $batchNumber из $($FullStatsBatches.Count)."

    $statsResponse = Invoke-WBGet `
        -BaseUri "https://advert-api.wildberries.ru/adv/v3/fullstats" `
        -Query @{
            ids       = ($batch -join ",")
            beginDate = $BeginDate
            endDate   = $EndDate
        }

    foreach ($campaignStat in @($statsResponse)) {
        $advertIdRaw = Get-Field $campaignStat "advertId" $null

        if ($null -eq $advertIdRaw) {
            continue
        }

        $advertId = [Int64]$advertIdRaw

        foreach ($day in @((Get-Field $campaignStat "days" @()))) {
            foreach ($app in @((Get-Field $day "apps" @()))) {
                foreach ($nm in @((Get-Field $app "nms" @()))) {
                    $nmIdRaw = Get-Field $nm "nmId" $null

                    if ($null -eq $nmIdRaw) {
                        $nmIdRaw = Get-Field $nm "nm" $null
                    }

                    if ($null -eq $nmIdRaw) {
                        continue
                    }

                    $nmId = [Int64]$nmIdRaw
                    $key = "$advertId|$nmId"

                    if (-not $StatsMap.ContainsKey($key)) {
                        $StatsMap[$key] = [pscustomobject]@{
                            AdvertId = $advertId
                            NmId = $nmId
                            Name = ""
                            Views = [Int64]0
                            Clicks = [Int64]0
                            Atbs = [Int64]0
                            Orders = [Int64]0
                            Canceled = [Int64]0
                            Spend = [double]0
                            Revenue = [double]0
                        }
                    }

                    $row = $StatsMap[$key]

                    $name = [string](Get-Field $nm "name" "")
                    if (-not [string]::IsNullOrWhiteSpace($name)) {
                        $row.Name = $name
                    }

                    $row.Views += [Int64](Get-Field $nm "views" 0)
                    $row.Clicks += [Int64](Get-Field $nm "clicks" 0)
                    $row.Atbs += [Int64](Get-Field $nm "atbs" 0)
                    $row.Orders += [Int64](Get-Field $nm "orders" 0)
                    $row.Canceled += [Int64](Get-Field $nm "canceled" 0)
                    $row.Spend += [double](Get-Field $nm "sum" 0)
                    $row.Revenue += [double](Get-Field $nm "sum_price" 0)
                }
            }
        }
    }
}


# ================================================================
# 12. ДОБАВЛЯЕМ НУЛЕВЫЕ ТОВАРЫ ПРОШЕДШИХ КАМПАНИЙ
# ================================================================
# Если товар есть в nm_settings, но в fullstats у него не было строки,
# создаём строку с нулевыми метриками.

foreach ($advertId in $FullStatsIds) {
    $campaignKey = [string]$advertId

    if (-not $CampaignInfoMap.ContainsKey($campaignKey)) {
        continue
    }

    $campaign = $CampaignInfoMap[$campaignKey]

    foreach ($nmSetting in @($campaign.NmSettings)) {
        $nmIdRaw = Get-Field $nmSetting "nm_id" $null

        if ($null -eq $nmIdRaw) {
            $nmIdRaw = Get-Field $nmSetting "nmId" $null
        }

        if ($null -eq $nmIdRaw) {
            continue
        }

        $nmId = [Int64]$nmIdRaw
        $key = "$advertId|$nmId"

        if (-not $StatsMap.ContainsKey($key)) {
            $StatsMap[$key] = [pscustomobject]@{
                AdvertId = [Int64]$advertId
                NmId = $nmId
                Name = ""
                Views = [Int64]0
                Clicks = [Int64]0
                Atbs = [Int64]0
                Orders = [Int64]0
                Canceled = [Int64]0
                Spend = [double]0
                Revenue = [double]0
            }
        }
    }
}


# ================================================================
# 13. ИТОГОВЫЕ СТРОКИ
# ================================================================

$Rows = @()

foreach ($stat in ($StatsMap.Values | Sort-Object AdvertId, NmId)) {
    $campaignKey = [string]$stat.AdvertId

    # В итог попадают только кампании, которые реально прошли spend filter.
    if (-not $SpendMap.ContainsKey($campaignKey)) {
        continue
    }

    if ($SpendMap[$campaignKey].Spend -lt $MinCampaignSpend) {
        continue
    }

    $campaign = $null

    if ($CampaignInfoMap.ContainsKey($campaignKey)) {
        $campaign = $CampaignInfoMap[$campaignKey]
    }

    $updInfo = $SpendMap[$campaignKey]

    $campaignName = $updInfo.CampaignName
    $bidTypeName = "Не определено"
    $paymentType = $updInfo.PaymentType
    $statusName = Get-StatusName $updInfo.Status
    $created = ""

    if ($null -ne $campaign) {
        if (-not [string]::IsNullOrWhiteSpace($campaign.Name)) {
            $campaignName = $campaign.Name
        }

        $bidTypeName = Get-BidTypeName $campaign.BidType

        if (-not [string]::IsNullOrWhiteSpace($campaign.PaymentType)) {
            $paymentType = $campaign.PaymentType
        }

        if ($null -ne $campaign.Status) {
            $statusName = Get-StatusName $campaign.Status
        }

        $created = $campaign.Created
    }

    $ctr = if ($stat.Views -gt 0) {
        [Math]::Round(($stat.Clicks / $stat.Views) * 100, 2)
    }
    else {
        0
    }

    $cr = if ($stat.Clicks -gt 0) {
        [Math]::Round(($stat.Orders / $stat.Clicks) * 100, 2)
    }
    else {
        0
    }

    $cpc = if ($stat.Clicks -gt 0) {
        [Math]::Round($stat.Spend / $stat.Clicks, 2)
    }
    else {
        0
    }

    $cpm = if ($stat.Views -gt 0) {
        [Math]::Round(($stat.Spend / $stat.Views) * 1000, 2)
    }
    else {
        0
    }

    $drr = if ($stat.Revenue -gt 0) {
        [Math]::Round(($stat.Spend / $stat.Revenue) * 100, 2)
    }
    else {
        0
    }

    $Rows += [pscustomobject][ordered]@{
        "Название кампании"      = $campaignName
        "Раздел (тип РК)"        = $bidTypeName
        "ID кампании"            = $stat.AdvertId
        "Артикул"                = $stat.NmId
        "Наименование"           = $stat.Name
        "Дата создания"          = $created
        "Статус"                 = $statusName
        "Модель оплаты"          = $paymentType
        "Период с"               = $BeginDate
        "Период по"              = $EndDate
        "Показы"                 = $stat.Views
        "Клики"                  = $stat.Clicks
        "Расход рекламы"         = [Math]::Round($stat.Spend, 2)
        "ДРР Реклама (acc.)"     = $drr
        "CTR"                    = $ctr
        "Корзины (acc.)"         = $stat.Atbs
        "Заказы (acc.)"          = $stat.Orders
        "Отмены заказов"         = $stat.Canceled
        "Выручка (acc.)"         = [Math]::Round($stat.Revenue, 2)
        "CR (acc.)"              = $cr
        "CPC"                    = $cpc
        "CPM"                    = $cpm
        "Расход кампании 7д upd" = [Math]::Round($updInfo.Spend, 2)
    }
}


Write-Log "Итоговых строк кампания + артикул: $($Rows.Count)."


# ================================================================
# 14. CSV
# ================================================================

$Headers = @(
    "Название кампании",
    "Раздел (тип РК)",
    "ID кампании",
    "Артикул",
    "Наименование",
    "Дата создания",
    "Статус",
    "Модель оплаты",
    "Период с",
    "Период по",
    "Показы",
    "Клики",
    "Расход рекламы",
    "ДРР Реклама (acc.)",
    "CTR",
    "Корзины (acc.)",
    "Заказы (acc.)",
    "Отмены заказов",
    "Выручка (acc.)",
    "CR (acc.)",
    "CPC",
    "CPM",
    "Расход кампании 7д upd"
)


if ($Rows.Count -gt 0) {
    $csvLines = $Rows |
        Select-Object $Headers |
        ConvertTo-Csv -NoTypeInformation -Delimiter ";"
}
else {
    $empty = [ordered]@{}

    foreach ($header in $Headers) {
        $empty[$header] = ""
    }

    $tempCsv = [pscustomobject]$empty |
        ConvertTo-Csv -NoTypeInformation -Delimiter ";"

    $csvLines = @($tempCsv[0])
}


$utf8Bom = New-Object System.Text.UTF8Encoding($true)

[System.IO.File]::WriteAllLines(
    $TempPath,
    $csvLines,
    $utf8Bom
)


# ================================================================
# 15. БЕЗОПАСНАЯ ЗАМЕНА CSV
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
    $fallback = Join-Path $ScriptDir ("WB_Реклама_7дней_NEW_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

    if (Test-Path -LiteralPath $TempPath) {
        Move-Item -LiteralPath $TempPath -Destination $fallback -Force
    }

    throw "WB_Реклама_7дней.csv заблокирован. Новые данные сохранены: $fallback"
}


Write-Log "ГОТОВО."
Write-Log "CSV: $OutputPath"
Write-Log "Период: $BeginDate - $EndDate."
Write-Log "Кампаний с затратами: $($SpendMap.Count)."
Write-Log "Кампаний с расходом >= $MinCampaignSpend руб.: $($QualifiedIds.Count)."
Write-Log "Кампаний отправлено в fullstats: $($FullStatsIds.Count)."
Write-Log "Строк: $($Rows.Count)."

Write-Host ""
Write-Host "ОБНОВЛЕНИЕ РЕКЛАМЫ ЗАВЕРШЕНО" -ForegroundColor Green
Write-Host "Период: $BeginDate - $EndDate"
Write-Host "Расход-фильтр: >= $MinCampaignSpend руб."
Write-Host "Кампаний с расходом >= порога: $($QualifiedIds.Count)"
Write-Host "Запросов fullstats: $($FullStatsBatches.Count)"
Write-Host "Файл: $OutputPath"
