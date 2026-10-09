# ================================================================
# WB РЕКЛАМА — СТАТИСТИКА ЗА 7 ДНЕЙ
#
# Архитектура:
#   1) /adv/v1/upd -> история фактических затрат за 7 полных дней
#   2) суммируем updSum по campaign ID
#   3) оставляем все кампании с фактическим расходом > 0 руб.
#   4) для них получаем сведения о кампании
#   5) для них вызываем /adv/v3/fullstats
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

# В рабочей версии берём ВСЕ кампании, у которых за период есть фактический расход > 0.
# Это нужно, чтобы не терять малые кампании и ассоциативные конверсии.
# Защита от неожиданного роста количества запросов fullstats.
# 1 пакет = максимум 50 ID. Между пакетами выдерживается пауза 22 сек.
# Лимит поднят до 20 пакетов (до 1000 кампаний), чтобы рабочая выгрузка
# не останавливалась при естественном росте числа кампаний.
# При превышении 20 пакетов скрипт остановится ДО fullstats как защита от аномалии.
$MaxFullStatsBatches = 20

# Историю затрат /adv/v1/upd загружаем ПО ДНЯМ.
# Это уменьшает размер каждого ответа и защищает от зависания длинного 7-дневного запроса.
$UpdPauseMilliseconds = 1200
$UpdTimeoutMilliseconds = 60000
$UpdMaxAttempts = 3

# fullstats: 3 запроса в минуту, официальный интервал 20 секунд.
$FullStatsPauseSeconds = 22


# ================================================================
# 2. ПУТИ
# ================================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$TokenPath  = Join-Path $ScriptDir "wb_adv_token.txt"
$OutputPath    = Join-Path $ScriptDir "WB_Реклама_7дней.csv"
$TempPath      = Join-Path $ScriptDir "WB_Реклама_7дней.tmp.csv"
$MultiPath     = Join-Path $ScriptDir "WB_Реклама_МУЛЬТИАРТИКУЛ.csv"
$MultiTempPath = Join-Path $ScriptDir "WB_Реклама_МУЛЬТИАРТИКУЛ.tmp.csv"
$LogPath       = Join-Path $ScriptDir "WB_Реклама_7дней.log"


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
        [int]$MaxAttempts = 8,
        [int]$TimeoutMilliseconds = 180000
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
            $request.Timeout = $TimeoutMilliseconds
            $request.ReadWriteTimeout = $TimeoutMilliseconds
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

            if ($_.Exception.Status -eq [System.Net.WebExceptionStatus]::Timeout -and $attempt -lt $MaxAttempts) {
                $wait = 10
                Write-Log "WB API не ответил за $([Math]::Round($TimeoutMilliseconds / 1000)) сек. Попытка $attempt из $MaxAttempts. Повтор через $wait сек."
                Start-Sleep -Seconds $wait
                continue
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
Write-Log "Старт выгрузки рекламной статистики WB за 7 дней."
Write-Log "Период: $BeginDate - $EndDate."
Write-Log "Шаг 1: история затрат -> берём все кампании с фактическим расходом > 0 руб."


# ================================================================
# 7. ИСТОРИЯ ФАКТИЧЕСКИХ ЗАТРАТ — ПО ОДНОМУ ДНЮ
# ================================================================
#
# GET /adv/v1/upd?from=YYYY-MM-DD&to=YYYY-MM-DD
#
# Раньше весь 7-дневный период запрашивался одним запросом.
# На больших кабинетах этот запрос может долго отвечать и уйти в timeout.
#
# Теперь делаем 7 небольших запросов — по одному календарному дню.
# Между запросами пауза 1,2 сек., чтобы соблюдать лимит API.
# Если отдельный день не ответил за 60 сек., запрос автоматически повторяется.

$UpdUrl = "https://advert-api.wildberries.ru/adv/v1/upd"
$updRows = @()

for ($dayIndex = 0; $dayIndex -lt 7; $dayIndex++) {
    $day = $DateFrom.AddDays($dayIndex)
    $dayText = $day.ToString("yyyy-MM-dd")
    $humanIndex = $dayIndex + 1

    if ($dayIndex -gt 0) {
        Start-Sleep -Milliseconds $UpdPauseMilliseconds
    }

    Write-Log "История затрат: день $humanIndex из 7 ($dayText). Отправляю запрос..."

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    $dayResponse = Invoke-WBGet `
        -BaseUri $UpdUrl `
        -Query @{
            from = $dayText
            to   = $dayText
        } `
        -MaxAttempts $UpdMaxAttempts `
        -TimeoutMilliseconds $UpdTimeoutMilliseconds

    $stopwatch.Stop()

    $dayRows = @($dayResponse)
    $updRows += $dayRows

    Write-Log "История затрат: день $humanIndex из 7 получен за $([Math]::Round($stopwatch.Elapsed.TotalSeconds, 1)) сек. Строк: $($dayRows.Count)."
}

Write-Log "Строк истории затрат получено за 7 дней: $($updRows.Count)."


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
        Where-Object { $_.Spend -gt 0 } |
        Sort-Object AdvertId
)

$QualifiedIds = @(
    $QualifiedBySpend |
        ForEach-Object { [Int64]$_.AdvertId }
)

Write-Log "Кампаний с любыми затратами за период: $($SpendMap.Count)."
Write-Log "Кампаний с фактическим расходом > 0 руб.: $($QualifiedIds.Count)."


if ($QualifiedIds.Count -eq 0) {
    Write-Log "Нет кампаний с фактическим расходом > 0 руб."
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
# которые имеют фактический расход > 0 руб.
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

    # Нас интересуют только кампании, уже прошедшие spend > 0.
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

# Для каждой кампании отдельно запоминаем товары, которые действительно
# находятся в её nm_settings. Это важно: fullstats может вернуть nmId товара,
# который был КУПЛЕН после рекламы другой кампании. Такой nmId является
# ассоциативной конверсией, а не рекламируемым товаром этой кампании.
$CampaignNmSetMap = @{}
$CampaignsWithoutNmSettings = 0

foreach ($campaignKey in $CampaignInfoMap.Keys) {
    $nmSet = @{}

    foreach ($nmSetting in @($CampaignInfoMap[$campaignKey].NmSettings)) {
        $nmIdRaw = Get-Field $nmSetting "nm_id" $null

        if ($null -eq $nmIdRaw) {
            $nmIdRaw = Get-Field $nmSetting "nmId" $null
        }

        if ($null -eq $nmIdRaw) {
            continue
        }

        $nmSet[[string][Int64]$nmIdRaw] = $true
    }

    if ($nmSet.Count -eq 0) {
        $CampaignsWithoutNmSettings++
    }

    $CampaignNmSetMap[$campaignKey] = $nmSet
}

Write-Log "Кампаний без nm_settings: $CampaignsWithoutNmSettings."

# ================================================================
# 10. ID ДЛЯ FULLSTATS
# ================================================================
#
# Так как предыдущий запрос уже был ограничен statuses=7,9,11,
# в fullstats отправляем только пересечение:
#
#   расход > 0
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
# 11. FULLSTATS ПО ВСЕМ КАМПАНИЯМ С ФАКТИЧЕСКИМ РАСХОДОМ > 0
# ================================================================

$FullStatsBatches = Split-IntoBatches -Items $FullStatsIds -Size 50

Write-Log "Запросов fullstats потребуется: $($FullStatsBatches.Count)."

if ($FullStatsBatches.Count -gt $MaxFullStatsBatches) {
    throw "ЗАЩИТА API: требуется $($FullStatsBatches.Count) запросов fullstats, разрешено максимум $MaxFullStatsBatches. Fullstats НЕ запускался."
}

Write-Log "Защита лимита: максимум fullstats-пакетов в рабочем запуске = $MaxFullStatsBatches."

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
                    $campaignKey = [string]$advertId
                    $isDirect = $false

                    if ($CampaignNmSetMap.ContainsKey($campaignKey)) {
                        $isDirect = $CampaignNmSetMap[$campaignKey].ContainsKey([string]$nmId)
                    }

                    if (-not $StatsMap.ContainsKey($key)) {
                        $StatsMap[$key] = [pscustomobject]@{
                            AdvertId = $advertId
                            NmId = $nmId
                            IsDirect = $isDirect
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
# 12. АГРЕГАЦИЯ FULLSTATS ДО УРОВНЯ КАМПАНИИ
# ================================================================
#
# Ключевая бизнес-логика:
#   1 строка итогового CSV = 1 рекламная кампания.
#
# Артикул кампании определяется ТОЛЬКО по nm_settings самой кампании.
# Все метрики fullstats (включая ассоциативные конверсии по другим nmId)
# суммируются внутрь общей статистики этой рекламной кампании.
#
# Таким образом чужой nmId из ассоциативной конверсии больше не создаёт
# отдельную строку "кампания + чужой артикул".

$CampaignTotalsMap = @{}

foreach ($advertId in $FullStatsIds) {
    $campaignKey = [string]$advertId

    $CampaignTotalsMap[$campaignKey] = [pscustomobject]@{
        AdvertId = [Int64]$advertId
        Views = [Int64]0
        Clicks = [Int64]0
        Atbs = [Int64]0
        Orders = [Int64]0
        Canceled = [Int64]0
        Spend = [double]0
        Revenue = [double]0

        # Диагностика: какая часть заказов пришла по nmId,
        # которого нет в nm_settings самой кампании.
        AssocAtbs = [Int64]0
        AssocOrders = [Int64]0
        AssocCanceled = [Int64]0
        AssocRevenue = [double]0
    }
}

foreach ($stat in $StatsMap.Values) {
    $campaignKey = [string]$stat.AdvertId

    if (-not $CampaignTotalsMap.ContainsKey($campaignKey)) {
        continue
    }

    $total = $CampaignTotalsMap[$campaignKey]

    $total.Views += [Int64]$stat.Views
    $total.Clicks += [Int64]$stat.Clicks
    $total.Atbs += [Int64]$stat.Atbs
    $total.Orders += [Int64]$stat.Orders
    $total.Canceled += [Int64]$stat.Canceled
    $total.Spend += [double]$stat.Spend
    $total.Revenue += [double]$stat.Revenue

    if (-not $stat.IsDirect) {
        $total.AssocAtbs += [Int64]$stat.Atbs
        $total.AssocOrders += [Int64]$stat.Orders
        $total.AssocCanceled += [Int64]$stat.Canceled
        $total.AssocRevenue += [double]$stat.Revenue
    }
}


# ================================================================
# 13. ФОРМИРУЕМ 1 СТРОКУ НА 1 РЕКЛАМНУЮ КАМПАНИЮ
# ================================================================

$Rows = @()
$MultiCampaignRows = @()

foreach ($advertId in ($FullStatsIds | Sort-Object)) {
    $campaignKey = [string]$advertId

    if (-not $CampaignInfoMap.ContainsKey($campaignKey)) {
        continue
    }

    if (-not $SpendMap.ContainsKey($campaignKey)) {
        continue
    }

    $campaign = $CampaignInfoMap[$campaignKey]
    $updInfo = $SpendMap[$campaignKey]
    $total = $CampaignTotalsMap[$campaignKey]

    # Реальные рекламируемые артикулы берём только из nm_settings кампании.
    $nmIds = @(
        $CampaignNmSetMap[$campaignKey].Keys |
            ForEach-Object { [Int64]$_ } |
            Sort-Object -Unique
    )

    $nmCount = $nmIds.Count
    $article = ""
    $articleList = ""
    $campaignCheck = "OK"

    if ($nmCount -eq 1) {
        $article = [string]$nmIds[0]
        $articleList = [string]$nmIds[0]
    }
    elseif ($nmCount -gt 1) {
        $articleList = ($nmIds -join ", ")
        $campaignCheck = "ВНИМАНИЕ: несколько артикулов"

        Write-Log "ВНИМАНИЕ: кампания $advertId содержит $nmCount артикулов: $articleList"

        $MultiCampaignRows += [pscustomobject][ordered]@{
            "ID кампании"       = [Int64]$advertId
            "Название кампании" = $campaign.Name
            "Тип РК"            = (Get-BidTypeName $campaign.BidType)
            "Артикулов в РК"    = $nmCount
            "Артикулы"          = $articleList
            "Статус"            = (Get-StatusName $campaign.Status)
            "Проверка"          = "Требует разбиения: в одной РК несколько артикулов"
        }
    }
    else {
        $campaignCheck = "ВНИМАНИЕ: артикул не найден в nm_settings"
        Write-Log "ВНИМАНИЕ: у кампании $advertId не найден артикул в nm_settings."
    }

    # Наименование товара пытаемся взять из прямой строки fullstats.
    # Дополнительных API-запросов ради названия не делаем.
    $productName = ""

    if ($nmCount -eq 1) {
        $directKey = "$advertId|$($nmIds[0])"

        if ($StatsMap.ContainsKey($directKey)) {
            $productName = [string]$StatsMap[$directKey].Name
        }
    }

    $campaignName = $updInfo.CampaignName
    if (-not [string]::IsNullOrWhiteSpace($campaign.Name)) {
        $campaignName = $campaign.Name
    }

    $bidTypeName = Get-BidTypeName $campaign.BidType

    $paymentType = $updInfo.PaymentType
    if (-not [string]::IsNullOrWhiteSpace($campaign.PaymentType)) {
        $paymentType = $campaign.PaymentType
    }

    $statusName = Get-StatusName $updInfo.Status
    if ($null -ne $campaign.Status) {
        $statusName = Get-StatusName $campaign.Status
    }

    # Все показатели ниже — ОБЩИЕ по кампании.
    # Ассоциативные конверсии уже включены в Orders / Revenue / Atbs.
    $acceptedOrders = [Int64]$total.Orders - [Int64]$total.Canceled

    $ctr = if ($total.Views -gt 0) {
        [Math]::Round(($total.Clicks / $total.Views) * 100, 2)
    }
    else { 0 }

    $cr = if ($total.Clicks -gt 0) {
        [Math]::Round(($total.Orders / $total.Clicks) * 100, 2)
    }
    else { 0 }

    $cpc = if ($total.Clicks -gt 0) {
        [Math]::Round($total.Spend / $total.Clicks, 2)
    }
    else { 0 }

    $cpm = if ($total.Views -gt 0) {
        [Math]::Round(($total.Spend / $total.Views) * 1000, 2)
    }
    else { 0 }

    $cpo = if ($total.Orders -gt 0) {
        [Math]::Round($total.Spend / $total.Orders, 2)
    }
    else { 0 }

    $drr = if ($total.Revenue -gt 0) {
        [Math]::Round(($total.Spend / $total.Revenue) * 100, 2)
    }
    else { 0 }

    $Rows += [pscustomobject][ordered]@{
        "Название кампании"          = $campaignName
        "Раздел (тип РК)"            = $bidTypeName
        "ID кампании"                = [Int64]$advertId
        "Артикул"                    = $article
        "Артикулы РК"                = $articleList
        "Количество артикулов в РК"  = $nmCount
        "Проверка РК"                = $campaignCheck
        "Наименование"               = $productName
        "Дата создания"              = $campaign.Created
        "Статус"                     = $statusName
        "Модель оплаты"              = $paymentType
        "Период с"                   = $BeginDate
        "Период по"                  = $EndDate

        # Метрики кампании, включая ассоциативные конверсии.
        "Показы"                     = $total.Views
        "Клики"                      = $total.Clicks
        "Расход рекламы"             = [Math]::Round($total.Spend, 2)
        "ДРР Реклама (acc.)"         = $drr
        "CTR"                        = $ctr
        "Корзины (acc.)"             = $total.Atbs
        "Заказы (acc.)"              = $total.Orders
        "Принятые заказы"            = $acceptedOrders
        "Отмены заказов"             = $total.Canceled
        "Выручка (acc.)"             = [Math]::Round($total.Revenue, 2)
        "CR (acc.)"                  = $cr
        "CPC"                        = $cpc
        "CPM"                        = $cpm
        "CPO"                        = $cpo

        # Контрольные поля рабочей версии.
        # Они показывают, какая часть итогов пришла от ассоциативных nmId.
        "Ассоц. заказы (контроль)"   = $total.AssocOrders
        "Ассоц. отмены (контроль)"   = $total.AssocCanceled
        "Ассоц. выручка (контроль)"  = [Math]::Round($total.AssocRevenue, 2)

        "Расход кампании 7д upd"     = [Math]::Round($updInfo.Spend, 2)
    }
}

Write-Log "Итоговых строк: $($Rows.Count) (1 строка = 1 кампания)."
Write-Log "Кампаний с несколькими артикулами: $($MultiCampaignRows.Count)."


# ================================================================
# 14. ОСНОВНОЙ CSV
# ================================================================

$Headers = @(
    "Название кампании",
    "Раздел (тип РК)",
    "ID кампании",
    "Артикул",
    "Артикулы РК",
    "Количество артикулов в РК",
    "Проверка РК",
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
    "Принятые заказы",
    "Отмены заказов",
    "Выручка (acc.)",
    "CR (acc.)",
    "CPC",
    "CPM",
    "CPO",
    "Ассоц. заказы (контроль)",
    "Ассоц. отмены (контроль)",
    "Ассоц. выручка (контроль)",
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
# 15. ОТДЕЛЬНЫЙ СПИСОК КАМПАНИЙ С НЕСКОЛЬКИМИ АРТИКУЛАМИ
# ================================================================

$MultiHeaders = @(
    "ID кампании",
    "Название кампании",
    "Тип РК",
    "Артикулов в РК",
    "Артикулы",
    "Статус",
    "Проверка"
)

if ($MultiCampaignRows.Count -gt 0) {
    $multiCsvLines = $MultiCampaignRows |
        Select-Object $MultiHeaders |
        ConvertTo-Csv -NoTypeInformation -Delimiter ";"
}
else {
    $emptyMulti = [ordered]@{}

    foreach ($header in $MultiHeaders) {
        $emptyMulti[$header] = ""
    }

    $multiTempCsv = [pscustomobject]$emptyMulti |
        ConvertTo-Csv -NoTypeInformation -Delimiter ";"

    $multiCsvLines = @($multiTempCsv[0])
}

[System.IO.File]::WriteAllLines(
    $MultiTempPath,
    $multiCsvLines,
    $utf8Bom
)

if (Test-Path -LiteralPath $MultiPath) {
    Remove-Item -LiteralPath $MultiPath -Force -ErrorAction SilentlyContinue
}

Move-Item -LiteralPath $MultiTempPath -Destination $MultiPath -Force


# ================================================================
# 16. БЕЗОПАСНАЯ ЗАМЕНА ОСНОВНОГО CSV
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
Write-Log "Мультиартикульные РК: $MultiPath"
Write-Log "Период: $BeginDate - $EndDate."
Write-Log "Кампаний с затратами: $($SpendMap.Count)."
Write-Log "Кампаний с фактическим расходом > 0 руб.: $($QualifiedIds.Count)."
Write-Log "Кампаний отправлено в fullstats: $($FullStatsIds.Count)."
Write-Log "Строк итогового отчёта: $($Rows.Count)."
Write-Log "Кампаний с несколькими артикулами: $($MultiCampaignRows.Count)."

Write-Host ""
Write-Host "ОБНОВЛЕНИЕ РЕКЛАМЫ ЗАВЕРШЕНО" -ForegroundColor Green
Write-Host "Период: $BeginDate - $EndDate"
Write-Host "Одна строка итогового CSV = одна рекламная кампания."
Write-Host "Ассоциативные конверсии включены в общие показатели кампании."
Write-Host "Кампаний с несколькими артикулами: $($MultiCampaignRows.Count)"
Write-Host "Запросов fullstats: $($FullStatsBatches.Count)"
Write-Host "Файл: $OutputPath"
Write-Host "Мультиартикульные РК: $MultiPath"
