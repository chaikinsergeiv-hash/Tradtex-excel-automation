# ================================================================
# WB РЕКЛАМА — RAW DUMP API v2
#
# Цель:
#   посмотреть, какие поля реально возвращает WB API,
#   прежде чем решать, как строить итоговый CSV.
#
# Скрипт сохраняет:
#   01_UPD_RAW.json              — ответ /adv/v1/upd "как есть"
#   01_UPD_FLAT.csv              — тот же ответ в плоском CSV
#
#   02_ADVERTS_RAW.json          — ответ /api/advert/v2/adverts "как есть"
#   02_ADVERTS_TOP.csv           — верхний уровень кампаний
#
#   03_FULLSTATS_SAMPLE_RAW.json — пример ответа /adv/v3/fullstats
#   03_FULLSTATS_SAMPLE_FLAT.csv — развёртка campaign/day/app/nm
#
# ВАЖНО:
#   WB API возвращает JSON, а не CSV.
#   Поэтому именно .json — настоящий "сырой" формат API.
#
# Нужен wb_adv_token.txt категории "Продвижение".
# ================================================================

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# -----------------------------
# 1. НАСТРОЙКИ
# -----------------------------

$MinCampaignSpend = 100

# Для просмотра структуры fullstats нам не нужны все 260 кампаний.
# Берём небольшой пример. Если нужно — увеличь число.
$FullStatsSampleCampaigns = 10

# -----------------------------
# 2. ПУТИ
# -----------------------------

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TokenPath = Join-Path $ScriptDir "wb_adv_token.txt"

$UpdRawPath       = Join-Path $ScriptDir "01_UPD_RAW.json"
$UpdCsvPath       = Join-Path $ScriptDir "01_UPD_FLAT.csv"

$AdvertsRawPath   = Join-Path $ScriptDir "02_ADVERTS_RAW.json"
$AdvertsCsvPath   = Join-Path $ScriptDir "02_ADVERTS_TOP.csv"

$FullRawPath      = Join-Path $ScriptDir "03_FULLSTATS_SAMPLE_RAW.json"
$FullCsvPath      = Join-Path $ScriptDir "03_FULLSTATS_SAMPLE_FLAT.csv"

# -----------------------------
# 3. ПЕРИОД
# -----------------------------

$DateTo   = (Get-Date).Date.AddDays(-1)
$DateFrom = $DateTo.AddDays(-6)

$BeginDate = $DateFrom.ToString("yyyy-MM-dd")
$EndDate   = $DateTo.ToString("yyyy-MM-dd")

# -----------------------------
# 4. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# -----------------------------

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

function Save-Utf8Bom {
    param(
        [string]$Path,
        [string[]]$Lines
    )

    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllLines($Path, $Lines, $utf8Bom)
}

function Invoke-WBRawGet {
    param(
        [Parameter(Mandatory=$true)][string]$BaseUri,
        [hashtable]$Query = @{}
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

    $request = [System.Net.HttpWebRequest]::Create($Uri)
    $request.Method = "GET"
    $request.Timeout = 180000
    $request.ReadWriteTimeout = 180000
    $request.Accept = "application/json"
    $request.Headers["Authorization"] = "Bearer $script:Token"
    $request.AutomaticDecompression = `
        [System.Net.DecompressionMethods]::GZip -bor `
        [System.Net.DecompressionMethods]::Deflate

    try {
        $response = [System.Net.HttpWebResponse]$request.GetResponse()

        try {
            $stream = $response.GetResponseStream()
            $memory = New-Object System.IO.MemoryStream

            try {
                $stream.CopyTo($memory)
                $bytes = $memory.ToArray()
            }
            finally {
                if ($null -ne $stream) {
                    $stream.Dispose()
                }

                $memory.Dispose()
            }

            $text = [System.Text.Encoding]::UTF8.GetString($bytes)

            return [pscustomobject]@{
                Uri        = $Uri
                StatusCode = [int]$response.StatusCode
                RawText    = $text
                Json       = if ([string]::IsNullOrWhiteSpace($text)) { $null } else { $text | ConvertFrom-Json }
            }
        }
        finally {
            $response.Dispose()
        }
    }
    catch [System.Net.WebException] {
        $status = ""
        $body = ""

        if ($null -ne $_.Exception.Response) {
            $webResponse = [System.Net.HttpWebResponse]$_.Exception.Response
            $status = [int]$webResponse.StatusCode

            try {
                $stream = $webResponse.GetResponseStream()
                $memory = New-Object System.IO.MemoryStream

                try {
                    $stream.CopyTo($memory)
                    $bytes = $memory.ToArray()
                    $body = [System.Text.Encoding]::UTF8.GetString($bytes)
                }
                finally {
                    if ($null -ne $stream) {
                        $stream.Dispose()
                    }

                    $memory.Dispose()
                }
            }
            catch {
            }
        }

        throw "WB API HTTP $status`nURL: $Uri`nОтвет: $body"
    }
}

# -----------------------------
# 5. ТОКЕН
# -----------------------------

if (-not (Test-Path -LiteralPath $TokenPath)) {
    throw "Не найден wb_adv_token.txt рядом со скриптом."
}

$script:Token = (Get-Content -LiteralPath $TokenPath -Raw).Trim()

if ([string]::IsNullOrWhiteSpace($script:Token)) {
    throw "wb_adv_token.txt пустой."
}

Write-Host "Период: $BeginDate - $EndDate"
Write-Host ""

# ================================================================
# 6. /adv/v1/upd
# ================================================================

Write-Host "1/3 Получаю /adv/v1/upd..."

$upd = Invoke-WBRawGet `
    -BaseUri "https://advert-api.wildberries.ru/adv/v1/upd" `
    -Query @{
        from = $BeginDate
        to   = $EndDate
    }

# RAW JSON сохраняем максимально близко к тому, что прислал WB.
Save-Utf8Bom -Path $UpdRawPath -Lines @($upd.RawText)

$updRows = @($upd.Json)

# CSV — верхний плоский массив ответа /upd.
$updCsvLines = $updRows |
    ConvertTo-Csv -NoTypeInformation -Delimiter ";"

Save-Utf8Bom -Path $UpdCsvPath -Lines $updCsvLines

Write-Host "   RAW JSON: $UpdRawPath"
Write-Host "   CSV:      $UpdCsvPath"

# ================================================================
# 7. СЧИТАЕМ РАСХОД И ВЫБИРАЕМ SAMPLE ID
# ================================================================

$spendMap = @{}

foreach ($row in $updRows) {
    $idRaw = Get-Field $row "advertId" $null

    if ($null -eq $idRaw) {
        continue
    }

    $id = [Int64]$idRaw
    $sum = [double](Get-Field $row "updSum" 0)
    $key = [string]$id

    if (-not $spendMap.ContainsKey($key)) {
        $spendMap[$key] = [double]0
    }

    $spendMap[$key] += $sum
}

$qualifiedIds = @(
    $spendMap.Keys |
        Where-Object { $spendMap[$_] -ge $MinCampaignSpend } |
        ForEach-Object { [Int64]$_ } |
        Sort-Object
)

Write-Host ""
Write-Host "Кампаний с расходом >= $MinCampaignSpend руб.: $($qualifiedIds.Count)"

# ================================================================
# 8. /api/advert/v2/adverts?statuses=7,9,11
# ================================================================

Write-Host ""
Write-Host "2/3 Получаю /api/advert/v2/adverts?statuses=7,9,11..."

$adverts = Invoke-WBRawGet `
    -BaseUri "https://advert-api.wildberries.ru/api/advert/v2/adverts" `
    -Query @{
        statuses = "7,9,11"
    }

Save-Utf8Bom -Path $AdvertsRawPath -Lines @($adverts.RawText)

$adItems = @()

if ($null -ne $adverts.Json.PSObject.Properties["adverts"]) {
    $adItems = @($adverts.Json.adverts)
}
else {
    $adItems = @($adverts.Json)
}

# TOP CSV:
# вложенные settings / timestamps / nm_settings остаются JSON-текстом,
# чтобы ничего не потерять.
$adTop = @()

foreach ($campaign in $adItems) {
    $flat = [ordered]@{}

    foreach ($prop in $campaign.PSObject.Properties) {
        $value = $prop.Value

        if ($null -eq $value) {
            $flat[$prop.Name] = ""
        }
        elseif (($value -is [System.Management.Automation.PSCustomObject]) -or (($value -is [System.Collections.IEnumerable]) -and -not ($value -is [string]))) {
            $flat[$prop.Name] = ($value | ConvertTo-Json -Depth 20 -Compress)
        }
        else {
            $flat[$prop.Name] = [string]$value
        }
    }

    $adTop += [pscustomobject]$flat
}

$adCsvLines = $adTop |
    ConvertTo-Csv -NoTypeInformation -Delimiter ";"

Save-Utf8Bom -Path $AdvertsCsvPath -Lines $adCsvLines

Write-Host "   RAW JSON: $AdvertsRawPath"
Write-Host "   CSV:      $AdvertsCsvPath"

# ================================================================
# 9. /adv/v3/fullstats — SAMPLE
# ================================================================

$sampleIds = @($qualifiedIds | Select-Object -First $FullStatsSampleCampaigns)

if ($sampleIds.Count -eq 0) {
    Write-Host ""
    Write-Host "3/3 Fullstats sample пропущен: нет кампаний после фильтра."
}
else {
    Write-Host ""
    Write-Host "3/3 Получаю fullstats SAMPLE по $($sampleIds.Count) кампаниям..."

    $full = Invoke-WBRawGet `
        -BaseUri "https://advert-api.wildberries.ru/adv/v3/fullstats" `
        -Query @{
            ids       = ($sampleIds -join ",")
            beginDate = $BeginDate
            endDate   = $EndDate
        }

    Save-Utf8Bom -Path $FullRawPath -Lines @($full.RawText)

    # Разворачиваем nested JSON:
    # campaign -> days -> apps -> nms
    # Один ряд CSV = один товар в одной app/day внутри кампании.
    #
    # При этом сохраняем ВСЕ найденные поля nm динамически.

    $flatRows = @()

    foreach ($campaign in @($full.Json)) {
        $advertId = Get-Field $campaign "advertId" ""

        foreach ($day in @((Get-Field $campaign "days" @()))) {
            $date = Get-Field $day "date" ""

            foreach ($app in @((Get-Field $day "apps" @()))) {
                $appType = Get-Field $app "appType" ""

                foreach ($nm in @((Get-Field $app "nms" @()))) {
                    $flat = [ordered]@{
                        advertId = $advertId
                        date     = $date
                        appType  = $appType
                    }

                    foreach ($prop in $nm.PSObject.Properties) {
                        $value = $prop.Value

                        if ($null -eq $value) {
                            $flat[$prop.Name] = ""
                        }
                        elseif (($value -is [System.Management.Automation.PSCustomObject]) -or (($value -is [System.Collections.IEnumerable]) -and -not ($value -is [string]))) {
                            $flat[$prop.Name] = ($value | ConvertTo-Json -Depth 20 -Compress)
                        }
                        else {
                            $flat[$prop.Name] = $value
                        }
                    }

                    $flatRows += [pscustomobject]$flat
                }
            }
        }
    }

    if ($flatRows.Count -gt 0) {
        $fullCsvLines = $flatRows |
            ConvertTo-Csv -NoTypeInformation -Delimiter ";"

        Save-Utf8Bom -Path $FullCsvPath -Lines $fullCsvLines
    }
    else {
        Save-Utf8Bom -Path $FullCsvPath -Lines @("advertId;date;appType")
    }

    Write-Host "   RAW JSON: $FullRawPath"
    Write-Host "   CSV:      $FullCsvPath"
}

Write-Host ""
Write-Host "ГОТОВО." -ForegroundColor Green
Write-Host ""
Write-Host "Сначала открой:"
Write-Host "  01_UPD_FLAT.csv"
Write-Host "  02_ADVERTS_TOP.csv"
Write-Host "  03_FULLSTATS_SAMPLE_FLAT.csv"
Write-Host ""
Write-Host "А если нужно увидеть настоящий ответ WB без преобразований — открывай соответствующий *_RAW.json."
