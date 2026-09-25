# ================================================================
# WB РЕКЛАМА — ДИАГНОСТИКА ЛИМИТОВ API
#
# Делает ОДИН запрос к:
#   GET https://advert-api.wildberries.ru/adv/v1/upd
#
# Цель:
#   увидеть фактический HTTP-статус и rate-limit заголовки WB:
#   X-Ratelimit-Limit
#   X-Ratelimit-Remaining
#   X-Ratelimit-Reset
#   X-Ratelimit-Retry
#   Retry-After
#
# ВАЖНО:
#   - повторных запросов НЕТ;
#   - токен в лог НЕ записывается;
#   - используется тот же wb_adv_token.txt.
# ================================================================

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TokenPath = Join-Path $ScriptDir "wb_adv_token.txt"
$LogPath   = Join-Path $ScriptDir "WB_Реклама_API_Диагностика.txt"

# Последние 7 полных дней, заканчивая вчера.
$DateTo   = (Get-Date).Date.AddDays(-1)
$DateFrom = $DateTo.AddDays(-6)

$BeginDate = $DateFrom.ToString("yyyy-MM-dd")
$EndDate   = $DateTo.ToString("yyyy-MM-dd")

function Write-Both {
    param([string]$Text = "")

    Write-Host $Text
    Add-Content -LiteralPath $LogPath -Value $Text -Encoding UTF8
}

function Get-HeaderValue {
    param(
        [System.Net.WebHeaderCollection]$Headers,
        [string]$Name
    )

    if ($null -eq $Headers) {
        return ""
    }

    $value = $Headers[$Name]

    if ($null -eq $value) {
        return ""
    }

    return [string]$value
}

# ------------------------------------------------
# 1. Token
# ------------------------------------------------
if (-not (Test-Path -LiteralPath $TokenPath)) {
    throw "Не найден wb_adv_token.txt рядом со скриптом."
}

$Token = (Get-Content -LiteralPath $TokenPath -Raw).Trim()

if ([string]::IsNullOrWhiteSpace($Token) -or ($Token -match "ВСТАВ|TOKEN|ТОКЕН")) {
    throw "wb_adv_token.txt пустой или содержит шаблон."
}

# Старый диагностический лог удаляем.
if (Test-Path -LiteralPath $LogPath) {
    Remove-Item -LiteralPath $LogPath -Force
}

$Uri = "https://advert-api.wildberries.ru/adv/v1/upd?from=$BeginDate&to=$EndDate"

Write-Both "============================================================"
Write-Both "WB Promotion API — диагностика rate limit"
Write-Both ("Время запуска: " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
Write-Both ("Период: " + $BeginDate + " - " + $EndDate)
Write-Both ("URL: " + $Uri)
Write-Both ""
Write-Both "Выполняю РОВНО ОДИН запрос..."
Write-Both ""

$response = $null

try {
    $request = [System.Net.HttpWebRequest]::Create($Uri)
    $request.Method = "GET"
    $request.Timeout = 120000
    $request.ReadWriteTimeout = 120000
    $request.Accept = "application/json"
    $request.Headers["Authorization"] = "Bearer $Token"
    $request.AutomaticDecompression = `
        [System.Net.DecompressionMethods]::GZip -bor `
        [System.Net.DecompressionMethods]::Deflate

    $response = [System.Net.HttpWebResponse]$request.GetResponse()

    $statusCode = [int]$response.StatusCode
    $statusText = [string]$response.StatusDescription
    $headers = $response.Headers

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

    $body = [System.Text.Encoding]::UTF8.GetString($rawBytes)

    Write-Both ("HTTP статус: " + $statusCode + " " + $statusText)
    Write-Both ""
    Write-Both "RATE-LIMIT ЗАГОЛОВКИ:"
    Write-Both ("X-Ratelimit-Limit:     " + (Get-HeaderValue $headers "X-Ratelimit-Limit"))
    Write-Both ("X-Ratelimit-Remaining: " + (Get-HeaderValue $headers "X-Ratelimit-Remaining"))
    Write-Both ("X-Ratelimit-Reset:     " + (Get-HeaderValue $headers "X-Ratelimit-Reset"))
    Write-Both ("X-Ratelimit-Retry:     " + (Get-HeaderValue $headers "X-Ratelimit-Retry"))
    Write-Both ("Retry-After:           " + (Get-HeaderValue $headers "Retry-After"))
    Write-Both ("Date:                  " + (Get-HeaderValue $headers "Date"))
    Write-Both ""
    Write-Both "ВСЕ ЗАГОЛОВКИ ОТВЕТА:"

    foreach ($key in $headers.AllKeys) {
        Write-Both ($key + ": " + $headers[$key])
    }

    Write-Both ""

    if ([string]::IsNullOrWhiteSpace($body)) {
        Write-Both "Тело ответа пустое."
    }
    else {
        $previewLength = [Math]::Min(1200, $body.Length)
        Write-Both "Начало тела ответа (до 1200 символов):"
        Write-Both $body.Substring(0, $previewLength)
    }

    Write-Both ""
    Write-Both "РЕЗУЛЬТАТ: запрос успешен."
}
catch [System.Net.WebException] {
    $webResponse = $null
    $statusCode = ""
    $statusText = ""
    $headers = $null
    $body = ""

    try {
        $webResponse = [System.Net.HttpWebResponse]$_.Exception.Response
    }
    catch {
    }

    if ($null -ne $webResponse) {
        try {
            $statusCode = [int]$webResponse.StatusCode
            $statusText = [string]$webResponse.StatusDescription
            $headers = $webResponse.Headers
        }
        catch {
        }

        try {
            $stream = $webResponse.GetResponseStream()
            $memory = New-Object System.IO.MemoryStream

            try {
                $stream.CopyTo($memory)
                $rawBytes = $memory.ToArray()
                $body = [System.Text.Encoding]::UTF8.GetString($rawBytes)
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

    Write-Both ("HTTP статус: " + $statusCode + " " + $statusText)
    Write-Both ""
    Write-Both "RATE-LIMIT ЗАГОЛОВКИ:"
    Write-Both ("X-Ratelimit-Limit:     " + (Get-HeaderValue $headers "X-Ratelimit-Limit"))
    Write-Both ("X-Ratelimit-Remaining: " + (Get-HeaderValue $headers "X-Ratelimit-Remaining"))
    Write-Both ("X-Ratelimit-Reset:     " + (Get-HeaderValue $headers "X-Ratelimit-Reset"))
    Write-Both ("X-Ratelimit-Retry:     " + (Get-HeaderValue $headers "X-Ratelimit-Retry"))
    Write-Both ("Retry-After:           " + (Get-HeaderValue $headers "Retry-After"))
    Write-Both ("Date:                  " + (Get-HeaderValue $headers "Date"))
    Write-Both ""
    Write-Both "ВСЕ ЗАГОЛОВКИ ОТВЕТА:"

    if ($null -ne $headers) {
        foreach ($key in $headers.AllKeys) {
            Write-Both ($key + ": " + $headers[$key])
        }
    }
    else {
        Write-Both "(заголовки не получены)"
    }

    Write-Both ""

    if ([string]::IsNullOrWhiteSpace($body)) {
        Write-Both ("Тело ошибки отсутствует. Сообщение Windows: " + $_.Exception.Message)
    }
    else {
        Write-Both "Тело ошибки WB:"
        Write-Both $body
    }

    Write-Both ""
    Write-Both "РЕЗУЛЬТАТ: запрос НЕ выполнен."
}
finally {
    if ($null -ne $response) {
        $response.Dispose()
    }
}

Write-Both ""
Write-Both ("Диагностический файл: " + $LogPath)
Write-Host ""
Write-Host "Диагностика завершена. Пришли мне скрин окна ИЛИ содержимое WB_Реклама_API_Диагностика.txt." -ForegroundColor Green
