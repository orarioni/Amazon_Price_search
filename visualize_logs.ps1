# Log visualization/diagnostic tool (English version to avoid encoding issues)
param(
    [switch]$Realtime,
    [switch]$Summary,
    [switch]$Errors,
    [int]$LastLines = 50
)

$logPath = Join-Path $PSScriptRoot 'logs\run.log'
$metricsPath = Join-Path $PSScriptRoot 'logs\metrics.jsonl'

function Show-LogSummary {
    Write-Host "==== Log Summary ====" -ForegroundColor Cyan
    Write-Host ""
    if (Test-Path $logPath) {
        $content = Get-Content $logPath -Encoding UTF8 | Out-String
        $warnCount = ([regex]::Matches($content, '\[WARN\]')).Count
        $infoCount = ([regex]::Matches($content, '\[INFO\]')).Count
        $errorCount = ([regex]::Matches($content, '\[ERROR\]')).Count
        $http403Count = ([regex]::Matches($content, 'status=403')).Count
        Write-Host "Log stats:" -ForegroundColor Green
        Write-Host "  INFO: $infoCount"
        Write-Host "  WARN: $warnCount"
        Write-Host "  ERROR: $errorCount"
        Write-Host ""
        Write-Host "API errors:" -ForegroundColor Yellow
        Write-Host "  HTTP 403: $http403Count"
        Write-Host ""
        $lastLine = Get-Content $logPath -Encoding UTF8 -Tail 1
        if ($lastLine -and $lastLine -match '(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})') {
            Write-Host "Last log timestamp: $($matches[1])" -ForegroundColor Cyan
            $now = Get-Date
            try {
                $lastLogTime = [datetime]::ParseExact(($lastLine -split '\[')[0].Trim(), 'yyyy-MM-ddTHH:mm:ss.fffffffzzz', $null)
                $diff = $now - $lastLogTime
                Write-Host "Time since last log: $($diff.Minutes)m $($diff.Seconds)s" -ForegroundColor Yellow
            } catch {
                # ignore parse errors
            }
        } else {
            Write-Host "Last log timestamp: unknown" -ForegroundColor Cyan
        }
    }
}

function Show-Errors {
    Write-Host "==== Error Details ====" -ForegroundColor Red
    Write-Host ""
    if (Test-Path $logPath) {
        $lines = Get-Content $logPath -Encoding UTF8
        $errorLines = $lines | Where-Object { $_ -match '\[WARN\].*status=403' }
        Write-Host "[Recent HTTP 403 errors]" -ForegroundColor Yellow
        $errorLines | Select-Object -Last 5 | ForEach-Object { Write-Host "$_" }
        Write-Host ""
        Write-Host "Possible causes for 403 errors:" -ForegroundColor Yellow
        Write-Host "  - LWA token expired"
        Write-Host "  - Catalog API permission issue"
        Write-Host "  - retry limit reached"
    }
}

function Show-Metrics {
    Write-Host "==== Execution Metrics ====" -ForegroundColor Cyan
    Write-Host ""
    if (Test-Path $metricsPath) {
        $metrics = Get-Content $metricsPath -Encoding UTF8 | ConvertFrom-Json
        if ($metrics) {
            $latest = $metrics | Sort-Object { [datetime]$_.ts } | Select-Object -Last 1
            Write-Host "Latest run:" -ForegroundColor Green
            Write-Host "  Timestamp: $($latest.ts)"
            Write-Host "  Input rows: $($latest.input_rows)"
            Write-Host "  Unique ASIN: $($latest.unique_asin)"
            Write-Host "  API calls: $($latest.api_total_calls)"
            Write-Host "  Retry count: $($latest.retry_count)"
            Write-Host "  HTTP 429: $($latest.http429_count)"
            Write-Host "  Wait total sec: $($latest.total_wait_sec)"
        }
    }
}

function Show-RealtimeTail {
    Write-Host "==== Realtime log tail ====" -ForegroundColor Cyan
    Write-Host "Ctrl+C to exit" -ForegroundColor Gray
    Write-Host ""
    if (Test-Path $logPath) {
        Get-Content $logPath -Encoding UTF8 -Wait -Tail 10
    }
}

function Show-LastLines {
    param([int]$Lines = 50)
    Write-Host "==== Last $Lines log lines ====" -ForegroundColor Cyan
    Write-Host ""
    if (Test-Path $logPath) {
        $content = Get-Content $logPath -Encoding UTF8 -Tail $Lines
        $content | ForEach-Object {
            if ($_ -match '\[ERROR\]') { Write-Host $_ -ForegroundColor Red }
            elseif ($_ -match '\[WARN\]') { Write-Host $_ -ForegroundColor Yellow }
            elseif ($_ -match '\[INFO\]') { Write-Host $_ -ForegroundColor Green }
            else { Write-Host $_ }
        }
    }
}

# Main execution
if ($Realtime) {
    Show-RealtimeTail
} elseif ($Summary) {
    Show-LogSummary
    Write-Host ""
    Show-Metrics
} elseif ($Errors) {
    Show-Errors
} else {
    Show-LogSummary
    Write-Host ""
    Show-Metrics
    Write-Host ""
    Show-LastLines -Lines $LastLines
}