# ============================================================
#  SharePoint File Downloader
# ============================================================
# Usage:
#   .\SharePoint-Downloader.ps1
#   .\SharePoint-Downloader.ps1 -Url "https://..." -Destination "C:\Downloads"
#   .\SharePoint-Downloader.ps1 -SourceFile "sources.txt" -Destination "C:\Downloads"
#
# sources.txt format:
#   One URL per line
#   Lines starting with # are treated as comments
#   Empty lines are skipped
# ============================================================

param(
    [string]$Url,
    [string]$Destination,
    [string]$SourceFile
)

function Format-FileSize {
    param([long]$Bytes)
    if    ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    elseif ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    else                    { return "$Bytes B" }
}

function Resolve-FileName {
    param([string]$Url)
    try   { $response = Invoke-WebRequest -Uri $Url -Method Head -ErrorAction Stop }
    catch { $response = Invoke-WebRequest -Uri $Url -Method Get  -ErrorAction Stop }

    $contentDisp = $response.Headers["Content-Disposition"]
    $fileName    = $null

    if ($contentDisp) {
        if    ($contentDisp -match "filename\*=UTF-8''([^;]+)")  { $fileName = [uri]::UnescapeDataString($Matches[1]) }
        elseif ($contentDisp -match 'filename="([^"]+)"')        { $fileName = $Matches[1].Trim() }
        elseif ($contentDisp -match 'filename=([^;]+)')          { $fileName = $Matches[1].Trim() }
    }
    if ($fileName) { return $fileName }
    return "download.bin"
}

# ─── Header ────────────────────────────────────────────────
Write-Host ""
Write-Host "  ┌────────────────────────────────┐" -ForegroundColor Cyan
Write-Host "  │   SharePoint File Downloader   │" -ForegroundColor Cyan
Write-Host "  └────────────────────────────────┘" -ForegroundColor Cyan
Write-Host ""

# ─── Collect URLs ──────────────────────────────────────────
$urls = @()

if ($SourceFile) {
    if (!(Test-Path $SourceFile)) {
        Write-Host "  ✖  Source file not found: $SourceFile" -ForegroundColor Red
        Write-Host ""
        exit 1
    }
    $urls = Get-Content $SourceFile |
            Where-Object { $_.Trim() -ne '' -and -not $_.Trim().StartsWith('#') } |
            ForEach-Object { $_.Trim() }

    if ($urls.Count -eq 0) {
        Write-Host "  ✖  No valid URLs found in $SourceFile" -ForegroundColor Red
        Write-Host ""
        exit 1
    }
    Write-Host "  Loaded $($urls.Count) URL(s) from $(Split-Path $SourceFile -Leaf)" -ForegroundColor DarkGray
    Write-Host ""

} elseif ($Url) {
    $urls = @($Url)

} else {
    Write-Host "  Paste your SharePoint link:" -ForegroundColor White
    $urls = @(Read-Host "  >")
    Write-Host ""
}

# ─── Collect Destination ───────────────────────────────────
if (-not $Destination) {
    $defaultPath = Join-Path $env:USERPROFILE "Downloads"
    Write-Host "  Save location:" -ForegroundColor White
    Write-Host "  (Press Enter for: $defaultPath)" -ForegroundColor DarkGray
    $userPath = Read-Host "  >"
    $Destination = if ($userPath) { $userPath } else { $defaultPath }
    Write-Host ""
}

if (!(Test-Path $Destination)) {
    New-Item -ItemType Directory -Path $Destination | Out-Null
}

# ─── Download Loop ─────────────────────────────────────────
$total   = $urls.Count
$success = 0
$failed  = 0

for ($i = 0; $i -lt $total; $i++) {
    $currentUrl = $urls[$i].Trim()
    if ($currentUrl -eq '') { continue }

    try {
        $fileName        = Resolve-FileName -Url $currentUrl
        $destinationFile = Join-Path $Destination $fileName

        # Header per file
        if ($total -gt 1) {
            Write-Host "  ─── [$($i + 1) / $total] ──────────────────────────" -ForegroundColor DarkGray
        } else {
            Write-Host "  ─────────────────────────────────────" -ForegroundColor DarkGray
        }
        Write-Host -NoNewline "  File      " -ForegroundColor DarkGray
        Write-Host $fileName                  -ForegroundColor White
        Write-Host -NoNewline "  Folder    " -ForegroundColor DarkGray
        Write-Host $Destination               -ForegroundColor DarkGray
        Write-Host "  Downloading..." -ForegroundColor DarkGray

        # Download
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        Invoke-WebRequest -Uri $currentUrl -OutFile $destinationFile -UseBasicParsing
        $timer.Stop()

        $size = Format-FileSize (Get-Item $destinationFile).Length
        $secs = $timer.Elapsed.TotalSeconds
        $time = if ($secs -lt 1) { "< 1 sec" } else { "{0:N1} sec" -f $secs }

        Write-Host "  ✔  $fileName  ·  $size  ·  $time" -ForegroundColor Green
        Write-Host ""
        $success++

    } catch {
        Write-Host "  ✖  Failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        $failed++
    }
}

# ─── Summary (multi-file only) ─────────────────────────────
if ($total -gt 1) {
    Write-Host "  ═════════════════════════════════════" -ForegroundColor DarkGray
    if ($failed -eq 0) {
        Write-Host "  ✔  All $total file(s) downloaded successfully." -ForegroundColor Green
    } else {
        Write-Host "  $success of $total file(s) downloaded successfully." -ForegroundColor White
        Write-Host "  ✖  $failed file(s) failed."                         -ForegroundColor Red
    }
    Write-Host "     Saved to $Destination" -ForegroundColor DarkGray
    Write-Host ""
}
