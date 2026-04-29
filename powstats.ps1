$ErrorActionPreference = "Stop"

# Configuration
$ConfigDir = "$env:APPDATA\powstats"
$CredFile  = "$ConfigDir\strava_credentials.xml"
$TokenFile = "$ConfigDir\auth_tokens.json"
$CacheFile = "$ConfigDir\activity_cache.json"
$RedirectUri = "http://localhost:9876/callback"

# OPTIONAL: Set to a season (e.g. "2025-2026") to process only that season's activities
$SeasonFilter = "2025-2026"  # Set to $null for all seasons

if (-not (Test-Path $ConfigDir)) {
    New-Item -ItemType Directory -Path $ConfigDir | Out-Null
}

# ── Functions ──────────────────────────────────────────────────────────────────

function Initialize-StravaAuth {
    Write-Host "`n=== Strava API Setup ===" -ForegroundColor Cyan
    Write-Host "You need to create a Strava API application first."
    Write-Host "1. Go to: https://www.strava.com/settings/api"
    Write-Host "2. Create an app with these settings:"
    Write-Host "   - Application Name: powstats (or whatever you want)"
    Write-Host "   - Category: Data Importer"
    Write-Host "   - Authorization Callback Domain: localhost`n"

    $ClientId     = Read-Host "Enter your Client ID"
    $ClientSecret = Read-Host "Enter your Client Secret" -AsSecureString

    $cred = New-Object System.Management.Automation.PSCredential($ClientId, $ClientSecret)
    $cred | Export-Clixml $CredFile

    Write-Host "`nCredentials saved. Starting OAuth flow..." -ForegroundColor Green

    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("$RedirectUri/")
    $listener.Start()

    $authUrl = "https://www.strava.com/oauth/authorize?client_id=$ClientId&response_type=code&redirect_uri=$RedirectUri&scope=activity:read_all"
    Write-Host "`nOpening browser for Strava authorization..."
    Start-Process $authUrl

    Write-Host "Waiting for authorization..." -ForegroundColor Yellow
    $context = $listener.GetContext()
    $code    = $context.Request.QueryString['code']

    $response = $context.Response
    $html     = "<html><body><h1>Authorization successful!</h1><p>You can close this window and return to PowerShell.</p></body></html>"
    $buffer   = [System.Text.Encoding]::UTF8.GetBytes($html)
    $response.ContentLength64 = $buffer.Length
    $response.OutputStream.Write($buffer, 0, $buffer.Length)
    $response.Close()
    $listener.Stop()

    if (-not $code) { throw "No authorization code received" }

    $ClientSecretPlain = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
    )

    $tokens = Invoke-RestMethod -Method Post https://www.strava.com/oauth/token -Body @{
        client_id     = $ClientId
        client_secret = $ClientSecretPlain
        code          = $code
        grant_type    = "authorization_code"
    }
    $tokens | ConvertTo-Json | Set-Content $TokenFile
    Write-Host "`nSetup complete! You can now run powstats.ps1" -ForegroundColor Green
}

function Import-TokenFile {
    if (Test-Path $TokenFile) { Get-Content $TokenFile | ConvertFrom-Json }
}

function Update-AccessToken($clientId, $clientSecret, $refreshToken) {
    $tokens = Invoke-RestMethod -Method Post https://www.strava.com/oauth/token -Body @{
        client_id     = $clientId
        client_secret = $clientSecret
        grant_type    = "refresh_token"
        refresh_token = $refreshToken
    }
    $tokens | ConvertTo-Json | Set-Content $TokenFile
    return $tokens
}

function Import-ActivityCache {
    if (Test-Path $CacheFile) {
        try {
            $cacheData = Get-Content $CacheFile | ConvertFrom-Json
            $cacheHash = @{}
            foreach ($item in $cacheData) {
                $cacheHash[[string]$item.activity_id] = $item
            }
            return $cacheHash
        }
        catch {
            Write-Host "Warning: Could not load cache, starting fresh" -ForegroundColor Yellow
        }
    }
    return @{}
}

function Export-ActivityCache($cache) {
    try {
        $cache.Values | ConvertTo-Json -Depth 10 | Set-Content $CacheFile
    }
    catch {
        Write-Host "Warning: Could not save cache - $_" -ForegroundColor Yellow
    }
}

function Test-ActivityStale($activity, $cache) {
    $activityId = [string]$activity.id
    if (-not $cache.ContainsKey($activityId)) { return $true }

    $cached = $cache[$activityId]
    if (-not $activity.updated_at -or -not $cached.updated_at) { return $true }

    try {
        return ([datetime]$activity.updated_at) -gt ([datetime]$cached.updated_at)
    }
    catch {
        return $true
    }
}

function Get-VerticalDescent($elevationData) {
    if (-not $elevationData -or $elevationData.Count -lt 2) { return 0 }

    $totalDescent = 0
    for ($i = 1; $i -lt $elevationData.Count; $i++) {
        $diff = $elevationData[$i] - $elevationData[$i-1]
        if ($diff -lt 0) { $totalDescent += [math]::Abs($diff) }
    }
    return $totalDescent
}

function Format-TimeHoursMinutes($seconds) {
    $hours   = [math]::Floor($seconds / 3600)
    $minutes = [math]::Floor(($seconds % 3600) / 60)
    return "${hours}h ${minutes}m"
}

function Get-ActivityDetails($activityId, $headers) {
    try {
        return Invoke-RestMethod "https://www.strava.com/api/v3/activities/$activityId" -Headers $headers
    }
    catch {
        if ($_.Exception.Response -and $_.Exception.Response.StatusCode.value__ -eq 429) {
            Write-Host "  Rate limit hit -- Strava allows 100 requests per 15 minutes. Wait and retry." -ForegroundColor Red
            throw
        }
        Write-Host "  Warning: Could not fetch details for activity $activityId - $_" -ForegroundColor Yellow
    }
    return $null
}

function Get-ActivityStreams($activityId, $headers) {
    try {
        $stream = Invoke-RestMethod "https://www.strava.com/api/v3/activities/$activityId/streams?keys=altitude&key_by_type=true" -Headers $headers
        if ($stream.altitude -and $stream.altitude.data) { return $stream.altitude.data }
    }
    catch {
        if ($_.Exception.Response -and $_.Exception.Response.StatusCode.value__ -eq 429) {
            Write-Host "  Rate limit hit -- Strava allows 100 requests per 15 minutes. Wait and retry." -ForegroundColor Red
            throw
        }
        Write-Host "  Warning: Could not fetch streams for activity $activityId - $_" -ForegroundColor Yellow
    }
    return $null
}

function Get-Season($d) {
    if ($d.Month -ge 7) { "$($d.Year)-$($d.Year+1)" }
    else { "$($d.Year-1)-$($d.Year)" }
}

function Get-ActivityStats($group, $label) {
    $totalMovingTime  = ($group | Measure-Object -Property moving_time  -Sum).Sum
    $totalElapsedTime = ($group | Measure-Object -Property elapsed_time -Sum).Sum
    $activityCount    = $group.Count
    $daysCount        = ($group.day | Sort-Object -Unique).Count
    $totalRuns        = ($group | Measure-Object -Property run_count          -Sum).Sum
    $totalDistance    = ($group | Measure-Object -Property distance           -Sum).Sum
    $totalVertical    = ($group | Measure-Object -Property vertical_drop      -Sum).Sum
    $totalUphill      = ($group | Measure-Object -Property total_elevation_gain -Sum).Sum

    [pscustomobject]@{
        Label                      = $label
        Days                       = $daysCount
        Activities                 = $activityCount
        Runs                       = [int]$totalRuns
        'Avg Runs Per Day'         = [int][math]::Round($totalRuns / $daysCount, 0)
        'Distance (km)'            = [math]::Round($totalDistance / 1000, 2)
        'Avg Distance Per Day (km)'= [math]::Round(($totalDistance / 1000) / $daysCount, 2)
        'Vertical Descent (km)'    = [math]::Round($totalVertical / 1000, 2)
        'Avg Vertical Descent (m)' = [int][math]::Round($totalVertical / $activityCount, 0)
        'Uphill Ascent (km)'       = [math]::Round($totalUphill / 1000, 2)
        'Avg Uphill Ascent (m)'    = [int][math]::Round($totalUphill / $activityCount, 0)
        'Total Moving Time (h)'    = [math]::Round($totalMovingTime / 3600, 2)
        'Total Elapsed Time (h)'   = [math]::Round($totalElapsedTime / 3600, 2)
        'Avg Moving Time'          = Format-TimeHoursMinutes ($totalMovingTime  / $activityCount)
        'Avg Elapsed Time'         = Format-TimeHoursMinutes ($totalElapsedTime / $activityCount)
        'Max Speed (km/h)'         = [math]::Round(($group | Measure-Object -Property max_speed     -Maximum).Maximum * 3.6, 2)
        'Avg Max Speed (km/h)'     = [math]::Round(($group | Measure-Object -Property max_speed     -Average).Average * 3.6, 2)
        'Avg Speed (km/h)'         = [math]::Round(($group | Measure-Object -Property average_speed -Average).Average * 3.6, 2)
    }
}

function Write-StatsTables($stat, $title) {
    Write-Host "`n=== ${title}: $($stat.Label) ===" -ForegroundColor Yellow

    Write-Host "`nRuns & Distance" -ForegroundColor Cyan
    [pscustomobject]@{
        Days                       = $stat.Days
        Runs                       = $stat.Runs
        'Avg Runs Per Day'         = $stat.'Avg Runs Per Day'
        'Distance (km)'            = $stat.'Distance (km)'
        'Avg Distance Per Day (km)'= $stat.'Avg Distance Per Day (km)'
    } | Format-Table -AutoSize

    Write-Host "Elevation" -ForegroundColor Cyan
    [pscustomobject]@{
        'Vertical Descent (km)'    = $stat.'Vertical Descent (km)'
        'Avg Vertical Descent (m)' = $stat.'Avg Vertical Descent (m)'
        'Uphill Ascent (km)'       = $stat.'Uphill Ascent (km)'
        'Avg Uphill Ascent (m)'    = $stat.'Avg Uphill Ascent (m)'
    } | Format-Table -AutoSize

    Write-Host "Time" -ForegroundColor Cyan
    [pscustomobject]@{
        'Total Moving Time (h)'  = $stat.'Total Moving Time (h)'
        'Total Elapsed Time (h)' = $stat.'Total Elapsed Time (h)'
        'Avg Moving Time'        = $stat.'Avg Moving Time'
        'Avg Elapsed Time'       = $stat.'Avg Elapsed Time'
    } | Format-Table -AutoSize

    Write-Host "Speed" -ForegroundColor Cyan
    [pscustomobject]@{
        'Max Speed (km/h)'     = $stat.'Max Speed (km/h)'
        'Avg Max Speed (km/h)' = $stat.'Avg Max Speed (km/h)'
        'Avg Speed (km/h)'     = $stat.'Avg Speed (km/h)'
    } | Format-Table -AutoSize
}

# ── Entry point ────────────────────────────────────────────────────────────────

if (-not (Test-Path $CredFile)) {
    Initialize-StravaAuth
    exit
}

$cred         = Import-Clixml $CredFile
$ClientId     = $cred.UserName
$ClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred.Password)
)

if (-not (Test-Path $TokenFile)) {
    Write-Host "No tokens found. Run the setup first." -ForegroundColor Red
    exit
}

$tokens  = Import-TokenFile
$tokens  = Update-AccessToken $ClientId $ClientSecret $tokens.refresh_token
$Headers = @{ Authorization = "Bearer $($tokens.access_token)" }

# Load cache
$cache = Import-ActivityCache
Write-Host "Loaded cache with $($cache.Count) activities" -ForegroundColor Cyan

# Fetch all activities (paginated)
Write-Host "Fetching activity list from Strava..." -ForegroundColor Cyan
$activities = @()
$page = 1
do {
    $batch = Invoke-RestMethod "https://www.strava.com/api/v3/athlete/activities?per_page=200&page=$page" -Headers $Headers
    if ($batch.Count -eq 0) { break }
    $activities += $batch
    $page++
} while ($true)

# Filter for winter sports
$winterSports = $activities | Where-Object {
    $_.sport_type -in @("Snowboard", "AlpineSki", "BackcountrySki", "NordicSki")
}

# Attach date and season
$winterSports | ForEach-Object {
    $_ | Add-Member -NotePropertyName date   -NotePropertyValue ([datetime]$_.start_date)
    $_ | Add-Member -NotePropertyName season -NotePropertyValue (Get-Season $_.date)
}

# Apply season filter
if ($SeasonFilter) {
    $winterSports = $winterSports | Where-Object { $_.season -eq $SeasonFilter }
    Write-Host "Found $($winterSports.Count) winter sport activities in season $SeasonFilter" -ForegroundColor Green
} else {
    Write-Host "Found $($winterSports.Count) winter sport activities" -ForegroundColor Green
}

# Determine which activities need processing
$toProcess = $winterSports | Where-Object { Test-ActivityStale $_ $cache }
$fromCache = $winterSports.Count - $toProcess.Count
Write-Host "  $fromCache activities loaded from cache" -ForegroundColor Gray
Write-Host "  $($toProcess.Count) activities require processing`n" -ForegroundColor Gray

# Process uncached activities
if ($toProcess.Count -gt 0) {
    Write-Host "Calculating vertical descent and run counts (this may take a few moments)..." -ForegroundColor Cyan
    $processedCount = 0

    foreach ($activity in $toProcess) {
        $processedCount++

        $details  = Get-ActivityDetails $activity.id $Headers
        $runCount = if ($details -and $details.laps) { $details.laps.Count } else { 0 }

        $elevationStream = Get-ActivityStreams $activity.id $Headers
        $verticalDescent = Get-VerticalDescent $elevationStream

        $cache[[string]$activity.id] = @{
            activity_id  = $activity.id
            updated_at   = $activity.updated_at
            run_count    = $runCount
            vertical_drop = $verticalDescent
        }

        if ($processedCount % 10 -eq 0) {
            Write-Host "  Processed $processedCount of $($toProcess.Count) activities..." -ForegroundColor Gray
        }
    }

    Write-Host "Processing complete`n" -ForegroundColor Green
    Export-ActivityCache $cache
    Write-Host "Cache updated with $($cache.Count) activities`n" -ForegroundColor Green
}

# Merge cached data into activity objects
$winterSports | ForEach-Object {
    $activityId = [string]$_.id
    $cached = if ($cache.ContainsKey($activityId)) { $cache[$activityId] } else { @{ run_count = 0; vertical_drop = 0 } }
    $_ | Add-Member -NotePropertyName run_count     -NotePropertyValue $cached.run_count    -Force
    $_ | Add-Member -NotePropertyName vertical_drop -NotePropertyValue $cached.vertical_drop -Force
    $_ | Add-Member -NotePropertyName day           -NotePropertyValue $_.date.Date
}

# Calculate season stats
$seasonStats = $winterSports | Group-Object season | ForEach-Object {
    Get-ActivityStats $_.Group $_.Name
} | Sort-Object Label -Descending

#OPTIONAL: Calendar year stats
<#
$yearStats = $winterSports | Group-Object { $_.date.Year } | ForEach-Object {
    Get-ActivityStats $_.Group $_.Name
} | Sort-Object Label -Descending
#>

# Output
Write-Host "                              __        __
    ____  ____ _      _______/ /_____ _/ /______
   / __ \/ __ \ | /| / / ___/ __/ __ `/ __/ ___/
  / /_/ / /_/ / |/ |/ (__  ) /_/ /_/ / /_(__  )
 / .___/\____/|__/|__/____/\__/\__,_/\__/____/
/_/ " -ForegroundColor Cyan

foreach ($stat in $seasonStats) {
    Write-StatsTables $stat "Season"
}

# OPTIONAL: Calendar year output
<#
Write-Host "`nBy Calendar Year"
foreach ($stat in $yearStats) {
    Write-StatsTables $stat "Year"
}
#>
