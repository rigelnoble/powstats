$ErrorActionPreference = "Stop"

# Configuration
$ConfigDir = "$env:APPDATA\powstats"
$CredFile = "$ConfigDir\strava_credentials.xml"
$TokenFile = "$ConfigDir\auth_tokens.json"
$CacheFile = "$ConfigDir\activity_cache.json"
$RedirectUri = "http://localhost:9876/callback"

# Ensure config directory exists
if (-not (Test-Path $ConfigDir)) {
    New-Item -ItemType Directory -Path $ConfigDir | Out-Null
}

# OPTIONAL: Set to a season (e.g. "2025-2026") to process only that season's activities
$SeasonFilter = "2025-2026"  # Change to "2025-2026" to filter that season, otherwise leave as "$null" for all.

# Setup function - run once
function Initialize-StravaAuth {
    Write-Host "`n=== Strava API Setup ===" -ForegroundColor Cyan
    Write-Host "You need to create a Strava API application first."
    Write-Host "1. Go to: https://www.strava.com/settings/api"
    Write-Host "2. Create an app with these settings:"
    Write-Host "   - Application Name: powstats (or whatever you want)"
    Write-Host "   - Category: Data Importer"
    Write-Host "   - Authorization Callback Domain: localhost`n"
    
    $ClientId = Read-Host "Enter your Client ID"
    $ClientSecret = Read-Host "Enter your Client Secret" -AsSecureString
    
    # Save credentials
    $cred = New-Object System.Management.Automation.PSCredential($ClientId, $ClientSecret)
    $cred | Export-Clixml $CredFile
    
    Write-Host "`nCredentials saved. Starting OAuth flow..." -ForegroundColor Green
    
    # Start temporary HTTP listener
    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("$RedirectUri/")
    $listener.Start()
    
    # Open browser for authorization
    $authUrl = "https://www.strava.com/oauth/authorize?client_id=$ClientId&response_type=code&redirect_uri=$RedirectUri&scope=activity:read_all"
    Write-Host "`nOpening browser for Strava authorization..."
    Start-Process $authUrl
    
    # Wait for callback
    Write-Host "Waiting for authorization..." -ForegroundColor Yellow
    $context = $listener.GetContext()
    $code = $context.Request.QueryString['code']
    
    # Send response to browser
    $response = $context.Response
    $html = "<html><body><h1>Authorization successful!</h1><p>You can close this window and return to PowerShell.</p></body></html>"
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
    $response.ContentLength64 = $buffer.Length
    $response.OutputStream.Write($buffer, 0, $buffer.Length)
    $response.Close()
    $listener.Stop()
    
    if (-not $code) {
        throw "No authorization code received"
    }
    
    # Exchange code for tokens
    $ClientSecretPlain = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
    )
    
    $tokens = Invoke-RestMethod -Method Post https://www.strava.com/oauth/token -Body @{
        client_id = $ClientId
        client_secret = $ClientSecretPlain
        code = $code
        grant_type = "authorization_code"
    }
    
    $tokens | ConvertTo-Json | Set-Content $TokenFile
    
    Write-Host "`nSetup complete! You can now run powstats.ps1" -ForegroundColor Green
}

# Load cache from disk
function Load-Cache {
    if (Test-Path $CacheFile) {
        try {
            $cacheData = Get-Content $CacheFile | ConvertFrom-Json
            # Convert array to hashtable for fast lookups by activity ID
            $cacheHash = @{}
            foreach ($item in $cacheData) {
                $cacheHash[[string]$item.activity_id] = $item
            }
            return $cacheHash
        }
        catch {
            Write-Host "Warning: Could not load cache, starting fresh" -ForegroundColor Yellow
            return @{}
        }
    }
    return @{}
}

# Save cache to disk
function Save-Cache($cache) {
    try {
        # Convert hashtable back to array for JSON storage
        $cacheArray = $cache.Values | ForEach-Object { $_ }
        $cacheArray | ConvertTo-Json -Depth 10 | Set-Content $CacheFile
    }
    catch {
        Write-Host "Warning: Could not save cache - $_" -ForegroundColor Yellow
    }
}

# Check if activity needs processing (not in cache or modified since last cache)
function Needs-Processing($activity, $cache) {
    $activityId = [string]$activity.id
    
    if (-not $cache.ContainsKey($activityId)) {
        return $true
    }
    
    $cached = $cache[$activityId]
    
    # Handle missing updated_at gracefully
    if (-not $activity.updated_at -or -not $cached.updated_at) {
        return $true
    }
    
    try {
        $activityUpdated = [datetime]$activity.updated_at
        $cacheUpdated = [datetime]$cached.updated_at
        
        return $activityUpdated -gt $cacheUpdated
    }
    catch {
        # If date parsing fails, reprocess to be safe
        return $true
    }
}

# Calculate vertical descent from elevation stream
function Get-VerticalDescent($elevationData) {
    if (-not $elevationData -or $elevationData.Count -lt 2) {
        return 0
    }
    
    $totalDescent = 0
    for ($i = 1; $i -lt $elevationData.Count; $i++) {
        $diff = $elevationData[$i] - $elevationData[$i-1]
        if ($diff -lt 0) {
            $totalDescent += [math]::Abs($diff)
        }
    }
    return $totalDescent
}

# Format seconds to hours and minutes (e.g. "1h 46m")
function Format-TimeHoursMinutes($seconds) {
    $hours = [math]::Floor($seconds / 3600)
    $minutes = [math]::Floor(($seconds % 3600) / 60)
    return "${hours}h ${minutes}m"
}

# Fetch detailed activity data including laps
function Get-ActivityDetails($activityId, $headers) {
    try {
        $activity = Invoke-RestMethod "https://www.strava.com/api/v3/activities/$activityId" -Headers $headers
        return $activity
    }
    catch {
        Write-Host "  Warning: Could not fetch details for activity $activityId - $_" -ForegroundColor Yellow
    }
    return $null
}

# Fetch streams for an activity
function Get-ActivityStreams($activityId, $headers) {
    try {
        $stream = Invoke-RestMethod "https://www.strava.com/api/v3/activities/$activityId/streams?keys=altitude&key_by_type=true" -Headers $headers
        if ($stream.altitude -and $stream.altitude.data) {
            return $stream.altitude.data
        }
    }
    catch {
        Write-Host "  Warning: Could not fetch streams for activity $activityId - $_" -ForegroundColor Yellow
    }
    return $null
}

function Get-Season($d) {
    if ($d.Month -ge 7) { "$($d.Year)-$($d.Year+1)" }
    else { "$($d.Year-1)-$($d.Year)" }
}

# Check if setup needed
if (-not (Test-Path $CredFile)) {
    Initialize-StravaAuth
    exit
}

# Normal script execution
$cred = Import-Clixml $CredFile
$ClientId = $cred.UserName
$ClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred.Password)
)

function Load-Tokens {
    if (Test-Path $TokenFile) {
        Get-Content $TokenFile | ConvertFrom-Json
    }
}

function Refresh-Token($refresh) {
    Invoke-RestMethod -Method Post https://www.strava.com/oauth/token -Body @{
        client_id = $ClientId
        client_secret = $ClientSecret
        grant_type = "refresh_token"
        refresh_token = $refresh
    }
}

# Token refresh
if (-not (Test-Path $TokenFile)) {
    Write-Host "No tokens found. Run the setup first." -ForegroundColor Red
    exit
}

$tokens = Load-Tokens
$tokens = Refresh-Token $tokens.refresh_token
$tokens | ConvertTo-Json | Set-Content $TokenFile
$Headers = @{ Authorization = "Bearer $($tokens.access_token)" }

# Load cache
$cache = Load-Cache
Write-Host "Loaded cache with $($cache.Count) activities" -ForegroundColor Cyan

# Download activities
Write-Host "Fetching activity list from Strava..." -ForegroundColor Cyan
$activities = @()
$page = 1
do {
    $batch = Invoke-RestMethod "https://www.strava.com/api/v3/athlete/activities?per_page=200&page=$page" -Headers $Headers
    $activities += $batch
    $page++
} while ($batch.Count -gt 0)

# Filter for winter sports: Snowboard, Alpine Ski, Backcountry Ski, Nordic Ski
# Strava sport_type values: "Snowboard", "AlpineSki", "BackcountrySki", "NordicSki"
$winterSports = $activities | Where-Object { 
    $_.sport_type -in @("Snowboard", "AlpineSki", "BackcountrySki", "NordicSki")
}

# Add dates first for season filtering
$winterSports | ForEach-Object {
    $_ | Add-Member -NotePropertyName date -NotePropertyValue ([datetime]$_.start_date)
    $_ | Add-Member -NotePropertyName season -NotePropertyValue (Get-Season $_.date)
}

# Apply season filter if specified
if ($SeasonFilter) {
    $winterSports = $winterSports | Where-Object { $_.season -eq $SeasonFilter }
    Write-Host "Found $($winterSports.Count) winter sport activities in season $SeasonFilter" -ForegroundColor Green
} else {
    Write-Host "Found $($winterSports.Count) winter sport activities" -ForegroundColor Green
}

# Determine which activities need processing
$toProcess = $winterSports | Where-Object { Needs-Processing $_ $cache }
$fromCache = $winterSports.Count - $toProcess.Count

Write-Host "  $fromCache activities loaded from cache" -ForegroundColor Gray
Write-Host "  $($toProcess.Count) activities require processing`n" -ForegroundColor Gray

# Process activities that need updating
if ($toProcess.Count -gt 0) {
    Write-Host "Calculating vertical descent and run counts (this may take a few moments)..." -ForegroundColor Cyan
    $processedCount = 0
    
    foreach ($activity in $toProcess) {
        $processedCount++
        
        # Fetch detailed activity data for lap count
        $details = Get-ActivityDetails $activity.id $Headers
        $runCount = if ($details -and $details.laps) { $details.laps.Count } else { 0 }
        
        # Fetch elevation stream and calculate vertical descent
        $elevationStream = Get-ActivityStreams $activity.id $Headers
        $verticalDescent = Get-VerticalDescent $elevationStream
        
        # Update cache with processed data
        $cache[[string]$activity.id] = @{
            activity_id = $activity.id
            updated_at = $activity.updated_at
            run_count = $runCount
            vertical_drop = $verticalDescent
        }
        
        # Progress indicator every 10 activities
        if ($processedCount % 10 -eq 0) {
            Write-Host "  Processed $processedCount of $($toProcess.Count) activities..." -ForegroundColor Gray
        }
    }
    
    Write-Host "Processing complete`n" -ForegroundColor Green
    
    # Save updated cache
    Save-Cache $cache
    Write-Host "Cache updated with $($cache.Count) activities`n" -ForegroundColor Green
}

# Merge cached data into all activities
$winterSports | ForEach-Object {
    $activityId = [string]$_.id
    if ($cache.ContainsKey($activityId)) {
        $_ | Add-Member -NotePropertyName run_count -NotePropertyValue $cache[$activityId].run_count -Force
        $_ | Add-Member -NotePropertyName vertical_drop -NotePropertyValue $cache[$activityId].vertical_drop -Force
    } else {
        # Shouldn't happen, but handle gracefully
        $_ | Add-Member -NotePropertyName run_count -NotePropertyValue 0 -Force
        $_ | Add-Member -NotePropertyName vertical_drop -NotePropertyValue 0 -Force
    }
}

$winterSports | ForEach-Object {
    $_ | Add-Member day $_.date.Date
}

# Season stats
$seasonStats = $winterSports | Group-Object season | ForEach-Object {
    $group = $_.Group
    $totalMovingTime = ($group | Measure-Object -Property moving_time -Sum).Sum
    $totalElapsedTime = ($group | Measure-Object -Property elapsed_time -Sum).Sum
    $activityCount = $group.Count
    $daysCount = ($group.day | Sort-Object -Unique).Count
    $totalRuns = ($group | Measure-Object -Property run_count -Sum).Sum
    $totalDistance = ($group | Measure-Object -Property distance -Sum).Sum
    $totalVertical = ($group | Measure-Object -Property vertical_drop -Sum).Sum
    $totalUphill = ($group | Measure-Object -Property total_elevation_gain -Sum).Sum
    
    [pscustomobject]@{
        Season = $_.Name
        Days = $daysCount
        Activities = $activityCount
        Runs = [int]$totalRuns
        'Avg Runs Per Day' = [int][math]::Round($totalRuns / $daysCount, 0)
        'Distance (km)' = [math]::Round($totalDistance / 1000, 2)
        'Avg Distance Per Day (km)' = [math]::Round(($totalDistance / 1000) / $daysCount, 2)
        'Vertical Descent (km)' = [math]::Round($totalVertical / 1000, 2)
        'Avg Vertical Descent (m)' = [int][math]::Round($totalVertical / $activityCount, 0)
        'Uphill Ascent (km)' = [math]::Round($totalUphill /1000, 2)
        'Avg Uphill Ascent (m)' = [int][math]::Round($totalUphill / $activityCount, 0)
        'Total Moving Time (h)' = [math]::Round($totalMovingTime / 3600, 2)
        'Total Elapsed Time (h)' = [math]::Round($totalElapsedTime / 3600, 2)
        'Avg Moving Time' = Format-TimeHoursMinutes ($totalMovingTime / $activityCount)
        'Avg Elapsed Time' = Format-TimeHoursMinutes ($totalElapsedTime / $activityCount)
        'Max Speed (km/h)' = [math]::Round(($group | Measure-Object -Property max_speed -Maximum).Maximum * 3.6, 2)
        'Avg Max Speed (km/h)' = [math]::Round((($group | Measure-Object -Property max_speed -Average).Average) * 3.6, 2)
        'Avg Speed (km/h)' = [math]::Round((($group | Measure-Object -Property average_speed -Average).Average) * 3.6, 2)
    }
} | Sort-Object Season -Descending

#OPTIONAL: Calendar year stats
<#
$yearStats = $winterSports | Group-Object { $_.date.Year } | ForEach-Object {
    $group = $_.Group
    $totalMovingTime = ($group | Measure-Object -Property moving_time -Sum).Sum
    $totalElapsedTime = ($group | Measure-Object -Property elapsed_time -Sum).Sum
    $activityCount = $group.Count
    $daysCount = ($group.day | Sort-Object -Unique).Count
    $totalRuns = ($group | Measure-Object -Property run_count -Sum).Sum
    $totalDistance = ($group | Measure-Object -Property distance -Sum).Sum
    $totalVertical = ($group | Measure-Object -Property vertical_drop -Sum).Sum
    $totalUphill = ($group | Measure-Object -Property total_elevation_gain -Sum).Sum
    
    [pscustomobject]@{
        Year = $_.Name
        Days = $daysCount
        Activities = $activityCount
        Runs = [int]$totalRuns
        'Avg Runs Per Day' = [int][math]::Round($totalRuns / $daysCount, 0)
        'Distance (km)' = [math]::Round($totalDistance / 1000, 2)
        'Avg Distance Per Day (km)' = [math]::Round(($totalDistance / 1000) / $daysCount, 2)
        'Vertical Descent (km)' = [math]::Round($totalVertical / 1000, 2)
        'Avg Vertical Descent (m)' = [int][math]::Round($totalVertical / $activityCount, 0)
        'Uphill Ascent (km)' = [math]::Round($totalUphill / 1000, 2)
        'Avg Uphill Ascent (m)' = [int][math]::Round($totalUphill / $activityCount, 0)
        'Total Moving Time (h)' = [math]::Round($totalMovingTime / 3600, 2)
        'Total Elapsed Time (h)' = [math]::Round($totalElapsedTime / 3600, 2)
        'Avg Moving Time' = Format-TimeHoursMinutes ($totalMovingTime / $activityCount)
        'Avg Elapsed Time' = Format-TimeHoursMinutes ($totalElapsedTime / $activityCount)
        'Max Speed (km/h)' = [math]::Round(($group | Measure-Object -Property max_speed -Maximum).Maximum * 3.6, 2)
        'Avg Max Speed (km/h)' = [math]::Round((($group | Measure-Object -Property max_speed -Average).Average) * 3.6, 2)
        'Avg Speed (km/h)' = [math]::Round((($group | Measure-Object -Property average_speed -Average).Average) * 3.6, 2)
    }
} | Sort-Object Year -Descending
#>

# Output
Write-Host "                              __        __      
    ____  ____ _      _______/ /_____ _/ /______
   / __ \/ __ \ | /| / / ___/ __/ __ `/ __/ ___/
  / /_/ / /_/ / |/ |/ (__  ) /_/ /_/ / /_(__  ) 
 / .___/\____/|__/|__/____/\__/\__,_/\__/____/  
/_/ " -ForegroundColor Cyan

foreach ($stat in $seasonStats) {
    Write-Host "`n=== Season: $($stat.Season) ===" -ForegroundColor Yellow
    
    # Table 1: Runs & Distance
    Write-Host "`nRuns & Distance" -ForegroundColor Cyan
    [pscustomobject]@{
        Days = $stat.Days
        Runs = $stat.Runs
        'Avg Runs Per Day' = $stat.'Avg Runs Per Day'
        'Distance (km)' = $stat.'Distance (km)'
        'Avg Distance Per Day (km)' = $stat.'Avg Distance Per Day (km)'
    } | Format-Table -AutoSize
    
    # Table 2: Elevation
    Write-Host "Elevation" -ForegroundColor Cyan
    [pscustomobject]@{
        'Vertical Descent (km)' = $stat.'Vertical Descent (km)'
        'Avg Vertical Descent (m)' = $stat.'Avg Vertical Descent (m)'
        'Uphill Ascent (km)' = $stat.'Uphill Ascent (km)'
        'Avg Uphill Ascent (m)' = $stat.'Avg Uphill Ascent (m)'
    } | Format-Table -AutoSize
    
    # Table 3: Time
    Write-Host "Time" -ForegroundColor Cyan
    [pscustomobject]@{
        'Total Moving Time (h)' = $stat.'Total Moving Time (h)'
        'Total Elapsed Time (h)' = $stat.'Total Elapsed Time (h)'
        'Avg Moving Time' = $stat.'Avg Moving Time'
        'Avg Elapsed Time' = $stat.'Avg Elapsed Time'
    } | Format-Table -AutoSize
    
    # Table 4: Speed
    Write-Host "Speed" -ForegroundColor Cyan
    [pscustomobject]@{
        'Max Speed (km/h)' = $stat.'Max Speed (km/h)'
        'Avg Max Speed (km/h)' = $stat.'Avg Max Speed (km/h)'
        'Avg Speed (km/h)' = $stat.'Avg Speed (km/h)'
    } | Format-Table -AutoSize
}

# OPTIONAL: Calendar year output
<#
Write-Host "`nBy Calendar Year"
foreach ($stat in $yearStats) {
    Write-Host "`n=== Year: $($stat.Year) ===" -ForegroundColor Yellow
    
    # Table 1: Runs & Distance
    Write-Host "`nRuns & Distance" -ForegroundColor Cyan
    [pscustomobject]@{
        Days = $stat.Days
        Runs = $stat.Runs
        'Avg Runs Per Day' = $stat.'Avg Runs Per Day'
        'Distance (km)' = $stat.'Distance (km)'
        'Avg Distance Per Day (km)' = $stat.'Avg Distance Per Day (km)'
    } | Format-Table -AutoSize
    
    # Table 2: Elevation
    Write-Host "Elevation" -ForegroundColor Cyan
    [pscustomobject]@{
        'Vertical Descent (km)' = $stat.'Vertical Descent (km)'
        'Avg Vertical Descent (m)' = $stat.'Avg Vertical Descent (m)'
        'Uphill Ascent (km)' = $stat.'Uphill Ascent (km)'
        'Avg Uphill Ascent (m)' = $stat.'Avg Uphill Ascent (m)'
    } | Format-Table -AutoSize
    
    # Table 3: Time
    Write-Host "Time" -ForegroundColor Cyan
    [pscustomobject]@{
        'Total Moving Time (h)' = $stat.'Total Moving Time (h)'
        'Total Elapsed Time (h)' = $stat.'Total Elapsed Time (h)'
        'Avg Moving Time' = $stat.'Avg Moving Time'
        'Avg Elapsed Time' = $stat.'Avg Elapsed Time'
    } | Format-Table -AutoSize
    
    # Table 4: Speed
    Write-Host "Speed" -ForegroundColor Cyan
    [pscustomobject]@{
        'Max Speed (km/h)' = $stat.'Max Speed (km/h)'
        'Avg Max Speed (km/h)' = $stat.'Avg Max Speed (km/h)'
        'Avg Speed (km/h)' = $stat.'Avg Speed (km/h)'
    } | Format-Table -AutoSize
}
#>