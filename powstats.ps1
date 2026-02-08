$ErrorActionPreference = "Stop"

# Configuration
$CredFile = "$env:APPDATA\powstats_strava.xml"
$TokenFile = "$env:APPDATA\powstats_token.json"
$RedirectUri = "http://localhost:9876/callback"

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

# Download activities
Write-Host "Downloading activities from Strava..." -ForegroundColor Cyan
$activities = @()
$page = 1
do {
    $batch = Invoke-RestMethod "https://www.strava.com/api/v3/athlete/activities?per_page=200&page=$page" -Headers $Headers
    $activities += $batch
    $page++
} while ($batch.Count -gt 0)

# Filter snowboarding
$snow = $activities | Where-Object { $_.type -eq "Snowboard" }

# Add dates first for season filtering
$snow | ForEach-Object {
    $_ | Add-Member -NotePropertyName date -NotePropertyValue ([datetime]$_.start_date)
    $_ | Add-Member -NotePropertyName season -NotePropertyValue (Get-Season $_.date)
}

# Apply season filter if specified
if ($SeasonFilter) {
    $snow = $snow | Where-Object { $_.season -eq $SeasonFilter }
    Write-Host "Found $($snow.Count) snowboarding activities in season $SeasonFilter`n" -ForegroundColor Green
} else {
    Write-Host "Found $($snow.Count) snowboarding activities`n" -ForegroundColor Green
}

# Parse dates and fetch elevation streams for accurate vertical descent
Write-Host "Calculating vertical descent and run counts (this may take a moment)..." -ForegroundColor Cyan
$processedCount = 0
$snow | ForEach-Object {
    $processedCount++
    
    # Fetch detailed activity data for lap count
    $details = Get-ActivityDetails $_.id $Headers
    $runCount = if ($details -and $details.laps) { $details.laps.Count } else { 0 }
    $_ | Add-Member -NotePropertyName run_count -NotePropertyValue $runCount
    
    # Fetch elevation stream and calculate vertical descent
    $elevationStream = Get-ActivityStreams $_.id $Headers
    $verticalDescent = Get-VerticalDescent $elevationStream
    
    $_ | Add-Member -NotePropertyName vertical_drop -NotePropertyValue $verticalDescent
    
    # Progress indicator every 10 activities
    if ($processedCount % 10 -eq 0) {
        Write-Host "  Processed $processedCount of $($snow.Count) activities..." -ForegroundColor Gray
    }
}
Write-Host "Processing complete`n" -ForegroundColor Green

$snow | ForEach-Object {
    $_ | Add-Member day $_.date.Date
}

# Season stats
$seasonStats = $snow | Group-Object season | ForEach-Object {
    $group = $_.Group
    [pscustomobject]@{
        Season = $_.Name
        Days = ($group.day | Sort-Object -Unique).Count
        Runs = [int](($group | Measure-Object -Property run_count -Sum).Sum)
        'Distance (km)' = [math]::Round(($group | Measure-Object -Property distance -Sum).Sum / 1000, 2)
        'Vertical Descent (km)' = [math]::Round(($group | Measure-Object -Property vertical_drop -Sum).Sum / 1000, 2)
        'Uphill Ascent (m)' = [math]::Round(($group | Measure-Object -Property total_elevation_gain -Sum).Sum, 2)
        'Moving Time (h)' = [math]::Round(($group | Measure-Object -Property moving_time -Sum).Sum / 3600, 2)
        'Elapsed Time (h)' = [math]::Round(($group | Measure-Object -Property elapsed_time -Sum).Sum / 3600, 2)
        'Max Speed (km/h)' = [math]::Round(($group | Measure-Object -Property max_speed -Maximum).Maximum * 3.6, 2)
        'Avg Max Speed (km/h)' = [math]::Round((($group | Measure-Object -Property max_speed -Average).Average) * 3.6, 2)
        'Avg Speed (km/h)' = [math]::Round((($group | Measure-Object -Property average_speed -Average).Average) * 3.6, 2)
    }
} | Sort-Object Season -Descending

#OPTIONAL: Calendar year stats
<#
$yearStats = $snow | Group-Object { $_.date.Year } | ForEach-Object {
    $group = $_.Group
    [pscustomobject]@{
        Year = $_.Name
        Days = ($group.day | Sort-Object -Unique).Count
        Runs = [int](($group | Measure-Object -Property run_count -Sum).Sum)
        'Distance (km)' = [math]::Round(($group | Measure-Object -Property distance -Sum).Sum / 1000, 2)
        'Vertical Descent (km)' = [math]::Round(($group | Measure-Object -Property vertical_drop -Sum).Sum / 1000, 2)
        'Uphill Ascent (m)' = [math]::Round(($group | Measure-Object -Property total_elevation_gain -Sum).Sum, 2)
        'Moving Time (h)' = [math]::Round(($group | Measure-Object -Property moving_time -Sum).Sum / 3600, 2)
        'Elapsed Time (h)' = [math]::Round(($group | Measure-Object -Property elapsed_time -Sum).Sum / 3600, 2)
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
Write-Host "`nBy Season"
$seasonStats | Format-Table -AutoSize
# OPTIONAL:
<#
Write-Host "By Calendar Year"
$yearStats | Format-Table -AutoSize
#>