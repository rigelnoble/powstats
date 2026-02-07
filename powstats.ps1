$ErrorActionPreference = "Stop"

# Configuration
$CredFile = "$env:APPDATA\powstats_strava.xml"
$TokenFile = "$env:APPDATA\powstats_token.json"
$RedirectUri = "http://localhost:9876/callback"

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

Write-Host "Total activities downloaded: $($activities.Count)" -ForegroundColor Green

# Filter snowboarding
$snow = $activities | Where-Object { $_.type -eq "Snowboard" }
Write-Host "Snowboarding activities: $($snow.Count)" -ForegroundColor Green

# Parse dates and fetch elevation streams for accurate vertical descent
Write-Host "`nFetching elevation data..." -ForegroundColor Cyan
$processedCount = 0
$snow | ForEach-Object {
    $processedCount++
    Write-Host "  Processing activity $processedCount of $($snow.Count): $($_.name)" -ForegroundColor Gray
    
    $_ | Add-Member -NotePropertyName date -NotePropertyValue ([datetime]$_.start_date)
    
    # Fetch elevation stream and calculate vertical descent
    $elevationStream = Get-ActivityStreams $_.id $Headers
    $verticalDescent = Get-VerticalDescent $elevationStream
    
    $_ | Add-Member -NotePropertyName vertical_drop -NotePropertyValue $verticalDescent
}

function Get-Season($d) {
    if ($d.Month -ge 7) { "$($d.Year)-$($d.Year+1)" }
    else { "$($d.Year-1)-$($d.Year)" }
}

$snow | ForEach-Object {
    $_ | Add-Member season (Get-Season $_.date)
    $_ | Add-Member day $_.date.Date
}

# Season stats
$seasonStats = $snow | Group-Object season | ForEach-Object {
    $group = $_.Group
    [pscustomobject]@{
        Season = $_.Name
        Days = ($group.day | Sort-Object -Unique).Count
        'Distance (km)' = [math]::Round(($group | Measure-Object -Property distance -Sum).Sum / 1000, 2)
        'Vertical Descent (m)' = [math]::Round(($group | Measure-Object -Property vertical_drop -Sum).Sum, 2)
        'Uphill Ascent (m)' = [math]::Round(($group | Measure-Object -Property total_elevation_gain -Sum).Sum, 2)
        'Moving Time (h)' = [math]::Round(($group | Measure-Object -Property moving_time -Sum).Sum / 3600, 2)
        'Elapsed Time (h)' = [math]::Round(($group | Measure-Object -Property elapsed_time -Sum).Sum / 3600, 2)
        'All Time Max Speed (km/h)' = [math]::Round(($group | Measure-Object -Property max_speed -Maximum).Maximum * 3.6, 2)
        'Avg Max Speed (km/h)' = [math]::Round((($group | Measure-Object -Property max_speed -Average).Average) * 3.6, 2)
        'Avg Speed (km/h)' = [math]::Round((($group | Measure-Object -Property average_speed -Average).Average) * 3.6, 2)
    }
} | Sort-Object Season -Descending

# OPTIONAL: Calendar year stats
<#
 $yearStats = $snow | Group-Object { $_.date.Year } | ForEach-Object {
    $group = $_.Group
    [pscustomobject]@{
        Year = $_.Name
        Days = ($group.day | Sort-Object -Unique).Count
        'Distance (km)' = [math]::Round(($group | Measure-Object -Property distance -Sum).Sum / 1000, 2)
        'Vertical (m)' = [math]::Round(($group | Measure-Object -Property vertical_drop -Sum).Sum, 2)
        'Non-Chairlift Uphill Gain (m)' = [math]::Round(($group | Measure-Object -Property total_elevation_gain -Sum).Sum, 2)
        'Moving Time (h)' = [math]::Round(($group | Measure-Object -Property moving_time -Sum).Sum / 3600, 2)
        'Elapsed Time (h)' = [math]::Round(($group | Measure-Object -Property elapsed_time -Sum).Sum / 3600, 2)
        'Max Speed (km/h)' = [math]::Round(($group | Measure-Object -Property max_speed -Maximum).Maximum * 3.6, 2)
        'Avg Max Speed (km/h)' = [math]::Round((($group | Measure-Object -Property max_speed -Average).Average) * 3.6, 2)
        'Avg Speed (km/h)' = [math]::Round((($group | Measure-Object -Property average_speed -Average).Average) * 3.6, 2)
    }
} | Sort-Object Year -Descending
#>

# Output
Write-Host "`npowstats" -ForegroundColor Cyan
Write-Host "`nBy Season (Northern Hemisphere Winter)"
$seasonStats | Format-Table -AutoSize
# OPTIONAL: Stats by calendar year
<#
Write-Host "By Calendar Year"
$yearStats | Format-Table -AutoSize
#>