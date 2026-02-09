# powstats

PowerShell script for enhanced analysis of Strava winter sport activity data.

## Functionality

Retrieves your personal Strava winter sport activities (snowboarding, skiing, backcountry touring and nordic skiing) and calculates the following statistics grouped by season (or optionally by calendar year):

**Runs & Distance:**
- Days on snow
- Total runs/laps
- Average runs per day
- Total distance travelled (km)
- Average distance per day (km)

**Elevation:**
- Total vertical descent (km) - calculated from GPS elevation data
- Average vertical descent per activity (m)
- Total uphill ascent (km) - bootpacking, hiking, skinning, etc.
- Average uphill ascent per activity (m)

**Time:**
- Total moving time
- Total elapsed time
- Average moving time per activity
- Average elapsed time per activity

**Speed:**
- All time maximum speed
- Average maximum speed per activity
- Average speed across all activities

## Setup

1. Create a Strava API application at https://www.strava.com/settings/api
   - Application Name: powstats (or your choice)
   - Category: Data Importer
   - Authorization Callback Domain: `localhost`

2. Run the script - it will prompt for your Client ID and Client Secret on first run

3. Authorise the application in your browser when prompted

4. Credentials and cache are saved locally in `%APPDATA%\powstats\`:
   - `strava_credentials.xml` - API credentials (encrypted via Windows DPAPI)
   - `auth_tokens.json` - OAuth tokens (plaintext, auto-refreshed)
   - `activity_cache.json` - Processed activity data (run counts, vertical descent)

## Supported activities

The script processes the following Strava activity types:
- **Snowboard** - resort snowboarding
- **AlpineSki** - resort skiing
- **BackcountrySki** - ski touring (no backcountry snowboarding activity in Strava >:( )
- **NordicSki** - cross-country skiing

## Season filtering
To process only a specific season's activities, edit this value in the script:
```powershell
$SeasonFilter = "2025-2026"  # Process only 2025-2026 season
$SeasonFilter = $null        # Process all seasons
```

## Performance and rate limits

The script makes three API calls per winter sport activity:
1. One call to fetch the activity list (shared across all activities)
2. One call per activity to fetch detailed data (for run/lap counts)
3. One call per activity to fetch GPS elevation streams (for accurate vertical descent)

**Strava API rate limits:**
- 100 requests per 15 minutes (read operations)
- 1,000 requests per day (read operations)

**Estimated API usage:**
- 50 activities = ~101 API calls (within 100/15min read limit)
- 75 activities = ~151 API calls (will exceed 100/15min limit, wait ~8 minutes for rate limit reset)
- 100 activities = ~201 API calls (will exceed both 100/15min and 200/15min overall limits)

For large activity counts (75+), use the season filter to process only recent activities.

## Caching
The script caches processed activity data (run counts, vertical descent) to `%APPDATA%\powstats\activity_cache.json`.

**How it works:**
- First run: Processes all activities and caches results
- Subsequent runs: Only re-processes activities that have been modified on Strava (checked via `updated_at` timestamp)
- New activities are always processed

**Benefits:**
- Avoids recalculating vertical descent for unchanged activities
- Useful when you've edited/corrected old activities and need to reprocess them

**Note:** Caching does not reduce API calls significantly, as the script must still check each activity's `updated_at` timestamp. The primary benefit is avoiding redundant GPS elevation calculations.

**Cache management:**
- Cache persists between runs automatically
- To clear cache: Delete `%APPDATA%\powstats\activity_cache.json`
- Cache includes: activity ID, last updated timestamp, run count, vertical descent

## Security

- No secrets hardcoded in script
- API credentials encrypted using Windows Data Protection API (DPAPI)
- Only accessible by your Windows user account
- OAuth tokens stored locally in `%APPDATA%\powstats\`
- Read-only access to your activities (`activity:read_all` scope)

## Limitations

- Vertical descent calculated from GPS elevation streams - requires one API call per activity
- GPS elevation data can have minor inaccuracies due to barometric sensor drift or GPS noise (typically ±1-2% variance from Strava's app calculations)
- Run/lap counts may be inaccurate if Strava's auto-lap detection failed or wasn't used
- Season defined as July-June (Northern hemisphere winter). For those in the southern hemisphere, recommend using the optional calendar year mode.
- Requires Windows (uses DPAPI for credential encryption)
- Personal use only - not designed for multi-user or production environments
- Rate limits apply - for 75+ activities, consider using season filter or running script in multiple sessions
- Only tested with regular resort snowboard and ski activities (for now)

## Data collected

Script fetches all your Strava activities via API, filters for winter sport types, retrieves GPS elevation data, and performs local calculations. No data is sent anywhere except to Strava's API for retrieval. All data is stored locally on your machine.

## Output format

Results are displayed in four tables per season:

1. **Runs & Distance** - Days, runs, distance metrics
2. **Elevation** - Vertical descent and uphill ascent metrics
3. **Time** - Moving and elapsed time totals and averages
4. **Speed** - Maximum and average speed metrics

## Troubleshooting

**Rate limit errors:** Wait 15 minutes and run again, or use season filter to reduce activities processed

**Cache not working:** Delete `%APPDATA%\powstats\activity_cache.json` to start fresh

## Disclaimer

- Personal use only - not designed for multi-user or production environments
- Vibe coded with Claude on a lazy morning
- Use at your own risk
