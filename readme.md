# **powstats**
PowerShell script to provide enhanced analysis of snowboarding activity data using Strava's API.

## What it does
Retrieves your personal Strava snowboarding activities and calculates statistics grouped by season, or optionally by calendar year.

Statistics include:

- Days snowboarded
- Distance traveled (km)
- Vertical descent (m) - calculated from GPS elevation data
- Uphill ascent (m) - bootpacking / hiking, uphill sections
- Moving time vs elapsed time (hours)
- All time max speed and average max speed per activity (km/h)
- Average speed (km/h)

## Setup

1. Create a Strava API application at https://www.strava.com/settings/api

- Application Name: powstats (or your choice)
- Category: `Data Importer`
- Authorization Callback Domain: `localhost`

2. Run the script - it will prompt for your Client ID and Client Secret on first run.
3. Authorise the application in your browser when prompted.
4. Credentials are saved locally in `%APPDATA%` as encrypted files:

- `powstats_strava.xml` - API credentials (encrypted via Windows DPAPI)
- `powstats_token.json` - OAuth tokens (plaintext, auto-refreshed)

## Performance and rate limits
The script makes two API calls per snowboarding activity:
- One call to fetch activity list (shared across all activities)
- One call per activity to fetch GPS elevation streams. This is required for accurate elevation calculations.

**Strava API rate limits:**
- 100 requests per 15 minutes (read operations)
- 1,000 requests per day (read operations)

### Expected usage
- 50 activities = ~51 API calls (~8 minutes to avoid rate limits)
- 100 activities = ~101 API calls (will hit 15-minute rate limit, script will need to be run in batches or with delays)

For large activity counts (100+), consider running the script in multiple sessions or adding rate limit handling.

## Security

- No secrets hardcoded in script
- API credentials encrypted using Windows Data Protection API
- Only accessible by your Windows user account
- OAuth tokens stored locally, not in version control
- Read-only access to your activities (`activity:read_all scope`)

## Limitations

- Vertical descent calculated from GPS elevation streams - requires one API call per activity, which can approach rate limits for users with 100+ activities
- GPS elevation data can have minor inaccuracies due to barometric sensor drift or GPS noise (typically ±1-2% variance from Strava's app calculations)
- Season defined as northern hemisphere winter
- Requires Windows (uses DPAPI for credential encryption)
- Will not count backcountry snowboarding or any ski activities although this is probably an easy future addition. Sorry skiers, one plank is much better.

## Disclaimer
- Personal use only - not designed for multi-user or production environments
- Vibe coded with Claude on a lazy morning
- Use at your own risk

## Data collected
Script fetches all your Strava activities via API, filters for snowboarding type, retrieves GPS elevation data, and performs local calculations. No data is sent anywhere except to Strava's API for retrieval.