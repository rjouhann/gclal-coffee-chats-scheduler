# Coffee Chat Scheduler for Google Calendar

This script schedules coffee chat meetings between 2 groups of people. It pairs people randomly based on a Google Sheets document, checks their availability using Google Calendar API, and creates events in a specified calendar.

## Features

-   **Pairing People**: Randomly pairs Group 1 (e.g. Product Managers) with Group 2 (e.g. Sales/CSM) based on data in a Google Sheet.
-   **Availability Check**: Uses Google Calendar API to ensure both members are free.
-   **Event Creation**: Schedules coffee chats in the team calendar.
-   **Timezone Support**: Automatically adjusts meeting times across supported time zones: Paris, Boston, Chicago, Seattle.
-   **Email Notifications**: Optionally disable email invites when scheduling events.
-   **Dry Run Support**: Preview pairings and scheduling without creating calendar events.
-   **Organizer Reminder**: Schedules an event to remind the organizer to set up the next batch of meetings.

## Prerequisites

1.  **Google API Credentials**:
    -   Enable Google Calendar API and Google Sheets API in your [Google Cloud Console](https://console.cloud.google.com/).
    -   Download `credentials.json` and place it in the same directory as the script.

2.  **Install Dependencies**:

    ```bash
    pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib pytz
    ```

## Configuration

Set these variables at the top of the script:

-   `TEAM_CALENDAR_ID`: ID of the calendar where events will be created.
-   `SPREADSHEET_ID`: ID of the Google Sheet with participant info.
-   `SHEET_RANGE`: Cell range where participant data is located.
-   `MAX_MEETINGS_PER_WEEK_GROUP1`: Maximum number of meetings per week for Group 1 members.
-   `ORGANIZER_EMAIL`: Email of the organizer who will receive the reminder.
-   `MEETING_DURATION_MINUTES`: Duration of each coffee chat meeting in minutes.
-   `DAYS_AHEAD`: Number of days ahead to schedule events.
-   
###   Expected Google Sheet Format

| Name       | Group  | Email Address          | Timezone |
| ---------- | ------ | ---------------------- | -------- |
| John Doe   | group1 | john.doe@example.com   | Paris    |
| Jane Smith | group2 | jane.smith@example.com | Boston   |
| ...        | ...    | ...                    | ...      |

-   `Group`: Use `group1` or `group2`.
-   `Timezone`: Must be one of: `Paris`, `Boston`, `Chicago`, `Seattle`.

##   Running the Script

1.  **Authenticate Google Services**:
    -   On the first run, the script opens a browser for authentication. Grant required permissions.

2.  **Schedule Coffee Chats**:

    ```bash
    python coffee-chats.py
    ```

3.  **Command-Line Options**:

| Option       | Description                                     |
| ------------ | ----------------------------------------------- |
| `--debug`    | Enable verbose logging for troubleshooting      |
| `--no-email` | Disable email invitations to participants       |
| `--dry-run`  | Simulate scheduling without creating any events |

Example:

```bash
python coffee-chats.py --debug --dry-run
```

##   Scheduling Logic

-   **Time Windows** (based on Paris timezone):

| Timezone Pair   | Start (Paris Time) | Duration (Hours) |
| --------------- | ------------------ | ---------------- |
| Paris - Boston  | 3:00 PM            | 3                |
| Paris - Chicago | 4:00 PM            | 2                |
| Paris - Seattle | 5:30 PM            | 1.5              |
| Paris - Paris   | 10:00 AM           | 7.5              |

-   **Lunch Break**: No meetings are scheduled between **12 PM and 2 PM Paris time**.
-   **Days**: Only Monday to Friday are considered.
-   **Group 1 Limit**: Each Group 1 member is scheduled for up to 2 chats per week.
-   **Organizer Reminder**: An event is created for the organizer on the day of the last scheduled meeting to remind them to set up the next batch.

##   Debugging

-   When using `--debug`, the script logs:
    -   Pairings and time suggestions
    -   Availability results from the Calendar API
    -   Details of event creation

-   `--dry-run` mode is useful for:
    -   Verifying pairings and scheduling logic
    -   Testing with live calendar data without creating events
