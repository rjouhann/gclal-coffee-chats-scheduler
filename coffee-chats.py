import os
import pickle
from datetime import datetime, timedelta, timezone
import pytz  # For handling time zones
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import argparse
from google_auth_oauthlib.flow import InstalledAppFlow
from collections import defaultdict
import random

# Set up the scope for Google APIs
SCOPES = ['https://www.googleapis.com/auth/calendar',
          'https://www.googleapis.com/auth/calendar.readonly',
          'https://www.googleapis.com/auth/spreadsheets.readonly']

# Variables
TEAM_CALENDAR_ID = 'abc@group.calendar.google.com'  # Replace with your team calendar ID
SPREADSHEET_ID = '123'  # Replace with your spreadsheet ID
SHEET_RANGE = 'Sheet1!A2:D100'  # Replace with your sheet range
MAX_MEETINGS_PER_WEEK_GROUP1 = 2
ORGANIZER_EMAIL = 'user@company.com'  # Email of the organizer
MEETING_DURATION_MINUTES = 20
DAYS_AHEAD = 90

# Authenticate and create Google API services
def authenticate_google_services():
    creds = None

    # Check if we already have a saved token
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there's no valid token, start the authentication flow
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            try:
                # Initiate the authentication process (browser-based)
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)  # Use browser-based auth to get the token
                creds = flow.run_local_server(port=0)

                # Save the credentials for future use (token.pickle)
                with open('token.pickle', 'wb') as token:
                    pickle.dump(creds, token)
                print("Token saved successfully!")  # Add this line for confirmation

            except Exception as e:
                print(f"Authentication failed: {e}")
                return None, None  # Return None if authentication fails

    # Build Google Calendar and Sheets services
    calendar_service = build('calendar', 'v3', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)

    return calendar_service, sheets_service

# Get data from Google Sheets
def get_people_data(sheets_service, spreadsheet_id, range_name):
    sheet = sheets_service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])

    return values

# Pair people randomly: one group1 with one group2
def pair_people(data):
    group1 = []  # e.g. Product Managers
    group2 = []  # e.g. Sales/CSM

    for row in data:
        group = row[1]  # Assuming the 2nd column is the "Group" column
        if group == 'group1':
            group1.append(row)
        elif group == 'group2':
            group2.append(row)

    if group1 and group2:
        print(f"Pairing: {len(group1)} from group 1 x {len(group2)} from group 2")
    else:
        print("No valid groups found!")

    return group1, group2

# Check availability of people for a given time, considering timezone differences
def check_availability(calendar_service, people_emails, start_time, debug):
    # Convert start_time to UTC for the freebusy query
    start_time_utc = start_time.astimezone(pytz.utc)  # Convert to UTC
    end_time = start_time + timedelta(minutes=MEETING_DURATION_MINUTES)
    end_time_utc = end_time.astimezone(pytz.utc)

    if debug:
        print(
            f"Checking availability for {people_emails} between {start_time_utc.isoformat()} and {end_time_utc.isoformat()}")

    busy_slots = []
    for email in people_emails:
        try:
            events_result = calendar_service.freebusy().query(body={
                "timeMin": start_time_utc.isoformat(),  # Use isoformat() for the query
                "timeMax": end_time_utc.isoformat(),
                "timeZone": 'UTC',  # Ensure the timezone is properly set to UTC for the query
                "items": [{"id": email}],
            }).execute()

            if debug:
                print(f"Freebusy result for {email}: {events_result}")
            busy_slots.append(events_result.get('calendars', {}).get(email, {}).get('busy', []))
        except HttpError as e:
            print(f"An error occurred: {e}")
            return False

    # Check for any overlap in UTC
    for slot in busy_slots:
        for event in slot:
            event_start = datetime.fromisoformat(event['start'].replace('Z', '+00:00'))
            event_end = datetime.fromisoformat(event['end'].replace('Z', '+00:00'))

            # Check for overlap:
            if (start_time_utc < event_end and end_time_utc > event_start) or \
               (event_start < end_time_utc and event_end > start_time_utc):
                if debug:
                    print(f"At least one person is busy during the meeting time.")
                return False

    return True

# Create a calendar event for the coffee chat in a different calendar
def create_calendar_event(calendar_service, person1, person2, start_time, end_time, debug, calendar_id=None,
                          send_email=True, is_organizer_reminder=False):
    if not calendar_id:
        calendar_id = TEAM_CALENDAR_ID  # Use default if no calendar_id is passed

    # Determine the event summary (title)
    if is_organizer_reminder:
        summary = '⏰ Coffee Chat Prep Reminder'  # Distinct title for organizer reminder
    else:
        summary = '☕ Chat Product - Sales'  # Default title

    event_body = {
    'summary': summary,
    'description': '''Some ideas:
• Discuss what is coming up in the next few weeks
• The roadmap: https://gitguardian.productboard.com/roadmap/8538423-launch-plan-saas
• Dive into a specific feature
• User Experience feedback
• Share insights from a recent customer interaction
• Chat about the 🌤️!

⚠️ If this time does not work, please feel free to reschedule!''',
    'start': {
        'dateTime': start_time.isoformat(),
        'timeZone': 'UTC',
    },
    'end': {
        'dateTime': end_time.isoformat(),
        'timeZone': 'UTC',
    },
    'attendees': [
        {'email': person1[2]},
        {'email': person2[2]},
    ],
    'reminders': {
        'useDefault': False,  # Disable default reminders
        'overrides': [
            {'method': 'popup', 'minutes': 10},  # Show a pop-up notification 10 minutes before the event
        ],
    },
    'sendUpdates': 'none' if not send_email else 'all',  # Disable emails if specified
    'guestsCanModify': True,  # Allow guests to modify the event
    'transparency': 'opaque',   # this makes it show as "busy"
    }

    event = calendar_service.events().insert(
        calendarId=calendar_id,  # Specify the calendar where you want the event to be created
        body=event_body
    ).execute()

    # Show details of the event that was created
    event_details = f"Event created: {event.get('htmlLink')} for {person1[0]} and {person2[0]} at {start_time.strftime('%Y-%m-%d %H:%M:%S')} UTC"

    if debug:
        print(event_details)
    else:
        print(event_details)  # Display even when not in debug mode
    return True

# Convert Paris time to UTC
def convert_france_to_utc(france_time):
    france_tz = pytz.timezone('Europe/Paris')

    # Localize France time if it is naive
    if france_time.tzinfo is None:
        france_time = france_tz.localize(france_time)

    # Convert to UTC time
    utc_time = france_time.astimezone(pytz.utc)
    return utc_time

# Ensure the scheduled time is a weekday (Monday to Friday)
def adjust_to_weekday(start_time):
    # Skip Saturday (5), and Sunday (6)
    while start_time.weekday() in [5, 6]:
        start_time += timedelta(days=0)

    return start_time

# Check if the time is within lunch hours (12 PM to 2 PM France time)
def is_lunch_time(start_time):
    # We need to check if the time is between 12 PM and 2 PM France time
    france_tz = pytz.timezone('Europe/Paris')
    local_start_time = start_time.astimezone(france_tz)

    # If the time is between 12 PM and 2 PM, return True
    if 12 <= local_start_time.hour < 14:
        return True
    return False

# Get the available time window for the pair
def get_time_window_for_pair(tz1, tz2, date):
    france_tz = pytz.timezone('Europe/Paris')

    # Define preferred start and time window durations (in hours)
    preferred_times = {
        ('Paris', 'Boston'): {'start': 15, 'duration': 3},  # 3:00 PM Paris, 3-hour window
        ('Boston', 'Paris'): {'start': 15, 'duration': 3},
        ('Paris', 'Chicago'): {'start': 16, 'duration': 2},  # 4:00 PM Paris, 2-hour window
        ('Chicago', 'Paris'): {'start': 16, 'duration': 2},
        ('Paris', 'Seattle'): {'start': 17.5, 'duration': 1.5},  # 5:30 PM Paris, 1.5-hour window
        ('Seattle', 'Paris'): {'start': 17.5, 'duration': 1.5},
        ('Paris', 'Paris'): {'start': 10, 'duration': 7.5},    # 10:00 AM Paris, 7.5-hour window
    }

    time_window = preferred_times.get((tz1, tz2))
    if not time_window:
        return None

    preferred_start_hour = time_window['start']
    preferred_duration = time_window['duration']

    preferred_start_time = france_tz.localize(
        datetime(date.year, date.month, date.day, int(preferred_start_hour),
                 int((preferred_start_hour - int(preferred_start_hour)) * 60)))

    return preferred_start_time, preferred_duration

def schedule_coffee_chats_with_tz(calendar_service, sheets_service, spreadsheet_id, range_name, days_ahead,
                              debug=False, send_email=True, dry_run=False):
    people_data = get_people_data(sheets_service, spreadsheet_id, range_name)
    group1, group2 = pair_people(people_data)

    expected_meetings = len(group1) * len(group2)
    print(f"Expected total number of meetings (full pairing): {expected_meetings}")

    start_date = datetime.now(timezone.utc)
    event_count = 0
    latest_event_date = None
    line_count = 0

    person_weekly_meetings = defaultdict(lambda: defaultdict(int))

    all_possible_pairs = list((p1, p2) for p1 in group1 for p2 in group2)
    unfulfilled_pairs = {}
    random.shuffle(all_possible_pairs)  # Shuffle to distribute meetings more evenly

    for p1, p2 in all_possible_pairs:
        current_date = start_date
        pair = tuple(sorted((p1[0], p2[0])))  # Create a consistent key for tracking
        slot_found = False  # Flag to track if a slot is found for the pair
        pair_unfulfilled_reasons = []  # Track reasons within the loop

        while current_date <= start_date + timedelta(days=days_ahead):
            if current_date.weekday() in [0, 1, 2, 3, 4]:  # Mon–Fri only
                week_start = current_date - timedelta(
                    days=current_date.weekday())  # Get the Monday of the current week
                current_week = week_start.date()

                p1_name = p1[0]

                # Check weekly limit for group1
                if p1[1] == 'group1' and person_weekly_meetings[p1_name][current_week] >= MAX_MEETINGS_PER_WEEK_GROUP1:
                    pair_unfulfilled_reasons.append(f"{p1_name} exceeded weekly limit ({MAX_MEETINGS_PER_WEEK_GROUP1})")
                    current_date += timedelta(days=1)
                    continue

                time_window = get_time_window_for_pair(p1[3], p2[3], current_date)
                if not time_window:
                    pair_unfulfilled_reasons.append("No overlapping time window between timezones")
                    current_date += timedelta(days=1)
                    continue

                preferred_start_time, preferred_duration = time_window
                slot_start = convert_france_to_utc(preferred_start_time)
                slot_end = slot_start + timedelta(minutes=MEETING_DURATION_MINUTES)

                # Iterate through possible meeting times within the window
                current_slot_start = slot_start
                while current_slot_start.astimezone(pytz.timezone('Europe/Paris')).hour < \
                        (preferred_start_time.hour + preferred_duration):
                    slot_end = current_slot_start + timedelta(minutes=MEETING_DURATION_MINUTES)
                    print(".", end="", flush=True)  # or even use logging.debug if preferred
                    # Check if the current slot is within the lunch break (12 PM to 2 PM Paris time)
                    paris_time = current_slot_start.astimezone(pytz.timezone('Europe/Paris'))
                    if 12 <= paris_time.hour < 14:
                        current_slot_start += timedelta(minutes=30)
                        continue

                    if check_availability(calendar_service, [p1[2], p2[2]], current_slot_start, debug):
                        line_count += 1
                        paris_tz = pytz.timezone('Europe/Paris')
                        local_slot = current_slot_start.astimezone(paris_tz)
                        if dry_run:
                            print(
                                f"[Dry Run {line_count}] Would schedule: {p1[0]} ↔ {p2[0]} at {local_slot.strftime('%Y-%m-%d %H:%M')} Paris time")
                        else:
                            paris_tz = pytz.timezone('Europe/Paris')
                            local_slot = current_slot_start.astimezone(paris_tz)
                            print(
                                f"[Run {line_count}] Schedule: {p1[0]} ↔ {p2[0]} at {local_slot.strftime('%Y-%m-%d %H:%M')} Paris time")

                            if not create_calendar_event(calendar_service, p1, p2, current_slot_start, slot_end, debug,
                                                        calendar_id=TEAM_CALENDAR_ID, send_email=send_email):
                                current_date += timedelta(days=1)
                                continue

                        # Update weekly meeting counts (only for group1)
                        if p1[1] == 'group1':
                            person_weekly_meetings[p1_name][current_week] += 1

                        event_count += 1
                        # Update latest_event_date if this event is later
                        if latest_event_date is None or current_slot_start.date() > latest_event_date:
                            latest_event_date = current_slot_start.date()
                        slot_found = True
                        break  # Break if a slot is found
                    else:
                        pair_unfulfilled_reasons.append(
                            f"No mutual availability at {current_slot_start.strftime('%Y-%m-%d %H:%M')}"
                        )

                    current_slot_start += timedelta(minutes=30)  # Increment by 30 minutes

                if slot_found:
                    break  # Break the while loop if a slot is found

            current_date += timedelta(days=1)

        # After checking all dates, if no slot was found, record the reasons
        if not slot_found and pair_unfulfilled_reasons:
            unfulfilled_pairs[pair] = "; ".join(set(pair_unfulfilled_reasons))

    print(f"Total events created: {event_count}")

    if latest_event_date:
        organizer_reminder_date = latest_event_date  # Reminder day of the last event
        organizer_reminder_time = pytz.timezone('Europe/Paris').localize(
            datetime(organizer_reminder_date.year, organizer_reminder_date.month, organizer_reminder_date.day, 17,
                     0))  # 5 PM Paris time
        organizer_reminder_time_utc = convert_france_to_utc(organizer_reminder_time)
        organizer_reminder_end_time_utc = organizer_reminder_time_utc + timedelta(minutes=30)

        if dry_run:
            paris_tz = pytz.timezone('Europe/Paris')
            local_slot = organizer_reminder_time_utc.astimezone(paris_tz)
            print(
                f"[Dry Run] Would schedule organizer reminder for {ORGANIZER_EMAIL} on {local_slot.strftime('%Y-%m-%d %H:%M')} Paris time")
        else:
            print(
                f"[Run] Schedule organizer reminder for {ORGANIZER_EMAIL} on {local_slot.strftime('%Y-%m-%d %H:%M')} Paris time")
            create_calendar_event(
                calendar_service,
                [None, None, ORGANIZER_EMAIL],  # Dummy person data for the organizer
                [None, None, ORGANIZER_EMAIL],  # Dummy person data for the organizer
                organizer_reminder_time_utc,
                organizer_reminder_end_time_utc,
                debug,
                calendar_id=TEAM_CALENDAR_ID,
                send_email=send_email,
                is_organizer_reminder=True
            )

    if latest_event_date:
        print(f"Last event scheduled on: {latest_event_date.strftime('%Y-%m-%d')}")

    # Summary of skipped pairs
    print("\n=== Unfulfilled Pairings Summary ===")
    total_unfulfilled = len(unfulfilled_pairs)  # Count the keys, not the reasons
    for pair, reason in unfulfilled_pairs.items():
        print(f"❌ {pair[0]} ↔ {pair[1]} — {reason}")
    print(f"\nTotal unfulfilled pairings: {total_unfulfilled}")

# Parse command-line arguments
def parse_arguments():
    parser = argparse.ArgumentParser(description="Schedule coffee chats")
    parser.add_argument('--debug', action='store_true', help="Enable debug mode for detailed logs")
    parser.add_argument('--no-email', action='store_true', help="Disable sending email invites")
    parser.add_argument('--dry-run', action='store_true', help="Run without creating calendar events")
    return parser.parse_args()

# Usage Example
args = parse_arguments()
calendar_service, sheets_service = authenticate_google_services()
schedule_coffee_chats_with_tz(calendar_service, sheets_service, SPREADSHEET_ID, SHEET_RANGE, DAYS_AHEAD,
                              debug=args.debug, send_email=not args.no_email, dry_run=args.dry_run)