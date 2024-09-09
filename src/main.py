#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# version 1.0.0
# https://github.com/danydavila/Python_Push_Event_To_gCalendar

import yaml
import os
import datetime
from zoneinfo import ZoneInfo
from typing import List, Dict, Any, Optional
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.discovery import Resource

# Define constants
YAML_FILE_PATH: str = './events.yaml'  # Path to the YAML file containing events
SCOPES: List[str] = ['https://www.googleapis.com/auth/calendar']  # Scope to access Google Calendar
DATE_FORMAT: str = '%Y/%m/%d %I:%M %p'
REQUIRED_EVENT_FIELDS: List[str] = ['title', 'description', 'event start time', 'event end time']


def authenticate_google() -> Resource:
    """Authenticate the user using OAuth 2.0 and return the Google Calendar service."""
    creds: Optional[Credentials] = None
    # Check if token.json exists and contains valid credentials
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    # If credentials are not available or are invalid, go through OAuth flow
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())  # Refresh the token
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json',
                                                             SCOPES)  # Run OAuth flow using credentials.json
            creds = flow.run_local_server(port=0)  # Open browser to authorize the app

        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    # Return an authenticated service object
    service: Resource = build('calendar', 'v3', credentials=creds)
    return service


def load_yaml_file(file_path: str) -> Dict[str, Any]:
    """Loads and parses the YAML file."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"YAML file {file_path} does not exist.")

    with open(file_path, 'r') as file:
        return yaml.safe_load(file)


def validate_event(event: Dict[str, Any]) -> None:
    """Validates the structure of an event and its date/time format."""
    missing_fields: List[str] = [field for field in REQUIRED_EVENT_FIELDS if field not in event]
    if missing_fields:
        raise ValueError(f"Event is missing the following required fields: {', '.join(missing_fields)}")

    # Validate datetime format for start and end time
    try:
        datetime.datetime.strptime(event['event start time'], DATE_FORMAT)
    except ValueError:
        raise ValueError(
            f"Invalid format for 'event start time' in event '{event['title']}'. Expected format is '{DATE_FORMAT}'.")

    try:
        datetime.datetime.strptime(event['event end time'], DATE_FORMAT)
    except ValueError:
        raise ValueError(
            f"Invalid format for 'event end time' in event '{event['title']}'. Expected format is '{DATE_FORMAT}'.")


def validate_yaml_data(data: Dict[str, Any]) -> None:
    """Validates the calendar name, timezone, and events from the YAML file."""
    if 'calendar_name' not in data:
        raise ValueError("Missing 'calendar_name' in YAML file.")

    if 'timezone' not in data:
        raise ValueError("Missing 'timezone' in YAML file.")

    if 'events' not in data or not isinstance(data['events'], list):
        raise ValueError("Missing or invalid 'events' list in YAML file.")

    # Validate each event
    for event in data['events']:
        validate_event(event)


def get_or_create_calendar(service: Resource, calendar_name: str) -> Optional[str]:
    """Fetches the custom calendar or creates it if it does not exist."""
    try:
        # List all calendars
        calendars = service.calendarList().list().execute()

        # Search for the custom calendar
        for calendar in calendars['items']:
            if calendar['summary'] == calendar_name:
                return calendar['id']

        # If not found, create the calendar
        calendar = {
            'summary': calendar_name,
            'timeZone': TIMEZONE
        }
        created_calendar = service.calendars().insert(body=calendar).execute()
        return created_calendar['id']

    except HttpError as error:
        print(f"An error occurred: {error}")
        return None


def create_event(service: Resource, calendar_id: str, event_data: Dict[str, Any], timezone_str: str) -> None:
    """Creates an event in the Google Calendar with optional attendees."""
    try:
        # Use zoneinfo to localize the event times to the given timezone
        timezone = ZoneInfo(timezone_str)

        # Convert the start and end time strings from the YAML into valid datetime objects
        start_datetime = datetime.datetime.strptime(event_data['event start time'], DATE_FORMAT)
        end_datetime = datetime.datetime.strptime(event_data['event end time'], DATE_FORMAT)

        # Localize the datetime to the given timezone using zoneinfo
        start_datetime = start_datetime.replace(tzinfo=timezone)
        end_datetime = end_datetime.replace(tzinfo=timezone)

        # Prepare the event data
        event = {
            'summary': event_data['title'],
            'description': event_data['description'],  # Supports HTML for description
            'start': {
                'dateTime': start_datetime.isoformat(),
                'timeZone': timezone_str,
            },
            'end': {
                'dateTime': end_datetime.isoformat(),
                'timeZone': timezone_str,
            }
        }

        # Add attendees only if they are present in the YAML data
        if 'attendees' in event_data:
            attendees = [{'email': email} for email in event_data['attendees']]
            event['attendees'] = attendees

        # Insert event into the specified calendar
        created_event = service.events().insert(calendarId=calendar_id, body=event).execute()
        print(f"Event created: {created_event.get('htmlLink')}")

    except KeyError as error:
        print(f"Missing required field: {error}")
    except HttpError as error:
        print(f"An error occurred: {error}")
    except Exception as e:
        print(f"Unexpected error: {e}")


def main() -> None:
    """Main function to load events and create Google Calendar entries."""
    try:
        # Load the events and calendar name from the YAML file
        data = load_yaml_file(YAML_FILE_PATH)

        # Validate the data from YAML
        validate_yaml_data(data)

        calendar_name = data.get('calendar_name', 'Work')  # Default to 'Work' custom calendar if not specified
        timezone_str = data.get('timezone', 'America/New_York')  # Default to 'America/New_York' if not specified

        # Authenticate and build the Google Calendar service
        service = authenticate_google()

        # Get or create the custom calendar
        calendar_id = get_or_create_calendar(service, calendar_name)
        if not calendar_id:
            print("Failed to find or create the calendar.")
            return

        # Loop through events and create them
        for event_data in data['events']:
            create_event(service, calendar_id, event_data, timezone_str)

    except FileNotFoundError as fnf_error:
        print(fnf_error)
    except ValueError as val_error:
        print(f"Validation Error: {val_error}")
    except Exception as error:
        print(f"An unexpected error occurred: {error}")


# Prevents main() from being executed during imports.
if __name__ == "__main__":
    """ This is executed when run from the command line """
    main()
