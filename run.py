#!/usr/bin/env python3
import pandas as pd
import pygsheets
import sys
import os
from datetime import datetime, timedelta

import pickle
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from typing import Union, Callable

# Toggle this for using the TEST sheet
DEBUG = False

dict_type = Union[dict, pd.Series, pd.DataFrame]


def get_google_cal_handler():
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/calendar']

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.

    dir_name = os.path.dirname(os.path.realpath(sys.argv[0]))

    calendar_token_loc = os.path.join(dir_name,
                                      'creds_google/calendar-token.pickle')
    credentials_loc = os.path.join(dir_name, 'creds_google/credentials.json')

    if os.path.exists(calendar_token_loc):
        with open(calendar_token_loc, 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_loc, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(calendar_token_loc, 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)
    return service


service = get_google_cal_handler()


class GoogleEvent():
    mandatory_entries: "list[str]" = ['Name', 'Duration',
                                      'Time', 'Start Date']

    def __init__(self, event: dict_type):
        self.event: pd.Series = event
        self.event_code: str = self.event['Google Calendar Invite Code']
        self.event_name = event['Name']
        self.action: str = str(event['Control Action']).strip()
        self.event['Control Action'] = ''

    def execute(self) -> dict_type:
        # Match to the right control action
        control_action_to_funcs: "dict[str,Callable]" = {
            "create": self.create_event,
            "update": self.update_event,
            "delete": self.delete_event,

        }

        func = control_action_to_funcs.get(self.action)
        if func is not None:
            func()

        return self.event

    def create_event(self):

        if self.event_code != '':
            print(f"Event '{self.event_name}' already scheduled")
            # It will return the same event, though not Nonetype as originally
            # conceived
            return self.event

        event_bluprnt: dict = create_event_json(self.event)
        print(event_bluprnt)

        """
        event = {
        'summary': 'Appointment',
        'location': 'Somewhere',
        'start': {
            'dateTime': '2021-01-20T10:00:00.000',
            'timeZone': 'GMT+3:00'
        },
        'end': {
            'dateTime': '2021-01-20T10:25:00.000',
            'timeZone': 'GMT+3:00'
        },
        'recurrence': [
            'RRULE:FREQ=DAILY;UNTIL=20210122T235959Z',
        ],
        }
        """
        updated_event = service.events().insert(calendarId='primary',
                                                body=event_bluprnt)
        updated_event = updated_event.execute()
        self.event['Google Calendar Invite Code'] = updated_event['id']

    def update_event(self):

        if self.event_code == '':
            print(f"Can't update Event '{self.event_name}'."
                  "It doesn't seem to be scheduled!")
            return None

        proposed_event = create_event_json(self.event)
        event_code = self.event_code

        # Get the last updated event
        current_event = service.events().get(calendarId='primary',
                                             eventId=event_code).execute()

        # Compare both events
        diffkeys = [k for k in proposed_event
                    if proposed_event[k] != current_event[k]]

        if len(diffkeys) > 0:
            # Update the dict object
            current_event.update(proposed_event)
            updated_event = service.events().update(calendarId='primary',
                                                    eventId=event_code,
                                                    body=current_event)
            # Not to be confused for this class's bound method `execute`
            updated_event = updated_event.execute()

            print(f"The event was updated at {updated_event['updated']}")

            for k in diffkeys:
                print(f"Field {k}: {current_event[k]}")
        else:
            print(f"No change to {current_event['summary']} was necessary")

    def delete_event(self):

        if self.event_code == '':
            print(f"Can't delete Event '{self.event_name}'."
                  "It doesn't seem to be scheduled!")
            return None

        try:
            service.events().delete(calendarId='primary',
                                    eventId=self.event_code).execute()
        except HttpError as e:
            print(e)
            error = e.__getattribute__("error_details")
            print(f"Event '{self.event_name}: {error}'")

            if error == "Resource has been deleted":
                # Get rid of the invite code
                self.event['Google Calendar Invite Code'] = ""
        else:
            print(f"Event {self.event_name} DELETED")
            self.event['Google Calendar Invite Code'] = ""

    @classmethod
    def _mandatory_entries_missing(cls, event: dict_type) -> bool:
        missing: pd.Series = (event[cls.mandatory_entries] == '')
        return bool(missing.any())

    # Factory Method
    @classmethod
    def from_series(cls, event: dict_type):
        event_name = event['Name']

        if pd.Series(event == "").all():
            return None

        # check for missing entries
        if cls._mandatory_entries_missing(event) is True:
            print(f"Fields are missing for '{event_name}'")
            return None

        return cls(event)


# Helper Functions
def create_event_json(series: dict) -> dict:
    # Create a dictionary of the event's details

    # Extract the Date, Time & Duration
    start_date = datetime.strptime(series['Time'] + " "
                                   + series['Start Date'],
                                   "%I:%M %p %d-%b-%Y")
    duration = datetime.strptime(series['Duration'], "%H:%M")

    # Calculate the End Date
    end_date = start_date + timedelta(minutes=duration.minute,
                                      hours=duration.hour)

    timezone = '+03:00'

    # Convert to Text ISO Format
    start_date = datetime.isoformat(start_date, timespec='seconds') + timezone
    end_date = datetime.isoformat(end_date, timespec='seconds') + timezone

    event_bluprnt = {
      'summary': series['Name'],
      'location': 'Virtual',
      'start': {
        'dateTime': start_date,
        'timeZone': 'GMT' + timezone
      },
      'end': {
        'dateTime': end_date,
        'timeZone': 'GMT' + timezone
      },
      'reminders': {
            'useDefault': False,
            # Overrides can be set if and only if useDefault is false.
            'overrides': [
                {
                    'method': 'popup',
                    'minutes': 10
                },
            ]
            }
    }

    # Generate RRULE standard frequency pattern
    rrule = generate_rrule_pattern(series["Frequency"], series['End Date'])
    if rrule != "":
        event_bluprnt['recurrence'] = [rrule]

    return event_bluprnt


def generate_rrule_pattern(freq_code, end_date=''):
    # sUnday, Monday, Tuesday, Wednesday, tHursday, Friday, Saturday
    # FREQ=WEEKLY;BYDAY=SU,MO,TU,WE,TH,FR,SA
    DAY_CODES = {'U': 'SU', 'M': 'MO',
                 'T': 'TU', 'W': 'WE', 'H': 'TH',
                 'F': 'FR', 'S': 'SA'}

    freq_code = freq_code.upper()
    rrule = ''

    if freq_code not in ["", "ONCE"]:
        if freq_code in ["WEEKLY", "DAILY"]:
            rrule = f'RRULE:FREQ={freq_code}'
        else:
            repet_pattern = []
            for letter in DAY_CODES:
                if letter in freq_code:
                    repet_pattern.append(DAY_CODES[letter])

            if len(repet_pattern) > 0:
                repet_pattern = ",".join(repet_pattern)
                rrule = f'RRULE:FREQ=WEEKLY;BYDAY={repet_pattern};INTERVAL=1'

    if end_date != '':
        # Force Daily Pattern if "end date" exists
        if rrule == '':
            rrule = 'RRULE:FREQ=DAILY'

        # Add something like ;UNTIL=20210124T210000Z
        end_date = datetime.strptime(end_date, "%d-%b-%Y")
        end_date = datetime.strftime(end_date, "%Y%m%dT235959Z")
        rrule += f";UNTIL={end_date}"

    return rrule


if __name__ == "__main__":
    SHEET_NAME = "STUDY-PLAN"

    if DEBUG is True:
        SHEET_NAME = "TEST | STUDY-PLAN"

    SPREADSHEET_ID = '1GDXzzTD1dBnXWcpjPXqfIjpLwxHtXVQtjwfLhCaZNHA'

    # Get real path of executable directory
    dir_name = os.path.dirname(os.path.realpath(sys.argv[0]))

    secretpath = os.path.join(dir_name, 'creds_google/service_account.json')

    gc = pygsheets.authorize(service_file=secretpath)

    workbook = gc.open_by_key(SPREADSHEET_ID)

    study_sheet = workbook.worksheet_by_title(SHEET_NAME)

    # Retrieve the data from the sheets
    df = study_sheet.get_as_df(start="A", end="I")
    # Have an empty dataframe to write back to the sheets
    new_df = pd.DataFrame(columns=df.columns)

    # Process the table
    for index in range(len(df)):
        row = df.loc[index]
        event = GoogleEvent.from_series(row)
        if event is not None:
            # Process according to its control action
            row = event.execute()
        new_df = new_df.append(row, ignore_index=True)

    # Write back to the sheet
    study_sheet.set_dataframe(new_df, start='A2', copy_head=False)
