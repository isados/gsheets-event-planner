#!/usr/bin/env python3
import pandas as pd
import pygsheets
import sys
import os
from datetime import datetime, timedelta
import util_calendar

# Toggle this for using the TEST sheet
DEBUG = False


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
            'useDefault': 'False',
            # Overrides can be set if and only if useDefault is false.
            'overrides': [
                {
                    'method': 'popup',
                    'minutes': '10'
                },
            ]
            }
    }

    # Generate RRULE standard frequency pattern
    rrule = generate_rrule_pattern(series["Frequency"], series['End Date'])
    if rrule != "":
        event_bluprnt['recurrence'] = [rrule]

    return event_bluprnt


def initiate_event(event_bluprnt):
    # Pubish the event
    cal_code = util_calendar.insert_event(event_bluprnt)
    return cal_code


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
    df = study_sheet.get_as_df(start="A", end="H")
    # Have an empty dataframe to write back to the sheets
    new_df = pd.DataFrame(columns=df.columns)

    # Process the table
    for index in range(len(df)):
        event = df.loc[index]
        event_code = event['Google Calendar Invite Code']
        control_action = event['Control Action'].lower()

        if event_code == '':
            mandatory_entries: pd.DataFrame = event[['Name', 'Duration',
                                                    'Time', 'Start Date']]
            if (mandatory_entries == "").any():
                print(f"Fields are missing for {event['Name']}")
            else:
                # Create Event
                event_bluprnt = create_event_json(event)
                print(event_bluprnt)
                event_code = initiate_event(event_bluprnt)

                event['Google Calendar Invite Code'] = event_code

        else:
            if control_action == 'delete':
                util_calendar.delete_event(event_code)
                event['Google Calendar Invite Code'] = ""
                event['Control Action'] = ''
                print(f"Event : {event['Name']} was deleted")

            if control_action == 'update':
                event_bluprnt = create_event_json(event)
                util_calendar.update_event(event_bluprnt, event_code)
                event['Control Action'] = ''

        new_df = new_df.append(event, ignore_index=True)

    # Retain the size of the dataframe as the old one.
    for _ in range(0, len(df)-len(new_df)):
        new_df = new_df.append(pd.Series(), ignore_index=True)
        new_df.iloc[-1] = ""

    # Write back to the sheet
    study_sheet.set_dataframe(new_df, start='A2', copy_head=False)
