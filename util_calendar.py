
from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']

"""Shows basic usage of the Google Calendar API.
Prints the start and name of the next 10 events on the user's calendar.
"""
creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
calendar_token_loc = 'creds_google/calendar-token.pickle'
credentials_loc = 'creds_google/credentials.json'

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


def insert_event(event: dict)-> str:
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
    recurring_event = service.events().insert(calendarId='primary', body=event).execute()

    return recurring_event['id']

