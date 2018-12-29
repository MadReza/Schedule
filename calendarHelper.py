from __future__ import print_function
import datetime
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/calendar'

def create_event(calendarID, summary, location, description, startTime, endTime, timeZone="America/Montreal", colorId="NA"):
    """
    returns events resource: https://developers.google.com/calendar/v3/reference/events#resource
    More Info: https://developers.google.com/calendar/v3/reference/events/insert
    Recurrence info: https://developers.google.com/calendar/concepts/events-calendars#recurring_events
    """
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('calendar', 'v3', http=creds.authorize(Http()))

    event = {
        'summary': summary,
        'location': location,
        'description': description,
        'start': {
            'dateTime': startTime, #'2015-05-28T09:00:00'
            'timeZone': timeZone,
        },
        'end': {
            'dateTime': endTime, #'2015-05-28T17:00:00'
            'timeZone': timeZone,
        },
        'reminders': {
            'useDefault': True,
        },
    }

    if colorId != "NA":
        event['colorId'] = colorId

    event = service.events().insert(calendarId=calendarID, body=event).execute()
    print('Event created: %s' % (event.get('htmlLink')))

def create_calendar(summary, description, timeZone="America/Montreal"):
    """
    returns calendar object: https://developers.google.com/calendar/v3/reference/calendars#resource
    """
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('calendar', 'v3', http=creds.authorize(Http()))

    calendar = {
        "summary": summary,
        "description": description,
        "timeZone": timeZone,
    }

    created_calendar = service.calendars().insert(body=calendar).execute()

    return created_calendar

def get_calendars():
    """
    returns dictionary of summary/title as key and id as value
    """
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('calendar', 'v3', http=creds.authorize(Http()))

    page_token = None
    calendars = {}
    while True:
        calendar_list = service.calendarList().list(pageToken=page_token).execute()
        for calendar_list_entry in calendar_list['items']:
            calendars[calendar_list_entry['summary']] = calendar_list_entry['id']
        page_token = calendar_list.get('nextPageToken')
        if not page_token:
            break

    return calendars

def main():
    """
    Used for Testing Only
    Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('calendar', 'v3', http=creds.authorize(Http()))


    #c = create_calendar("GIM", "Cegep Gim Courses", "America/Montreal")

    #calendar_list = get_calendars()

    #calendar = service.calendars().get(calendarId=calendar_list['GIM']).execute()
    #print(calendar['summary'])


    #create_event(calendar_list['GIM'], "Test", "1001 Sherbr", "Description stuff", "2019-01-01T09:00:00", "2019-01-01T10:00:00","America/Montreal", "NoColor")

    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    print('Getting the upcoming 10 events')
    events_result = service.events().list(calendarId='primary', timeMin=now,
                                        maxResults=10, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        print('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        print(start, event['summary'])

if __name__ == '__main__':
    main()
