import json
import urllib2

from lxml import html
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
from pyasn1.compat.dateandtime import strptime

SCOPES = ['https://www.googleapis.com/auth/calendar']


def get_calendar(service):
    page_token = None
    while True:
        calendar_list = service.calendarList().list(pageToken=page_token).execute()
        for calendar_list_entry in calendar_list['items']:
            if calendar_list_entry['summary'] == "CSTA Conference":
                return calendar_list_entry['id']
        page_token = calendar_list.get('nextPageToken')
        if not page_token:
            break

    calendar = {
        'summary': 'CSTA Conference',
        'timeZone': 'America/New_York'
    }

    created_calendar = service.calendars().insert(body=calendar).execute()
    return created_calendar['id']


def get_service():
    credentials = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            credentials = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            credentials = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(credentials, token)

    return build('calendar', 'v3', credentials=credentials)


def get_csta_events(service, csta_calendar):
    all_events = []
    page_token = None
    while True:
        page_events = service.events().list(calendarId=csta_calendar, pageToken=page_token).execute()
        all_events.extend(page_events['items'])
        page_token = page_events.get('nextPageToken')
        if not page_token:
            break
    return all_events


def find_event_by_name(event_list, title):
    for event in event_list:
        if event['summary'] == title:
            return event
    return None


def date_time_from(date, time):
    return strptime(date + time, "%m/%d/%y%I:%M %p").isoformat()


def get_website_events():
    events = []
    response = urllib2.urlopen(
        "https://www.cvent.com/events/2020-csta-annual-conference/agenda-236d288a403041f8a7a935b0bd74131c.aspx")
    tree = html.fromstring(response.read().decode('utf-8'))
    sessions = tree.xpath('//div[@class="reg-matrix-header-container"]')

    for session in sessions:
        title = session.xpath('div/h3/text()')[0].strip()
        description = session.xpath('div/div[@class="session-description"]/text()')[0].strip()

        website_event = {'title': title, 'description': description}
        keys = session.xpath('div/div[@class="session-info"]/p/span')
        for key in keys:
            value = key.xpath('span/text()')
            if value:
                website_event[key.xpath('text()')[0]] = value[0]
            elif key.get('class') == 'speaker-name':
                if 'presenters' not in website_event:
                    website_event['presenters'] = []
                for presenter in key.xpath('a/text()'):
                    website_event['presenters'].append(presenter)

        events.append(website_event)
    return events


def get_categories():
    with open('categories.json') as f:
        categories = json.load(f)

    for category in categories:
        category['key'] = u'{}|'.format(category['name'].lower().replace(' ', '-'))
        category['label'] = u'{}: '.format(category['name'])
        category['values'] = []

    return categories


def put_filter_options(categories):
    with open('filter.json', 'w') as f:
        json.dump(categories, f)


def main():
    service = get_service()
    csta_calendar = get_calendar(service)
    events = get_csta_events(service, csta_calendar)
    events_to_remove = list(events)
    website_events = get_website_events()
    categories = get_categories()

    for website_event in website_events:
        existing_event = find_event_by_name(events, website_event['title'])
        if existing_event:
            events_to_remove.remove(existing_event)

        description = u'{}\n\n'.format(website_event['description'])
        if 'presenters' in website_event:
            description += u'{}\t{}\n'.format("Presenter(s):", ', '.join(website_event['presenters']))

        for category in categories:
            if category['label'] in website_event:
                description += u'{}\t{}\n'.format(category['label'], website_event[category['label']])

        event = {
            'summary': website_event['title'],
            'description': description,
            'transparency': 'transparent',
            'reminders': {
                'useDefaults': False
            },
            'start': {
                'dateTime': date_time_from(website_event['Start Date: '], website_event['Start Time: ']),
                'timeZone': 'America/New_York'
            },
            'end': {
                'dateTime': date_time_from(website_event['Start Date: '], website_event['End Time: ']),
                'timeZone': 'America/New_York'
            }
        }

        if 'Location: ' in website_event:
            event['location'] = website_event['Location: ']

        shared_properties = []
        for category in categories:
            if category['label'] in website_event:
                category['values'] = list(
                    set(category['values']) |
                    set([value.strip() for value in website_event[category['label']].split(",")])
                )
                for category_value in category['values']:
                    shared_properties.append(u'{}{}'.format(category['key'], category_value))
            else:
                category['values'] = list(
                    set(category['values']) |
                    set(['Undefined'])
                )
                shared_properties.append(u'{}{}'.format(category['key'], 'Undefined'))

        put_filter_options(categories)

        if shared_properties:
            event['extendedProperties'] = {
                'shared': {
                }
            }
            for property_name in shared_properties:
                event['extendedProperties']['shared'][property_name] = 'yes'

        if existing_event:
            event = service.events().patch(calendarId=csta_calendar, eventId=existing_event['id'], body=event).execute()
            print 'Event updated: %s' % (event.get('htmlLink'))
        else:
            event = service.events().insert(calendarId=csta_calendar, body=event).execute()
            print 'Event created: %s' % (event.get('htmlLink'))

    if events_to_remove:
        for event in events_to_remove:
            service.events().delete(calendarId=csta_calendar, eventId=event['id']).execute()
            print 'Event deleted: %s' % (event.get('summary'))


if __name__ == '__main__':
    main()
