"""
Google Calendar backup script
by sondo01@gmail.com

This script is based on multiple suggestions from forums and especially from Google Developer page and https://stackoverflow.com/a/30811628

USE WITH YOUR OWN RISK.

29-Aug-2019: Scratch script for backup Google Calendar (similar to Export function from Google Calendar)

This script will:
a) Connect to Google Calendar
b) Export all calendars or calendars those are in Backup_calendars tuple (if provided)
c) Each calendar will be save in iCalendar format (.ics)
d) Reminders are handled
e) This script is tested on Linux, Mac OS X and supposed to work on Windows. Carriage return keys ("\r\n") could cause trouble on Windows. I don't have Windows to test.

This script will NOT handle (implement them yourself if needed):
a) ATTENDEE
b) Conference and other info

To setup, I followed this reference:
https://developers.google.com/google-apps/calendar/quickstart/python

To run:
$ python gg_backup.py

References:
1. https://developers.google.com/google-apps/calendar/quickstart/python
2. https://developers.google.com/google-apps/calendar/v3/reference/events/list
3. https://developers.google.com/google-apps/calendar/v3/reference/calendarList/list
"""


from __future__ import print_function
import httplib2
import re, os
import pytz, datetime
import textwrap
from pathlib import Path

# from apiclient import discovery
# Commented above import statement and replaced it below because of
# reader Vishnukumar's comment
# Src: https://stackoverflow.com/a/30811628

import googleapiclient.discovery as discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# Provide explite Calendar names if don't want to backup everything

Backup_calendars=[]

# Relative folder to this script. Change to suit your need
BACKUPDIR = "GGBK"


# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
CLIENT_SECRET_FILE = 'Path_to_your/credentials.json'
APPLICATION_NAME = 'Reminder'

now = datetime.datetime.now()
DTSTAMP = now.strftime ("%Y-%m-%dT%H:%M:%S") + "+10:00"

def dateTime_to_Z(t):
    # Google export event to Zulu time so we have to convert event time to this format
    st = t.split('+')
    local = pytz.timezone ("Australia/Melbourne")
    naive = datetime.datetime.strptime (st[0], "%Y-%m-%dT%H:%M:%S")
    local_dt = local.localize(naive, is_dst=None)
    utc_dt = local_dt.astimezone(pytz.utc)
    return utc_dt.strftime ("%Y%m%dT%H%M%SZ")

def date_to_ics(t):
    return t.replace("-", "")
def dtstamp_to_ics(t):
    st = t.split(".")
    st1 = st[0].replace("-", "")
    return st1.replace(":", "") + "Z"
def newline_to_text(x):
    y = x.replace(",", "\,")
    y1 = y.replace(";", "\;")
    y2 = y1.replace("\n", "\\n")
    y3 = '\r\n'.join(textwrap.wrap(y2,60))
    return y3
def get_reminder(x, email):
    rem_ics = ""
    y = x['overrides']
    try:
        reminder_lenth = len(y)
    except:
        reminder_lenth = 0
    if ( reminder_lenth > 0 ):
        for elm in y:
            if ( elm['method'] == 'email'):
                action = 'EMAIL'
                xemail = "SUMMARY:Alarm notification\r\n"
                xemail += "ATTENDEE:mailto:" + email + "\r\n"
            else:
                action = 'DISPLAY'
                xemail = ""
            hm = divmod(elm['minutes'], 60)
            min = hm[1]
            dh = divmod(hm[0], 24)
            day = dh[0]
            hr = dh[1]
            trigger_time = "-P%s" % dh[0] +"D"
            if ( dh[1] != 0 or hm[1] !=0 ):
                trigger_time += "T%s" % dh[1] +"H%s" % hm[1] +"M0S"
            rem_ics += "BEGIN:VALARM\r\n"
            rem_ics += "ACTION:" + action + "\r\n"
            rem_ics += "DESCRIPTION:This is an event reminder\r\n"
            rem_ics += xemail
            rem_ics += "TRIGGER:" + trigger_time  + "\r\n"
            rem_ics += "END:VALARM\r\n"
        # print(rem_ics)
    return rem_ics

def gg_to_ics(ev):
    ics_ev = "BEGIN:VEVENT\r\n"
    if ( ev['start'].keys()[0] == "dateTime" ):
        ics_ev += "DTSTART:" + dateTime_to_Z(ev['start'].get('dateTime')) + "\r\n"
        ics_ev += "DTEND:" + dateTime_to_Z(ev['end'].get('dateTime')) + "\r\n"
    else:
        ics_ev += "DTSTART;VALUE=DATE:" + date_to_ics(ev['start'].get('date')) + "\r\n"
        ics_ev += "DTEND;VALUE=DATE:" + date_to_ics(ev['end'].get('date')) + "\r\n"
    ics_ev += "DTSTAMP:" + dateTime_to_Z(DTSTAMP) + "\r\n"
    ics_ev += "UID:" + ev['iCalUID'] + "\r\n"
    ics_ev += "CREATED:" + dtstamp_to_ics(ev['created']) + "\r\n"
    try:
        description = newline_to_text(ev['description'])
    except:
        description = ""
    ics_ev += "DESCRIPTION:" + description + "\r\n"
    ics_ev += "LAST-MODIFIED:" + dtstamp_to_ics(ev['updated']) + "\r\n"
    ics_ev += "LOCATION:" + "\r\n"
    ics_ev += "SEQUENCE:%s" % ev['sequence'] + "\r\n"
    ics_ev += "STATUS:" + ev['status'].upper() + "\r\n"
    ics_ev += "SUMMARY:" + ev['summary'] + "\r\n"
    try:
        transp = ev['transparency'].upper()
    except:
        transp = "OPAQUE"
    ics_ev += "TRANSP:" + transp + "\r\n"
    try:
        check_reminder = len(ev['reminders']['overrides'])
    except:
        check_reminder = 0
    if ( check_reminder == 0 ):
        ics_ev += "X-APPLE-TRAVEL-ADVISORY-BEHAVIOR:AUTOMATIC\r\n"
    else:
        ics_ev += get_reminder(ev['reminders'], ev['creator']['email'])
    ics_ev += "END:VEVENT" + "\r\n"
    return ics_ev

def get_header(x):
    h = "BEGIN:VCALENDAR\r\n"
    h += "PRODID:-//Google Inc//Google Calendar 70.9054//EN\r\n"
    h += "VERSION:2.0\r\n"
    h += "CALSCALE:GREGORIAN\r\n"
    h += "METHOD:PUBLISH\r\n"
    h += "X-WR-CALNAME:" + x + "\r\n"
    h += "X-WR-TIMEZONE:Australia/Melbourne\r\n"
    return h

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,'calendar-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def getEventsFromGoogle():
    """Shows basic usage of the Google Calendar API.

    Creates a Google Calendar API service object and outputs a list of the next
    10 events on the user's calendar.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)
    page_token = None
    #
    calendar_ids = {}
    while True:
        calendar_list = service.calendarList().list(pageToken=page_token).execute()
        for calendar_list_entry in calendar_list['items']:
            if (len(Backup_calendars) > 0):
                if calendar_list_entry['summary'] in Backup_calendars:
                    calendar_ids[calendar_list_entry['summary']] = calendar_list_entry['id']
            else:
                calendar_ids[calendar_list_entry['summary']] = calendar_list_entry['id']

        page_token = calendar_list.get('nextPageToken')
        if not page_token:
            break

    for calendar_name, calendar_id in calendar_ids.items():
        bak_folder = os.path.dirname(__file__)
        ffname = calendar_name + "-" + now.strftime ("%Y-%m-%d") + ".ics"
        fname = os.path.join(bak_folder, BACKUPDIR, ffname)
        print(fname)
        f = open(fname, "a")
        header_file = get_header(calendar_name)
        f.write(header_file)
        eventsResult = service.events().list(
            calendarId=calendar_id,
            timeZone='Australia/Melbourne').execute()
        events = eventsResult.get('items', [])
        for event in events:
            try:
                a = gg_to_ics(event)
                f.write(a)
            except:
                a=""

        f.write("END:VCALENDAR")
        f.close()

def main():
    getEventsFromGoogle()

if __name__ == '__main__':
    main()
