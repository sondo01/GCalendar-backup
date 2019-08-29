# GCalendar-backup
Python script to Backup Google Calendar to .ics file (similar to Google Calendar Export).

This script is based on multiple suggestions from different forums and especially from Google Developer page and https://stackoverflow.com/a/30811628

USE WITH YOUR OWN RISK.

29-Aug-2019: Scratch script for backup Google Calendar (similar to Export function from Google Calendar)

This script will:

1. Connect to Google Calendar
2. Export all calendars or calendars those are in Backup_calendars tuple (if provided)
3. Each calendar will be save in iCalendar format (.ics)
4. Reminders are handled
5. This script is tested on Linux, Mac OS X and supposed to work on Windows. Carriage return keys ("\r\n") could cause trouble on Windows. I don't have Windows to test.

This script will NOT handle (implement them yourself if needed):
1. ATTENDEE
2. Conference and other info

To setup, I followed this reference:
https://developers.google.com/google-apps/calendar/quickstart/python

To run:
1. Open the script, edit BACKUPDIR to point to where you want to save the .ics files
2. Specify which Calendar you want to backup in Backup_calendars tuple. By default, all calendars will be backup
3. run: $ python gg_backup.py

References:
1. https://developers.google.com/google-apps/calendar/quickstart/python
2. https://developers.google.com/google-apps/calendar/v3/reference/events/list
3. https://developers.google.com/google-apps/calendar/v3/reference/calendarList/list
