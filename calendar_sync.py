import win32com.client
from time import strptime, strftime, time, gmtime
from datetime import date, timedelta, datetime
from getpass import getpass
from os import environ
from ConfigParser import SafeConfigParser
import gdata.calendar.service, gdata.service, gdata.calendar, gdata.calendar.client
import atom.service, atom, atom.data

import pdb


class Event(object):

    outlook_date_format = "%m/%d/%y %H:%M:%S"
    google_cal_format = "%Y-%m-%dT%H:%M:%S.000"
    google_date_format = "%Y-%m-%dT%H:%M:%S-04:00"

    def __init__(self, start = None, end = None, text="Busy!"):
        # Format to date object
        self.start = start
        self.end = end
        self.text = text

    def from_outlook_fmt(self):
        # Format to date object
        self.start = strptime(self.start, self.outlook_date_format)
        self.end = strptime(self.end, self.outlook_date_format)

    def from_gcal_fmt(self):
        self.start = "-".join(self.start.split("-")[0:-1])
        self.end = "-".join(self.end.split("-")[0:-1])
        self.start = strptime(self.start, self.google_cal_format)
        self.end = strptime(self.end, self.google_cal_format)

    def to_google_fmt(self):
        # Format to google calendar format
        self.start = strftime(self.google_date_format, self.start)
        self.end = strftime(self.google_date_format, self.end)

    def outlook_to_google(self):
        self.from_outlook_fmt()
        self.to_google_fmt()

    def gcal_to_google(self):
        self.from_gcal_fmt()
        self.to_google_fmt()

    def __str__(self):
        return self.start+self.end+self.text


class GCal(object):

    def log_in(self, username, password):
        self.cal_client = gdata.calendar.client.CalendarClient(source='Google-Calendar-sync')
        self.cal_client.ClientLogin(username, password, self.cal_client.source)
        return self.cal_client

    def insert_event(self, start_time=None, end_time=None):
        event = gdata.calendar.data.CalendarEventEntry()
        event.title = atom.data.Title(text="Busy!")
        event.content = atom.data.Content(text="")

        if start_time is None:
            print "No start time, can not sync event"
        event.when.append(gdata.data.When(start=start_time, end=end_time))

        # Enter event into google calendar
        new_event = self.cal_client.InsertEvent(event)

        # Add an alert
        g_reminder = gdata.data.Reminder(minutes='60')
        g_reminder.method = 'alert'
        new_event.when[0].reminder.append(g_reminder)
        self.cal_client.Update(new_event)


        print 'New single event inserted: %s' % (new_event.id.text,)
        #print '\tEvent edit URL: %s' % (new_event.GetEditLink().href,)
        #print '\tEvent HTML URL: %s' % (new_event.GetHtmlLink().href,)

def evt_exist(evt, evt_list):
    for event in evt_list:
        if evt.__str__() == event.__str__():
            return True
    return False

def approve_appointment(appointment):
    print "-------------------------------------------------------------"
    print "appointment info:",appointment.text
    print "start time:",appointment.start
    print "end time:", appointment.end
    response = raw_input("\nSync this event? (Y/n): ").lower()
    if response != "y" and response != "n":
        response = "y"
    if response == "y":
        return True
    else:
        return False

CALENDAR = 9

def get_dates_in_range(cal_client, start_date, end_date):
    # date format is yyyy-mm-dd

    gevent_list = []

    print 'Date range query for events on Primary Calendar: %s to %s' % (
        start_date, end_date,)
    query = gdata.calendar.client.CalendarEventQuery(start_min=start_date, start_max=end_date)
    feed = cal_client.GetCalendarEventFeed(q=query)
    for event in feed.entry:
        title = event.title.text
        e = Event(text=title)
        for e_when in event.when:
            e.start = e_when.start
            e.end = e_when.end
        e.gcal_to_google()
        gevent_list.append(e)
    sorted_gevents = sorted(gevent_list, key=lambda x:x.__str__())
    return sorted_gevents


if __name__ == "__main__":


    config = SafeConfigParser()
    config.read("config")
    username = config.get("user", "username")
    http_proxy = config.get("proxy", "http_proxy")
    https_proxy = config.get("proxy", "https_proxy")
    environ['http_proxy']=http_proxy
    environ['https_proxy']=https_proxy

    outlook_event_list = []

    gcal = GCal()
    password = getpass("Enter Google password: ")
    gclient = gcal.log_in(username, password)

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    appointments = namespace.GetDefaultFolder(CALENDAR).Items
    appointments.Sort("[Start]")
    appointments.IncludeRecurrences = True
    today = date.today()
    outlook_span_start = strftime("%m/%d/%y 00:00 %p", today.timetuple())
    outlook_span_stop = strftime("%m/%d/%y 00:00 %p", (today+timedelta(weeks=2)).timetuple())
    restriction_span = "[Start] >= '"+outlook_span_start+"' and [Start] <= '"+outlook_span_stop+"'"
    appointments = appointments.Restrict(restriction_span)

    google_span_start = strftime("%Y-%m-%d", today.timetuple())
    google_span_stop = strftime("%Y-%m-%d", (today+timedelta(weeks=2)).timetuple())
    google_events = get_dates_in_range(gclient, google_span_start, google_span_stop)




    for appointment in appointments:
        start = appointment.start.Format()
        end = appointment.end.Format()
        event = Event(start, end)
        event.outlook_to_google()
        outlook_event_list.append(event)

    for evt in outlook_event_list:
        if not evt_exist(evt, google_events):
            if approve_appointment(evt):
                gcal.insert_event(evt.start, evt.end)
            else:
                print "Not syncing"

    print "All events synchronized"






