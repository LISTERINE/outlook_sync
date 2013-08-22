import win32com.client
from time import strptime, strftime, time, gmtime
from datetime import date, timedelta, datetime
from getpass import getpass

import gdata.calendar.service, gdata.service, atom.service, gdata.calendar, atom, atom.data

import pdb


"""

this is what is called by GCal.insert_event
see if we can swap this out useing the put method that this InsertEvent uses.
should be something like cal_client.put(...)

def InsertEvent(self, new_event, insert_uri, url_params=None, escape_params=True):
    #Adds an event to Google Calendar.

    #Args: 
    #  new_event: atom.Entry or subclass A new event which is to be added to 
    #            Google Calendar.
    #  insert_uri: the URL to post new events to the feed
    #  url_params: dict (optional) Additional URL parameters to be included
    #              in the insertion request. 
    #  escape_params: boolean (optional) If true, the url_parameters will be
    #                 escaped before they are included in the request.

    #Returns:
    #  On successful insert,  an entry containing the event created
    #  On failure, a RequestError is raised of the form:
    #    {'status': HTTP status code from server, 
    #     'reason': HTTP reason from the server, 
    #     'body': HTTP body of the server's response}
    #

    return self.Post(new_event, insert_uri, url_params=url_params,
                     escape_params=escape_params, 
                     converter=gdata.calendar.CalendarEventEntryFromString)



------------------------------------
Methods inherited from atom.service.AtomService:
PrepareConnection(self, full_uri)
Opens a connection to the server based on the full URI.
 
Examines the target URI and the proxy settings, which are set as 
environment variables, to open a connection with the server. This 
connection is used to make an HTTP request.
 
Args:
  full_uri: str Which is the target relative (lacks protocol and host) or
  absolute URL to be opened. Example:
  'https://www.google.com/accounts/ClientLogin' or
  'base/feeds/snippets' where the server is set to www.google.com.
 
Returns:
  A tuple containing the httplib.HTTPConnection and the full_uri for the
  request.












"""




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
        print self.start
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
        self.cal_client = gdata.calendar.service.CalendarService()
        self.cal_client.email = username
        self.cal_client.password = password
        self.cal_client.source = 'Google-Calendar-sync'
        self.cal_client.ProgrammaticLogin()
        return self.cal_client

    def insert_event(self, start_time=None, end_time=None):
        event = gdata.calendar.CalendarEventEntry()
        event.title = atom.Title(text="Busy!")
        event.content = atom.Content(text="")

        if start_time is None:
            print "No start time, can not sync event"
        when = gdata.calendar.When(start_time=start_time, end_time=end_time)
        event.when.append(when)
        # Add an alert
        reminder = gdata.calendar.Reminder(minutes='60')
        reminder._attributes['method'] = 'method'
        reminder.method = 'alert'
        event.when[0].reminder.append(reminder)

        # Enter event into google calendar
        new_event = self.cal_client.InsertEvent(event, '/calendar/feeds/default/private/full')

        print 'New single event inserted: %s' % (new_event.id.text,)
        #print '\tEvent edit URL: %s' % (new_event.GetEditLink().href,)
        #print '\tEvent HTML URL: %s' % (new_event.GetHtmlLink().href,)

def evt_exist(evt, evt_list):
    for event in evt_list:
        if evt.__str__() == event.__str__():
            return True
    return False

def approve_appointment(appointment):
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

def get_dates_in_range(cal_client, start_date='2013-08-19', end_date='2013-09-02'):
    # date format is yyyy-mm-dd

    gevent_list = []

    print 'Date range query for events on Primary Calendar: %s to %s' % (
        start_date, end_date,)
    query = gdata.calendar.service.CalendarEventQuery('default', 'private', 'full')
    query.start_min=start_date
    query.start_max=end_date
    query.params="orderby=starttime"
    feed = cal_client.CalendarQuery(query)
    for event in feed.entry:
        title = event.title.text
        e = Event(text=title)
        for e_when in event.when:
            e.start = e_when.start_time
            e.end = e_when.end_time
        e.gcal_to_google()
        gevent_list.append(e)
    sorted_gevents = sorted(gevent_list, key=lambda x:x.__str__())
    return sorted_gevents


if __name__ == "__main__":

    event_list = []

    gcal = GCal()
    username = raw_input("Enter google username: ")
    password = getpass("Enter Password: ")
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
        event_list.append(event)

    for evt in event_list:
        if not evt_exist(evt, google_events):
            if approve_appointment(evt):
                gcal.insert_event(evt.start, evt.end)
            else:
                print "Not syncing"






