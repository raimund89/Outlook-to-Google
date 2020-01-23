using System;
using System.IO;
using System.Collections.Generic;
using NetOffice.OutlookApi;
using TimeZoneConverter;

namespace OutlookToGoogle
{
    public class OutlookICS
    {
        // NOTE: Do not call Quit() on the application, as that will also
        // close the Outlook desktop application if it is opened.

        private List<AppointmentItem> items = new List<AppointmentItem>();
        private Application application;
        private NameSpace mapiNamespace;

        public OutlookICS()
        {
            application = new Application();
            mapiNamespace = (NameSpace)application.GetNamespace("MAPI");
        }

        public string GetVersion()
        {
            return application.Version;
        }

        public string GetDefaultProfile()
        {
            return mapiNamespace.CurrentUser.Name;
        }

        public void Cleanup()
        {
            items.Clear();
        }

        public void ReadCalendar()
        {
            Console.WriteLine("Reading calendar...");

            MAPIFolder calendarFolder = mapiNamespace.GetDefaultFolder(NetOffice.OutlookApi.Enums.OlDefaultFolders.olFolderCalendar);
            Items calendarItems = (Items)calendarFolder.Items;
            calendarItems.IncludeRecurrences = true;

            DateTime start = DateTime.Today.AddMonths(Properties.Settings.Default.rangeStart);
            DateTime end = DateTime.Today.AddMonths(Properties.Settings.Default.rangeEnd);

            foreach (AppointmentItem item in calendarItems)
            {
                if(item.IsRecurring)
                {
                    RecurrencePattern rp = item.GetRecurrencePattern();
                    if(rp.PatternStartDate.CompareTo(end) < 0 && (rp.NoEndDate || rp.PatternEndDate.CompareTo(start) > 0))
                        items.Add(item);
                }
                else
                {
                    if (item.Start.CompareTo(start) > 0 && item.Start.CompareTo(end) < 0)
                        items.Add(item);
                }
            }
        }

        // TODO: Finish/check priority, participant type, meeting status
        // TODO: Summary: add language tag
        // TODO: olResponseOrganized doesn't have the right partstat
        // TODO: Sensitivity private doesn't have the right class
        // TODO: Enable recognition and use of any timezone
        // TODO: olImportanceLow gives the wrong number

        // TODO: Any exception is now processed as a cancellation, but it can also be a rescheduled event

        // TODO: Recurrences are implemented. However, several recurrenceTypes are not. Also, what's with the MONTLY + BYSETPOS? And implement YEARLY
        // TODO: Exceptions don't work yet. Cancelled events are still visible in google calendar

        public void WriteICS(String filename)
        {
            Console.WriteLine("Writing ICS...");

            using (StreamWriter sw = new StreamWriter(filename, false))
            {
                string created = DateTime.Now.ToUniversalTime().ToString(@"yyyyMMdd\THHmmssZ");

                sw.WriteLine("BEGIN:VCALENDAR");
                sw.WriteLine("VERSION:2.0");
                sw.WriteLine("PRODID:-//OutlookToGoogle/ICS");
                sw.WriteLine("METHOD:PUBLISH");
                sw.WriteLine("X-WR-CALNAME:" + GetDefaultProfile());
                sw.WriteLine("X-WR-TIMEZONE:Europe/Amsterdam");

                // Force recognition of Europe/Amsterdam timezone for now
                sw.WriteLine("BEGIN:VTIMEZONE");
                sw.WriteLine("TZID:Europe/Amsterdam");
                sw.WriteLine("X-LIC-LOCATION:Europe/Amsterdam");
                sw.WriteLine("BEGIN:DAYLIGHT");
                sw.WriteLine("TZOFFSETFROM:+0100");
                sw.WriteLine("TZOFFSETTO:+0200");
                sw.WriteLine("TZNAME:CEST");
                sw.WriteLine("DTSTART:19700329T020000");
                sw.WriteLine("RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=-1SU;BYMONTH=3");
                sw.WriteLine("END:DAYLIGHT");
                sw.WriteLine("BEGIN:STANDARD");
                sw.WriteLine("TZOFFSETFROM:+0200");
                sw.WriteLine("TZOFFSETTO:+0100");
                sw.WriteLine("TZNAME:CET");
                sw.WriteLine("DTSTART:19701025T030000");
                sw.WriteLine("RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=-1SU;BYMONTH=10");
                sw.WriteLine("END:STANDARD");
                sw.WriteLine("END:VTIMEZONE");

                foreach (AppointmentItem item in this.items)
                {
                    string exceptionevents = "";

                    // Recursively write events and their optional cancelled occurrences
                    sw.Write(FormattedEvent(item, created, false, true));

                    // SEQUENCE

                    // BEGIN:VALARM/END:VALARM
                    // BUNCH OF X-ALT and X-MICROSOFT tags

                    sw.Write(exceptionevents);
                }

                sw.WriteLine("END:VCALENDAR");
            }

            Console.WriteLine("Finished writing!");
        }

        private String FormattedEvent(AppointmentItem item, String created, bool cancelled, bool expandRecurring)
        {
            String str = "BEGIN:VEVENT\r\n";

            str += "CREATED:" + created + "\r\n";

            // If we're not expanding recurring events, this is an exception, which means it's cancelled
            if(!expandRecurring)
                str += "TRANSP:TRANSPARENT\r\n";
            else
                str += "TRANSP:OPAQUE\r\n";

            str += wrapString("DESCRIPTION:" + item.Body.Replace("\r\n", "\\n").Replace(",", "\\,") + "\\n") + "\r\n";

            str += wrapString("SUMMARY:" + item.Subject) + "\r\n";

            str += wrapString("UID:" + item.GlobalAppointmentID) + "\r\n";

            str += wrapString("ORGANIZER;CN=\"" + item.GetOrganizer().Name + "\":mailto:" + item.GetOrganizer().PropertyAccessor.GetProperty(@"http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString()) + "\r\n";

            Recipients recipients = item.Recipients;
            if (recipients.Count > 1)
            {
                foreach (Recipient recipient in recipients)
                {
                    String recip = "ATTENDEE;CN=\"" + recipient.Name + "\"";
                    switch (recipient.Type)
                    {
                        case (int)NetOffice.OutlookApi.Enums.OlMeetingRecipientType.olOptional:
                            recip += ";ROLE=OPT-PARTICIPANT";
                            break;
                        case (int)NetOffice.OutlookApi.Enums.OlMeetingRecipientType.olRequired:
                            recip += ";ROLE=REQ-PARTICIPANT";
                            break;
                        case (int)NetOffice.OutlookApi.Enums.OlMeetingRecipientType.olOrganizer:
                            recip += ";ROLE=CHAIR";
                            break;
                        case (int)NetOffice.OutlookApi.Enums.OlMeetingRecipientType.olResource:
                            recip += ";ROLE=NON-PARTICIPANT";
                            break;
                    }
                    if (item.ResponseRequested)
                        recip += ";RSVP=TRUE";
                    switch (recipient.MeetingResponseStatus)
                    {
                        case NetOffice.OutlookApi.Enums.OlResponseStatus.olResponseAccepted:
                            recip += ";PARTSTAT=ACCEPTED";
                            break;
                        case NetOffice.OutlookApi.Enums.OlResponseStatus.olResponseDeclined:
                            recip += ";PARTSTAT=DECLINED";
                            break;
                        case NetOffice.OutlookApi.Enums.OlResponseStatus.olResponseNotResponded:
                            recip += ";PARTSTAT=NEEDS-ACTION";
                            break;
                        case NetOffice.OutlookApi.Enums.OlResponseStatus.olResponseOrganized:
                            recip += ";PARTSTAT=ACCEPTED";
                            break;
                        case NetOffice.OutlookApi.Enums.OlResponseStatus.olResponseTentative:
                            recip += ";PARTSTAT=TENTATIVE";
                            break;
                    }
                    str += wrapString(recip + ":mailto:" + recipient.PropertyAccessor.GetProperty(@"http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString()) + "\r\n";
                }
            }

            if (item.Location.Length > 0)
                str += wrapString("LOCATION:" + item.Location) + "\r\n";

            {
                str += "PRIORITY:";
                switch (item.Importance)
                {
                    case NetOffice.OutlookApi.Enums.OlImportance.olImportanceNormal:
                        str += "5" + "\r\n";
                        break;
                    case NetOffice.OutlookApi.Enums.OlImportance.olImportanceLow:
                        str += "6" + "\r\n";
                        break;
                    case NetOffice.OutlookApi.Enums.OlImportance.olImportanceHigh:
                        str += "1" + "\r\n";
                        break;
                }
            }
            {
                // SENSITIVITY
                switch (item.Sensitivity)
                {
                    case NetOffice.OutlookApi.Enums.OlSensitivity.olNormal:
                        str += "CLASS:PUBLIC" + "\r\n";
                        break;
                    case NetOffice.OutlookApi.Enums.OlSensitivity.olConfidential:
                        str += "CLASS:CONFIDENTIAL" + "\r\n";
                        break;
                    case NetOffice.OutlookApi.Enums.OlSensitivity.olPrivate:
                        str += "CLASS:PRIVATE" + "\r\n";
                        break;
                    case NetOffice.OutlookApi.Enums.OlSensitivity.olPersonal:
                        str += "CLASS:PRIVATE" + "\r\n";
                        break;
                }
            }

            if (item.AllDayEvent)
            {
                str += "DTSTART;VALUE=DATE:" + item.Start.ToString(@"yyyyMMdd") + "\r\n";
                str += "DTEND;VALUE=DATE:" + item.End.ToString(@"yyyyMMdd") + "\r\n";
            }
            else
            {
                str += "DTSTART;TZID=" + TZConvert.WindowsToIana(item.StartTimeZone.ID, "NL") + ":" + item.StartInStartTimeZone.ToString(@"yyyyMMdd\THHmmss") + "\r\n";
                str += "DTEND;TZID=" + TZConvert.WindowsToIana(item.EndTimeZone.ID, "NL") + ":" + item.EndInEndTimeZone.ToString(@"yyyyMMdd\THHmmss") + "\r\n";
            }

            str += "LAST-MODIFIED:" + item.LastModificationTime.ToUniversalTime().ToString(@"yyyyMMdd\THHmmssZ") + "\r\n";
            str += "DTSTAMP:" + item.CreationTime.ToUniversalTime().ToString(@"yyyyMMdd\THHmmssZ") + "\r\n";


            if (expandRecurring && item.IsRecurring)
            {
                Console.WriteLine("Recurring meeting: " + item.Subject);

                RecurrencePattern rp = item.GetRecurrencePattern();

                str += "RRULE:";
                switch (rp.RecurrenceType)
                {
                    case NetOffice.OutlookApi.Enums.OlRecurrenceType.olRecursDaily:
                        str += "FREQ=DAILY";
                        break;
                    case NetOffice.OutlookApi.Enums.OlRecurrenceType.olRecursMonthly:
                        str += "FREQ=MONTHLY";
                        break;
                    case NetOffice.OutlookApi.Enums.OlRecurrenceType.olRecursMonthNth:
                        str += "FREQ=MONTHLY";
                        break;
                    case NetOffice.OutlookApi.Enums.OlRecurrenceType.olRecursWeekly:
                        str += "FREQ=WEEKLY";
                        break;
                    case NetOffice.OutlookApi.Enums.OlRecurrenceType.olRecursYearly:
                        str += "FREQ=YEARLY";
                        break;
                    case NetOffice.OutlookApi.Enums.OlRecurrenceType.olRecursYearNth:
                        break;
                }

                if (!rp.NoEndDate)
                {
                    // There is an end-date, either in occurrences or an end-date
                    if (rp.PatternEndDate != null)
                    {
                        str += ";UNTIL=" + rp.PatternEndDate.ToUniversalTime().ToString(@"yyyyMMdd\THHmmssZ");
                    }
                    else
                    {
                        str += ";COUNT=" + rp.Occurrences;
                    }
                }

                if (rp.Interval > 1)
                {
                    str += ";INTERVAL=" + rp.Interval;
                }

                string days = "";

                if ((rp.DayOfWeekMask & NetOffice.OutlookApi.Enums.OlDaysOfWeek.olMonday) == NetOffice.OutlookApi.Enums.OlDaysOfWeek.olMonday)
                {
                    if (days.Length > 0)
                        days += ",";
                    days += "MO";
                }
                if ((rp.DayOfWeekMask & NetOffice.OutlookApi.Enums.OlDaysOfWeek.olTuesday) == NetOffice.OutlookApi.Enums.OlDaysOfWeek.olTuesday)
                {
                    if (days.Length > 0)
                        days += ",";
                    days += "TU";
                }
                if ((rp.DayOfWeekMask & NetOffice.OutlookApi.Enums.OlDaysOfWeek.olWednesday) == NetOffice.OutlookApi.Enums.OlDaysOfWeek.olWednesday)
                {
                    if (days.Length > 0)
                        days += ",";
                    days += "WE";
                }
                if ((rp.DayOfWeekMask & NetOffice.OutlookApi.Enums.OlDaysOfWeek.olThursday) == NetOffice.OutlookApi.Enums.OlDaysOfWeek.olThursday)
                {
                    if (days.Length > 0)
                        days += ",";
                    days += "TH";
                }
                if ((rp.DayOfWeekMask & NetOffice.OutlookApi.Enums.OlDaysOfWeek.olFriday) == NetOffice.OutlookApi.Enums.OlDaysOfWeek.olFriday)
                {
                    if (days.Length > 0)
                        days += ",";
                    days += "FR";
                }
                if ((rp.DayOfWeekMask & NetOffice.OutlookApi.Enums.OlDaysOfWeek.olSaturday) == NetOffice.OutlookApi.Enums.OlDaysOfWeek.olSaturday)
                {
                    if (days.Length > 0)
                        days += ",";
                    days += "SA";
                }
                if ((rp.DayOfWeekMask & NetOffice.OutlookApi.Enums.OlDaysOfWeek.olSunday) == NetOffice.OutlookApi.Enums.OlDaysOfWeek.olSunday)
                {
                    if (days.Length > 0)
                        days += ",";
                    days += "SU";
                }

                str += ";BYDAY=" + days;

                if (rp.RecurrenceType == NetOffice.OutlookApi.Enums.OlRecurrenceType.olRecursMonthNth)
                {
                    str += ";BYSETPOS=1";
                }

                str += "\r\n";

                str += "END:VEVENT" + "\r\n";

                Exceptions exceptions = rp.Exceptions;

                foreach (NetOffice.OutlookApi.Exception exc in exceptions)
                {
                    Console.WriteLine("Found an exception to item " + item.Subject);
                    try
                    {
                        str += FormattedEvent(exc.AppointmentItem, created, true, false);
                        Console.WriteLine("Wrote the exception");
                    }
                    catch (System.Runtime.InteropServices.COMException e)
                    {
                        // Let's ignore any errors for the moment, they come from items out of the current range i think...
                        Console.WriteLine("Error writing exception: " + e.Message);
                    }
                }
            }
            else
            {
                str += "END:VEVENT" + "\r\n";
            }

            return str;
        }

        private String wrapString(String input)
        {
            String result = input;
            int maxlength = 75;
            if (result.Length > maxlength)
            {
                for (int i = maxlength; i < result.Length; i += maxlength)
                {
                    result = result.Insert(i, "\r\n\t");
                }
            }

            return result;
        }
    }
}
