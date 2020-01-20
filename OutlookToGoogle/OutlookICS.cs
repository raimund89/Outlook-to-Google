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

            DateTime start = DateTime.Today.AddDays(-30);
            DateTime end = DateTime.Today.AddDays(90);

            foreach (AppointmentItem item in calendarItems)
            {
                if (item.IsRecurring)
                {
                    RecurrencePattern rp = item.GetRecurrencePattern();
                    AppointmentItem recur = null;

                    DateTime first = start.AddHours(item.Start.Hour).AddMinutes(item.Start.Minute);
                    DateTime last = end.AddHours(item.Start.Hour).AddMinutes(item.Start.Minute);

                    for (DateTime cur = first; cur <= last; cur = cur.AddDays(1))
                    {
                        try
                        {
                            recur = rp.GetOccurrence(cur);
                            items.Add(recur);
                        }
                        catch
                        {

                        }
                    }
                }
                else
                {
                    if (item.Start.CompareTo(start) > 0 && item.Start.CompareTo(end) < 0)
                    {
                        items.Add(item);
                    }
                }
            }
        }

        // TODO: Recurring items should not be recorded separately, but using RRULE tag. Lots of double UIDs now. But what if one instance is cancelled?
        // TODO: Finish/check priority, participant type, meeting status
        // TODO: Summary: add language tag
        // TODO: olResponseOrganized doesn't have the right partstat
        // TODO: Sensitivity private doesn't have the right class
        // TODO: Europe/Amsterdam is not valid, according to ICS validator

        public void WriteICS(String filename)
        {
            Console.WriteLine("Writing ICS...");

            using (StreamWriter sw = new StreamWriter(filename, false))
            {
                string created = DateTime.Now.ToUniversalTime().ToString(@"yyyyMMdd\THHmmssZ");

                sw.WriteLine("BEGIN:VCALENDAR");
                sw.WriteLine("VERSION:2.0");
                sw.WriteLine("PRODID:-//Microsoft Corporation//Outlook 16.0 MIMEDIR//EN");

                foreach (AppointmentItem item in this.items)
                {
                    sw.WriteLine("BEGIN:VEVENT");

                    sw.WriteLine("CREATED:" + created);

                    sw.WriteLine(wrapString("DESCRIPTION:" + item.Body.Replace("\r\n", "\\n").Replace(",", "\\,") + "\\n"));

                    sw.WriteLine(wrapString("SUMMARY:" + item.Subject));

                    sw.WriteLine(wrapString("UID:" + item.GlobalAppointmentID));

                    sw.WriteLine(wrapString("ORGANIZER;CN=\"" + item.GetOrganizer().Name + "\":mailto:" + item.GetOrganizer().PropertyAccessor.GetProperty(@"http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString()));

                    Recipients recipients = item.Recipients;
                    if (recipients.Count > 1)
                    {
                        foreach (Recipient recipient in recipients)
                        {
                            String recip = "ATTENDEE;CN=\"" + recipient.Name + "\";";
                            switch (recipient.Type)
                            {
                                case (int)NetOffice.OutlookApi.Enums.OlMeetingRecipientType.olOptional:
                                    recip += "ROLE=OPT-PARTICIPANT;";
                                    break;
                                case (int)NetOffice.OutlookApi.Enums.OlMeetingRecipientType.olRequired:
                                    recip += "ROLE=REQ-PARTICIPANT;";
                                    break;
                                case (int)NetOffice.OutlookApi.Enums.OlMeetingRecipientType.olOrganizer:
                                    recip += "ROLE=CHAIR;";
                                    break;
                                case (int)NetOffice.OutlookApi.Enums.OlMeetingRecipientType.olResource:
                                    recip += "ROLE=NON-PARTICIPANT;";
                                    break;
                            }
                            if (item.ResponseRequested)
                                recip += "RSVP=TRUE;";
                            switch (recipient.MeetingResponseStatus)
                            {
                                case NetOffice.OutlookApi.Enums.OlResponseStatus.olResponseAccepted:
                                    recip += "PARTSTAT=ACCEPTED;";
                                    break;
                                case NetOffice.OutlookApi.Enums.OlResponseStatus.olResponseDeclined:
                                    recip += "PARTSTAT=DECLINED;";
                                    break;
                                case NetOffice.OutlookApi.Enums.OlResponseStatus.olResponseNotResponded:
                                    recip += "PARTSTAT=NEEDS-ACTION;";
                                    break;
                                case NetOffice.OutlookApi.Enums.OlResponseStatus.olResponseOrganized:
                                    recip += "PARTSTAT=ACCEPTED;";
                                    break;
                                case NetOffice.OutlookApi.Enums.OlResponseStatus.olResponseTentative:
                                    recip += "PARTSTAT=TENTATIVE;";
                                    break;
                            }
                            sw.WriteLine(wrapString(recip + ":mailto:" + recipient.PropertyAccessor.GetProperty(@"http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString()));
                        }
                    }

                    if (item.Location.Length > 0)
                        sw.WriteLine(wrapString("LOCATION:" + item.Location));

                    {
                        sw.Write("PRIORITY:");
                        switch (item.Importance)
                        {
                            case NetOffice.OutlookApi.Enums.OlImportance.olImportanceNormal:
                                sw.WriteLine("5");
                                break;
                            case NetOffice.OutlookApi.Enums.OlImportance.olImportanceLow:
                                sw.WriteLine("6");
                                break;
                            case NetOffice.OutlookApi.Enums.OlImportance.olImportanceHigh:
                                sw.WriteLine("1");
                                break;
                        }
                    }
                    {
                        // SENSITIVITY
                        switch (item.Sensitivity)
                        {
                            case NetOffice.OutlookApi.Enums.OlSensitivity.olNormal:
                                sw.WriteLine("CLASS:PUBLIC");
                                break;
                            case NetOffice.OutlookApi.Enums.OlSensitivity.olConfidential:
                                sw.WriteLine("CLASS:CONFIDENTIAL");
                                break;
                            case NetOffice.OutlookApi.Enums.OlSensitivity.olPrivate:
                                sw.WriteLine("CLASS:PRIVATE");
                                break;
                            case NetOffice.OutlookApi.Enums.OlSensitivity.olPersonal:
                                sw.WriteLine("CLASS:PRIVATE");
                                break;
                        }
                    }

                    if(item.AllDayEvent)
                    {
                        sw.WriteLine("DTSTART;VALUE=DATE:" + item.Start.ToString(@"yyyyMMdd"));
                        sw.WriteLine("DTEND;VALUE=DATE:" + item.End.ToString(@"yyyyMMdd"));
                    } 
                    else
                    {
                        sw.WriteLine("DTSTART;TZID=" + TZConvert.WindowsToIana(item.StartTimeZone.ID, "NL") + ":" + item.StartInStartTimeZone.ToString(@"yyyyMMdd\THHmmss"));
                        sw.WriteLine("DTEND;TZID=" + TZConvert.WindowsToIana(item.EndTimeZone.ID, "NL") + ":" + item.EndInEndTimeZone.ToString(@"yyyyMMdd\THHmmss"));
                    }
                    
                    sw.WriteLine("LAST-MODIFIED:" + item.LastModificationTime.ToUniversalTime().ToString(@"yyyyMMdd\THHmmssZ"));
                    sw.WriteLine("DTSTAMP:" + item.CreationTime.ToUniversalTime().ToString(@"yyyyMMdd\THHmmssZ"));
                    // TRANSP
                    // RRULE
                    // SEQUENCE

                    // BEGIN:VALARM/END:VALARM
                    // BUNCH OF X-ALT and X-MICROSOFT tags

                    sw.WriteLine("END:VEVENT");
                }

                sw.WriteLine("END:VCALENDAR");
            }

            Console.WriteLine("Finished writing!");
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
