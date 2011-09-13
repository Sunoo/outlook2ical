using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using Microsoft.Office.Interop.Outlook;

namespace Outlook2iCal
{
    class Outlook2iCal
    {
        private static void FtpUpload(string input)
        {
            try
            {
                //TODO: Don't hardcode location
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create("FTPSERVERPATH");
                request.Method = WebRequestMethods.Ftp.UploadFile;

                //TODO: Don't hardcode credentials
                request.Credentials = new NetworkCredential("FTPUSERNAME", "FTPPASSWORD");

                byte[] fileContents = Encoding.UTF8.GetBytes(input);
                request.ContentLength = fileContents.Length;

                Stream requestStream = request.GetRequestStream();
                requestStream.Write(fileContents, 0, fileContents.Length);
                requestStream.Close();

                FtpWebResponse response = (FtpWebResponse)request.GetResponse();

                Console.WriteLine("Upload File Complete, status {0}", response.StatusDescription);

                response.Close();
            }
            catch (WebException ex)
            {
                Console.WriteLine("Upload File Failed, {0}", ex.Message);
            }
        }

        private static string FormatDate(DateTime date)
        {
            //NOTE: The original code uses month + 1 instead of month. No idea why yet.
            return date.ToString("yyyyMMdd");
        }

        private static string FormatDateTime(DateTime date)
        {
            //NOTE: The original code uses month + 1 instead of month. No idea why yet.
            return date.ToString("yyyyMMdd'T'HHmmss");
        }

        private static string CleanLineEndings(string input)
        {
            if (input == null)
            {
                return String.Empty;
            }
            else
            {
                return input.Replace("\r", "\n").Replace("\n\n", "\n").Replace("\n", "\\n").Replace(@",", @"\,");
            }
        }

        private static string DaysOfWeek(OlDaysOfWeek mask, string week)
        {
            string days = String.Empty;
            if ((mask & OlDaysOfWeek.olMonday) == OlDaysOfWeek.olMonday)
            {
                days += week + "MO";
            }
            if ((mask & OlDaysOfWeek.olTuesday) == OlDaysOfWeek.olTuesday)
            {
                if (days.Length > 0)
                {
                    days += ",";
                }
                days += week + "TU";
            }
            if ((mask & OlDaysOfWeek.olWednesday) == OlDaysOfWeek.olWednesday)
            {
                if (days.Length > 0)
                {
                    days += ",";
                }
                days += week + "WE";
            }
            if ((mask & OlDaysOfWeek.olThursday) == OlDaysOfWeek.olThursday)
            {
                if (days.Length > 0)
                {
                    days += ",";
                }
                days += week + "TH";
            }
            if ((mask & OlDaysOfWeek.olFriday) == OlDaysOfWeek.olFriday)
            {
                if (days.Length > 0)
                {
                    days += ",";
                }
                days += week + "FR";
            }
            if ((mask & OlDaysOfWeek.olSaturday) == OlDaysOfWeek.olSaturday)
            {
                if (days.Length > 0)
                {
                    days += ",";
                }
                days += week + "SA";
            }
            if ((mask & OlDaysOfWeek.olSunday) == OlDaysOfWeek.olSunday)
            {
                if (days.Length > 0)
                {
                    days += ",";
                }
                days += week + "SU";
            }

            return days;
        }

        private static string DaysOfWeek(OlDaysOfWeek mask)
        {
            return DaysOfWeek(mask, "");
        }

        private static string WeekNum(int week)
        {
            if (week == 5)
            {
                return "-1";
            }
            else
            {
                return "0" + week;
            }
        }

        private static string CreateReoccuringEvent(AppointmentItem item)
        {
            string recurEvent = "RRULE:";
            
            RecurrencePattern pattern = item.GetRecurrencePattern();
            OlRecurrenceType patternType = pattern.RecurrenceType;

            if (patternType == OlRecurrenceType.olRecursDaily)
            {
                recurEvent += "FREQ=DAILY";
                if (!pattern.NoEndDate)
                {
                    recurEvent += ";UNTIL=" + FormatDateTime(pattern.PatternEndDate);
                }
                recurEvent += ";INTERVAL=" + pattern.Interval;
            }
            else if (patternType == OlRecurrenceType.olRecursMonthly)
            {
                recurEvent += "FREQ=MONTHLY";
                if (!pattern.NoEndDate)
                {
                    recurEvent += ";UNTIL=" + FormatDateTime(pattern.PatternEndDate);
                }
                recurEvent += ";INTERVAL=" + pattern.Interval;
                recurEvent += ";BYMONTHDAY=" + pattern.DayOfMonth;
            }
            else if (patternType == OlRecurrenceType.olRecursMonthNth)
            {
                recurEvent += "FREQ=MONTHLY";
                if (!pattern.NoEndDate)
                {
                    recurEvent += ";UNTIL=" + FormatDateTime(pattern.PatternEndDate);
                }
                recurEvent += ";INTERVAL=" + pattern.Interval;
                recurEvent += ";BYDAY=" + DaysOfWeek(pattern.DayOfWeekMask, WeekNum(pattern.Instance));
            }
            else if (patternType == OlRecurrenceType.olRecursWeekly)
            {
                recurEvent += "FREQ=WEEKLY";
                if (!pattern.NoEndDate)
                {
                    recurEvent += ";UNTIL=" + FormatDateTime(pattern.PatternEndDate);
                }
                recurEvent += ";INTERVAL=" + pattern.Interval;
                recurEvent += ";BYDAY=" + DaysOfWeek(pattern.DayOfWeekMask);
            }
            else if (patternType == OlRecurrenceType.olRecursYearly)
            {
                recurEvent += "FREQ=YEARLY";
                if (!pattern.NoEndDate)
                {
                    recurEvent += ";UNTIL=" + FormatDateTime(pattern.PatternEndDate);
                }
                recurEvent += ";INTERVAL=1";
                recurEvent += ";BYDAY=" + DaysOfWeek(pattern.DayOfWeekMask);
            }
            else if (patternType == OlRecurrenceType.olRecursYearNth)
            {
                recurEvent += "FREQ=YEARLY";
                if (!pattern.NoEndDate)
                {
                    recurEvent += ";UNTIL=" + FormatDateTime(pattern.PatternEndDate);
                }
                recurEvent += ";BYMONTH=" + pattern.MonthOfYear;
                recurEvent += ";BYDAY=" + DaysOfWeek(pattern.DayOfWeekMask, WeekNum(pattern.Instance));
            }

            recurEvent += "\n";

            if (pattern.Exceptions.Count > 0)
            {
                recurEvent += "EXDATE:";
                //NOTE: I need to think of a better way to do this, but this works for now.
                bool firstExcept = true;
                foreach (Microsoft.Office.Interop.Outlook.Exception except in pattern.Exceptions)
                {
                    if (!firstExcept)
                    {
                        recurEvent += ",";
                    }
                    recurEvent += FormatDateTime(except.OriginalDate);
                    firstExcept = false;
                }
                recurEvent += "\n";
            }

            return recurEvent;
        }

        private static string CreateEvent(AppointmentItem item, bool notRecurring)
        {
            string icsEvent = "BEGIN:VEVENT\n";

            if (item.AllDayEvent)
            {
                icsEvent += "DTSTART;VALUE=DATE:" + FormatDate(item.Start) + "\n";
                if (notRecurring && !item.IsRecurring)
                {
                    icsEvent += "DTEND;VALUE=DATE:" + FormatDate(item.End) + "\n";
                }
            }
            else
            {
                icsEvent += "DTSTART:" + FormatDateTime(item.Start) + "\n";
                icsEvent += "DTEND:" + FormatDateTime(item.End) + "\n";
            }

            if (!notRecurring && item.IsRecurring)
            {
                icsEvent += CreateReoccuringEvent(item);
            }

            icsEvent += "LOCATION:" + item.Location + "\n";
            icsEvent += "SUMMARY:" + item.Subject + "\n";

            if (item.Categories != null)
            {
                icsEvent += "CATEGORIES:" + item.Categories + "\n";
            }

            //NOTE: This causes issues with Google Calendar, so it's not worth it
            /*if (item.Sensitivity == OlSensitivity.olNormal)
            {
                icsEvent += "CLASS:PUBLIC\n";
            }
            else if (item.Sensitivity == OlSensitivity.olPersonal)
            {
                icsEvent += "CLASS:CONFIDENTIAL\n";
            }
            else
            {
                icsEvent += "CLASS:PRIVATE\n";
            }*/

            icsEvent += "DESCRIPTION:" + CleanLineEndings(item.Body) + "\n";

            if (item.ReminderMinutesBeforeStart > 0)
            {
                icsEvent += "BEGIN:VALARM\n";
                icsEvent += "TRIGGER:-PT" + item.ReminderMinutesBeforeStart + "M\n";
                icsEvent += "ACTION:DISPLAY\n" +
                    "DESCRIPTION:Reminder\n" +
                    "END:VALARM\n";
            }

            Dictionary<string, string> emails = new Dictionary<string, string>();
            Dictionary<string, OlResponseStatus> status = new Dictionary<string, OlResponseStatus>();

            foreach (Recipient recip in item.Recipients)
            {
                string email = recip.Address.Substring(recip.Address.LastIndexOf('=') + 1);
                email += "@travelers.com";
                emails.Add(recip.Name, email);
                status.Add(recip.Name, recip.MeetingResponseStatus);
            }

            if (item.Organizer != null)
            {
                icsEvent += "ORGANIZER;CN=" + item.Organizer.Replace(",", @"\,");
                if (emails.ContainsKey(item.Organizer))
                {
                    icsEvent += ":MAILTO:" + emails[item.Organizer] + "\n";
                }
                else
                {
                    icsEvent += ":MAILTO:noreply@example.com\n";
                }
            }

            //NOTE: Attendees do not appear on the iPhone when an Organizer is set. This is a strange behavior...
            //NOTE: Actually it only happens sometimes... Even stranger...
            if (item.RequiredAttendees != null)
            {
                string[] required = item.RequiredAttendees.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string req in required)
                {
                    icsEvent += "ATTENDEE;ROLE=REQ-PARTICIPANT;";
                    if (status.ContainsKey(req))
                    {
                        switch (status[req])
                        {
                            case OlResponseStatus.olResponseAccepted:
                                icsEvent += "PARTSTAT=ACCEPTED;";
                                break;
                            case OlResponseStatus.olResponseDeclined:
                                icsEvent += "PARTSTAT=DECLINED;";
                                break;
                            case OlResponseStatus.olResponseTentative:
                                icsEvent += "PARTSTAT=TENTATIVE;";
                                break;
                        }
                    }
                    icsEvent += "CN=" + req.Replace(",", @"\,");
                    if (emails.ContainsKey(req))
                    {
                        icsEvent += ":MAILTO:" + emails[req] + "\n";
                    }
                    else
                    {
                        icsEvent += ":MAILTO:noreply@example.com\n";
                    }
                }
            }

            if (item.OptionalAttendees != null)
            {
                string[] optional = item.OptionalAttendees.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string opt in optional)
                {
                    icsEvent += "ATTENDEE;";
                    if (status.ContainsKey(opt))
                    {
                        switch (status[opt])
                        {
                            case OlResponseStatus.olResponseAccepted:
                                icsEvent += "PARTSTAT=ACCEPTED;";
                                break;
                            case OlResponseStatus.olResponseDeclined:
                                icsEvent += "PARTSTAT=DECLINED;";
                                break;
                            case OlResponseStatus.olResponseTentative:
                                icsEvent += "PARTSTAT=TENTATIVE;";
                                break;
                        }
                    }
                    icsEvent += "CN=" + opt.Replace(",", @"\,");
                    if (emails.ContainsKey(opt))
                    {
                        icsEvent += ":MAILTO:" + emails[opt] + "\n";
                    }
                    else
                    {
                        icsEvent += ":MAILTO:noreply@example.com\n";
                    }
                }
            }

            icsEvent += "DTSTAMP:" + FormatDateTime(item.CreationTime.ToUniversalTime()) + "\n";
            icsEvent += "LAST-MODIFIED:" + FormatDateTime(item.LastModificationTime.ToUniversalTime()) + "\n";
            //NOTE: Since this is the same for all recurring items, it was breaking the iPhone calendar
            //icsEvent += "UID:" + item.EntryID + "\n";

            icsEvent += "END:VEVENT\n";

            if (!notRecurring && item.IsRecurring)
            {
                RecurrencePattern pattern = item.GetRecurrencePattern();
                foreach (Microsoft.Office.Interop.Outlook.Exception except in pattern.Exceptions)
                {
                    if (!except.Deleted)
                    {
                        icsEvent += CreateEvent(except.AppointmentItem, true);
                    }
                }
            }

            return icsEvent;
        }

        private static string CreateEvent(AppointmentItem item)
        {
            return CreateEvent(item, false);
        }

        private static string GenerateIcs()
        {
            var ics = "BEGIN:VCALENDAR\n" +
                "X-WR-CALNAME:Outlook\n" +
                "X-WR-CALDESC:Outlook\n" +
                "X-WR-TIMEZONE:America/New_York" +
                "PRODID:-//David Maher/NONSGML Outlook2iCal 1.0//EN" +
                "VERSION:2.0\n";

            Items calendarItems = new Application().GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderCalendar).Items;
            calendarItems.IncludeRecurrences = true;

            foreach (AppointmentItem item in calendarItems)
            {
                ics += CreateEvent(item);
            }

            ics += "END:VCALENDAR\n";

            return ics;
        }

        static void Main(string[] args)
        {
            string ics = GenerateIcs();
            //Console.Write(ics);
            FtpUpload(ics);
            //Console.Read();
        }
    }
}
