using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using Microsoft.Office.Interop.Outlook;
using DDay.iCal;
using DDay.iCal.Serialization.iCalendar;
using System.Security.Cryptography;

namespace Outlook2iCal
{
    class Outlook2iCal
    {
        private static string tzid;

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

        private static void DaysOfWeek(DDay.iCal.RecurrencePattern pattern, OlDaysOfWeek mask, int week)
        {
            FrequencyOccurrence occur;
            switch (week)
            {
                case 1:
                    occur = FrequencyOccurrence.Last;
                    break;
                case 2:
                    occur = FrequencyOccurrence.Last;
                    break;
                case 3:
                    occur = FrequencyOccurrence.Last;
                    break;
                case 4:
                    occur = FrequencyOccurrence.Last;
                    break;
                case 5:
                    occur = FrequencyOccurrence.Last;
                    break;
                default:
                    occur = FrequencyOccurrence.None;
                    break;
            }
            if ((mask & OlDaysOfWeek.olMonday) == OlDaysOfWeek.olMonday)
            {
                pattern.ByDay.Add(new WeekDay(DayOfWeek.Monday, occur));
            }
            if ((mask & OlDaysOfWeek.olTuesday) == OlDaysOfWeek.olTuesday)
            {
                pattern.ByDay.Add(new WeekDay(DayOfWeek.Tuesday, occur));
            }
            if ((mask & OlDaysOfWeek.olWednesday) == OlDaysOfWeek.olWednesday)
            {
                pattern.ByDay.Add(new WeekDay(DayOfWeek.Wednesday, occur));
            }
            if ((mask & OlDaysOfWeek.olThursday) == OlDaysOfWeek.olThursday)
            {
                pattern.ByDay.Add(new WeekDay(DayOfWeek.Thursday, occur));
            }
            if ((mask & OlDaysOfWeek.olFriday) == OlDaysOfWeek.olFriday)
            {
                pattern.ByDay.Add(new WeekDay(DayOfWeek.Friday, occur));
            }
            if ((mask & OlDaysOfWeek.olSaturday) == OlDaysOfWeek.olSaturday)
            {
                pattern.ByDay.Add(new WeekDay(DayOfWeek.Saturday, occur));
            }
            if ((mask & OlDaysOfWeek.olSunday) == OlDaysOfWeek.olSunday)
            {
                pattern.ByDay.Add(new WeekDay(DayOfWeek.Sunday, occur));
            }
        }

        private static void CreateReoccuringEvent(Event icsEvent, AppointmentItem item)
        {
            DDay.iCal.RecurrencePattern newPatt = new DDay.iCal.RecurrencePattern();

            Microsoft.Office.Interop.Outlook.RecurrencePattern pattern = item.GetRecurrencePattern();
            OlRecurrenceType patternType = pattern.RecurrenceType;

            if (patternType == OlRecurrenceType.olRecursDaily)
            {
                newPatt.Frequency = FrequencyType.Daily;
                if (!pattern.NoEndDate)
                {
                    newPatt.Until = pattern.PatternEndDate;
                }
                newPatt.Interval = pattern.Interval;
            }
            else if (patternType == OlRecurrenceType.olRecursMonthly)
            {
                newPatt.Frequency = FrequencyType.Monthly;
                if (!pattern.NoEndDate)
                {
                    newPatt.Until = pattern.PatternEndDate;
                }
                newPatt.Interval = pattern.Interval;
                newPatt.ByMonthDay.Add(pattern.DayOfMonth);
            }
            else if (patternType == OlRecurrenceType.olRecursMonthNth)
            {
                newPatt.Frequency = FrequencyType.Monthly;
                if (!pattern.NoEndDate)
                {
                    newPatt.Until = pattern.PatternEndDate;
                }
                newPatt.Interval = pattern.Interval;
                DaysOfWeek(newPatt, pattern.DayOfWeekMask, pattern.Instance);
            }
            else if (patternType == OlRecurrenceType.olRecursWeekly)
            {
                newPatt.Frequency = FrequencyType.Weekly;
                if (!pattern.NoEndDate)
                {
                    newPatt.Until = pattern.PatternEndDate;
                }
                newPatt.Interval = pattern.Interval;
                DaysOfWeek(newPatt, pattern.DayOfWeekMask, 0);
            }
            else if (patternType == OlRecurrenceType.olRecursYearly)
            {
                newPatt.Frequency = FrequencyType.Yearly;
                if (!pattern.NoEndDate)
                {
                    newPatt.Until = pattern.PatternEndDate;
                }
                newPatt.Interval = 1;
                DaysOfWeek(newPatt, pattern.DayOfWeekMask, 0);
            }
            else if (patternType == OlRecurrenceType.olRecursYearNth)
            {
                newPatt.Frequency = FrequencyType.Yearly;
                if (!pattern.NoEndDate)
                {
                    newPatt.Until = pattern.PatternEndDate;
                }
                newPatt.ByMonth.Add(pattern.MonthOfYear);
                DaysOfWeek(newPatt, pattern.DayOfWeekMask, pattern.Instance);
            }

            icsEvent.RecurrenceRules.Add(newPatt);

            if (pattern.Exceptions.Count > 0)
            {
                PeriodList list = new PeriodList();
                foreach (Microsoft.Office.Interop.Outlook.Exception except in pattern.Exceptions)
                {
                    list.Add(new iCalDateTime(except.OriginalDate, tzid));
                }
                icsEvent.ExceptionDates.Add(list);
            }
        }

        private static void CreateEvent(iCalendar ics, AppointmentItem item, bool notRecurring)
        {
            Event icsEvent = ics.Create<Event>();

            if (item.AllDayEvent)
            {
                icsEvent.IsAllDay = true;
                icsEvent.DTStart = new iCalDateTime(item.Start, tzid);
                if (notRecurring && !item.IsRecurring)
                {
                    icsEvent.DTStart = new iCalDateTime(item.End, tzid);
                }
            }
            else
            {
                icsEvent.DTStart = new iCalDateTime(item.Start, tzid);
                icsEvent.DTEnd = new iCalDateTime(item.End, tzid);
            }

            if (!notRecurring && item.IsRecurring)
            {
                CreateReoccuringEvent(icsEvent, item);
            }

            icsEvent.Location = item.Location;
            icsEvent.Summary = item.Subject;

            if (item.Categories != null)
            {
                string[] cats = item.Categories.Split(',');
                foreach (string cat in cats)
                {
                    icsEvent.Categories.Add(cat);
                }
            }

            //NOTE: This causes issues with Google Calendar, so it's not worth it
            /*switch (item.Sensitivity)
            {
                case OlSensitivity.olNormal:
                    icsEvent.Class = "PUBLIC";
                    break;
                case OlSensitivity.olPersonal:
                    icsEvent.Class = "CONFIDENTIAL";
                    break;
                default:
                    icsEvent.Class = "PRIVATE";
                    break;
            }*/

            icsEvent.Description = item.Body;

            if (item.ReminderMinutesBeforeStart > 0)
            {
                Alarm alarm = new Alarm();
                alarm.Trigger = new Trigger(new TimeSpan(0, 0 - item.ReminderMinutesBeforeStart, 0));
                alarm.Action = AlarmAction.Display;
                alarm.Description = "Reminder";
                icsEvent.Alarms.Add(alarm);
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
                if (emails.ContainsKey(item.Organizer))
                {
                    icsEvent.Organizer = new Organizer("MAILTO:" + emails[item.Organizer]);
                }
                else
                {
                    icsEvent.Organizer = new Organizer();
                }
                icsEvent.Organizer.CommonName = item.Organizer;
            }

            if (item.RequiredAttendees != null)
            {
                string[] required = item.RequiredAttendees.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string req in required)
                {
                    Attendee attend;
                    if (emails.ContainsKey(req))
                    {
                        attend = new Attendee("MAILTO:" + emails[req]);
                    }
                    else
                    {
                        attend = new Attendee();
                    }
                    attend.Role = "REQ-PARTICIPANT";
                    if (status.ContainsKey(req))
                    {
                        switch (status[req])
                        {
                            case OlResponseStatus.olResponseAccepted:
                                attend.ParticipationStatus = "ACCEPTED";
                                break;
                            case OlResponseStatus.olResponseDeclined:
                                attend.ParticipationStatus = "DECLINED";
                                break;
                            case OlResponseStatus.olResponseTentative:
                                attend.ParticipationStatus = "TENTATIVE";
                                break;
                        }
                    }
                    attend.CommonName = req;
                    icsEvent.Attendees.Add(attend);
                }
            }

            if (item.OptionalAttendees != null)
            {
                string[] optional = item.OptionalAttendees.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string opt in optional)
                {
                    Attendee attend;
                    if (emails.ContainsKey(opt))
                    {
                        attend = new Attendee("MAILTO:" + emails[opt]);
                    }
                    else
                    {
                        attend = new Attendee();
                    }
                    if (status.ContainsKey(opt))
                    {
                        switch (status[opt])
                        {
                            case OlResponseStatus.olResponseAccepted:
                                attend.ParticipationStatus = "ACCEPTED";
                                break;
                            case OlResponseStatus.olResponseDeclined:
                                attend.ParticipationStatus = "DECLINED";
                                break;
                            case OlResponseStatus.olResponseTentative:
                                attend.ParticipationStatus = "TENTATIVE";
                                break;
                        }
                    }
                    attend.CommonName = opt;
                    icsEvent.Attendees.Add(attend);
                }
            }

            icsEvent.DTStamp = new iCalDateTime(item.CreationTime, tzid);
            icsEvent.LastModified = new iCalDateTime(item.LastModificationTime, tzid);
            //NOTE: Since this is the same for all recurring items, it was breaking the iPhone calendar
            //icsEvent.UID = item.EntryID;
            byte[] buf = Encoding.Default.GetBytes(item.EntryID + item.Start.ToFileTimeUtc());
            SHA1CryptoServiceProvider sha = new SHA1CryptoServiceProvider();
            icsEvent.UID = BitConverter.ToString(sha.ComputeHash(buf)).Replace("-", "");

            if (!notRecurring && item.IsRecurring)
            {
                Microsoft.Office.Interop.Outlook.RecurrencePattern pattern = item.GetRecurrencePattern();
                foreach (Microsoft.Office.Interop.Outlook.Exception except in pattern.Exceptions)
                {
                    if (!except.Deleted)
                    {
                        CreateEvent(ics, except.AppointmentItem, true);
                    }
                }
            }
        }

        private static iCalendar GenerateIcs()
        {
            iCalendar ics = new iCalendar();

            ITimeZone tz = ics.AddLocalTimeZone();
            tzid = tz.TZID;

            ics.ProductID = "-//David Maher/NONSGML Outlook2iCal 2.0//EN";
            
            Items calendarItems = new Application().GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderCalendar).Items;
            calendarItems.IncludeRecurrences = true;

            foreach (AppointmentItem item in calendarItems)
            {
                CreateEvent(ics, item, false);
            }
            
            return ics;
        }

        static void Main(string[] args)
        {
            iCalendar ics = GenerateIcs();
            iCalendarSerializer serializer = new iCalendarSerializer();
            string output = serializer.SerializeToString(ics);
            //Console.Write(output);
            FtpUpload(output);
            //Console.Read();
        }
    }
}
