using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.IO;
using Microsoft.Office.Interop.Outlook;
using DDay.iCal;
using DDay.iCal.Serialization.iCalendar;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace Outlook2iCal
{
    public partial class Outlook2iCal : Form
    {
        private string tzid;

        private void FtpUpload(string input)
        {
            try
            {
                exceptLabel.Text = "Upload:";
                currentBox.Text = "Uploading...";
                byte[] fileContents = Encoding.UTF8.GetBytes(input);
                exceptBar.Maximum = fileContents.Length;
                
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(Properties.Settings.Default.FtpUrl);
                request.Method = WebRequestMethods.Ftp.UploadFile;

                request.Credentials = new NetworkCredential(Properties.Settings.Default.FtpUser, Properties.Settings.Default.FtpPass);

                request.ContentLength = fileContents.Length;

                using (Stream fileStream = new MemoryStream(fileContents))
                using (Stream requestStream = request.GetRequestStream())
                {
                    var buffer = new byte[1024];
                    int totalReadBytesCount = 0;
                    int readBytesCount;
                    while ((readBytesCount = fileStream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        requestStream.Write(buffer, 0, readBytesCount);
                        totalReadBytesCount += readBytesCount;
                        exceptBar.Value = totalReadBytesCount;
                        //exceptText.Text = totalReadBytesCount + "/" + fileContents.Length;
                    }
                }

                FtpWebResponse response = (FtpWebResponse)request.GetResponse();

                currentBox.Text = "Upload Complete";

                response.Close();
            }
            catch (WebException)
            {
                currentBox.Text = "Upload Failed";
            }
        }

        private void DaysOfWeek(DDay.iCal.RecurrencePattern pattern, OlDaysOfWeek mask, int week)
        {
            FrequencyOccurrence occur;
            switch (week)
            {
                case 1:
                    occur = FrequencyOccurrence.First;
                    break;
                case 2:
                    occur = FrequencyOccurrence.Second;
                    break;
                case 3:
                    occur = FrequencyOccurrence.Third;
                    break;
                case 4:
                    occur = FrequencyOccurrence.Fourth;
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

        private string CleanupName(string rawName)
        {
            if ((rawName[0] == '\'') && (rawName[rawName.Length - 1] == '\''))
            {
                return rawName.Substring(1, rawName.Length - 2).Replace(", ", ",");
            }
            else
            {
                return rawName;
            }
        }

        private void CreateReoccuringEvent(Event icsEvent, AppointmentItem item)
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
                    DateTime exDate = except.OriginalDate;
                    if (exDate.TimeOfDay.TotalSeconds > 0)
                    {
                        list.Add(new iCalDateTime(exDate, tzid));
                    }
                    else
                    {
                        list.Add(new iCalDateTime(exDate.Add(item.Start.TimeOfDay), tzid));
                    }
                }
                icsEvent.ExceptionDates.Add(list);
            }
        }

        private void CreateEvent(iCalendar ics, AppointmentItem item, bool notRecurring)
        {
            currentBox.Text = item.Subject;
            Event icsEvent = new Event();

            if (item.Categories != null)
            {
                string[] cats = item.Categories.Split(',');
                foreach (string cat in cats)
                {
                    icsEvent.Categories.Add(cat);
                    if (Properties.Settings.Default.SkipCategories.Contains(cat))
                    {
                        return;
                    }
                }
            }

            if (item.AllDayEvent)
            {
                icsEvent.IsAllDay = true;
                icsEvent.DTStart = new iCalDateTime(item.Start, tzid);
                if (notRecurring && !item.IsRecurring)
                {
                    icsEvent.DTEnd = new iCalDateTime(item.End, tzid);
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

            string summary = item.Subject;
            if (summary.StartsWith("Canceled: ", StringComparison.CurrentCultureIgnoreCase))
            {
                return;
            }
            foreach (string filter in Properties.Settings.Default.CleanSubjects)
            {
                if (summary.StartsWith(filter, StringComparison.CurrentCultureIgnoreCase))
                {
                    summary = summary.Substring(filter.Length);
                }
            }
            icsEvent.Summary = summary;

            if (Properties.Settings.Default.IncludeClass)
            {
                switch (item.Sensitivity)
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
                }
            }

            string descr = item.Body;
            if (descr != null)
            {
                // I have no idea what the proper way of handling hyperlinks in iCal feeds is, I'll implement this if I learn it.
                /*if (descr.IndexOf("HYPERLINK") > -1)
                {
                    Regex regex = new Regex(@"HYPERLINK ""(?<url>.*)""(?<link>.*)\r", RegexOptions.Multiline);
                    descr = regex.Replace(descr, new MatchEvaluator(Outlook2iCal.ReplaceHyperlink));
                }*/

                int startIndex = descr.IndexOf(Properties.Settings.Default.DescriptionStart);
                int endIndex = descr.IndexOf(Properties.Settings.Default.DescriptionEnd);
                if (startIndex == -1)
                {
                    startIndex = 0;
                }
                else
                {
                    startIndex += Properties.Settings.Default.DescriptionStart.Length;
                }
                if (endIndex == -1)
                {
                    icsEvent.Description = descr.Substring(startIndex).Trim();
                }
                else
                {
                    icsEvent.Description = descr.Substring(startIndex, endIndex - startIndex).Trim();
                }
            }

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
                string email = "MAILTO:" + recip.Address.Substring(recip.Address.LastIndexOf('=') + 1);
                if (email.IndexOf('@') == -1)
                {
                    email += Properties.Settings.Default.MailDomain;
                }
                emails.Add(recip.Name, email);
                status.Add(recip.Name, recip.MeetingResponseStatus);
            }

            if (item.Organizer != null)
            {
                if (emails.ContainsKey(item.Organizer))
                {
                    icsEvent.Organizer = new Organizer(emails[item.Organizer]);
                }
                else
                {
                    icsEvent.Organizer = new Organizer();
                }
                icsEvent.Organizer.CommonName = CleanupName(item.Organizer);
            }

            if (item.RequiredAttendees != null)
            {
                string[] required = item.RequiredAttendees.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string req in required)
                {
                    Attendee attend;
                    if (emails.ContainsKey(req))
                    {
                        attend = new Attendee(emails[req]);
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
                    attend.CommonName = CleanupName(req);
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
                        attend = new Attendee(emails[opt]);
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
                    attend.CommonName = CleanupName(opt);
                    icsEvent.Attendees.Add(attend);
                }
            }

            if (item.ResponseStatus == OlResponseStatus.olResponseAccepted)
            {
                icsEvent.Status = EventStatus.Confirmed;
            }
            else
            {
                icsEvent.Status = EventStatus.Tentative;
            }

            icsEvent.DTStamp = new iCalDateTime(item.CreationTime, tzid);
            icsEvent.LastModified = new iCalDateTime(item.LastModificationTime, tzid);

            byte[] buf = Encoding.Default.GetBytes(item.EntryID + item.Start.ToFileTimeUtc());
            SHA1CryptoServiceProvider sha = new SHA1CryptoServiceProvider();
            icsEvent.UID = BitConverter.ToString(sha.ComputeHash(buf)).Replace("-", "");

            if (!notRecurring && item.IsRecurring)
            {
                Microsoft.Office.Interop.Outlook.RecurrencePattern pattern = item.GetRecurrencePattern();

                int count = 0;
                exceptBar.Maximum = pattern.Exceptions.Count;

                foreach (Microsoft.Office.Interop.Outlook.Exception except in pattern.Exceptions)
                {
                    count++;
                    exceptBar.Value = count;
                    //exceptText.Text = count + "/" + pattern.Exceptions.Count;
                    if (!except.Deleted)
                    {
                        CreateEvent(ics, except.AppointmentItem, true);
                    }
                }

                exceptBar.Value = 0;
                //exceptText.Text = String.Empty;
            }

            ics.Events.Add(icsEvent);
        }

        // I have no idea what the proper way of handling hyperlinks in iCal feeds is, I'll implement this if I learn it.
        /*public static string ReplaceHyperlink(Match match)
        {
            string url = match.Groups["url"].Captures[0].Value;
            string link = match.Groups["link"].Captures[0].Value;
            return "<a href=\"" + url + "\">" + link + "</a>";
        }*/

        private iCalendar GenerateIcs()
        {
            iCalendar ics = new iCalendar();

            ITimeZone tz = ics.AddLocalTimeZone();
            tzid = tz.TZID;

            ics.ProductID = "-//David Maher/NONSGML Outlook2iCal 2.0//EN";

            Items calendarItems = new Microsoft.Office.Interop.Outlook.Application().GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderCalendar).Items;
            calendarItems.IncludeRecurrences = true;

            int count = 0;
            eventBar.Maximum = calendarItems.Count;

            foreach (AppointmentItem item in calendarItems)
            {
                count++;
                eventBar.Value = count;
                //eventText.Text = count + "/" + calendarItems.Count;
                CreateEvent(ics, item, false);
            }

            return ics;
        }

        public string SeperateExdates(string ics, string tzid)
        {
            //I freely admit this is a kludge, but whatever, it works
            string output = String.Empty;
            string exdate = String.Empty;
            StringReader reader = new StringReader(ics);
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                if (exdate.Length == 0)
                {
                    if (line.Length > 7 && line.Substring(0, 7) == "EXDATE:")
                    {
                        exdate = line.Substring(7);
                    }
                    else
                    {
                        output += line + "\r\n";
                    }
                }
                else
                {
                    if (line[0] == ' ')
                    {
                        exdate += line.Substring(1);
                    }
                    else
                    {
                        string[] exdates = exdate.Split(',');
                        foreach (string curexdate in exdates)
                        {
                            output += "EXDATE;TZID=" + tzid + ":" + curexdate + "\r\n";
                        }
                        exdate = String.Empty;
                        output += line + "\r\n";
                    }
                }
            }
            return output;
        }
        
        public Outlook2iCal()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        private void startButton_Click(object sender, EventArgs e)
        {
            currentBox.ForeColor = Color.Black;
            startButton.Enabled = false;
            exceptLabel.Text = "Exceptions:";
            //exceptText.Text = String.Empty;
            exceptBar.Value = 0;
            currentBox.Text = String.Empty;
            backgroundWorker.RunWorkerAsync();
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            iCalendar ics = GenerateIcs();
            iCalendarSerializer serializer = new iCalendarSerializer();
            string output = serializer.SerializeToString(ics);
            output = SeperateExdates(output, ics.TimeZones[0].TZID);
            FtpUpload(output);
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            startButton.Enabled = true;
        }

        private void eventBar_ValueChanged(object sender, EventArgs e)
        {
            if (eventBar.Value != 0)
            {
                eventText.Text = eventBar.Value + "/" + eventBar.Maximum;
            }
            else
            {
                eventText.Text = String.Empty;
            }
        }

        private void exceptBar_ValueChanged(object sender, EventArgs e)
        {
            if (exceptBar.Value != 0)
            {
                exceptText.Text = exceptBar.Value + "/" + exceptBar.Maximum;
            }
            else
            {
                exceptText.Text = String.Empty;
            }
        }
    }
}
