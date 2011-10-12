using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Outlook2iCal
{
    class Configuration
    {
        public static readonly string FtpUrl = "ftp://SERVER/PATH/outlook.ics";
        public static readonly string FtpUser = "USERNAME";
        public static readonly string FtpPass = "PASSWORD";

        public static readonly string MailDomain = "@example.com";

        public static readonly bool IncludeClass = false;

        public static readonly string[] CleanSubjects = new string[] { "Updated: ", "FW: " };
        public static readonly string DescriptionStart = "*~*~*~*~*~*~*~*~*~*";
        public static readonly string DescriptionEnd = "-+-----+-----+-----+-----+-----+-----+-----+-----+-";

        public static readonly string[] SkipCategories = new string[] { "Reminder", "Vacation" };
    }
}
