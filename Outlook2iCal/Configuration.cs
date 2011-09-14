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
    }
}
