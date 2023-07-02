using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Diagnostics;

namespace TaskbarMeetingNotifier
{
    public class OutlookCliHelper
    {
        public static void OpenAppointment(string EntryID)
        {
            var outlookExecutable = OutlookCliHelper.GetOutlookExecutable();
            var Arguments = "/select outlook:" + EntryID;
            Process.Start(outlookExecutable, Arguments);
        }

        public static string GetOutlookExecutable()
        {
            var AppPathKey = @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE";
            return Registry.GetValue(AppPathKey, null, null).ToString();
        }
    }
}