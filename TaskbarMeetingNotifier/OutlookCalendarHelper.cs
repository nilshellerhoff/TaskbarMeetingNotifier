using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System;
using System.Reflection; // to use Missing.Value

using Outlook = Microsoft.Office.Interop.Outlook;


namespace TaskbarMeetingNotifier
{
    public class OutlookCalendarHelper
    {
        private Outlook.Application oApp;
        private Outlook.NameSpace oNS;
        private Outlook.Items oItems;
        private Outlook.MAPIFolder oCalendar;

        public static int Main()
        {
            var next = OutlookCalendarHelper.getNextAppointment();
            Console.WriteLine(next.ToString());
            return 0;
        }

        public static Dictionary<String, Object> AppointmentToDict(Outlook.AppointmentItem app)
        {
            var Now = OutlookCalendarHelper.GetNowMinusTime();
            Outlook.AppointmentItem appointment;
            //if (app.IsRecurring)
            //{
            //    // recurring appointments must always be at the same time of day
            //    // construct the target DateTime from the Time of the reccurrent item and the date of today
            //    var timeOfDay = app.Start.TimeOfDay;
            //    var TargetDateTime = Now.Date.Add(timeOfDay);
            //    var rec = app.GetRecurrencePattern();
            //    appointment = rec.GetOccurrence(TargetDateTime);
            //}
            //else
            //{
            //    appointment = app;
            //}

            return new Dictionary<string, object>(){
                { "Subject", app.Subject },
                { "Start", app.StartUTC },
                { "End", app.EndUTC },
                { "TimeTillMinutes", (app.StartUTC - DateTime.UtcNow).TotalMinutes },
                //{ "Body", app.Body }, // accessing Body triggers Outlook warning (maybe)
                { "EntryID", app.EntryID }
            };
        }

        public static DateTime GetNowMinusTime()
        {
            DateTime now = DateTime.Now;
            var fiveMinutesAgo = now.Add(new TimeSpan(0, -5, 0));
            return fiveMinutesAgo;
        }

        public static Dictionary<String, Object> getNextAppointment()
        {
            var helper = new OutlookCalendarHelper();
            helper.OutlookLogin();

            // Get the first item.
            //Outlook.AppointmentItem oAppt = (Outlook.AppointmentItem)oItems

            var fiveMinutesAgo = OutlookCalendarHelper.GetNowMinusTime();
            var timeFilter = "([Start] >= " + fiveMinutesAgo.ToString("ddddd h:mm tt") + ")";
            var recurringFilter = "([IsRecurring] = False)";
            var sFilter = timeFilter + "AND" + recurringFilter;

            helper.oItems.IncludeRecurrences = false;
            var appointment = helper.oItems.Find(sFilter);

            helper.OutlookLogoff();

            return OutlookCalendarHelper.AppointmentToDict(appointment);
        }

        private void OutlookLogin()
        {
            // connect to the Outlook Application
            // Create the Outlook application.
            this.oApp = new Outlook.Application();

            // Get the NameSpace and Logon information.
            // Outlook.NameSpace oNS = (Outlook.NameSpace)oApp.GetNamespace("mapi");
            this.oNS = oApp.GetNamespace("mapi");

            //Log on by using a dialog box to choose the profile.
            this.oNS.Logon(Missing.Value, Missing.Value, true, true);

            //Alternate logon method that uses a specific profile.
            // TODO: If you use this logon method, 
            // change the profile name to an appropriate value.
            //oNS.Logon("YourValidProfile", Missing.Value, false, true); 

            // Get the Calendar folder.
            this.oCalendar = this.oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

            // Get the Items (Appointments) collection from the Calendar folder.
            this.oItems = oCalendar.Items;
            this.oItems.Sort("[Start]");
        }

        private void OutlookLogoff()
        {
            // Done. Log off.
            this.oNS.Logoff();

            // Clean up.
            this.oItems = null;
            this.oCalendar = null;
            this.oNS = null;
            this.oApp = null;
        }
    }
}
