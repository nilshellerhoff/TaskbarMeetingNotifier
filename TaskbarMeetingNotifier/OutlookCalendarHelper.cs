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
        public static int Main()
        {
            var next = OutlookCalendarHelper.getNextAppointment();
            Console.WriteLine(next.ToString());
            return 0;
        }

        public static Dictionary<String, Object> AppointmentToDict(Outlook.AppointmentItem app)
        {
            return new Dictionary<string, object>(){
                { "Subject", app.Subject },
                { "Start", app.Start },
                { "End", app.End },
                { "Body", app.Body }
            };
        }

        public static Dictionary<String, Object> getNextAppointment()
        {
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();

                // Get the NameSpace and Logon information.
                // Outlook.NameSpace oNS = (Outlook.NameSpace)oApp.GetNamespace("mapi");
                Outlook.NameSpace oNS = oApp.GetNamespace("mapi");

                //Log on by using a dialog box to choose the profile.
                oNS.Logon(Missing.Value, Missing.Value, true, true);

                //Alternate logon method that uses a specific profile.
                // TODO: If you use this logon method, 
                // change the profile name to an appropriate value.
                //oNS.Logon("YourValidProfile", Missing.Value, false, true); 

                // Get the Calendar folder.
                Outlook.MAPIFolder oCalendar = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

                // Get the Items (Appointments) collection from the Calendar folder.
                Outlook.Items oItems = oCalendar.Items;
                oItems.Sort("[Start]");


            // Get the first item.
            //Outlook.AppointmentItem oAppt = (Outlook.AppointmentItem)oItems

            DateTime now = DateTime.Now;
                var fiveMinutesAgo = now.Add(new TimeSpan(0, -5, 0));

                var sFilter = "([Start] >= " + fiveMinutesAgo.ToString("ddddd h:mm tt") + ")";

                var appointment = oItems.Find(sFilter);

                //// Show some common properties.
                //Console.WriteLine("Subject: " + oAppt.Subject);
                ////Console.WriteLine("Organizer: " + oAppt.Organizer);
                //Console.WriteLine("Body:" + oAppt.);
                ////Console.WriteLine(oAppt.)
                //Console.WriteLine("Start: " + oAppt.Start.ToString());
                //Console.WriteLine("End: " + oAppt.End.ToString());
                //Console.WriteLine("Location: " + oAppt.Location);
                //Console.WriteLine("Recurring: " + oAppt.IsRecurring);

                //Show the item to pause.
                //oAppt.Display(true);

                // Done. Log off.
                oNS.Logoff();

                // Clean up.
                //oAppt = null;
                oItems = null;
                oCalendar = null;
                oNS = null;
                oApp = null;

                return OutlookCalendarHelper.AppointmentToDict(appointment);
        }
    }
}
