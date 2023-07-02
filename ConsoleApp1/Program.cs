using System;
using TaskbarMeetingNotifier;

namespace ConsoleApp1
{
    class Program
    {
        static int Main(string[] args)
        {
            var next = OutlookCalendarHelper.getNextAppointment();
            Console.WriteLine("Subject: " + next["Subject"].ToString());
            Console.WriteLine("Start: " + next["Start"].ToString());
            Console.WriteLine("End: " + next["End"].ToString());
            //Console.WriteLine("Body: " + next["Body"]);
            Console.WriteLine("EntryID" + next["EntryID"]);

            OutlookCliHelper.OpenAppointment(next["EntryID"].ToString());

            return 0;

        }
    }
}
