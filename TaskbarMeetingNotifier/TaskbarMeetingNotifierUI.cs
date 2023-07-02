using System;
using System.Windows.Forms;
using System.Runtime;
using TaskbarMeetingNotifier;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace TaskbarMeetingNotifierUI
{
    public partial class TaskbarMeetingNotifierUI : UserControl
    {

        public Dictionary<String, Object> NextAppointment;

        public TaskbarMeetingNotifierUI()
        {
            InitializeComponent();
            Timer timer = new Timer();
            timer.Interval = 10*1000;
            timer.Tick += UpdateInterface;
            timer.Start();

        }

        public void UpdateClock(object sender, EventArgs e)
        {
            string TimeStr = DateTime.Now.ToString();
            meetingLabel.Text = TimeStr;
        }

        public void UpdateInterface(object sender, EventArgs e)
        {
            Task UpdateAppointment = new Task(this.UpdateAppointment);
            UpdateAppointment.Start();
        }

        private void UpdateAppointment()
        {
            NextAppointment = OutlookCalendarHelper.getNextAppointment();
            //meetingLabel.Text = NextAppointment["Subject"].ToString();
            var TimeTillNext = "";
            int TimeTillNextMinutes = Convert.ToInt32(NextAppointment["TimeTillMinutes"]);
            if (TimeTillNextMinutes > 24 * 60)
            {
                TimeTillNext = String.Format("in {0} days", TimeTillNextMinutes / 24 / 60);
            }
            else if (TimeTillNextMinutes > 60)
            {
                TimeTillNext = String.Format("in {0} hours", TimeTillNextMinutes / 60);
            }
            else if (TimeTillNextMinutes >= 0)
            {
                TimeTillNext = String.Format("in {0} minutes", TimeTillNextMinutes);
            } else
            {
                TimeTillNext = String.Format("{0} minutes ago", -1 * TimeTillNextMinutes);
            }
            //TimeTillNext = String.Format("in {0} minutes", TimeTillNextMinutes);
            meetingLabel.Text = TimeTillNext + ": " + NextAppointment["Subject"].ToString();
            //meetingLabel.Text = NextAppointment["TimeTillMinutes"].ToString() + ": " + NextAppointment["Subject"].ToString();
        }

        private void showButton_click(object sender, EventArgs e)
        {
            Task OpenAppointment = new Task(this.OpenAppointment);
            OpenAppointment.Start();
        }

        private void OpenAppointment()
        {
            OutlookCliHelper.OpenAppointment(NextAppointment["EntryID"].ToString());
        }
    }
}
