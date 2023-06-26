using System;
using System.Windows.Forms;
using System.Runtime;


namespace TaskbarMeetingNotifierUI
{
    public partial class TaskbarMeetingNotifierUI : UserControl
    {
        public TaskbarMeetingNotifierUI()
        {
            InitializeComponent();
            Timer timer = new System.Windows.Forms.Timer();
            timer.Interval = 1000;
            timer.Tick += UpdateClock;
            timer.Start();

        }

        public void UpdateClock(object sender, EventArgs e)
        {
            string TimeStr = DateTime.Now.ToString();
            timeLabel.Text = TimeStr;
        }
    }
}
