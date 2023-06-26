using SharpShell.Attributes;
using System.Runtime.InteropServices;
using SharpShell.SharpDeskBand;


namespace TaskbarMeetingNotifier

{
    [ComVisible(true)]
    [DisplayName("Taskbar Meeting Notifier")]


    public class TaskbarMeetingNotifierMain : SharpDeskBand
    {
        protected override System.Windows.Forms.UserControl CreateDeskBand()
        {
            return new TaskbarMeetingNotifierUI.TaskbarMeetingNotifierUI();
        }

        protected override BandOptions GetBandOptions()
        {
            return new BandOptions
            {
                HasVariableHeight = true,
                IsSunken = false,
                ShowTitle = true,
                Title = "Taskbar Meeting Notifier",
                UseBackgroundColour = false,
                AlwaysShowGripper = true
            };
        }
    }
}