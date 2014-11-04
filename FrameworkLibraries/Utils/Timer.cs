using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;

namespace FrameworkLibraries.Utils
{
    public class Timer
    {
        public static void SetTime(int timeOut)
        {
            var aTimer = new System.Timers.Timer(timeOut);
            aTimer.Start();
            aTimer.Elapsed += HandleTimerElapsed;
        }

        public static void HandleTimerElapsed(object sender, ElapsedEventArgs e)
        {
            MessageBox.Show(e.SignalTime.ToString());
        }
    }
}
