using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameworkLibraries.Utils
{
    public class OSOperations
    {
        public static void KillProcess(string processName)
        {
            foreach (Process p in Process.GetProcesses("."))
            {
              
                if (p.ProcessName.Contains(processName) || p.ProcessName.Contains(processName.ToUpper()))
                {
                    p.Kill();
                }
            }
        }

        public static void CommandLineExecute(string cmd)
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = "/c "+cmd;
            process.StartInfo = startInfo;
            process.Start();
        }

        public static void InvokeInstaller(string installDir, string exeName)
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.WorkingDirectory = Path.Combine(installDir);
            startInfo.FileName = exeName;
            process.StartInfo = startInfo;
            process.Start();
        }
    }


}
