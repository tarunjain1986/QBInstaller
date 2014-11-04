using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Forms;

namespace FrameworkLibraries.Utils
{
    public class KeyStrokeSimulator
    {
            [DllImport("user32.dll", EntryPoint = "FindWindowEx")]
            public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
            [DllImport("User32.dll")]
            public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, string lParam);

            public static void SendUsingProcess(string process, string message)
            {
                Process[] proc = Process.GetProcessesByName(process);
                if (proc.Length == 0)
                    return;
                if (proc[0] != null)
                {
                    Debug.WriteLine("found..");
                    IntPtr child = FindWindowEx(proc[0].MainWindowHandle, new IntPtr(0), null, null);
                    SendMessage(child, 0x000C, 0, message);
                }
            }

            public static void SendKeysAsCharacters(String input)
            {
                foreach (char c in input)
                {
                    SendKeys.SendWait(c.ToString());
                    Thread.Sleep(100);
                }
            }

            public static void SendKey(String input)
            {
                SendKeys.SendWait(input);
                Thread.Sleep(500);
            }

            public static void SendKeysAsNumeric(String input)
            {
                foreach (char c in input)
                {
                    SendKeys.SendWait("{"+c+"}");
                    Thread.Sleep(100);
                }
            }

    }
}
