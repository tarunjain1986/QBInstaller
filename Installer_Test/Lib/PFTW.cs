using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

using FrameworkLibraries;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.ActionLibs;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using ScreenShotDemo;
using TestStack.BDDfy;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems;

namespace Installer_Test
{

    public class PFTW
    {
        public static void Windiff_Compare(string Local_Windiff, string Local_B1Path,string Local_B2Path)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("Windiff comparison started:" + Local_B1Path + Local_B2Path + " - Started..");

            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
            string cmdText = "/c cd " + Local_Windiff + " && windiff " + Local_B1Path + " " + Local_B2Path + " -Sslrdx " + Local_Windiff + "Diff.txt";
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = cmdText;
            startInfo.UseShellExecute = true;
            process.StartInfo = startInfo;
            try
            {
                process.Start();
                process.WaitForExit();
                Logger.logMessage("Windiff comparison: " + Local_B1Path + Local_B2Path + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Windiff comparison: " + Local_B1Path + Local_B2Path + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                Logger.logMessage("------------------------------------------------------------------------------");
            }


        }
    }
}
