using System;
using FrameworkLibraries.Utils;
using System.Windows.Automation;
using System.Windows.Forms;
using FrameworkLibraries.ActionLibs;
using TestStack.White.UIItems.WindowItems;
using System.Threading;
using TestStack.White.UIItems.Finders;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using FrameworkLibraries;
using System.Collections.Generic;
using TestStack.White.UIItems;
using Xunit;
using TestStack.BDDfy;
using FrameworkLibraries.AppLibs.QBDT;
using System.IO;
using System.Reflection;
using Installer_Test;
using Installer_Test.Lib;


namespace AutoPatch
{
 
    public class AutoPatch
    {
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public Random rand = new Random();
        public string testName = "AutoPatch";

        [Given(StepTitle = "Given - QuickBooks App and Window instances are available")]
        public void Method()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
                        
            //// Close QuickBooks
            //qbApp = QuickBooks.GetApp("QuickBooks");
            //qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks " + "Enterprise");
            //Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
            //Actions.SelectMenu(qbApp, qbWindow, "File", "Exit");

            // Kill QuickBooks processes
            OSOperations.KillProcess("QBW32");
            OSOperations.KillProcess("qbupdate");

            Thread.Sleep(1000);

            string ver = "25.0",reg_ver = "bel", SKU = "Enterprise";
            string installed_datapath, service_dest, service_source = @"C:\Temp\", service_backup;

            installed_datapath = File_Functions.GetDataPath(ver, reg_ver);
            if (installed_datapath != "")
            {
                // Replace the local copy of the serviceguide.xml with the server copy
                service_dest = installed_datapath + @"Components\QBUpdate\serviceguide.xml";
                service_source = service_source + "serviceguide.xml";
                service_backup = installed_datapath + @"Components\QBUpdate\serviceguide_backup.xml";

                if (File.Exists(service_backup))
                {
                    File.SetAttributes(service_backup, FileAttributes.Normal);
                    File.Delete(service_backup);
                }

                System.IO.File.Move(service_dest, service_backup);
                File.Copy(service_source, service_dest);
                // Thread.Sleep(1000);
            }
            // File.Replace(service_source, service_dest, service_backup, true);
            // File.SetAttributes(service_dest, FileAttributes.Normal);

            // Launch QuickBooks
            string exe = conf.get("QBExePath");
            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            Thread.Sleep(2000);

            //////////////////////////////////////////////////////////////////////////////////////////////////////
            Boolean flag = false;

            flag = Actions.CheckDesktopWindowExists("QuickBooks Update Service");
            if (flag == true)
            {
                Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks Update Service"));
                SendKeys.SendWait("%l");
                Logger.logMessage("QuickBooks Update Service Window found.");
            }
  
            flag = false;

            while (flag == false)
            {
                flag = Actions.CheckDesktopWindowExists("QuickBooks " + SKU);

            }
            //////////////////////////////////////////////////////////////////////////////////////////////////////

            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
          

            var MainWindow = Actions.GetDesktopWindow("QuickBooks " + SKU);

            // Close QuickBook pop-up windows
            Install_Functions.CheckWindowsAndClose(SKU);

            Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
            // Actions.SelectMenu(qbApp, qbWindow, "Help", "Update QuickBooks...");

            string resultsPath = @"C:\Temp\AP\";
            Installer_Test.Help.ClickHelpUpdate_AutoPatch(qbApp, qbWindow, resultsPath);

            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");
            Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
            Actions.SelectMenu(qbApp, qbWindow, "File", "Exit");

            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);

            // Close QuickBook pop-up windows
            Install_Functions.CheckWindowsAndClose(SKU);

            Thread.Sleep(2000);
        }

        [Fact]
        public void Run_AntiVirusTest()
        {
            this.BDDfy();
        }
    }
}
