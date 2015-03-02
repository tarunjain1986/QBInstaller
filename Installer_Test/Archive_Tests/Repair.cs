using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Management;
using System.Reflection;
using System.Diagnostics;
//using System.Drawing;
//using System.Drawing.Imaging;
using System.Collections.Generic;
//using System.Runtime.InteropServices;

using FrameworkLibraries;
using FrameworkLibraries.Utils;
using FrameworkLibraries.ActionLibs;
using FrameworkLibraries.AppLibs.QBDT;
using TestStack.White.UIItems.WindowItems;
using FrameworkLibraries.ActionLibs.WhiteAPI;

using TestStack.BDDfy;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;

using Xunit;

using Installer_Test;
using Installer_Test.Lib;

//using ScreenShotDemo;

using Microsoft.Win32;

namespace Installer_Test.Archive_Tests
{
    public class Repair
    {
        //[DllImport("User32.dll")]
        //public static extern int SetForegroundWindow(IntPtr point);
        //[DllImport("User32.dll")]
        //private static extern IntPtr GetForegroundWindow();
        
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public static int Sync_Timeout = int.Parse(conf.get("SyncTimeOut"));
        public static string testName = "Repair";
        public string SKU, ver, reg_ver, installed_product, installed_path, installed_dir;
        string OS_Name = string.Empty;

    //    Object product, path;
  

        [Given (StepTitle = @"QuickBooks is installed")]
        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);

          // Install_Functions.CleanUp();


            //string readpath = @"C:\Temp\Parameters.txt";
            //File.WriteAllLines(readpath, File.ReadAllLines(readpath).Where(l => !string.IsNullOrWhiteSpace(l))); // Remove white space from the file

            //string[] lines = File.ReadAllLines(readpath);
            //var dic = lines.Select(line => line.Split('=')).ToDictionary(keyValue => keyValue[0], bits => bits[1]);

            //ver = dic["Version"];
            //reg_ver = dic["Registry Folder"];
            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");
            string Edition = qbWindow.Title;

            //string tobesearched = "code : "
            //string code = myString.Substring(myString.IndexOf(tobesearched) + tobesearched.Length);

            Edition = Edition.Substring (Edition.IndexOf ("Intuit ") + "Intuit ".Length);

            if (Edition.Contains ("Manufacturing and Wholesale"))
            {
                Edition = Edition.Replace("Manufacturing and Wholesale", "Mfg and Whsle");
            }


            string left_str  = Edition.Substring(0, Edition.LastIndexOf(' '));
            string right_str = Edition.Substring(Edition.LastIndexOf(' ') + 1);

            if (left_str != "QuickBooks Enterprise Solutions")
            {
                Edition = left_str + " Edition " + right_str;
            }

            Install_Functions.Add_Edition_Automation_Properties(Edition);

            //string readpath = "C:\\Temp\\Parameters.xlsm";

            //Dictionary<string, string> dic = new Dictionary<string, string>();
            ////dic = File_Functions.ReadExcelValues(readpath, "PostInstall", "B2:B4");
            ////ver = dic["B2"];
            ////reg_ver = dic["B3"];

            //dic = File_Functions.ReadExcelValues(readpath, "Install", "B9:B12");
            //ver = dic["B9"];
            //SKU = dic["B12"];

            //dic = File_Functions.ReadExcelValues(readpath, "Install", "E7");
            //reg_ver = File_Functions.GetRegVer(SKU);

            //OS_Name = File_Functions.GetOS();
            //installed_product = File_Functions.GetProduct(ver, reg_ver);

            //if (installed_product.Contains ("QuickBooks Premier"))
            //{
            //  installed_product = installed_product.Replace("QuickBooks Premier", "QuickBooks Premier Edition");
            //}

            //installed_path = File_Functions.GetPath(ver, reg_ver);         
            //installed_dir = Path.GetDirectoryName(installed_path); // Get the path (without the exe name)
        }

        [Then (StepTitle = "Delete dlls")]
        public void DeleteDLLs()
        {
            //qbApp = QuickBooks.GetApp("QuickBooks");
            //qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");

            //Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
            //Actions.SelectMenu(qbApp, qbWindow, "File", "Exit");
            //Thread.Sleep(10000);
            OSOperations.KillProcess("QBW32");
            Thread.Sleep(1000);
           // Install_Functions.Delete_QBDLLs(installed_dir);
        }

        //[AndThen(StepTitle = "Invoke QuickBooks")]
        //public void InvokeQB()
        //{
        //    // QuickBooks.Initialize(installed_path);

        //    Process proc = new Process();
        //    proc.StartInfo.FileName = installed_path;
        //    proc.Start();

        //    Thread.Sleep(1000);
        //    Boolean flag;
        //    flag = Actions.CheckDesktopWindowExists("Error");
        //    if (flag == true)
        //    {
        //        Actions.ClickElementByName(Actions.GetDesktopWindow("Error"), "OK");
        //    }
        //    Thread.Sleep(1000);
        //}

        [AndThen (StepTitle = "Repair QuickBooks")]
        public void RepairQB ()
        {
            //Repair
            // QuickBooks.RepairOrUnInstallQB(installed_product, true, false);
            conf.reload();
            installed_product = conf.get("Edition");
            QuickBooks.RepairOrUnInstallQB(installed_product, true, false);

            string exe = conf.get("QBExePath");
            // Invoke QB after Repair 
            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);

            Thread.Sleep(20000);
            qbApp.WaitWhileBusy();
            Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks " + SKU));

           

            // Close QuickBook pop-up windows
            Install_Functions.CheckWindowsAndClose(SKU);

            Thread.Sleep(1000);
            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks " + SKU);
            Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
            Actions.SelectMenu(qbApp, qbWindow, "File", "Exit");


            // Uninstall QuickBooks

            OSOperations.KillProcess("QBW32");
            Thread.Sleep(1000);

            // Get the QuickBooks Edition to Repair from the Automation.Properties file
            conf.reload();
            installed_product = conf.get("Edition");
            QuickBooks.RepairOrUnInstallQB(installed_product, false, true);
        }

        [AndThen(StepTitle = "Invoke QuickBooks after repair")]
        public void InvokeQB_afterRepair()
        {
           // QuickBooks.Initialize(installed_path);
        }

        [Fact]
        public void RunQBRepairTest()
        {
            this.BDDfy();
        }
    }
}
