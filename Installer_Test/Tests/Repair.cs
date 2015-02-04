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

using TestStack.BDDfy;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;

using Xunit;

using Installer_Test;
using Installer_Test.Lib;

//using ScreenShotDemo;

using Microsoft.Win32;

namespace QBInstall.Tests
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
        public string ver, reg_ver, installed_product, installed_path, installed_dir;
        string OS_Name = string.Empty;

        Object product, path;
  

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

            string readpath = "C:\\Temp\\Parameters.xlsm";

            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic = File_Functions.ReadExcelValues(readpath, "PostInstall", "B2:B4");
            ver = dic["B2"];
            reg_ver = dic["B4"];

            OS_Name = File_Functions.GetOS();
            installed_product = File_Functions.GetProduct(ver, reg_ver);
            installed_path = File_Functions.GetPath(ver, reg_ver);         
            installed_dir = Path.GetDirectoryName(installed_path); // Get the path (without the exe name)
        }

        [Then (StepTitle = "Delete dlls")]
        public void DeleteDLLs()
        {
            Install_Functions.Delete_QBDLLs(installed_dir);
        }

        [AndThen(StepTitle = "Invoke QuickBooks")]
        public void InvokeQB()
        {
            QuickBooks.Initialize(installed_path);
        }

        [AndThen (StepTitle = "Repair QuickBooks")]
        public void RepairQB ()
        {
            //Repair
            QuickBooks.RepairOrUnInstallQB(installed_product, true, false);
        }

        [AndThen(StepTitle = "Invoke QuickBooks after repair")]
        public void InvokeQB_afterRepair()
        {
            QuickBooks.Initialize(installed_path);
        }

        [Fact]
        public void RunQBRepairTest()
        {
            this.BDDfy();
        }
    }
}
