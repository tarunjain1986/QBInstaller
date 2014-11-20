using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Management;
using System.Reflection;
using System.Diagnostics;
using System.Collections.Generic;

using FrameworkLibraries;
using FrameworkLibraries.Utils;
using FrameworkLibraries.ActionLibs;
using TestStack.White.UIItems.WindowItems;

using TestStack.BDDfy;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;

using Xunit;

using Microsoft.Win32;

using FrameworkLibraries.AppLibs.QBDT;

using Installer_Test;
using Installer_Test.Lib;

namespace QBInstall.Tests
{
    public class Remove
    {
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public static int Sync_Timeout = int.Parse(conf.get("SyncTimeOut"));
        public static string testName = "Uninstall";
        public string ver, reg_ver, installed_product;
        string OS_Name = string.Empty;

        Object product;

        [Given]
        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);

            //string readpath = @"C:\Temp\Parameters.txt";
            //File.WriteAllLines(readpath, File.ReadAllLines(readpath).Where(l => !string.IsNullOrWhiteSpace(l))); // Remove white space from the file

            //string[] lines = File.ReadAllLines(readpath);
            //var dic = lines.Select(line => line.Split('=')).ToDictionary(keyValue => keyValue[0], bits => bits[1]);

            //ver = dic["Version"];
            //reg_ver = dic["Registry Folder"];

            string readpath = "C:\\Temp\\Parameters.xlsm"; 

            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic = File_Functions.ReadExcelValues(readpath, "PostInstall", "B2:B10");
            ver = dic["B7"];
            reg_ver = dic["B8"];

            OS_Name = File_Functions.GetOS();
            installed_product = Installer_Test.Lib.File_Functions.GetProduct(OS_Name, ver, reg_ver);
                        
            
            //Repair
            QuickBooks.RepairOrUnInstallQB(installed_product, false, true);
        }

        [Fact]
        public void RunQBRemoveTest()
        {
            this.BDDfy();
        }
    }
}
