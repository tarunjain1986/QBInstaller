using System;
using FrameworkLibraries.Utils;
using FrameworkLibraries.ActionLibs.QBDT;
using TestStack.White.UIItems.WindowItems;
using System.Threading;
using TestStack.White.UIItems.Finders;
using FrameworkLibraries.ActionLibs.QBDT.WhiteAPI;
using FrameworkLibraries;
using System.Collections.Generic;
using TestStack.White.UIItems;
using Xunit;
using System.Linq;
using TestStack.BDDfy;
using FrameworkLibraries.AppLibs.QBDT.WhiteAPI;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using Microsoft.Win32;
using System.Management;


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

            string readpath = @"C:\Temp\Parameters.txt";
            File.WriteAllLines(readpath, File.ReadAllLines(readpath).Where(l => !string.IsNullOrWhiteSpace(l))); // Remove white space from the file

            string[] lines = File.ReadAllLines(readpath);
            var dic = lines.Select(line => line.Split('=')).ToDictionary(keyValue => keyValue[0], bits => bits[1]);

            ver = dic["Version"];
            reg_ver = dic["Registry Folder"];

            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Caption FROM Win32_OperatingSystem");
            foreach (ManagementObject os in searcher.Get())
            {
                OS_Name = os["Caption"].ToString();
                break;
            }

            if (OS_Name.Contains("Windows 7"))
            {

                RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\Wow6432Node\\Intuit\\QuickBooks\\" + ver + "\\" + reg_ver);
                if (key != null)
                {
                    product = key.GetValue("Product");
                    if (product != null)
                    {
                        installed_product = product as string;
                    }
                }
               
            }
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
