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
