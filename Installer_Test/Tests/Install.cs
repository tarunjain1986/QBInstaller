﻿using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Reflection;
using System.Diagnostics;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

using FrameworkLibraries;
using FrameworkLibraries.Utils;
using FrameworkLibraries.ActionLibs;
using FrameworkLibraries.AppLibs.QBDT;

using TestStack.BDDfy;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.WindowItems;

using Xunit;

using Installer_Test;
using Installer_Test.Lib;


namespace Installer_Test.Tests
{
   
    public class Installer
    {
       // public TestStack.White.Application qbApp = null;
       // public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public static int Sync_Timeout = int.Parse(conf.get("SyncTimeOut"));
        public static string testName = "Install";
        public string targetPath, installPath, fileName, wkflow, customOpt, License_No, Product_No, UserID, Passwd, firstName, lastName;
        string [] LicenseNo, ProductNo;


        [Given(StepTitle = @"The parameters for installation are available at C:\Installation\Parameters.txt")]

        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
                
            string readpath = "C:\\Temp\\Parameters.xlsm"; // "C:\\Installation\\Sample.txt";

            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic = File_Functions.ReadExcelValues(readpath, "Path", "B2:B27");
                     
            targetPath = dic["B11"];
            targetPath = targetPath + @"QBooks\";
           
            customOpt = dic["B16"];
            wkflow = dic["B17"];
            License_No = dic["B18"];
            Product_No = dic["B19"];
            UserID = dic["B20"];
            Passwd = dic["B21"];
            firstName = dic["B22"];
            lastName = dic["B23"];

            installPath = dic["B26"];

            var regex = new Regex(@".{4}");
            string temp = regex.Replace(License_No, "$&" + "\n");
            LicenseNo = temp.Split('\n');

            regex = new Regex(@".{3}");
            temp = regex.Replace(Product_No, "$&" + "\n");
            ProductNo = temp.Split('\n');
            
        }

        [Then(StepTitle = "Then - Invoke QuickBooks installer")]
        public void InvokeQB()
        {
           OSOperations.InvokeInstaller(targetPath, "setup.exe");
        }


        [AndThen(StepTitle = "Then - Install QuickBooks")]
        public void RunInstallQB()
        {
            Install_Functions.Install_QB(targetPath, wkflow, customOpt, LicenseNo, ProductNo, UserID, Passwd, firstName, lastName, installPath);
        }

       [Fact]
        public void RunQBInstallTest()
        {
            this.BDDfy();
        }
    }
}
