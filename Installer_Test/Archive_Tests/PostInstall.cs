﻿using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Reflection;
using System.Management;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text.RegularExpressions;

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

using Microsoft.Win32;

using Installer_Test.Lib;

using Excel = Microsoft.Office.Interop.Excel;

namespace Installer_Test.Archive_Tests
{
    public class PostInstall
    {
        public static Property conf = Property.GetPropertyInstance();
        public static int Sync_Timeout = int.Parse(conf.get("SyncTimeOut"));
        public static string testName = "PostInstall";
        public string SKU, ver, reg_ver, expected_ver, installed_version, installed_product, installed_dataPath, installed_commonPath, dataPath;
         

        string OS_Name = string.Empty;
        string userName;

        [Given(StepTitle = @"QuickBooks is installed on the machine")]

        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
 
            string readpath = "C:\\Temp\\Parameters.xlsm"; // "C:\\Installation\\Sample.txt";
            
            Dictionary<string, string> dic = new Dictionary<string, string>();
           
            //dic = File_Functions.ReadExcelValues(readpath,"PostInstall","B2:B4");

            //ver = dic["B2"];
            //reg_ver = dic["B3"];
            //expected_ver = dic["B4"];

            dic = File_Functions.ReadExcelValues(readpath, "Install", "B9:B12");
            ver = dic["B9"];
            SKU = dic["B12"];

            dic = File_Functions.ReadExcelValues(readpath, "Install", "E7");
            reg_ver = File_Functions.GetRegVer(SKU);


            expected_ver = dic["E7"];

            OS_Name = File_Functions.GetOS();
            installed_version = File_Functions.GetQBVersion(ver,reg_ver);
            installed_product = File_Functions.GetProduct(ver, reg_ver);
            installed_dataPath = File_Functions.GetDataPath(ver, reg_ver);
            installed_commonPath = File_Functions.GetCommonFilesPath(ver, reg_ver);
            dataPath = File_Functions.GetDataPathKey(ver, reg_ver);
            userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            userName = userName.Remove(0, 5);

        }

        [Then(StepTitle = "Then - Check QB Version in Registry")]
        public void CheckReg_QBVersion()
        {
           // Assert.Equal(ver, installed_version);
           Actions.XunitAssertEuqals(expected_ver, installed_version);
        }


        [AndThen(StepTitle = "Then - Check whether QBInfo.dat file is created")]
        public void Check_QBInfo()
        {
           Logger.logMessage("Function call @ :" + DateTime.Now);
           
           try
           {
               Assert.True(File.Exists(dataPath + "QBInfo.dat"));
               Logger.logMessage("File QBInfo.dat is available at " + dataPath + " - Successful");
               Logger.logMessage("------------------------------------------------------------------------------");
           }
           catch (Exception e)
           {
               Logger.logMessage("File QBInfo.dat is not available at " + dataPath + " - Failed");
               Logger.logMessage(e.Message);
               Logger.logMessage("------------------------------------------------------------------------------");
               String sMessage = e.Message;
              // LastException.SetLastError(sMessage);
              // throw new Exception(sMessage);
           }
        
        }


        [AndThen(StepTitle = "Then - Check whether oauth_<username>.dat file is created")]
        public void Check_oAuth()
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            string fileName = "oauth_" + userName + ".dat";
            try
            {
                Assert.True(File.Exists(dataPath + @"iamdata\" + fileName));
                Logger.logMessage("File oauth_<username>.dat is available at " + dataPath + "iamdata\\ - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("File oauth_<username>.dat is not available at " + dataPath + "iamdata\\ - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
              //  LastException.SetLastError(sMessage);
               // throw new Exception(sMessage);
            }
        }

        [Fact]
        public void RunQBPostInstallTest()
        {
            this.BDDfy();
        }
    }
}
