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
using System.Linq;
using Xunit;
using TestStack.BDDfy;
using FrameworkLibraries.AppLibs.QBDT.WhiteAPI;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using System.Management;
using Excel = Microsoft.Office.Interop.Excel;

namespace Installer_Test.Tests
{
    public class PostInstall
    {
        public static Property conf = Property.GetPropertyInstance();
        public static int Sync_Timeout = int.Parse(conf.get("SyncTimeOut"));
        public static string testName = "PostInstall";
        public string ver, reg_ver, expected_ver, installed_version, installed_product, installed_dataPath, installed_commonPath, dataPath;
         

        string OS_Name = string.Empty;
        string userName;

        [Given(StepTitle = @"QuickBooks is installed on the machine")]

        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
 
            string readpath = "C:\\Temp\\Parameters.xlsx"; // "C:\\Installation\\Sample.txt";
            
            Dictionary<string, string> dic = new Dictionary<string, string>();
           
            dic = Installer_Test.Lib.File_Functions.ReadExcelValues(readpath,"Path","B2:B10");

            ver = dic["B7"];
            reg_ver = dic["B8"];
            expected_ver = dic["B10"];

            OS_Name = Installer_Test.Lib.File_Functions.GetOS();
            installed_version = Installer_Test.Lib.File_Functions.GetQBVersion(OS_Name,ver,reg_ver);
            installed_product = Installer_Test.Lib.File_Functions.GetProduct(OS_Name, ver, reg_ver);
            installed_dataPath = Installer_Test.Lib.File_Functions.GetDataPath(OS_Name, ver, reg_ver);
            installed_commonPath = Installer_Test.Lib.File_Functions.GetCommonFilesPath(OS_Name, ver, reg_ver);
            dataPath = Installer_Test.Lib.File_Functions.GetDataPathKey(OS_Name, ver, reg_ver);
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
