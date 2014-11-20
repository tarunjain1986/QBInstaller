using System;
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
using FrameworkLibraries.ActionLibs.WhiteAPI;
using Installer_Test.Lib;


namespace Installer_Test.Tests
{
   
    public class WebPatch
    {
       // public TestStack.White.Application qbApp = null;
       // public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public static string Sync_Timeout = conf.get("SyncTimeOut");
        public static string testName = "WebPatch";
        public string targetPath,patchpath, installPath, fileName, wkflow, customOpt, License_No, Product_No, UserID, Passwd, firstName, lastName;
        string [] LicenseNo, ProductNo;
        public string ver, reg_ver, installed_product;
        string OS_Name = string.Empty;
        string readpath = "C:\\Temp\\Parameters.xlsm";

        [Given(StepTitle = @"The parameters for installation are available at C:\Installation\Parameters.txt")]

        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
 
            
        }

        [Then(StepTitle = "Then - Get the Product Info from the excel")]

        public void Get_Product_info()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic = File_Functions.ReadExcelValues(readpath, "PostInstall", "B2:B10");
            ver = dic["B7"];
            reg_ver = dic["B8"];

            OS_Name = File_Functions.GetOS();
            installed_product = Installer_Test.Lib.File_Functions.GetProduct(OS_Name, ver, reg_ver);

        }    

        [AndThen(StepTitle = "Then - Check if the product is installed or not")]
        public void Check_Product_Installed()
        {
            try
            {
                FrameworkLibraries.Utils.OSOperations.CommandLineExecute("control appwiz.cpl");

                try
                {
                    Logger.logMessage("---------------Try-Catch Block------------------------");
                    Actions.WaitForWindow("Programs and Features", int.Parse(Sync_Timeout));
                }
                catch { }

                var controlPanelWindow = Actions.GetDesktopWindow("Programs and Features");
                var uiaWindow = Actions.UIA_GetAppWindow("Programs and Features");

                Actions.UIA_SetTextByName(uiaWindow, Actions.GetDesktopWindow("Programs and Features"), "Search Box", installed_product);

                try
                {
                    Logger.logMessage("---------------Try-Catch Block------------------------");
                    Actions.WaitForElementEnabled(Actions.GetDesktopWindow("Programs and Features"), installed_product, int.Parse(Sync_Timeout));
                }
                catch { }
            }
            catch { }
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

        [AndThen(StepTitle = "Then - Kill QuickBooks")]
        public void KillQB()
        {
            OSOperations.KillProcess("setup");

        }
        [Then(StepTitle = "Copy the web patch to local")]
        public void copyPatch()
        {
            Installer_Test.Install_Functions.Copy_WebPatch("BEL",patchpath);
        }

        [AndThen(StepTitle = "Then - Invoke Web Patch installer")]
        public void InvokeWP()
        {
            string targetPath = @"C:\Temp\WebPatch\";
            OSOperations.InvokeInstaller(targetPath, "en_qbwebpatch.exe");
            Logger.logMessage("Copied");
            Thread.Sleep(1000);
            Window patchWin= Actions.GetDesktopWindow("QuickBooks Update,Version");
            Thread.Sleep(1000);
            Actions.ClickElementByName(patchWin, "Install Now");
            Logger.logMessage("Installing webpatch");

        }
      


       [Fact]
        public void RunInstallWebPatch()
        {
            this.BDDfy();
        }
    }
}
