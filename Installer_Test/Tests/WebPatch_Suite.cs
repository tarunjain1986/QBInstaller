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
using FrameworkLibraries.ActionLibs.WhiteAPI;

using TestStack.BDDfy;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.WindowItems;

using Xunit;

using Installer_Test;
using Installer_Test.Lib;



namespace Installer_Test.Tests
{
   
    public class WebPatch_Suite
    {
       /// <summary>
       /// Install QB
       /// </summary>
        public static Property conf = Property.GetPropertyInstance();
        public static int Sync_Timeout = int.Parse(conf.get("SyncTimeOut"));
        public string readpath = "C:\\Temp\\Parameters.xlsm";
        public static string resultsPath;

        //public static Property conf = Property.GetPropertyInstance();
        //public static int Sync_Timeout = int.Parse(conf.get("SyncTimeOut"));


        public static string testName = "Install QuickBooks";
        public string country, targetPath, SKU;
        Dictionary<string, string> dic = new Dictionary<string, string>();
        

        /// <summary>
        /// Invoke QB
        /// </summary>
        string OS_Name = string.Empty;
        Dictionary<string, string> dic_InvokeQB = new Dictionary<string, string>();

        /// <summary>
        /// Check F2
        /// </summary>
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public TestStack.White.UIItems.WindowItems.Window appWizWindow = null;
        public string exe = conf.get("QBExePath");
        // public string exe = conf.get("QBExePath");


        Dictionary<string, string> dic_WebPatch = new Dictionary<string, string>();
        public string SKU_WebPatch, targetPath_WP;

        /// <summary>
        /// Create Company File
        /// </summary>
        Dictionary<String, String> keyvaluepairdic;

        /// <summary>
        /// Repair / Uninstall
        /// </summary>
        public static string installed_dir, installed_path, installed_product, ver, reg_ver;
        Dictionary<string, string> dic_Repair = new Dictionary<string, string>();

        [Given(StepTitle = @"The parameters for installation are available at C:\Installer\Parameters.xlsm")]

        public void Setup()
        {

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Read an initiliaze variables used in all scripts in this function (Setup ())
            //////////////////////////////////////////////////////////////////////////////////////////////////
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
                
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Invoke Installer and Install QB
            ///////////////////////////////////////////////////////////////////////////////////////////////////
         
            dic = Lib.File_Functions.ReadExcelValues(readpath, "Install", "B2:B27");
            country = dic["B5"];
            targetPath = dic["B12"];
            SKU = dic["B7"];
            targetPath = targetPath + @"QBooks\";


            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // WebPatch
            ///////////////////////////////////////////////////////////////////////////////////////////////////

            dic_WebPatch = Lib.File_Functions.ReadExcelValues(readpath, "WebPatch", "B2:B12");
            SKU_WebPatch = dic["B7"];
            targetPath_WP = dic["B11"];

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Create Company File
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            List<string> listHeader1 = new List<string>();
            List<string> ListValue1 = new List<string>();
            File_Functions.ReadExcelSheet(readpath, "CompanyFile", 1, ref listHeader1);
            File_Functions.ReadExcelSheet(readpath, "CompanyFile", 3, ref ListValue1);
            keyvaluepairdic = listHeader1.Zip(ListValue1, (k, v) => new { k, v })
                 .ToDictionary(x => x.k, x => x.v);

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Repair / Uninstall
            ///////////////////////////////////////////////////////////////////////////////////////////////////

            dic_Repair = File_Functions.ReadExcelValues(readpath, "PostInstall", "B2:B4");
            ver = dic_Repair["B2"];
            reg_ver = dic_Repair["B3"];
 
            ///////////////////////////////////////////////////////////////////////////////////////////////////
        }

        [Then(StepTitle = "Then - Invoke QuickBooks installer")]
        public void InvokeQB()
        {
            OSOperations.InvokeInstaller(targetPath, "setup.exe");
        }

        [AndThen(StepTitle = "Then - Install QuickBooks")]
        public void RunInstallQB()
        {
            switch (country)
            {
                case "US":
                    resultsPath = Install_Functions.Install_US();
                    break;

                case "UK":
                    Install_Functions.Install_UK();
                    break;

                case "CA":
                    Install_Functions.Install_CA();
                    break;
            }

            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);

            // Maximize QuickBooks before continuing
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.MaximizeQB(qbApp);

        }

        [AndThen(StepTitle = "Then - Perform PostInstall Tests")]
        public void Test_PostInstall()
        {
            Install_Functions.Post_Install();
        }

        [AndThen(StepTitle = "Then - Open F2")]
        public void CheckF2value()
        {
            // QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
            // Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");
            PostInstall_Functions.CheckF2value(qbApp, qbWindow, resultsPath);
        }

        [AndThen(StepTitle = "Then - Click on Help -> About")]
        public void HelpAbout()
        {
            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");
            Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
            Help.ClickHelpAbout(qbApp, qbWindow, resultsPath);
        }

        [AndThen(StepTitle = "Then - Perform WebPatch")]
        public void Web_Patch()
        {
            dic_WebPatch = Lib.File_Functions.ReadExcelValues(readpath, "WebPatch", "B2:B12");
            SKU_WebPatch = dic["B7"];
            targetPath_WP = dic["B11"];

            CloseQB();
            OSOperations.KillProcess("qbw32");

            if (SKU_WebPatch == "Enterprise" || SKU_WebPatch == "Enterprise Accountant")
                OSOperations.InvokeInstaller(targetPath_WP, "en_qbwebpatch.exe");
            else
                OSOperations.InvokeInstaller(targetPath_WP, "qbwebpatch.exe");
            
            WebPatch.ApplyWebPatch(resultsPath);
          
        }

        [AndThen(StepTitle = "Then - Create Company File")]
        public void CreateCompanyFile()
        {
            PostInstall_Functions.CreateCompanyFile(keyvaluepairdic);
        }

        [AndThen(StepTitle = "Then - Perform Money In Money Out")]
        public void PerformMIMO()
        {
            // QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");
            Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
            PostInstall_Functions.PerformMIMO(qbApp, qbWindow);
        }

        [AndThen(StepTitle = "Then - Perform Verify")]
        public void PerformVerify()
        {
            // QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");
            Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
            PostInstall_Functions.PerformVerify(qbApp, qbWindow);
        }

        [AndThen(StepTitle = "Then - Perform Rebuild")]
        public void PerformRebuild()
        {
            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");
            Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
            PostInstall_Functions.PerformRebuild(qbApp, qbWindow);
        }

        [AndThen(StepTitle = "Then - Perform Switch OR Toggle")]
        public void SwitchEdition_Enterprise()
        {
            switch (SKU)
            { 
                case "Enterprise":
                    if (country == "US" | country == "CA")
                    {
                        SwitchToggle.SwitchEdition("Enterprise");
                    }
                    break;

                case "Premier":
                    SwitchToggle.SwitchEdition("Premier");
                    break;

                case "Premier Plus":
                    if (country == "US")
                    {
                        SwitchToggle.SwitchEdition("Premier");
                    }
                    break;

                case "Enterprise Accountant":
                    if (country == "US" | country == "CA")
                    {
                        SwitchToggle.ToggleEdition("Enterprise");
                    }
                    break;

                case "Premier Accountant":
                    SwitchToggle.ToggleEdition("Premier");
                    break;
            }
        }

        [AndThen(StepTitle = "Close QuickBooks")]
        public void CloseQB ()
        {
            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");
            Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
            Actions.SelectMenu(qbApp, qbWindow, "File", "Exit");
        }

        [AndThen(StepTitle = "Repair QuickBooks")]
        public void RepairQB()
        {
            // OS_Name = File_Functions.GetOS();
            installed_product = File_Functions.GetProduct(ver, reg_ver);

            if (installed_product.Contains("QuickBooks Premier"))
            {
                installed_product = installed_product.Replace("QuickBooks Premier", "QuickBooks Premier Edition");
            }
            installed_path = File_Functions.GetPath(ver, reg_ver);
            installed_dir = Path.GetDirectoryName(installed_path); // Get the path (without the exe name)

            // Kill any existing QuickBooks process
            OSOperations.KillProcess("QBW32");
            Thread.Sleep(1000);
            
            // Delete DLLs
            Install_Functions.Delete_QBDLLs(installed_dir);

            // Invoke QuickBooks after deleting the dlls
            Process proc = new Process();
            proc.StartInfo.FileName = installed_path;
            proc.Start();

            Thread.Sleep(1000);

            Boolean flag;

            // Invoking QuickBooks after deleting the dlls gives an Error message
            flag = Actions.CheckDesktopWindowExists("Error");
            if (flag == true)
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("Error"), "OK");
            }
            Thread.Sleep(1000);


            //Repair
            QuickBooks.RepairOrUnInstallQB(installed_product, true, false);

            // Invoke QB after Repair : To be completed
            // QuickBooks.Initialize(installed_path);

            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);

            Thread.Sleep(20000);

            var MainWindow = Actions.GetDesktopWindow("QuickBooks");
            if (Actions.CheckWindowExists(MainWindow, "Register "))
            {
                Actions.ClickElementByName(Actions.GetChildWindow(MainWindow, "Register "), "Remind Me Later");
            }

            Thread.Sleep(1000);
            CloseQB();

        }

        [AndThen(StepTitle = "Uninstall QuickBooks")]
        public void UninstallQB()
        {
            // Kill any existing QuickBooks process before uninstalling
            OSOperations.KillProcess("QBW32");
            Thread.Sleep(1000);
            
            QuickBooks.RepairOrUnInstallQB(installed_product, false, true);

            // Delete existing traces of QuickBooks
            Install_Functions.CleanUp();
        }
        
       [Fact]
       [Category("Installer_Suite")]
        public void RunQBInstallSuite()
        {
            this.BDDfy();
        }
    }
}
