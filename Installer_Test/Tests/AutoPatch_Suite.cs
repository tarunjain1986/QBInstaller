using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Reflection;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Imaging;
using System.Collections.Generic;
using System.Runtime.InteropServices;
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
   
    public class AutoPatch_Suite
    {

        [DllImport("User32.dll")]
        public static extern int SetForegroundWindow(IntPtr point);
        [DllImport("User32.dll")]
        private static extern IntPtr GetForegroundWindow();

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
        //string OS_Name = string.Empty;
        //Dictionary<string, string> dic_InvokeQB = new Dictionary<string, string>();

        /// <summary>
        /// Check F2
        /// </summary>
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public TestStack.White.UIItems.WindowItems.Window appWizWindow = null;
        public string exe;

        /// <summary>
        /// Create Company File
        /// </summary>
        Dictionary<String, String> keyvaluepairdic;

        /// <summary>
        /// AutoPatch
        /// </summary>
        public static string installed_datapath, service_source, service_dest, service_backup;
        Dictionary<string, string> dic_AP = new Dictionary<string, string>();

        /// <summary>
        /// Repair / Uninstall
        /// </summary>

        public static string installed_dir, installed_path, installed_product, ver, reg_ver;
        // Dictionary<string, string> dic_Repair = new Dictionary<string, string>();

        [Given(StepTitle = @"The parameters for installation are available at C:\Installer\Parameters.xlsm")]

        public void Setup()
        {
            //string testName = "AutoPatch";
            //Logger log = new Logger(testName + "_" + DateTime.Now.ToString("yyyyMMdd"));

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Read an initiliaze variables used in all scripts in this function (Setup ())
            //////////////////////////////////////////////////////////////////////////////////////////////////
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
             
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Invoke Installer and Install QB
            ///////////////////////////////////////////////////////////////////////////////////////////////////
         
            //dic = Lib.File_Functions.ReadExcelValues(readpath, "Install", "B2:B27");
            //country = dic["B5"];
            //targetPath = dic["B12"];
            //SKU = dic["B7"];
            //targetPath = targetPath + @"QBooks\";

            dic = Lib.File_Functions.ReadExcelValues(readpath, "Install", "B7:B30");
            country = dic["B10"];
            targetPath = dic["B30"];
            SKU = dic["B12"];
            targetPath = targetPath + @"QBooks\";

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

            // dic_Repair = File_Functions.ReadExcelValues(readpath, "PostInstall", "B2:B4");
            //ver = dic_Repair["B2"];
            //reg_ver = dic_Repair["B3"];

            ver = dic["B8"];
            reg_ver = File_Functions.GetRegVer(SKU);

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // AutoPatch
            ///////////////////////////////////////////////////////////////////////////////////////////////////

            dic_AP = File_Functions.ReadExcelValues(readpath, "AutoPatch", "B2");
            service_source = dic_AP["B2"];
            // installed_datapath = File_Functions.GetDataPath(ver, reg_ver);

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
                 // resultsPath = Path on local machine where Logs and Screenshots will be stored
                    resultsPath = Install_Functions.Install_US(); 
                  
                    break;

                case "UK":
                    Install_Functions.Install_UK();
                    break;

                case "CA":
                    Install_Functions.Install_CA();
                    break;
            }

            conf.reload(); // Reload the property file
            exe = conf.get("QBExePath");
            
            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            
            //////////////////////////////////////////////////////////////////////////////////////////////////////
            Boolean flag = false;

            flag = Actions.CheckDesktopWindowExists("QuickBooks Update Service");
            if (flag == true)
            {
                Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks Update Service"));
                SendKeys.SendWait("%l");
                Logger.logMessage("QuickBooks Update Service Window found.");
            }

            flag = false;

            while (flag == false)
            {
                flag = Actions.CheckDesktopWindowExists("QuickBooks " + SKU);

            }
            //////////////////////////////////////////////////////////////////////////////////////////////////////

            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
            qbWindow.WaitWhileBusy();

            // Maximize QuickBooks before continuing
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.MaximizeQB(qbApp);

            // Close QuickBook pop-up windows
            Install_Functions.CheckWindowsAndClose(SKU);

            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks " + SKU);

            // Save the window title in the Automation.Properties file
            // This value will be used in Repair / Uninstall
            Install_Functions.Get_QuickBooks_Edition(qbApp, qbWindow);
        }

        [AndThen(StepTitle = "Then - Perform PostInstall Tests")]
        public void Test_PostInstall()
        {
            Install_Functions.Post_Install();
        }

        [AndThen(StepTitle = "Then - Open F2")]
        public void CheckF2value()
        {
            // Close QuickBook pop-up windows
            Install_Functions.CheckWindowsAndClose(SKU);

            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks " + SKU);
            qbWindow.WaitWhileBusy();

            // Maximize QuickBooks window and set it active
            if (!qbWindow.IsCurrentlyActive)
            {
                qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.MaximizeQB(qbApp);
                Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks " + SKU));
            }
            
            PostInstall_Functions.CheckF2value(qbApp, qbWindow, resultsPath, SKU);
        }

        [AndThen(StepTitle = "Then - Click on Help -> About")]
        public void HelpAbout()
        {
            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks " + SKU);

            // Close QuickBook pop-up windows
            Install_Functions.CheckWindowsAndClose(SKU);

            Help.ClickHelpAbout(qbApp, qbWindow, resultsPath);
        }

        [AndThen(StepTitle = "Then - Create Company File")]
        public void CreateCompanyFile()
        {
            PostInstall_Functions.CreateCompanyFile(keyvaluepairdic);
        }

        [AndThen(StepTitle = "Then - Perform AutoPatch")]
        public void AutoPatch()
        {
            
             // Close QuickBooks
             CloseQB();
            
            // Kill QuickBooks processes
            OSOperations.KillProcess("QBW32");
            OSOperations.KillProcess("qbupdate");

            Thread.Sleep(1000);
            installed_datapath = File_Functions.GetDataPath(ver, reg_ver);
            if (installed_datapath != "")
            {
                // Replace the local copy of the serviceguide.xml with the server copy
                service_dest = installed_datapath + @"Components\QBUpdate\serviceguide.xml";
                service_source = service_source + "serviceguide.xml";
                service_backup = installed_datapath + @"Components\QBUpdate\serviceguide_backup.xml";

                if (File.Exists(service_backup))
                {
                    File.SetAttributes(service_backup, FileAttributes.Normal);
                    File.Delete(service_backup);
                    Logger.logMessage("Earlier version of service_backup.xml file deleted.");
                }

                System.IO.File.Move(service_dest, service_backup);
                File.Copy(service_source, service_dest);
                Logger.logMessage("service_backup.xml file copied from " + service_source + " to " + service_dest);
            }

            conf.reload(); // Reload the property file
            exe = conf.get("QBExePath");

            // Launch QuickBooks
            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            qbApp.WaitWhileBusy();

            //////////////////////////////////////////////////////////////////////////////////////////////////////
            Boolean flag = false;

            flag = Actions.CheckDesktopWindowExists("QuickBooks Update Service");
            if (flag == true)
            {
                Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks Update Service"));
                SendKeys.SendWait("%l");
                Logger.logMessage("QuickBooks Update Service Window found.");
            }

            flag = false;

            while (flag == false)
            {
                flag = Actions.CheckDesktopWindowExists("QuickBooks " + SKU);

            }
            //////////////////////////////////////////////////////////////////////////////////////////////////////

            // Set focus on the QuickBooks window
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.MaximizeQB(qbApp);
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);

            // Close QuickBook pop-up windows
            Install_Functions.CheckWindowsAndClose(SKU);

            Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
            // Actions.SelectMenu(qbApp, qbWindow, "Help", "Update QuickBooks...");

            Help.ClickHelpUpdate_AutoPatch(qbApp, qbWindow, resultsPath);

        }

        [AndThen(StepTitle = "Then - Perform Money In Money Out")]
        public void Perform_MIMO()
        {
            // QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks " + SKU);

            // Close QuickBook pop-up windows
            Install_Functions.CheckWindowsAndClose(SKU);

            Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
            PostInstall_Functions.PerformMIMO(qbApp, qbWindow);
        }

        [AndThen(StepTitle = "Then - Perform Verify")]
        public void Perform_Verify()
        {
            // QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");

            // Close QuickBook pop-up windows
            Install_Functions.CheckWindowsAndClose(SKU);

            Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
            PostInstall_Functions.PerformVerify(qbApp, qbWindow);
        }

        [AndThen(StepTitle = "Then - Perform Rebuild")]
        public void Perform_Rebuild()
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
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks " + SKU);
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


            Installer_Test.Lib.ScreenCapture sc = new Installer_Test.Lib.ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();

            // Invoking QuickBooks after deleting the dlls gives an Error message
            flag = Actions.CheckDesktopWindowExists("Error");
            pointer = GetForegroundWindow();
            sc.CaptureWindowToFile(pointer, resultsPath + "Error_before_Repair.png", ImageFormat.Png);

            if (flag == true)
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("Error"), "OK");
            }

            // Get the QuickBooks Edition to Repair from the Automation.Properties file
            conf.reload();
            installed_product = conf.get("Edition");

            //Repair
            QuickBooks.RepairOrUnInstallQB(installed_product, true, false);

            conf.reload(); // Reload the property file
            exe = conf.get("QBExePath");

            // Invoke QB after Repair 
            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);

            //////////////////////////////////////////////////////////////////////////////////////////////////////
            flag = false;

            flag = Actions.CheckDesktopWindowExists("QuickBooks Update Service");
            if (flag == true)
            {
                Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks Update Service"));
                SendKeys.SendWait("%l");
                Logger.logMessage("QuickBooks Update Service Window found.");
            }

            flag = false;

            while (flag == false)
            {
                flag = Actions.CheckDesktopWindowExists("QuickBooks " + SKU);

            }
            //////////////////////////////////////////////////////////////////////////////////////////////////////

            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
            qbWindow.WaitWhileBusy();

            // Maximize QuickBooks window and set it active
            if (!qbWindow.IsCurrentlyActive)
            {
                qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.MaximizeQB(qbApp);
            }

            pointer = GetForegroundWindow();
            sc.CaptureWindowToFile(pointer, resultsPath + "QuickBooks_launched_after_Repair.png", ImageFormat.Png);

            // Close QuickBook pop-up windows
            Install_Functions.CheckWindowsAndClose(SKU);
            CloseQB();

        }

        [AndThen(StepTitle = "Uninstall QuickBooks")]
        public void UninstallQB()
        {
            // Kill any existing QuickBooks process before uninstalling
            OSOperations.KillProcess("QBW32");
            Thread.Sleep(1000);

            // Get the QuickBooks Edition to Repair from the Automation.Properties file
            conf.reload();
            installed_product = conf.get("Edition");
            QuickBooks.RepairOrUnInstallQB(installed_product, false, true);

            // Delete existing traces of QuickBooks
            Install_Functions.CleanUp();
        }
        
       [Fact]
       [Category("AutoPatch_Suite")]
        public void RunAPSuite()
        {
            this.BDDfy();
        }
    }
}
