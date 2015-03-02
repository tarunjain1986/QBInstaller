using System;
using System.Windows.Forms;
using Xunit;
//using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.BDDfy;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using System.Threading;
using System.Diagnostics;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
//using Microsoft.Office.Interop.Excel;

using Installer_Test.Lib;


namespace Installer_Test.Archive_Tests
{

    public class TestSwitch
    {

        public string testName = "Switch_Ent";
        public static string Bizname;
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public static string exe = conf.get("QBExePath");
       // String SearchText = "  - Intuit QuickBooks";


        [Given(StepTitle = "Given - QuickBooks App and Window instances are available")]

        public void Setup()
        {

            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
           
        }


        [Then(StepTitle = "Then - Switch Edition")]
        public void Switch_Edition()
        {
            //PostInstall_Functions.SwitchEdition(qbApp, dic, exe, Bizname, SearchText);
            SwitchToggle.SwitchEdition("Enterprise");
        }

        [AndThen(StepTitle = "Close QuickBooks")]
        public void CloseQB()
        {
            string SKU = "Enterprise";
            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks " + SKU);
            Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
            Actions.SelectMenu(qbApp, qbWindow, "File", "Exit");
        }

        [AndThen(StepTitle = "Repair QuickBooks")]
        public void RepairQB()
        {
            // OS_Name = File_Functions.GetOS();
            //installed_product = File_Functions.GetProduct(ver, reg_ver);

            //if (installed_product.Contains("QuickBooks Premier"))
            //{
            //    installed_product = installed_product.Replace("QuickBooks Premier", "QuickBooks Premier Edition");
            //}


            // Kill any existing QuickBooks process
            OSOperations.KillProcess("QBW32");
            Thread.Sleep(1000);
            string ver = "25.0", reg_ver = "bel", installed_path, installed_dir, installed_product;
            // Delete DLLs
            installed_path = File_Functions.GetPath(ver, reg_ver);
            installed_dir = Path.GetDirectoryName(installed_path); // Get the path (without the exe name)
            Install_Functions.Delete_QBDLLs(installed_dir);

            // Invoke QuickBooks after deleting the dlls
            Process proc = new Process();
            proc.StartInfo.FileName = installed_path;
            proc.Start();

            Thread.Sleep(1000);

            Boolean flag;


            //Installer_Test.Lib.ScreenCapture sc = new Installer_Test.Lib.ScreenCapture();
            //System.Drawing.Image img = sc.CaptureScreen();
            //IntPtr pointer = GetForegroundWindow();

            // Invoking QuickBooks after deleting the dlls gives an Error message
            flag = Actions.CheckDesktopWindowExists("Error");
            //pointer = GetForegroundWindow();
            //sc.CaptureWindowToFile(pointer, resultsPath + "Error_before_Repair.png", ImageFormat.Png);

            if (flag == true)
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("Error"), "OK");
            }
            Thread.Sleep(1000);

            // Get the QuickBooks Edition to Repair from the Automation.Properties file
            conf.reload();
            installed_product = conf.get("Edition");

            //Repair
            QuickBooks.RepairOrUnInstallQB(installed_product, true, false);

            // Invoke QB after Repair 
            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);

            Thread.Sleep(20000);
            qbApp.WaitWhileBusy();
            string SKU = "Enterprise";
            Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks " + SKU));

            //pointer = GetForegroundWindow();
            //sc.CaptureWindowToFile(pointer, resultsPath + "QuickBooks_launched_after_Repair.png", ImageFormat.Png);

            // Close QuickBook pop-up windows
            Install_Functions.CheckWindowsAndClose(SKU);

            Thread.Sleep(1000);
            CloseQB();

        }

        [AndThen(StepTitle = "Uninstall QuickBooks")]
        public void UninstallQB()
        {
            // Kill any existing QuickBooks process before uninstalling
            OSOperations.KillProcess("QBW32");
            Thread.Sleep(1000);

            if (Actions.CheckDesktopWindowExists("Programs and Features"))
            {
                Actions.SetFocusOnWindow(Actions.GetDesktopWindow("Programs and Features"));
                Actions.ClickElementByName(Actions.GetDesktopWindow("Programs and Features"), "Close");
            }
            string installed_product;
            // Get the QuickBooks Edition to Repair from the Automation.Properties file
            conf.reload();
            installed_product = conf.get("Edition");
            QuickBooks.RepairOrUnInstallQB(installed_product, false, true);

            if (Actions.CheckDesktopWindowExists("Programs and Features"))
            {
                Actions.SetFocusOnWindow(Actions.GetDesktopWindow("Programs and Features"));
                Actions.ClickElementByName(Actions.GetDesktopWindow("Programs and Features"), "Close");
            }
        }
        [Fact]
        public void Run_Switch_Ent()
        {
            this.BDDfy();
        }
    }
}
