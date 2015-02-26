using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Reflection;
using System.Diagnostics;
using System.Windows.Forms;
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
    // Class for QB Install,Repair & Uninstall Workflows
    public class Installer_Suite_Combined
    {


        //-------------------To enable screen capture functionality-----------------------
        [DllImport("User32.dll")]
        public static extern int SetForegroundWindow(IntPtr point);
        [DllImport("User32.dll")]
        private static extern IntPtr GetForegroundWindow();

        //-------------------Variable declerations for Install QB----------------------
        // Variable for reading from Automation.Property file, the value of varibale SyncTimeOut
        public static Property conf = Property.GetPropertyInstance();
        public static int Sync_Timeout = int.Parse(conf.get("SyncTimeOut"));
        //Variable storing the Input data file.
        public string readpath = "C:\\Temp\\Parameters.xlsm";
        public static string resultsPath, logFilePath, customOpt, wkflow;


        public static string testName = "Install QuickBooks";
        public string country, targetPath, SKU;
        //Dictionary object variable to read data from "Install" Sheet of Input file.
        Dictionary<string, string> dic = new Dictionary<string, string>();


        //-----Variable declerations for storing QB App and QB window variables and "Check F2" ----------------------
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public string exe;


        //----------------Variable declerations for "Create Company File" ----------------------
        Dictionary<String, String> keyvaluepairdic;


        //-----------------Variable declerations for "Repair / Uninstall" --------------------------------------------
        public static string installed_dir, installed_path, installed_product, ver, reg_ver;


        //-----------------Dictionary Object for Reading data from "Install Execution Flow" sheet.
        Dictionary<String, string> dictionaryExecutionFlow = new Dictionary<string, string>();
        //Variable to read the status of Exection flow.
        private string executionRequired = null;


        [Given(StepTitle = @"The parameters for installation are available at C:\Installer\Parameters.xlsm")]

        public void Setup()
        {


            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            //Reading content of Install Execution Flow sheet from Input File
            dictionaryExecutionFlow = Lib.File_Functions.ReadExcelCellValues(readpath, "Install Execution Flow");


            // Reading Input data from excel for Invoking Installer and Installing QB
            dic = Lib.File_Functions.ReadExcelCellValues(readpath, "Install");
            country = dic["Select Country:"];
            targetPath = dic["Build Location (Local):"];
            targetPath = targetPath + @"QBooks\";
            SKU = dic["Select SKU:"];

            //Reading values from Input Data excel to initialize logger
            customOpt = dic["Select Custom and Network Options:"];
            wkflow = dic["Select WorkFlow:"];
            resultsPath = @"C:\Temp\Results\Install_" + customOpt + "_" + wkflow + "_" + DateTime.Now.ToString("yyyyMMddmm") + @"\Screenshots\";
            logFilePath = @"C:\Temp\Results\Install_" + customOpt + "_" + wkflow + "_" + DateTime.Now.ToString("yyyyMMddmm") + @"\Logs\";
            Install_Functions.Add_Log_Automation_Properties(logFilePath);
            conf.reload();

            Logger log = new Logger(testName + "_" + DateTime.Now.ToString("yyyyMMddmm"));// + timeStamp);
            // Create a folder to save the Screenshots
            Install_Functions.Create_Dir(resultsPath);

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

            ver = dic["Select Version:"];
            reg_ver = File_Functions.GetRegVer(SKU); 
        }


        [Then(StepTitle = "Then - Invoke QuickBooks installer")]
        public void InvokeQB()
        {
            //Read Execution flow data from "Install Execution Flow" sheet
            dictionaryExecutionFlow.TryGetValue("InvokeQB", out executionRequired);
            if (executionRequired == "1")
            {
                OSOperations.InvokeInstaller(targetPath, "setup.exe");
            }

        }

        [AndThen(StepTitle = "Then - Install QuickBooks")]
        public void RunInstallQB()
        {

            //Read Execution flow data from "Install Execution Flow" sheet
            dictionaryExecutionFlow.TryGetValue("RunInstallQB", out executionRequired);
            if (executionRequired == "1")
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
            }
        }

        [AndThen(StepTitle = "Set Up after Install and  initialization of global variables.")]
        public void InitializeVariables()
        {
            Boolean Flag = false;
            int loopCounter = 0;

            // Reload the Automation.property file
            conf.reload();
            exe = conf.get("QBExePath");

            // Initializing variables for QB application and main QB window.
            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            qbApp.WaitWhileBusy();


            // Check for QuickBooks Update Service Window
            if (Actions.CheckDesktopWindowExists("QuickBooks Update Service")) ;
            {
                Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks Update Service"));
                SendKeys.SendWait("%l");
                Logger.logMessage("QuickBooks Update Service Window found.");
            }


            // Wait for Window to appear for Ceratin iterations and then Break out
            while (Flag == false || loopCounter < 20)
            {
                Flag = Actions.CheckDesktopWindowExists("QuickBooks " + SKU);
                Thread.Sleep(1000);
                loopCounter += 1;
            }

            // Get QB window and then Mazimize it.
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.MaximizeQB(qbApp);
            qbWindow.WaitWhileBusy();

            // Check for multiple QB windows QuickBook pop-up windows and eloading QB app and QB window variables.
            Install_Functions.CheckWindowsAndClose(SKU);
            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks " + SKU);

            // Save the window title in the Automation.Properties file
            Install_Functions.Get_QuickBooks_Edition(qbApp, qbWindow);
            conf.reload();


        }


        [AndThen(StepTitle = "Then - Perform PostInstall Tests")]
        public void Test_PostInstall()
        {
            //Read Execution flow data from "Install Execution Flow" sheet
            dictionaryExecutionFlow.TryGetValue("Test_PostInstall", out executionRequired);
            if (executionRequired == "1")
            {
                Install_Functions.Post_Install();
            }

        }


        [AndThen(StepTitle = "Then - Open F2")]
        public void CheckF2value()
        {
            //Read Execution flow data from "Install Execution Flow" sheet
            dictionaryExecutionFlow.TryGetValue("CheckF2value", out executionRequired);
            if (executionRequired == "1")
            {
                // Check for multiple QB windows QuickBook pop-up windows and eloading QB app and QB window variables.
                Install_Functions.CheckWindowsAndClose(SKU);
                qbApp = QuickBooks.GetApp("QuickBooks");
                qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks " + SKU);
                PostInstall_Functions.CheckF2value(qbApp, qbWindow, resultsPath, SKU);
            }
        }

        [AndThen(StepTitle = "Then - Click on Help -> About")]
        public void HelpAbout()
        {
            //Read Execution flow data from "Install Execution Flow" sheet
            dictionaryExecutionFlow.TryGetValue("HelpAbout", out executionRequired);
            if (executionRequired == "1")
            {
                // Check for multiple QB windows QuickBook pop-up windows and eloading QB app and QB window variables.
                Install_Functions.CheckWindowsAndClose(SKU);
                qbApp = QuickBooks.GetApp("QuickBooks");
                qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks " + SKU);
                Help.ClickHelpAbout(qbApp, qbWindow, resultsPath);
            }
        }


        [AndThen(StepTitle = "Then - Create Company File")]
        public void CreateCompanyFile()
        {
            //Read Execution flow data from "Install Execution Flow" sheet
            dictionaryExecutionFlow.TryGetValue("CreateCompanyFile", out executionRequired);
            if (executionRequired == "1")
            {
                // Check for multiple QB windows QuickBook pop-up windows and eloading QB app and QB window variables.
                Install_Functions.CheckWindowsAndClose(SKU);
                qbApp = QuickBooks.GetApp("QuickBooks");
                qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks " + SKU);
                PostInstall_Functions.CreateCompanyFile(keyvaluepairdic);
            }
        }

        [AndThen(StepTitle = "Then - Perform Money In Money Out")]

        public void PerformMIMO()
        {
            //Read Execution flow data from "Install Execution Flow" sheet
            dictionaryExecutionFlow.TryGetValue("PerformMIMO", out executionRequired);
            if (executionRequired == "1")
            {
                // Check for multiple QB windows QuickBook pop-up windows and eloading QB app and QB window variables.
                Install_Functions.CheckWindowsAndClose(SKU);
                qbApp = QuickBooks.GetApp("QuickBooks");
                qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks " + SKU);
                Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
                PostInstall_Functions.PerformMIMO(qbApp, qbWindow);
            }
        }

        [AndThen(StepTitle = "Then - Perform Verify")]

        public void PerformVerfiy()
        {
            //Read Execution flow data from "Install Execution Flow" sheet
            dictionaryExecutionFlow.TryGetValue("PerformVerfiy", out executionRequired);
            if (executionRequired == "1")
            {
                // Check for multiple QB windows QuickBook pop-up windows and eloading QB app and QB window variables.
                Install_Functions.CheckWindowsAndClose(SKU);
                qbApp = QuickBooks.GetApp("QuickBooks");
                qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks " + SKU);
                Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
                PostInstall_Functions.PerformVerify(qbApp, qbWindow);

            }
        }

        [AndThen(StepTitle = "Then - Perform Rebuild")]

        public void PerformRebuild()
        {
            //Read Execution flow data from "Install Execution Flow" sheet
            dictionaryExecutionFlow.TryGetValue("PerformRebuild", out executionRequired);
            if (executionRequired == "1")
            {
                // Check for multiple QB windows QuickBook pop-up windows and eloading QB app and QB window variables.
                Install_Functions.CheckWindowsAndClose(SKU);
                qbApp = QuickBooks.GetApp("QuickBooks");
                qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks " + SKU);
                Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
                PostInstall_Functions.PerformRebuild(qbApp, qbWindow);
            }
        }

        [AndThen(StepTitle = "Then - Perform Switch OR Toggle")]
        public void SwitchEdition_Enterprise()
        {
            //Read Execution flow data from "Install Execution Flow" sheet
            dictionaryExecutionFlow.TryGetValue("SwitchEdition_Enterprise", out executionRequired);
            if (executionRequired == "1")
            {

                switch (SKU)
                {
                    case "Enterprise":
                        if (country == "US" | country == "CA")
                        {
                            SwitchToggle.SwitchEdition("Enterprise");
                            // PostInstall_Functions.SwitchEdition(qbApp, dic_Switch_Enterprise, exe, Bizname, SearchText);
                        }
                        break;

                    case "Premier":
                        SwitchToggle.SwitchEdition("Premier");
                        // PostInstall_Functions.SwitchEdition(qbApp, dic_Switch_Premier, exe, Bizname, SearchText);
                        break;

                    case "Premier Plus":
                        if (country == "US")
                        {
                            SwitchToggle.SwitchEdition("Premier");
                            // PostInstall_Functions.SwitchEdition(qbApp, dic_Switch_Premier, exe, Bizname, SearchText);
                        }
                        break;

                    case "Enterprise Accountant":
                        if (country == "US" | country == "CA")
                        {
                            SwitchToggle.ToggleEdition("Enterprise");
                            // PostInstall_Functions.ToggleEdition(qbApp, dic_Toggle_Enterprise, exe, Bizname);
                        }
                        break;

                    case "Premier Accountant":
                        SwitchToggle.ToggleEdition("Premier");
                        // PostInstall_Functions.ToggleEdition(qbApp, dic_Toggle_Premier, exe, Bizname);
                        break;
                }

            }
        }

        [AndThen(StepTitle = "Exit QuickBooks")]
        public void CloseQB()
        {
            //Read Execution flow data from "Install Execution Flow" sheet
            dictionaryExecutionFlow.TryGetValue("CloseQB", out executionRequired);
            if (executionRequired == "1")
            {
                qbApp = QuickBooks.GetApp("QuickBooks");
                qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks " + SKU);
                Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
                Properties.Lib.QB_functions.CloseQBApplication(qbApp, qbWindow);

            }
        }

        [AndThen(StepTitle = "Repair QuickBooks")]
        public void RepairQB()
        {
            Boolean flag = false;
            int loopCounter = 0;
            //Read Execution flow data from "Install Execution Flow" sheet
            dictionaryExecutionFlow.TryGetValue("RepairQB", out executionRequired);
            if (executionRequired == "1")
            {

                // Kill any existing QuickBooks process
                OSOperations.KillProcess("QBW32");
                Thread.Sleep(1000);
               
                // Delete DLLs
                installed_path = File_Functions.GetPath(ver, reg_ver);
                installed_dir = Path.GetDirectoryName(installed_path); // Get the path (without the exe name)
                Install_Functions.Delete_QBDLLs(installed_dir);

                // Invoke QuickBooks after deleting the dlls
                Process proc = new Process();
                proc.StartInfo.FileName = installed_path;
                proc.Start();
                Thread.Sleep(1000);

                //---Taking screen shot --------------
                Installer_Test.Lib.ScreenCapture sc = new Installer_Test.Lib.ScreenCapture();
                System.Drawing.Image img = sc.CaptureScreen();
                IntPtr pointer = GetForegroundWindow();

                // QuickBooks after deleting the dlls gives an Error message
                //flag = Actions.CheckDesktopWindowExists("Error");
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "Error_before_Repair.png", ImageFormat.Png);

                if (Actions.CheckDesktopWindowExists("Error"))
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

                // Check for QuickBooks Update Service Window
                if (Actions.CheckDesktopWindowExists("QuickBooks Update Service")) ;
                {
                    Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks Update Service"));
                    SendKeys.SendWait("%l");
                    Logger.logMessage("QuickBooks Update Service Window found.");
                }

                // Wait for Window to appear for Ceratin iterations and then Break out
                while (flag == false || loopCounter < 20)
                {
                    flag = Actions.CheckDesktopWindowExists("QuickBooks " + SKU);
                    Thread.Sleep(1000);
                    loopCounter += 1;
                }

                // Get QB window and then Mazimize it.
                qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
                qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.MaximizeQB(qbApp);
                qbWindow.WaitWhileBusy();

                Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks " + SKU));
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "QuickBooks_launched_after_Repair.png", ImageFormat.Png);

                // Close QuickBook pop-up windows
                Install_Functions.CheckWindowsAndClose(SKU);
                Thread.Sleep(1000);
                CloseQB();
            }
        }

        // get all code from pooja and put it in a function.
        [AndThen(StepTitle = "Uninstall QuickBooks")]
        public void UninstallQB()
        {
            //Read Execution flow data from "Install Execution Flow" sheet
            dictionaryExecutionFlow.TryGetValue("UninstallQB", out executionRequired);
            if (executionRequired == "1")
            {
                // Kill any existing QuickBooks process before uninstalling
                OSOperations.KillProcess("QBW32");
                Thread.Sleep(1000);

                if (Actions.CheckDesktopWindowExists("Programs and Features"))
                {
                    Actions.SetFocusOnWindow(Actions.GetDesktopWindow("Programs and Features"));
                    Actions.ClickElementByName(Actions.GetDesktopWindow("Programs and Features"), "Close");
                }

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
        }

        [Fact]
        [Category("Installer_Suite")]
        public void RunQBInstallSuite()
        {
            this.BDDfy();
        }

    }
}

