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
using Installer_Test.Lib;




namespace Installer_Test.Tests
{
   
    public class Installer_Suite
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
        /// 
        
        string OS_Name = string.Empty;
        Dictionary<string, string> dic_InvokeQB = new Dictionary<string, string>();

        /// <summary>
        /// Check F2
        /// </summary>
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public string exe = conf.get("QBExePath");
        // public string exe = conf.get("QBExePath");

        /// <summary>
        /// Create Company File
        /// </summary>
        Dictionary<String, String> keyvaluepairdic;

        /// <summary>
        /// Switch / Toggle
        /// </summary>
        public static string Bizname;
        String SearchText = "  - Intuit QuickBooks"; 
        Dictionary<String, String> dic_Switch_Enterprise;
        Dictionary<String, String> dic_Switch_Premier;
        Dictionary<String, String> dic_Toggle_Enterprise;
        Dictionary<String, String> dic_Toggle_Premier;

        /// <summary>
        /// Repair / Uninstall
        /// </summary>

        public static string installed_dir, installed_path, installed_product, ver, reg_ver;
        Dictionary<string, string> dic_Repair = new Dictionary<string, string>();

        [Given(StepTitle = @"The parameters for installation are available at C:\Installer\Parameters.xlsm")]

        public void Setup()
        {

            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
          //  Logger log = new Logger(testName + "_" + timeStamp);
      
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Invoke Installer and Install QB
            ///////////////////////////////////////////////////////////////////////////////////////////////////
         
            dic = Lib.File_Functions.ReadExcelValues(readpath, "Install", "B2:B27");
            country = dic["B5"];
            targetPath = dic["B12"];
            SKU = dic["B7"];
            targetPath = targetPath + @"QBooks\";

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Invoke QB
            ///////////////////////////////////////////////////////////////////////////////////////////////////

            List<string> listHeader = new List<string>();
            List<string> ListValue = new List<string>();
            dic_InvokeQB = new Dictionary<string, string>();
            File_Functions.ReadExcelSheet(readpath, "InvokeQB", 1, ref listHeader);
            File_Functions.ReadExcelSheet(readpath, "InvokeQB", 2, ref ListValue);
            dic_InvokeQB = listHeader.Zip(ListValue, (k, v) => new { k, v })
                 .ToDictionary(x => x.k, x => x.v);

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Check F2
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            //qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            //qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);

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
            // Switch / Toggle
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            Bizname = File_Functions.ReadExcelBizName(readpath);
            dic_Switch_Enterprise = File_Functions.ReadExcelCellValues(readpath, "Ent-Switch");
            dic_Switch_Premier = File_Functions.ReadExcelCellValues(readpath, "Pre-Switch");
            dic_Toggle_Enterprise = File_Functions.ReadExcelCellValues(readpath, "Ent-Toggle");
            dic_Toggle_Premier = File_Functions.ReadExcelCellValues(readpath, "Pre-Toggle");

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Repair / Uninstall
            ///////////////////////////////////////////////////////////////////////////////////////////////////

            dic_Repair = File_Functions.ReadExcelValues(readpath, "PostInstall", "B2:B4");
            ver = dic["B2"];
            reg_ver = dic["B3"];
 
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

            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.MaximizeQB(qbApp);
        }

        [AndThen(StepTitle = "Then - Open F2")]
        public void CheckF2value()
        {
            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
            QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
            PostInstall_Functions.CheckF2value(qbApp, qbWindow, resultsPath);
        }

        [AndThen(StepTitle = "Then - Click on Help -> About")]
        public void HelpAbout()
        {
           Help.ClickHelpAbout(qbApp, qbWindow, resultsPath);
        }


        [AndThen(StepTitle = "Then - Create Company File")]
        public void CreateCompanyFile()
        {
            PostInstall_Functions.CreateCompanyFile(keyvaluepairdic);
        }

        [AndThen(StepTitle = "Then - Perform Money In Money Out")]

        public void PerformMIMO()
        {
            QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
            PostInstall_Functions.PerformMIMO(qbApp, qbWindow);
        }

        [AndThen(StepTitle = "Then - Perform Verify")]

        public void PerformVerfiy()
        {
            QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
            PostInstall_Functions.PerformVerify(qbApp, qbWindow);
        }

        [AndThen(StepTitle = "Then - Perform Rebuild")]

        public void PerformRebuild()
        {
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
                        PostInstall_Functions.SwitchEdition(qbApp, dic_Switch_Enterprise, exe, Bizname, SearchText);
                    }
                    break;

                case "Premier":
                    PostInstall_Functions.SwitchEdition(qbApp, dic_Switch_Premier, exe, Bizname, SearchText);
                    break;

                case "Premier Plus":
                    if (country == "US")
                    {
                        PostInstall_Functions.SwitchEdition(qbApp, dic_Switch_Premier, exe, Bizname, SearchText);
                    }
                    break;

                case "Enterprise Accountant":
                    if (country == "US" | country == "CA")
                    {
                        PostInstall_Functions.ToggleEdition(qbApp, dic_Toggle_Enterprise, exe, Bizname);
                    }
                    break;

                case "Premier Accountant":
                    PostInstall_Functions.ToggleEdition(qbApp, dic_Toggle_Premier, exe, Bizname);
                    break;
            }
        }

        [AndThen(StepTitle = "Repair QuickBooks")]
        public void RepairQB()
        {
            OS_Name = File_Functions.GetOS();
            installed_product = File_Functions.GetProduct(ver, reg_ver);
            installed_path = File_Functions.GetPath(ver, reg_ver);
            installed_dir = Path.GetDirectoryName(installed_path); // Get the path (without the exe name)
            
            // Delete DLLs
            Install_Functions.Delete_QBDLLs(installed_dir);

            // Invoke QB
            QuickBooks.Initialize(installed_path);

            //Repair
            QuickBooks.RepairOrUnInstallQB(installed_product, true, false);

            // Invoke QB after Repair : To be completed
            // QuickBooks.Initialize(installed_path);
        }

        [AndThen(StepTitle = "Uninstall QuickBooks")]
        public void UninstallQB()
        {
            QuickBooks.RepairOrUnInstallQB(installed_product, false, true);
            Install_Functions.CleanUp();
        }
        
       [Fact]
       [Category("P1")]
        public void RunQBInstallTest()
        {
            this.BDDfy();
        }
    }
}
