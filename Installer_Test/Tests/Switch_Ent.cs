using System;
using System.Windows.Forms;
using Xunit;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.BDDfy;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Installer_Test.Lib;

namespace Installer_Test.Tests
{

    public class Switch_Ent
    {

        public string testName = "Switch_Ent";
        public static string Bizname;
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public static string exe = conf.get("QBExePath");
        Dictionary<String, String> dic = new Dictionary<string, string>();
        String SearchText = "  - Intuit QuickBooks"; 


        [Given(StepTitle = "Given - QuickBooks App and Window instances are available")]

        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
            //QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
            string readpath = "C:\\Temp\\Parameters.xlsm";
          Bizname = File_Functions.ReadExcelBizName(readpath);
          
            
            dic = File_Functions.ReadExcelCellValues(readpath, "Ent-Switch");
        }


        [Then(StepTitle = "Then - Switch Edition")]
        public void SwitchEdition()
        {
            //Actions.SelectMenu(qbApp, qbWindow, "File", "New Company...");
            PostInstall_Functions.SwitchEdition(qbApp, dic, exe,Bizname, SearchText);

        }
        [Fact]
        public void Run_Switch_Ent()
        {
            this.BDDfy();
        }
    }
}
