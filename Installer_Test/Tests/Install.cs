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
   
    public class Installer
    {
       // public TestStack.White.Application qbApp = null;
       // public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public static int Sync_Timeout = int.Parse(conf.get("SyncTimeOut"));
        public static string testName = "Install";
        public string targetPath;


        [Given(StepTitle = @"The parameters for installation are available at C:\Installation\Parameters.txt")]

        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
            string readpath = "C:\\Temp\\Parameters.xlsm"; // "C:\\Installation\\Sample.txt";

            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic = Lib.File_Functions.ReadExcelValues(readpath, "Install", "B2:B27");
            targetPath = dic["B12"];
        }

        [Then(StepTitle = "Then - Invoke QuickBooks installer")]
        public void InvokeQB()
        {
           OSOperations.InvokeInstaller(targetPath, "setup.exe");
        }


        [AndThen(StepTitle = "Then - Install QuickBooks")]
        public void RunInstallQB()
        {
            Install_Functions.Install_UK();
        }

       [Fact]
       [Category("P1")]
        public void RunQBInstallTest()
        {
            this.BDDfy();
        }
    }
}
