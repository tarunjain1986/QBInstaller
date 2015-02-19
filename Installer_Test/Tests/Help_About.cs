using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Windows.Automation;

using FrameworkLibraries;
using FrameworkLibraries.Utils;
using FrameworkLibraries.ActionLibs;
using FrameworkLibraries.AppLibs.QBDT;

using TestStack.BDDfy;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.WindowItems;
using FrameworkLibraries.ActionLibs.WhiteAPI;

using Xunit;

using Installer_Test;

//using Microsoft.VisualStudio.TestTools.UnitTesting;


namespace Installer_Test.Tests
{
    public class Help_About
    {
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public string exe = conf.get("QBExePath");
        public Random rand = new Random();
        public string testName = "Help_About";        
        
        
        [Given(StepTitle = "Given - QuickBooks App and Window instances are available")]
        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
            //qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            //qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
            //QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
        }

        [Then(StepTitle = "Then - click on Help -> About")]
        public void HelpAbout()
        {
            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");

            // Close QuickBook pop-up windows
            // Install_Functions.CheckWindowsAndClose(SKU);
            Help.ClickHelpAbout(qbApp, qbWindow, @"C:\Temp\");
        }
        
        //[AndThen(StepTitle = "AndThen - Perform tear down activities to ensure that there are no on-screen exceptions")]
        //public void TearDown()
        //{
        //    QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
        //}

        [Fact]
        public void RunHelpAboutTest()
        {
            this.BDDfy();
        }
    }
}
