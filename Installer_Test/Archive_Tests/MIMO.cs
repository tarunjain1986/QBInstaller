using System;
//using Microsoft.VisualStudio.TestTools.UnitTesting;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.Utils;
using TestStack.BDDfy;
using Xunit;
using Installer_Test.Lib;


namespace Installer_Test.Archive_Tests
{

    public class MIMO
    {
        public string testName = "MIMO";
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public static string exe = conf.get("QBExePath");

        [Given(StepTitle = "Given - QuickBooks App and Window instances are available")]

        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
           //  QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
        }

         [Then(StepTitle = "Then - Perform MIMO")]  

        public void PerformMIMO()
        {

            PostInstall_Functions.PerformMIMO(qbApp, qbWindow);

        }
            
            
        [Fact]
        public void Run_MIMO()
        {
            this.BDDfy();
        }
    }
}



