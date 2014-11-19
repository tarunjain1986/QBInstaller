using System;
using FrameworkLibraries.Utils;
using System.Windows.Automation;
using System.Windows.Forms;
using FrameworkLibraries.ActionLibs;
using TestStack.White.UIItems.WindowItems;
using System.Threading;
using TestStack.White.UIItems.Finders;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using FrameworkLibraries;
using System.Collections.Generic;
using TestStack.White.UIItems;
using Xunit;
using TestStack.BDDfy;
using FrameworkLibraries.AppLibs.QBDT;
using System.IO;
using System.Reflection;
using Installer_Test.Lib;

namespace Installer_Test.Tests
{
    
    public class VerifyRebuild
    {
       public string testName = "VerifyRebuild";
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
            QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
        }

         [Then(StepTitle = "Then - Perform Verify")]  

        public void PerformVerfiy()
        {

            PostInstall_Functions.PerformVerify(qbApp, qbWindow);

        }

         [AndThen(StepTitle = "Then - Perform Rebuild")]

         public void PerformRebuild()
         {

             PostInstall_Functions.PerformRebuild(qbApp, qbWindow);

         }
      [Fact]
        public void Run_verifyBuild()
        {
            this.BDDfy();
        }
        
        
    }
}