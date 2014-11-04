using System;
using System.IO;
using System.Threading;
using System.Reflection;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Collections.Generic;

using FrameworkLibraries;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.ActionLibs.WhiteAPI;

using TestStack.White.UIItems;
using TestStack.BDDfy;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.WindowItems;

using Xunit;

namespace BATS.Tests
{
    public class CreateCompany
    {
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public string exe = conf.get("QBExePath");
        public Random rand = new Random();
        public string testName = "CreateAndCloseCompany";

        [Given(StepTitle = "Given - QuickBooks App and Window instances are available")]
        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName+"_"+timeStamp);
            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
            QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
        }

        [Then(StepTitle = "Then - Create new company file should be successful")]
        public void CreateAndCloseCompany()
        {
            string businessName = "White" + rand.Next(1234, 8976);
            QuickBooks.CreateCompany(qbApp, qbWindow, businessName, "Information Technology");
            QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
            var winTitleOfNewCompany = qbWindow.Title;
            Actions.XunitAssertEuqals(winTitleOfNewCompany, qbWindow.Title);
        }

        [AndThen(StepTitle = "AndThen - Perform tear down activities to ensure that there are no on-screen exceptions")]
        public void TearDown()
        {
            QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
        }

        [Fact]
        [Category("P1")]
        public void RunCreateCompanyTest()
        {
            this.BDDfy();
        }
    }
}
