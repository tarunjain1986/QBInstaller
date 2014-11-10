﻿using System;
using System.Windows.Forms;
using Xunit;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.BDDfy;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.ActionLibs.WhiteAPI;

namespace Installer_Test.Tests
{
    [TestClass]
    public class CheckF2
    {

        public string testName = "F2 Check";
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public string exe = conf.get("QBExePath");
        public Random rand = new Random();


        [Given(StepTitle = "Given - QuickBooks App and Window instances are available")]

        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
            QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
        }

        [Then(StepTitle = "Then - Open F2")]
        public void CheckF2value()
        {
            //Actions.SelectMenu(qbApp, qbWindow, "File", "New Company...");
            Install_Functions.CheckF2value(qbApp, qbWindow);

        }
        [Fact]
        public void Run_CheckF2()
        {
            this.BDDfy();
        }
    }
}