using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Management;
using System.Reflection;
using System.Diagnostics;
using System.Collections.Generic;

using FrameworkLibraries;
using FrameworkLibraries.Utils;
using FrameworkLibraries.ActionLibs;
using TestStack.White.UIItems.WindowItems;
using FrameworkLibraries.ActionLibs.WhiteAPI;

using TestStack.BDDfy;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;

using Xunit;

using Microsoft.Win32;

using FrameworkLibraries.AppLibs.QBDT;

using Installer_Test;
using Installer_Test.Lib;

namespace Installer_Test.Archive_Tests
{
    public class Remove
    {
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public static int Sync_Timeout = int.Parse(conf.get("SyncTimeOut"));
        public static string testName = "Uninstall";
        public string SKU, ver, reg_ver, installed_product;
        string OS_Name = string.Empty;

        [Given]
        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);

            //// Kill any existing QuickBooks process before uninstalling
            //OSOperations.KillProcess("QBW32");
            //Thread.Sleep(1000);

            //if (Actions.CheckDesktopWindowExists("Programs and Features"))
            //{
            //    Actions.SetFocusOnWindow(Actions.GetDesktopWindow("Programs and Features"));
            //    Actions.ClickElementByName(Actions.GetDesktopWindow("Programs and Features"), "Close");
            //}

            //// Get the QuickBooks Edition to Repair from the Automation.Properties file
            //conf.reload();
            //installed_product = conf.get("Edition");
            //QuickBooks.RepairOrUnInstallQB(installed_product, false, true);

            Install_Functions.CleanUp();
        }

        [Fact]
        public void RunQBRemoveTest()
        {
            this.BDDfy();
        }
    }
}
