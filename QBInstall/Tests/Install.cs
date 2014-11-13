using System;
using FrameworkLibraries.Utils;
using FrameworkLibraries.ActionLibs;
using TestStack.White.UIItems.WindowItems;
using System.Threading;
using TestStack.White.UIItems.Finders;

using FrameworkLibraries;
using System.Collections.Generic;
using TestStack.White.UIItems;
using Xunit;
using TestStack.BDDfy;
using FrameworkLibraries.AppLibs.QBDT;
using System.IO;
using System.Reflection;
using System.Diagnostics;


namespace QBInstall.Tests
{
    public class Install
    {
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public static int Sync_Timeout = int.Parse(conf.get("SyncTimeOut"));
        public static string testName = "Install";

        [Given]
        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);

            //Install
            QuickBooks.InstallQB(@"\\banfsalab01\Builds\QuickBooks\MangoR1_US\US_R1_45.1.db\CD_SPRO\QBooks\", "setup.exe", "9840-6473-1929-402", "169-744");
        }

        [Fact]
        public void RunQBInstallTest()
        {
            this.BDDfy();
        }
    }
}
