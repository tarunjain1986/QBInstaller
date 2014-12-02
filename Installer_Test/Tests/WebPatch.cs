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
using FrameworkLibraries.ActionLibs.WhiteAPI;
using Installer_Test.Lib;


namespace Installer_Test.Tests
{
   
    public class WebPatch
    {
       // public TestStack.White.Application qbApp = null;
       // public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public static string Sync_Timeout = conf.get("SyncTimeOut");
        public static string testName = "WebPatch";
        
        string readpath = "C:\\Temp\\Parameters.xlsm";

        [Given(StepTitle = @"The parameters for installation are available at C:\Installation\Parameters.txt")]

        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
 
            
        }

      

       
        [Then(StepTitle = "Then - Kill QuickBooks")]
        public void KillQB()
        {
            OSOperations.KillProcess("qbw32");

        }
        [Then(StepTitle = "Copy the web patch to local")]
        public void copyPatch()
        {

            // Installer_Test.Install_Functions.Copy_WebPatch("BEL",patchpath);

            File_Functions.Copy_WebPatch();

        }

        [AndThen(StepTitle = "Then - Invoke Web Patch installer")]
        public void InvokeWP()
        {
            string targetPath = @"C:\Temp\WebPatch\";
            OSOperations.InvokeInstaller(targetPath, "en_qbwebpatch.exe");
            Logger.logMessage("Copied");
            Thread.Sleep(1000);
            Window patchWin= Actions.GetDesktopWindow("QuickBooks Update,Version");
            Thread.Sleep(1000);
            Actions.ClickElementByName(patchWin, "Install Now");
            Logger.logMessage("Installing webpatch");

        }
      


       [Fact]
        public void RunInstallWebPatch()
        {
            this.BDDfy();
        }
    }
}
