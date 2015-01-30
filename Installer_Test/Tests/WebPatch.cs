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
using System.Drawing;
using System.Drawing.Imaging;


namespace Installer_Test.Tests
{
   
    public class WebPatch
    {
       
        public static Property conf = Property.GetPropertyInstance();
        public static string Sync_Timeout = conf.get("SyncTimeOut");
        public static string testName = "WebPatch";
        private static extern IntPtr GetForegroundWindow();
        String resultsPath = @"C:\Temp\Results\WebPatch_" + DateTime.Now.ToString("yyyyMMddHHmm") + @"\Screenshots\";
        
        string readpath = "C:\\Temp\\Parameters.xlsm";
        string targetPath, sku;

        [Given(StepTitle = @"The parameters for installation are available at C:\Installation\Parameters.txt")]

        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic = Lib.File_Functions.ReadExcelValues(readpath, "WebPatch", "B2:B12");
            sku = dic["B7"];
            targetPath = dic["B11"];
 
            
        }

       
        [Then(StepTitle = "Then - Kill QuickBooks")]
        public void KillQB()
        {
            try
            {
            OSOperations.KillProcess("qbw32");
            Logger.logMessage("QuickBooks process killed successfully");
            }
            catch(Exception e)
            {
                Logger.logMessage("Unable to Kill process QBW32" + e.ToString());
            }
        }
        
        [Then(StepTitle = "Then - Invoke Web Patch installer")]
        public void InvokeWP()
        {
            if (sku == "Enterprise" || sku == "Enterprise Accountant")
            OSOperations.InvokeInstaller(targetPath, "en_qbwebpatch.exe");
            else
            OSOperations.InvokeInstaller(targetPath, "qbwebpatch.exe");
            

            try
            {
                Actions.WaitForWindow("QuickBooks Update", 10000);
                if(Actions.CheckDesktopWindowExists("QuickBooks Update"))
                {
                    ScreenCapture sc = new ScreenCapture();
                    System.Drawing.Image img = sc.CaptureScreen();
                    IntPtr pointer = GetForegroundWindow();
                    pointer = GetForegroundWindow();
                    sc.CaptureWindowToFile(pointer, resultsPath + "Wrong_WebPatch_Error.png", ImageFormat.Png);
                    Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Update"), "OK");
                }

            }
            catch(Exception e)
            {
                Logger.logMessage("Wrong Patch" + e.ToString());
            }
            try
            {
                Actions.WaitForWindow("QuickBooks Update,Version", 60000);
                Window patchWin = Actions.GetDesktopWindow("QuickBooks Update,Version");
                Actions.WaitForElementEnabled(patchWin, "Install Now", 60000);
                Actions.ClickElementByName(patchWin, "Install Now");
                Logger.logMessage("Installing webpatch");
                Actions.WaitForWindow("QuickBooks Update,Version", 60000);
                Window patchWin1 = Actions.GetDesktopWindow("QuickBooks Update,Version");
                Window updatecomp = Actions.GetChildWindow(patchWin1, "Update complete");
                Actions.ClickElementByName(updatecomp, "OK");

            }
            catch(Exception e)
            {
                Logger.logMessage("Patch Failed" + e.ToString());

            }


           

        }
      


       [Fact]
        public void RunInstallWebPatch()
        {
            this.BDDfy();
        }
    }
}
