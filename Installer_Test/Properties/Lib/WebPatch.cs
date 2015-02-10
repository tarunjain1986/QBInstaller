using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

using FrameworkLibraries;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.ActionLibs.WhiteAPI;


using ScreenShotDemo;
using TestStack.BDDfy;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.WindowItems;

namespace Installer_Test.Lib
{

    public class WebPatch
    {

        [DllImport("User32.dll")]
        public static extern int SetForegroundWindow(IntPtr point);
        [DllImport("User32.dll")]
        private static extern IntPtr GetForegroundWindow();

        public static Property conf = Property.GetPropertyInstance();
        public static string Sync_Timeout = conf.get("SyncTimeOut");

        public static void ApplyWebPatch(string resultsPath)
        {

            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("Click Help -> Update" + " - Started..");

            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();
            
            try
            {
                Actions.WaitForWindow("QuickBooks Update", 30000);
                if (Actions.CheckDesktopWindowExists("QuickBooks Update"))
                {
                    pointer = GetForegroundWindow();
                    sc.CaptureWindowToFile(pointer, resultsPath + "Wrong_WebPatch_Error.png", ImageFormat.Png);
                    Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Update"), "OK");
                }

            }
            catch (Exception e)
            {
                Logger.logMessage("Wrong Patch" + e.ToString());
            }

            try
            {
                Actions.WaitForWindow("QuickBooks Update,Version", 60000);
                if (Actions.CheckDesktopWindowExists("QuickBooks Update,Version"))
                {
                    Window patchWin = Actions.GetDesktopWindow("QuickBooks Update,Version");
                    Actions.WaitForElementEnabled(patchWin, "Install Now", 60000);
                    Actions.ClickElementByName(patchWin, "Install Now");
                    Logger.logMessage("Installing webpatch");
                }
            }
                catch (Exception e)
                {
                    Logger.logMessage ("Patch application - Failed");
                    Logger.logMessage (e.Message);
                    Logger.logMessage ("--------------------------------------------------------------------");
                }

            try
            {
                    Actions.WaitForWindow("QuickBooks Update,Version", 60000);
                    Window patchWin1 = Actions.GetDesktopWindow("QuickBooks Update,Version");
                    Window updatecomp = Actions.GetChildWindow(patchWin1, "Update complete");
                    pointer = GetForegroundWindow();
                    sc.CaptureWindowToFile(pointer, resultsPath + "Patch_Applied_Succes.png", ImageFormat.Png);
                    Actions.ClickElementByName(updatecomp, "OK");

                    Logger.logMessage("Patch Application - Successful");
                    Logger.logMessage("-----------------------------------------------------------------------");
                
            }
            catch (Exception e)
            {
                Logger.logMessage("Patch Application - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("-----------------------------------------------------------------------");
            }
        }
    }
}
