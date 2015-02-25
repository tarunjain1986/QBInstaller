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

using Installer_Test.Properties.Lib;


namespace Installer_Test
{

    public class Help
    {
        [DllImport("User32.dll")]
        public static extern int SetForegroundWindow(IntPtr point);
        [DllImport("User32.dll")]
        private static extern IntPtr GetForegroundWindow();

        public static Property conf = Property.GetPropertyInstance();
        public static string Sync_Timeout = conf.get("SyncTimeOut");
        public static string SKU;

        public static void ClickHelpUpdate_ULIP(TestStack.White.Application qbApp, Window qbWindow, string resultsPath)
        {
           // Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("Click Help -> Update" + " - Started..");

            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();

            try
            {
                Actions.SelectMenu(qbApp, qbWindow, "Help", "Update QuickBooks...");
                Actions.WaitForChildWindow(qbWindow, "Update QuickBooks", int.Parse(Sync_Timeout));
                SendKeys.SendWait("%n");
                SendKeys.SendWait("{DOWN}");
                SendKeys.SendWait("{DOWN}");
                SendKeys.SendWait(" ");

                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "01_Help_Update.png", ImageFormat.Png);
                Logger.logMessage("ULIP: Help -> Update Now -> Uncheck Maintenance Release - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("ULIP: Help -> Update Now -> Uncheck Maintenance Release - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            try
            {
                SendKeys.SendWait("%g"); // Click on Get Updates
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "Get_Updates.png", ImageFormat.Png);
                Logger.logMessage("Click Get Updates - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click Get Updates - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            try
            {
                //Actions.WaitForChildWindow(qbWindow, "QuickBooks Information", int.Parse(Sync_Timeout));
                //Actions.WaitForChildWindow(qbWindow, "QuickBooks Information", int.Parse(Sync_Timeout));

                Window win1 = Actions.GetDesktopWindow("QuickBooks");

                Boolean flag = false;
                while (flag == false)
                {
                    flag = Actions.CheckWindowExists(win1, "QuickBooks Information");
                    Thread.Sleep(2000);
                }
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "QB_Information.png", ImageFormat.Png);
                Window CloseQBInfo = Actions.GetChildWindow(qbWindow, "QuickBooks Information");
                Actions.ClickElementByName(CloseQBInfo, "OK"); // Click on OK
                Logger.logMessage("Click QuickBooks Information -> OK - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click QuickBooks Information -> OK - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            try
            {
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "04_Updates_Installed.png", ImageFormat.Png);
                Window CloseUpdateQB = Actions.GetChildWindow(qbWindow, "Update QuickBooks");
                Actions.ClickElementByName(CloseUpdateQB, "Close"); // Click on Close
                Logger.logMessage("Click Update QuickBooks -> Close - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            catch (Exception e)
            {
                Logger.logMessage("Click Update QuickBooks -> Close - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

        }

        public static void ClickHelpUpdate_AutoPatch(TestStack.White.Application qbApp, Window qbWindow, string resultsPath)
        {
            // Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("Click Help -> Update" + " - Started..");

            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();

            try
            {
                Actions.SelectMenu(qbApp, qbWindow, "Help", "Update QuickBooks...");
                Actions.WaitForChildWindow(qbWindow, "Update QuickBooks", int.Parse(Sync_Timeout));
                SendKeys.SendWait("%n");

                ///////////////////////////////////////////////
                //SendKeys.SendWait("{DOWN}");
                //SendKeys.SendWait("{DOWN}");
                //SendKeys.SendWait(" ");
                ///////////////////////////////////////////////

                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "Help_Update.png", ImageFormat.Png);
                Logger.logMessage("AutoPatch: Help -> Update Now - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("AutoPatch: Help -> Update Now - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            try
            {
                SendKeys.SendWait("%g"); // Click on Get Updates
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "AutoPatch_Get_Updates.png", ImageFormat.Png);
                Logger.logMessage("Click Get Updates - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click Get Updates - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            try
            {
                //Actions.WaitForChildWindow(qbWindow, "QuickBooks Information", int.Parse(Sync_Timeout));
                //Actions.WaitForChildWindow(qbWindow, "QuickBooks Information", int.Parse(Sync_Timeout));

                Dictionary<string, string> dic = new Dictionary<string, string>();
                dic = Lib.File_Functions.ReadExcelValues("C:\\Temp\\Parameters.xlsm", "Install", "B9:B12");
                SKU = dic["B12"];

                Window win1 = Actions.GetDesktopWindow("QuickBooks " + SKU);

                Boolean flag = false;
                while (flag == false)
                {
                    flag = Actions.CheckWindowExists(win1, "QuickBooks Information");
                    Thread.Sleep(10000);
                }
                //pointer = GetForegroundWindow();
                //sc.CaptureWindowToFile(pointer, resultsPath + "QB_Information.png", ImageFormat.Png);
                Window CloseQBInfo = Actions.GetChildWindow(win1, "QuickBooks Information");
                Actions.ClickElementByName(CloseQBInfo, "OK"); // Click on OK
                Logger.logMessage("Click QuickBooks Information -> OK - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click QuickBooks Information -> OK - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            try
            {
                //pointer = GetForegroundWindow();
                //sc.CaptureWindowToFile(pointer, resultsPath + "Updates_Installed.png", ImageFormat.Png);
                Window CloseUpdateQB = Actions.GetChildWindow(Actions.GetDesktopWindow ("QuickBooks " + SKU), "Update QuickBooks");
                Actions.ClickElementByName(CloseUpdateQB, "Close"); // Click on Close
                Logger.logMessage("Click Update QuickBooks -> Close - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            catch (Exception e)
            {
                Logger.logMessage("Click Update QuickBooks -> Close - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

        }
        public static void ClickHelpAbout(TestStack.White.Application qbApp, Window qbWindow, string resultsPath)
        {
            Logger.logMessage("-----------------------------------------------------");
            Logger.logMessage("Click Help -> About - Started");

            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();

            try
            {
                string OS_Name = Installer_Test.Lib.File_Functions.GetOS();
                Dictionary<string, string> dic = new Dictionary<string, string>();

                //dic = Installer_Test.Lib.File_Functions.ReadExcelValues("C:\\Temp\\Parameters.xlsm", "PostInstall", "B2:B3");

                //string ver = dic["B2"];
                //string reg_ver = dic["B3"];

                dic = Lib.File_Functions.ReadExcelValues("C:\\Temp\\Parameters.xlsm", "Install", "B8:B12");
                string SKU = dic["B12"];
                string ver = dic["B8"];
                string reg_ver = Lib.File_Functions.GetRegVer(SKU);

                string product = Installer_Test.Lib.File_Functions.GetProduct(ver, reg_ver);
                string menu = "About Intuit " + product + "...";
                Actions.SelectMenu(qbApp, qbWindow, "Help", menu );
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "Help_About.png", ImageFormat.Png);
                Logger.logMessage("Help -> About - Successful");
                Logger.logMessage("-----------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Help -> About - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            Thread.Sleep(2000);
            SendKeys.SendWait(" "); // To close the About dialog

        }
    }
}
 