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

using Microsoft.VisualStudio.TestTools.UnitTesting;
using ScreenShotDemo;
using TestStack.BDDfy;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.WindowItems;





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

        public static void ClickHelpUpdate(TestStack.White.Application qbApp, Window qbWindow)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("Click Help -> Update" + " - Started..");

            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();
            string resultsPath = @"C:\Temp\Results\Help_Update_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "\\";

            if (!Directory.Exists(resultsPath))
            {
                try
                {
                    Directory.CreateDirectory(resultsPath);
                    Logger.logMessage("Directory " + resultsPath + " created - Successful");
                    Logger.logMessage("------------------------------------------------------------------------------");
                }
                catch (Exception e)
                {
                    Logger.logMessage("Directory " + resultsPath + " could not be created - Failed");
                    Logger.logMessage(e.Message);
                    Logger.logMessage("------------------------------------------------------------------------------");
                }
            }
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
                Logger.logMessage("Help -> Update Now -> Uncheck Maintenance Release - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Help -> Update Now -> Uncheck Maintenance Release - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            try
            {
                SendKeys.SendWait("%g"); // Click on Get Updates
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "02_Get_Updates.png", ImageFormat.Png);
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
                Actions.WaitForChildWindow(qbWindow, "QuickBooks Information", int.Parse(Sync_Timeout));
                Actions.WaitForChildWindow(qbWindow, "QuickBooks Information", int.Parse(Sync_Timeout));
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "03_QB_Information.png", ImageFormat.Png);
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

        public static void ClickHelpAbout(TestStack.White.Application qbApp, Window qbWindow)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("Click Help -> About" + " - Started..");

            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();
            string resultsPath = @"C:\Temp\Results\Help_About_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "\\";

            if (!Directory.Exists(resultsPath))
            {
                try
                {
                    Directory.CreateDirectory(resultsPath);
                    Logger.logMessage("Directory " + resultsPath + " created - Successful");
                    Logger.logMessage("------------------------------------------------------------------------------");
                }
                catch (Exception e)
                {
                    Logger.logMessage("Directory " + resultsPath + " could not be created - Failed");
                    Logger.logMessage(e.Message);
                    Logger.logMessage("------------------------------------------------------------------------------");
                }
            }
            try
            {
                string OS_Name = Installer_Test.Lib.File_Functions.GetOS();
                Dictionary<string, string> dic = new Dictionary<string, string>();


                dic = Installer_Test.Lib.File_Functions.ReadExcelValues("C:\\Temp\\Parameters.xlsx", "Path", "B2:B10");

                string ver = dic["B7"];
                string reg_ver = dic["B8"];
                string product = Installer_Test.Lib.File_Functions.GetProduct(OS_Name, ver, reg_ver);
                string menu = "About Intuit " + product + "...";
                Actions.SelectMenu(qbApp, qbWindow, "Help", menu );
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "01_Help_About.png", ImageFormat.Png);
                Logger.logMessage("Help -> About - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Help -> About - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            SendKeys.SendWait(" ");

        }
    }
}
 