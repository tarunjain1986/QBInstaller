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
using FrameworkLibraries.ActionLibs;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using FrameworkLibraries.AppLibs.QBDT;
using TestStack.White.UIItems.WindowItems;

using Excel = Microsoft.Office.Interop.Excel;

using ScreenShotDemo;
using Installer_Test.Properties.Lib;

using TestStack.BDDfy;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.InputDevices;

namespace Installer_Test
{
    public class AntiVirus_Functions
    {
        [DllImport("User32.dll")]
        public static extern int SetForegroundWindow(IntPtr point);
        [DllImport("User32.dll")]
        private static extern IntPtr GetForegroundWindow();
        public static Property conf = Property.GetPropertyInstance();
        public static string Sync_Timeout = conf.get("SyncTimeOut");

        public static void Copy_AVSoftware(string SWName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("Copy AntiVirus software started:" + SWName + " - Started..");

            string AVPath = @"\\banfsalab02\Users\RajSunder\AntiVirus-Trial\";
            string targetPath = @"C:\Temp\AntiVirus\";

            if (!Directory.Exists(targetPath))
            {
                try
                {
                    Directory.CreateDirectory(targetPath);
                    Logger.logMessage("Directory " + targetPath + " created - Successful");
                    Logger.logMessage("------------------------------------------------------------------------------");
                }
                catch (Exception e)
                {
                    Logger.logMessage("Directory " + targetPath + " could not be created - Failed");
                    Logger.logMessage(e.Message);
                    Logger.logMessage("------------------------------------------------------------------------------");
                }
            }
            if (!File.Exists(targetPath + SWName))
            {
                try
                {
                    File.Copy(AVPath + SWName, targetPath + SWName);
                    Logger.logMessage("File " + SWName + " copied to " + targetPath + " - Successful");
                    Logger.logMessage("------------------------------------------------------------------------------");
                }
                catch (Exception e)
                {
                    Logger.logMessage("File " + SWName + " could not be copied to " + targetPath + " - Failed");
                    Logger.logMessage(e.Message);
                    Logger.logMessage("------------------------------------------------------------------------------");
                }
            }
        }

        public static void Copy_WebPatch(string sku, string wppath)
        {
            string exename;
            wppath = wppath + sku + "\\qbwebpatch\\";
            Logger.logMessage("Function call @ :" + DateTime.Now);

            if (sku == "BEL")
            {
                exename = "en_qbwebpatch.exe";

            }
            else
            {
                exename = "qbwebpatch.exe";

            }

            Logger.logMessage("Copy" + sku + " WebPatch- Started..");

            string targetPath = @"C:\Temp\WebPatch\";

            if (!Directory.Exists(targetPath))
            {
                try
                {
                    Directory.CreateDirectory(targetPath);
                    Logger.logMessage("Directory " + targetPath + " created - Successful");
                    Logger.logMessage("------------------------------------------------------------------------------");
                }
                catch (Exception e)
                {
                    Logger.logMessage("Directory " + targetPath + " could not be created - Failed");
                    Logger.logMessage(e.Message);
                    Logger.logMessage("------------------------------------------------------------------------------");
                }
            }
            if (!File.Exists(targetPath + exename))
            {
                try
                {
                    File.Copy(wppath + exename, targetPath + exename);
                    Logger.logMessage("File " + exename + " copied to " + targetPath + " - Successful");
                    Logger.logMessage("------------------------------------------------------------------------------");
                }
                catch (Exception e)
                {
                    Logger.logMessage("File " + exename + " could not be copied to " + targetPath + " - Failed");
                    Logger.logMessage(e.Message);
                    Logger.logMessage("------------------------------------------------------------------------------");
                }
            }
        }

        public static void Install_AVSoftware(string SWName)
        {
            // Call the respective function
            switch (SWName)
            {
                case "MSEInstall.exe":
                    Install_MSEInstaller(SWName);
                    break;

                case "eset_nod32_antivirus_live_installer_.exe":
                    Install_Nod32(SWName);
                    break;

                case "avast_internet_security_setup.exe":
                    Install_Avast(SWName);
                    break;
            }
        }

        public static void Install_MSEInstaller(string SWName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("Install AntiVirus software started:" + SWName + " - Started..");

            string targetPath = @"C:\Temp\AntiVirus\";
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            string cmdText = "/c cd " + targetPath + " && ren " + SWName + " " + SWName + ".bak && type " + SWName + ".bak > " + SWName + " && del " + SWName + ".bak";
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = cmdText;
            startInfo.UseShellExecute = false;
            process.StartInfo = startInfo;
            process.Start();
            process.WaitForExit();

            try
            {
                OSOperations.InvokeInstaller(targetPath, SWName);
                Logger.logMessage("Open installer " + SWName + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Open installer " + SWName + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            Actions.WaitForAppWindow("Microsoft Security Essentials", int.Parse(Sync_Timeout));
            Actions.WaitForElementEnabled(Actions.GetDesktopWindow("Microsoft Security Essentials"), "Next >", int.Parse(Sync_Timeout));
            Actions.ClickElementByName(Actions.GetDesktopWindow("Microsoft Security Essentials"), "Next >");
            Actions.ClickElementByName(Actions.GetDesktopWindow("Microsoft Security Essentials"), "I accept");
            Actions.ClickElementByName(Actions.GetDesktopWindow("Microsoft Security Essentials"), "I do not want to join the program at this time");
            Actions.ClickElementByName(Actions.GetDesktopWindow("Microsoft Security Essentials"), "Next >");
            Actions.ClickElementByName(Actions.GetDesktopWindow("Microsoft Security Essentials"), "Turn on automatic sample submission.");
            Actions.ClickElementByName(Actions.GetDesktopWindow("Microsoft Security Essentials"), "Next >");

            // Actions.WaitForElementVisible(Actions.GetDesktopWindow("Microsoft Security Essentials"), "Install >", int.Parse(Sync_Timeout));
            // Actions.ClickElementByName(Actions.GetDesktopWindow("Microsoft Security Essentials"), "Cancel");
            Boolean flag = false;

            while (flag == false)
            {
                flag = Actions.CheckElementExistsByName(Actions.GetDesktopWindow("Microsoft Security Essentials"), "Install >");
            }

            Actions.WaitForElementEnabled(Actions.GetDesktopWindow("Microsoft Security Essentials"), "Cancel", int.Parse(Sync_Timeout));
            Actions.ClickElementByName(Actions.GetDesktopWindow("Microsoft Security Essentials"), "Install >");
            flag = false;
            while (flag == false)
            {
                flag = Actions.CheckElementExistsByName(Actions.GetDesktopWindow("Microsoft Security Essentials"), "Finish");
            }
            Actions.ClickElementByName(Actions.GetDesktopWindow("Microsoft Security Essentials"), "Scan my computer for potential threats after getting the latest updates.");
            Actions.ClickElementByName(Actions.GetDesktopWindow("Microsoft Security Essentials"), "Finish");
            Actions.WaitForAppWindow("Microsoft Security Essentials", int.Parse(Sync_Timeout));
            Actions.ClickElementByName(Actions.GetDesktopWindow("Microsoft Security Essentials"), "Close");
        }

        public static void Install_Nod32(string SWName)
        {
            string targetPath = @"C:\Temp\AntiVirus\";
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            string cmdText = "/c cd " + targetPath + " && ren " + SWName + " " + SWName + ".bak && type " + SWName + ".bak > " + SWName + " && del " + SWName + ".bak";
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = cmdText;
            startInfo.UseShellExecute = false;
            process.StartInfo = startInfo;
            process.Start();
            process.WaitForExit();
            OSOperations.InvokeInstaller(targetPath, SWName);

            var temp = Actions.GetDesktopWindow("Install ESET NOD32 Antivirus");


            //    Actions.WaitForAppWindow("Install ESET NOD32 Antivirus", int.Parse(Sync_Timeout));
            Boolean flag = false;

            while (flag == false)
            {
                flag = Actions.CheckElementExistsByName(Actions.GetDesktopWindow("Install ESET NOD32 Antivirus"), "Next");
            }

            // Actions.WaitForElementEnabled(Actions.GetDesktopWindow("Install ESET NOD32 Antivirus"), "Next", int.Parse(Sync_Timeout));
            Actions.ClickButtonByAutomationID(Actions.GetDesktopWindow("Install ESET NOD32 Antivirus"), "12324");
            // Actions.ClickElementByName(Actions.GetDesktopWindow("Install ESET NOD32 Antivirus"), "Next");

            flag = false;

            while (flag == false)
            {
                flag = Actions.CheckElementExistsByName(Actions.GetDesktopWindow("Install ESET NOD32 Antivirus"), "I accept");
            }
            FrameworkLibraries.ActionLibs.WhiteAPI.Actions.ClickElementByName(Actions.GetDesktopWindow("Install ESET NOD32 Antivirus"), "I accept");

            Actions.ClickElementByName(Actions.GetDesktopWindow("Install ESET NOD32 Antivirus"), "Enable detection of potentially unwanted applications");
            Actions.ClickElementByName(Actions.GetDesktopWindow("Install ESET NOD32 Antivirus"), "Install");
            // Actions.ClickElementByName(Actions.GetDesktopWindow("Microsoft Security Essentials"), "Turn on automatic sample submission.");
            // Actions.ClickElementByName(Actions.GetDesktopWindow("Microsoft Security Essentials"), "Next >");

        }

        public static void Install_Avast(string SWName)
        {
            
            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("Install AntiVirus software started:" + SWName + " - Started..");


            string targetPath = @"C:\Temp\AntiVirus\";
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            string cmdText = "/c cd " + targetPath + " && ren " + SWName + " " + SWName + ".bak && type " + SWName + ".bak > " + SWName + " && del " + SWName + ".bak";
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = cmdText;
            startInfo.UseShellExecute = false;
            process.StartInfo = startInfo;
            process.Start();
            process.WaitForExit();
            try
            {
                OSOperations.InvokeInstaller(targetPath, SWName);
                Logger.logMessage("Open installer " + SWName + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Open installer " + SWName + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            try
            {
                Thread.Sleep(15000);
                Process p = Process.GetProcessesByName("instup")[0];
                IntPtr pointer = p.MainWindowHandle;
                SetForegroundWindow(pointer);

                SendKeys.SendWait("%");
                SendKeys.SendWait("e");
                Thread.Sleep(3000);
                SendKeys.SendWait("%");
                SendKeys.SendWait("y");
                Thread.Sleep(1000);
                SendKeys.SendWait("%");
                SendKeys.SendWait("c");
                Thread.Sleep(1000);
                SendKeys.SendWait("%");
                SendKeys.SendWait("c");
                Logger.logMessage("Installed AntiVirus software " + SWName + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Installed AntiVirus software " + SWName + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
        }

        public static void Scan_AVSoftware(string SWName)
        {
            switch (SWName)
            {
                //case "MSEInstall.exe":
                //    Scan_MSEInstaller(SWName);
                //    break;

                //case "eset_nod32_antivirus_live_installer_.exe":
                //    Scan_Nod32(SWName);
                //    break;

                case "avast_internet_security_setup.exe":
                    Scan_Avast(SWName);
                    break;
            }
        }

        public static void Scan_Avast(string SWName)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("Scanning with AntiVirus software started:" + SWName + " - Started..");

            string antiVirusPath = @"C:\Program Files\AVAST Software\Avast\";
            string targetPath = @"C:\Installer_Build\";
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
            string cmdText = "/c cd " + antiVirusPath + " && ashCmd.exe " + targetPath;
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = cmdText;
            startInfo.UseShellExecute = true;
            process.StartInfo = startInfo;
            try
            {
                process.Start();
                process.WaitForExit();
                Logger.logMessage("Scanning with AntiVirus software " + SWName + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Scanning with AntiVirus software " + SWName + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                Logger.logMessage("------------------------------------------------------------------------------");
            }


        }
    }
}
