﻿
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
using Installer_Test.Lib;

using Microsoft.Win32;

using Excel = Microsoft.Office.Interop.Excel;

using Installer_Test.Properties.Lib;

using TestStack.BDDfy;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.InputDevices;

using Xunit;

namespace Installer_Test
{

    public class Install_Functions
    {
        [DllImport("User32.dll")]
        public static extern int SetForegroundWindow(IntPtr point);
        [DllImport("User32.dll")]
        private static extern IntPtr GetForegroundWindow();

        public static Property conf = Property.GetPropertyInstance();
        public static string Sync_Timeout = conf.get("SyncTimeOut");
        public string line;
        public static string custname, vendorname, itemname,backuppath;
        public static Random _r = new Random();
        public static string resultsPath, LogFilePath, industryEdition;

        public static string country, SKU, installType, targetPath, installPath, wkflow, customOpt, License_No, Product_No, UserID, Passwd, firstName, lastName;
        public static string[] LicenseNo, ProductNo;
 
        public static string ver, reg_ver, expected_ver, installed_version, installed_dataPath, dataPath, regPath;

        public static string testName = "Installer Test Suite";

        public static TestStack.White.Application qbApp = null;
        public static TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        public static string Install_US()
        {
            //string country, SKU, installType, targetPath, installPath, wkflow, customOpt, License_No, Product_No, UserID, Passwd, firstName, lastName;
            //string[] LicenseNo, ProductNo;
 

            Installer_Test.Lib.ScreenCapture sc = new Installer_Test.Lib.ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();
  
            string readpath = "C:\\Temp\\Parameters.xlsm";

            // Read all the input values from "C:\\Temp\\Parameters.xlsm"
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic = Lib.File_Functions.ReadExcelValues(readpath, "Install", "B7:B30");
            country = dic["B10"];
            SKU = dic["B12"];
            installType = dic["B13"];

            targetPath = dic["B30"];
            targetPath = targetPath + @"QBooks\";

            customOpt = dic["B14"];
            wkflow = dic["B15"];
            License_No = dic["B16"];
            Product_No = dic["B17"];
            UserID = dic["B18"];
            Passwd = dic["B19"];
            firstName = dic["B20"];
            lastName = dic["B21"];

            installPath = dic["B24"];

            var regex = new Regex(@".{4}");
            string temp = regex.Replace(License_No, "$&" + "\n");
            LicenseNo = temp.Split('\n');

            regex = new Regex(@".{3}");
            temp = regex.Replace(Product_No, "$&" + "\n");
            ProductNo = temp.Split('\n');
                       
            resultsPath = @"C:\Temp\Results\Install_" + customOpt + "_" + wkflow + "_" + DateTime.Now.ToString("yyyyMMddmm") + @"\Screenshots\";
            LogFilePath = @"C:\Temp\Results\Install_" + customOpt + "_" + wkflow + "_" + DateTime.Now.ToString("yyyyMMddmm") + @"\Logs\";
            
            // Add the LogFilePath created at runtime in the Automation.Properties file
            Add_Log_Automation_Properties(LogFilePath);
            conf.reload(); // Reload the property file

            Logger log = new Logger(testName + "_" + DateTime.Now.ToString("yyyyMMddmm"));

            // Create a folder to save the Screenshots
            Create_Dir(resultsPath);

            Logger.logMessage ("----------------------------------------------------------------------");
            Logger.logMessage ("Installation of QuickBooks at " + targetPath + " - Started");
            Logger.logMessage ("License Number: " + License_No);
            Logger.logMessage ("Product Number " + Product_No);

            Logger.logMessage("Function call @ :" + DateTime.Now);
           
            try
            {
                //////////////////////////////////////////////////////////////////////////////////////////////
                //Accept License Agreement
                //////////////////////////////////////////////////////////////////////////////////////////////
                Accept_License_Agreement(country);

                //////////////////////////////////////////////////////////////////////////////////////////////
                // Continue the installation based on the selected SKU
                //////////////////////////////////////////////////////////////////////////////////////////////
               
                switch (SKU)
                {
                    case "Enterprise":
                    case "Enterprise Accountant":
                    Select_Option(customOpt, targetPath, installPath);
                    break;

                    case "Pro":
                    case "Pro Plus":
                    case "Premier":
                    case "Premier Plus":
                    case "Premier Accountant":

                    Select_InstallType(installType, customOpt, targetPath, installPath);

                     if (installType == "Express")
                     {
                         customOpt = "";
                     }
                     break;
                 }

                if (customOpt != "Server")
                {

                    // Click on Use Your User ID instead
                    if (wkflow == "Signin")
                    {
                        SignIn_Flow(UserID, Passwd);
                    }

                    else // If wkflow != "Signin"
                    {

                      Enter_License(LicenseNo, ProductNo);

                      if (customOpt == "Local" | customOpt == "Shared")
                      {
                          Choose_Install_Location(installPath);
                      }

                      Install_QB();


                      if (wkflow == "Skip")
                      {
                          Skip_Flow();
                      }

                      if (wkflow == "Signup")
                      {
                          SignUp_Flow(UserID, Passwd, firstName, lastName);
                      }
                    }
                }
   
               // Click on "Open QuickBooks" on the last installer page
               Open_QB(targetPath);

               // Minimize all open applications before launching QuickBooks
               Shell32.Shell shell = new Shell32.Shell();
               shell.MinimizeAll();

               // Update the Automation.Properties with the new properties
               File_Functions.Update_Automation_Properties();
               conf.reload(); // Added

                // Launch QuickBooks after installation
               Install_Functions.Launch_QB(SKU);
        
               Logger.logMessage ("**********************************************************************");
               Logger.logMessage ("**********************************************************************");
               Logger.logMessage ("QuickBooks installation and Launch - Successful");
               Logger.logMessage ("**********************************************************************");
               Logger.logMessage ("**********************************************************************");

            }

            catch (Exception e)
            {
              Logger.logMessage ("**********************************************************************");
              Logger.logMessage ("**********************************************************************");
              Logger.logMessage ("Installation of QuickBooks at " + targetPath + " - Failed");
              Logger.logMessage (e.Message);
              Logger.logMessage ("**********************************************************************");
              Logger.logMessage ("**********************************************************************");

            }
            return resultsPath;
        }

        public static void Install_UK()
        {
            string country, SKU, installType, targetPath, installPath, wkflow, customOpt, License_No, Product_No, UserID, Passwd, firstName, lastName;
            string[] LicenseNo, ProductNo;
            Logger.logMessage("Function call @ :" + DateTime.Now);

            //ScreenCapture sc = new ScreenCapture();
            //System.Drawing.Image img = sc.CaptureScreen();
            //IntPtr pointer = GetForegroundWindow();

            string readpath = "C:\\Temp\\Parameters.xlsm";

            ////////////////////////////////////////////////////////////////////////////////////////
            //Dictionary<string, string> dic = new Dictionary<string, string>();
            //dic = Lib.File_Functions.ReadExcelValues(readpath, "Install", "B2:B27");
            //country = dic["B5"];
            //SKU = dic["B7"];
            //installType = dic["B8"];

            //targetPath = dic["B12"];
            //targetPath = targetPath + @"QBooks\";

            //customOpt = dic["B17"];
            //wkflow = dic["B18"];
            //License_No = dic["B19"];
            //Product_No = dic["B20"];
            //UserID = dic["B21"];
            //Passwd = dic["B22"];
            //firstName = dic["B23"];
            //lastName = dic["B24"];

            //installPath = dic["B27"];
            //////////////////////////////////////////////////////////////////////////////////////////

            // Read all the input values from "C:\\Temp\\Parameters.xlsm"
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic = Lib.File_Functions.ReadExcelValues(readpath, "Install", "B7:B30");
            country = dic["B10"];
            SKU = dic["B12"];
            installType = dic["B13"];

            targetPath = dic["B30"];
            targetPath = targetPath + @"QBooks\";

            customOpt = dic["B14"];
            wkflow = dic["B15"];
            License_No = dic["B16"];
            Product_No = dic["B17"];
            UserID = dic["B18"];
            Passwd = dic["B19"];
            firstName = dic["B20"];
            lastName = dic["B21"];

            installPath = dic["B24"];

            var regex = new Regex(@".{4}");
            string temp = regex.Replace(License_No, "$&" + "\n");
            LicenseNo = temp.Split('\n');

            regex = new Regex(@".{3}");
            temp = regex.Replace(Product_No, "$&" + "\n");
            ProductNo = temp.Split('\n');

            Logger.logMessage("InstallQB " + targetPath + " - Started..");
            Logger.logMessage("License Number: " + License_No);
            Logger.logMessage("Product Number " + Product_No);
            resultsPath = @"C:\Temp\Results\Install_" + customOpt + "_" + wkflow + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "\\";

            // Create a folder to save the Results
            Create_Dir(resultsPath);

            try
            {
                 Accept_License_Agreement(country);

                 /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                 // Select Install Type: Express OR Custom
                 /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                 Select_InstallType(installType, customOpt , targetPath, installPath);

                 /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                 // Enter the License and Product Numbers
                 /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                 Enter_License(LicenseNo, ProductNo);

                 if (customOpt == "Local" | customOpt == "Shared")
                 {
                      Choose_Install_Location(installPath);
                 }

                 Install_QB();

                 Open_QB (targetPath);
                                            
             }
 
            catch (Exception e)
            {
                Logger.logMessage("InstallQB " + targetPath + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                Logger.logMessage("------------------------------------------------------------------------------");
            }

        }

        public static void Install_CA()
        {
            string country, SKU, installType, targetPath, installPath, wkflow, customOpt, License_No, Product_No, UserID, Passwd, firstName, lastName;
            string[] LicenseNo, ProductNo;
            Logger.logMessage("Function call @ :" + DateTime.Now);


            Installer_Test.Lib.ScreenCapture sc = new Installer_Test.Lib.ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();

            string readpath = "C:\\Temp\\Parameters.xlsm";

            //////////////////////////////////////////////////////////////////////////////////////////
            //Dictionary<string, string> dic = new Dictionary<string, string>();
            //dic = Lib.File_Functions.ReadExcelValues(readpath, "Install", "B2:B27");
            //country = dic["B5"];
            //SKU = dic["B7"];
            //installType = dic["B8"];

            //targetPath = dic["B12"];
            //targetPath = targetPath + @"QBooks\";

            //customOpt = dic["B17"];
            //wkflow = dic["B18"];
            //License_No = dic["B19"];
            //Product_No = dic["B20"];
            //UserID = dic["B21"];
            //Passwd = dic["B22"];
            //firstName = dic["B23"];
            //lastName = dic["B24"];

            //installPath = dic["B27"];
            ////////////////////////////////////////////////////////////////////////////////////////////

            // Read all the input values from "C:\\Temp\\Parameters.xlsm"
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic = Lib.File_Functions.ReadExcelValues(readpath, "Install", "B7:B30");
            country = dic["B10"];
            SKU = dic["B12"];
            installType = dic["B13"];

            targetPath = dic["B30"];
            targetPath = targetPath + @"QBooks\";

            customOpt = dic["B14"];
            wkflow = dic["B15"];
            License_No = dic["B16"];
            Product_No = dic["B17"];
            UserID = dic["B18"];
            Passwd = dic["B19"];
            firstName = dic["B20"];
            lastName = dic["B21"];

            installPath = dic["B24"];

            var regex = new Regex(@".{4}");
            string temp = regex.Replace(License_No, "$&" + "\n");
            LicenseNo = temp.Split('\n');

            regex = new Regex(@".{3}");
            temp = regex.Replace(Product_No, "$&" + "\n");
            ProductNo = temp.Split('\n');

            Logger.logMessage("InstallQB " + targetPath + " - Started..");
            Logger.logMessage("License Number: " + License_No);
            Logger.logMessage("Product Number " + Product_No);
            resultsPath = @"C:\Temp\Results\Install_" + customOpt + "_" + wkflow + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "\\";

            // Create a folder to save the Results
            Create_Dir(resultsPath);

            try
            {
                //////////////////////////////////////////////////////////////////////////////////////////////
                //Accept License Agreement
                //////////////////////////////////////////////////////////////////////////////////////////////
                Accept_License_Agreement(country);

                //////////////////////////////////////////////////////////////////////////////////////////////
                // Continue the installation based on the selected SKU
                //////////////////////////////////////////////////////////////////////////////////////////////

                switch (SKU)
                {
                    case "Enterprise":
                    case "Enterprise Accountant":
                        Select_Option(customOpt, targetPath, installPath);
                        break;

                    case "Pro":
                    case "Premier":
                    case "Premier Accountant":

                        Select_InstallType(installType, customOpt, targetPath, installPath);

                        if (installType == "Express")
                        {
                            customOpt = "";
                        }
                        break;
                }

                if (customOpt != "Server")
                {
                      Enter_License(LicenseNo, ProductNo);

                      if (customOpt == "Local" | customOpt == "Shared")
                      {
                          Choose_Install_Location(installPath);
                      }
                    
                      Install_QB();
                }

                Open_QB(targetPath);
            }

            catch (Exception e)
            {
                Logger.logMessage("InstallQB " + targetPath + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                Logger.logMessage("------------------------------------------------------------------------------");

            }

        }

        public static void Install_QB ()
        {
            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();

            // Click on Ready to Install -> Print
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Print");
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "Ready_to_Install_Print.png", ImageFormat.Png);
                Logger.logMessage("Click on Print - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Print - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            // Click on Ready to Install -> Print -> No
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "No");
                Logger.logMessage("Click on No - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on No - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // FOR UK: ADD SERVICE AND SUPPORT SHORTCUTS????
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Enter code here

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Click on Install
            try
            {
                Thread.Sleep(1000);
                Actions.ClickElementByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1"); // Click on Install
                Actions.ClickElementByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1"); // Click on Install
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "Installation_Progress_01.png", ImageFormat.Png);
                Logger.logMessage("Click on Install - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Install - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
        }

        public static void Create_Dir(string resultsPath)
        {
            if (!Directory.Exists(resultsPath))
            {
                try
                {
                    Directory.CreateDirectory(resultsPath);
                    Thread.Sleep(1000);
                    Logger.logMessage("----------------------------------------------------------------------------------------------------");
                    Logger.logMessage("Directory " + resultsPath + " created - Successful");
                    Logger.logMessage("----------------------------------------------------------------------------------------------------");
                }
                catch (Exception e)
                {
                    Logger.logMessage("----------------------------------------------------------------------------------------------------");
                    Logger.logMessage("Directory " + resultsPath + " could not be created - Failed");
                    Logger.logMessage(e.Message);
                    Logger.logMessage("----------------------------------------------------------------------------------------------------");
                }
            }

        }

        public static void Add_Log_Automation_Properties(string LogFilePath)
        {
            string curr_dir, aut_file;
            curr_dir = Directory.GetCurrentDirectory();
            aut_file = curr_dir + @"\Automation.Properties";
            List<string> prop_value = new List<string>(File.ReadAllLines(aut_file));


            int lineIndex = prop_value.FindIndex(line => line.StartsWith("LogDirectory="));
            if (lineIndex != -1)
            {
                prop_value[lineIndex] = "LogDirectory=" + LogFilePath;
                File.WriteAllLines(aut_file, prop_value);
            }

        }

        public static void Accept_License_Agreement(string country)
        {
            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();

            // Wait for the QuickBooks Installation dialog to show up
            Actions.WaitForAppWindow("QuickBooks Installation", int.Parse(Sync_Timeout));

            // Wait for the Next button to be enabled
            Actions.WaitForElementEnabled(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >", int.Parse(Sync_Timeout));
            pointer = GetForegroundWindow();
            sc.CaptureWindowToFile(pointer, resultsPath + "Install.png", ImageFormat.Png);

            // Verify that the installer dialog box is displayed and click Next 
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >"); // Click on Next
                Thread.Sleep(1000);
                Logger.logMessage("Click Next - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click Next - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            // License Agreement Page
            try
            {
                if (country == "US")
                {
                    Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "I accept the terms of the license agreement");
                }

                else
                {
                    Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "I accept the terms of the licence agreement");
                }
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "License_Agreement.png", ImageFormat.Png);
                Logger.logMessage("License agreement accepted - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Unable to accept License Agreement - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            // Test Case 3: Verfy that the Next button is disabled


            // Click on License Agreement Page -> Print
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Print");
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "03_License_Agreement_Print.png", ImageFormat.Png);
                Logger.logMessage("Click on License agreement -> Print - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on License Agreement -> Print - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            // Test Case 4a: Click on Print -> Yes


            // Click on License Agreement Page -> Print -> No
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "No"); // Click on No
                Logger.logMessage("Click on License agreement -> Print -> No - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on License Agreement -> Print -> No - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            // Click on License Agreement Page -> Next
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >"); // Click on Next
                Thread.Sleep(2000);
                Logger.logMessage("Click on License agreement -> Next button - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on License agreement -> Next button - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            // Thread.Sleep(5000);
        }

        public static void Select_Option(string customOpt, string targetPath, string installPath)
        {
            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();

            // Click on Explain these options in detail
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Explain these options in detail");
                Logger.logMessage("Click on Explain these options in detail - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Explain these options in detail - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            if (customOpt == "Server")
            {
                // If the workflow option is 'Server', there is an entry to be made in
                AddEntry(targetPath, @"[Languages]");

                // Complete the Server Flow
                Server_Flow(installPath);

            }
            if (customOpt == "Local" | customOpt == "Shared")
            {
                switch (customOpt)
                {

                    case "Local":
                        try
                        {
                            // Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks Installation"));
                            // Select the first radio button to install QB on "this" machine
                            Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "I'll be using QuickBooks on this computer."); // 1016
                            pointer = GetForegroundWindow();
                            sc.CaptureWindowToFile(pointer, resultsPath + "InstallationType_Local.png", ImageFormat.Png);
                            Logger.logMessage("Click on I'll be using QuickBooks on this computer - Successful");
                            Logger.logMessage("------------------------------------------------------------------------------");
                        }
                        catch (Exception e)
                        {
                            Logger.logMessage("Click on I'll be using QuickBooks on this computer - Failed");
                            Logger.logMessage(e.Message);
                            Logger.logMessage("------------------------------------------------------------------------------");
                        }
                        break;

                    /////////////////////////////////////////////////////////////////////////////////////////////////////// 
                    case "Shared":
                        try
                        {
                            // Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks Installation"));
                            // Select the second radio button for a shared installation of QB
                            Actions.ClickElementByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1018"); // Select the second radio button
                            pointer = GetForegroundWindow();
                            sc.CaptureWindowToFile(pointer, resultsPath + "InstallationType_Shared.png", ImageFormat.Png);
                            Logger.logMessage("Click on Shared option - Successful");
                            Logger.logMessage("------------------------------------------------------------------------------");
                        }
                        catch (Exception e)
                        {
                            Logger.logMessage("Click on Shared option - Failed");
                            Logger.logMessage(e.Message);
                            Logger.logMessage("------------------------------------------------------------------------------");
                        }
                        break;
                }
            }

            try
            {
                // Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks Installation"));
                // Click on Next
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >");
                Thread.Sleep(1000);
                Logger.logMessage("Click on Next - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Next - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
        }

        public static void Select_InstallType(string installType, string customOpt, string targetPath, string installPath)
        {

            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();

            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Explain these choices in detail");
                Logger.logMessage("Click on Explain these choices in detail - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Explain these choices in detail - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            switch (installType)
            {
                case "Express":
                    try
                    {
                        // Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks Installation"));
                        // Select the first radio button
                        Actions.ClickElementByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "504");
                        pointer = GetForegroundWindow();
                        sc.CaptureWindowToFile(pointer, resultsPath + "InstallationType_Express.png", ImageFormat.Png);
                        Logger.logMessage("Click on Express option - Successful");
                        Logger.logMessage("------------------------------------------------------------------------------");
                    }
                    catch (Exception e)
                    {
                        Logger.logMessage("Click on Express option - Failed");
                        Logger.logMessage(e.Message);
                        Logger.logMessage("------------------------------------------------------------------------------");
                    }

                    try
                    {
                        // Click on Next
                        Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >");
                        Thread.Sleep(1000);
                    }
                    catch (Exception e)
                    {
                        Logger.logMessage("Click on Next - Failed");
                        Logger.logMessage(e.Message);
                        Logger.logMessage("------------------------------------------------------------------------------");
                    }
                    break;

                case "Custom and Network Options":
                    try
                    {
                        // Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks Installation"));
                        // Select the second radio button
                        Actions.ClickElementByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1006");
                        pointer = GetForegroundWindow();
                        sc.CaptureWindowToFile(pointer, resultsPath + "InstallationType_CustomAndNetworkOptions.png", ImageFormat.Png);
                        Logger.logMessage("Click on Custom and Network option - Successful");
                        Logger.logMessage("------------------------------------------------------------------------------");
                    }
                    catch (Exception e)
                    {
                        Logger.logMessage("Click on Custom and Network option - Failed");
                        Logger.logMessage(e.Message);
                        Logger.logMessage("------------------------------------------------------------------------------");
                    }
                    try
                    {
                        // Click on Next
                        Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >");
                        Thread.Sleep(1000);
                    }
                    catch (Exception e)
                    {
                        Logger.logMessage("Click on Next - Failed");
                        Logger.logMessage(e.Message);
                        Logger.logMessage("------------------------------------------------------------------------------");
                    }

                    // Select Custom Option: Local, Shared or Server
                    Select_Option(customOpt, targetPath, installPath);
                    break;
            }
        }

        public static void Enter_License(string[] LicenseNo, string[] ProductNo)
        {

            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();
            // Enter the License Number
            try
            {
                Actions.SetTextByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1054", LicenseNo[0]);
                Actions.SetTextByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1055", LicenseNo[1]);
                Actions.SetTextByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1056", LicenseNo[2]);
                Actions.SetTextByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1057", LicenseNo[3]);
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "License_Number.png", ImageFormat.Png);
                Logger.logMessage("Enter License Numbers - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Enter License Numbers - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            // Enter the Product Number
            try
            {
                Actions.SetTextByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1059", ProductNo[0]);
                Actions.SetTextByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1060", ProductNo[1]);
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "Product_Number.png", ImageFormat.Png);
                Logger.logMessage("Enter Product Numbers - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                // Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks Installation"));
            }
            catch (Exception e)
            {
                Logger.logMessage("Enter Product Numbers - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            // Test Case 6b: User your User ID instead

            // Click on I can't find these numbers
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "I can't find these numbers");
                Logger.logMessage("Click on I can't find these numbers - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on I can't find these numbers - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            // Click on Next
            try
            {
                // Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks Installation"));
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >");
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "Default_Installation_Location.png", ImageFormat.Png);
                Logger.logMessage("Click on Next - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Next - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
        }

        //public static void Open_QB (string targetPath)
        //{
        //   // This function clicks on 'Open QuickBooks' in the final installer dialog box
        //    ScreenCapture sc = new ScreenCapture();
        //    System.Drawing.Image img = sc.CaptureScreen();
        //    IntPtr pointer = GetForegroundWindow();

        //    Boolean flag = false;
        //    int loopCounter = 0;
        //    // Click on Open QuickBooks
        //    try
        //    {
        //        flag = false;

        //        // Wait for the QuickBooks Installation dialog to show up
        //        Actions.WaitForAppWindow("QuickBooks Installation", int.Parse(Sync_Timeout));


        //        // Wait for the "Open QuickBooks" button for max 10 mins
        //        while (flag == false && loopCounter < 120)
        //        {
        //            flag = Actions.CheckElementExistsByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Open QuickBooks"); //Window name changed by maneet
        //            Thread.Sleep(5000);
        //            loopCounter += 1;
        //        }

        //        Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Open QuickBooks"); // Window name changed by maneet
        //        Logger.logMessage("Click on Open QuickBooks - Successful");
        //        Logger.logMessage("------------------------------------------------------------------------------");
        //    }
        //    catch (Exception e)
        //    {
        //        Logger.logMessage("Click on Open QuickBooks - Failed");
        //        Logger.logMessage(e.Message);
        //        Logger.logMessage("------------------------------------------------------------------------------");
        //    }

        //    Logger.logMessage("Installation of Quickbooks " + targetPath + " - Successful");
        //    Logger.logMessage("----------------------------------------------------------------------");
        //}

        public static void Open_QB(string targetPath)
        {
            // This function clicks on 'Open QuickBooks' in the final installer dialog box
            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();

            Boolean flag = false;
            int loopCounter = 0;
            // Click on Open QuickBooks
            try
            {
                flag = false;

                // Wait for the QuickBooks Installation dialog to show up
                Actions.WaitForAppWindow("QuickBooks Installation", int.Parse(Sync_Timeout));

                // Wait for the "Open QuickBooks" button for max 10 mins
                while (flag == false && loopCounter < 120)
                {
                    try
                    {
                        flag = Actions.CheckElementExistsByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Open QuickBooks");
                        Logger.logMessage("CheckElementExistsByName - QuickBooks Installation " + loopCounter);
                        if (flag == true) { break; }
                    }
                    catch
                    {
                        //MessageBox.Show("CheckElementExistsByName - QuickBooks Installation - Get desktop window failed.");
                    }

                    try
                    {
                        flag = Actions.CheckElementExistsByName(Actions.GetDesktopWindow("Intuit QuickBooks Installer"), "Open QuickBooks");
                        Logger.logMessage("CheckElementExistsByName - Intuit QuickBooks Installer  " + loopCounter);
                        Thread.Sleep(5000);
                        loopCounter += 1;
                    }
                    catch
                    {
                        // MessageBox.Show("CheckElementExistsByName - Intuit QuickBooks Installer - Get desktop window failed.");
                    }

                }

                if (Actions.CheckDesktopWindowExists("QuickBooks Installation"))
                {
                    Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Open QuickBooks");
                    MessageBox.Show("QuickBooks Installation Found");
                }
                if (Actions.CheckDesktopWindowExists("Intuit QuickBooks Installer"))
                {
                    Actions.ClickElementByName(Actions.GetDesktopWindow("Intuit QuickBooks Installer"), "Open QuickBooks");
                    MessageBox.Show("Intuit QuickBooks Installer Found");
                }

                //Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Open QuickBooks"); 
                Logger.logMessage("Click on Open QuickBooks - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Open QuickBooks - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
        }


        public static void AddEntry(string targetPath, String SearchStr)
        {
            string lineToAdd = "Server=PDS\n";
            string fileName = targetPath + "Setup.ini";
            List<string> txtLines = new List<string>();

            //Fill a List<string> with the lines from the txt file.
            foreach (string str in File.ReadAllLines(fileName))
            {
                txtLines.Add(str);
            }

            //Insert the line you want to add last above the tag 'Languages'.
            txtLines.Insert(txtLines.IndexOf(SearchStr), lineToAdd);

            //Clear the file. The using block will close the connection immediately.
            using (File.Create(fileName)) { }

            //Add the lines including the new one.
            foreach (string str in txtLines)
            {
                File.AppendAllText(fileName, str + Environment.NewLine);
            }
        }

        public static void Server_Flow(string installPath)
        {
            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();

            try
            {
                // Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks Installation"));
                // Select the third radio button
                Actions.ClickElementByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1019");
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "04_InstallationType_Server.png", ImageFormat.Png);
                Logger.logMessage("Click on Server option - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Server option - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            try
            {
                // Click on Next
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >");
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "05_InstallationPath.png", ImageFormat.Png);
                Logger.logMessage("Click on Next - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Next - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            try
            {
                // Click on Explain these options in detail link
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Explain these options in detail");
                Logger.logMessage("Click on Explain these options in detail - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Explain these options in detail - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Change Install Location
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////

            if (installPath != "")
            {
                Change_Install_Location(installPath);
            }
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////

            try
            {
                // Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks Installation"));
                // Click on Next
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >");
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "06_ReadyToInstall.png", ImageFormat.Png);
                Logger.logMessage("Click on Next - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Next - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            try
            {
                // Click on Ready to Install -> Print
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Print");
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "07_ReadyToInstall_Print.png", ImageFormat.Png);
                Logger.logMessage("Click on Print - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Print - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            try
            {
                // Click on Ready to Install -> Print -> No
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "No");
                Logger.logMessage("Click on Print -> No - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Print -> No - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            try
            {
                // Click on Install
                Actions.ClickElementByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1");
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "08_Installation_Progress.png", ImageFormat.Png);
                Logger.logMessage("Click on Install - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Install - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
        }

        public static void Skip_Flow ()
        {
            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();

            // Click on Skip this
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Skip this");
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "Installation_Progress_after_Skip.png", ImageFormat.Png);
                Logger.logMessage("Click on Skip this - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Skip this - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
        }

        public static void SignIn_Flow (string UserID, string Passwd)
        {
            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();

            try
            {
               Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Use your user ID instead");
               pointer = GetForegroundWindow();
               sc.CaptureWindowToFile(pointer, resultsPath + "05_Sign_In_User_ID.png", ImageFormat.Png);
               Logger.logMessage("Click on Skip this - Successful");
               Logger.logMessage("------------------------------------------------------------------------------");
             }
 
            catch (Exception e)
            {
               Logger.logMessage("Click on Skip this - Failed");
               Logger.logMessage(e.Message);
               Logger.logMessage("------------------------------------------------------------------------------");
            }

            try
            {
               Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Need sign in help?");
               Logger.logMessage("Click on Need sign in help? - Successful");
               Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
               Logger.logMessage("Click on Need sign in help? - Failed");
               Logger.logMessage(e.Message);
               Logger.logMessage("------------------------------------------------------------------------------");
            }

            try
            {
               Actions.SetTextByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1136", UserID);
               Actions.SetTextByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1138", Passwd);
               pointer = GetForegroundWindow();
               sc.CaptureWindowToFile(pointer, resultsPath + "05_SignIn_UserID_Password.png", ImageFormat.Png);
               Logger.logMessage("Enter User ID / Password - Successful");
               Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
               Logger.logMessage("Enter User ID / Password - Failed");
               Logger.logMessage(e.Message);
               Logger.logMessage("------------------------------------------------------------------------------");
            }
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >");
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "07_Default_Installation_Location.png", ImageFormat.Png);
                Logger.logMessage("Click on Next - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Next - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
        }

        public static void SignUp_Flow (string UserID, string Passwd, string firstName, string lastName)
        {
            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();
                        
            // Enter User ID for Signup
            try
            {
                Actions.SetTextByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1135", UserID);
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "12_Signup_UserID.png", ImageFormat.Png);
                Logger.logMessage("Enter User ID - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Enter User ID - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            // Click on Validate
            try
            {
                Actions.ClickButtonByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Validate");
                Logger.logMessage("Click on Validate - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Validate - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            // Click on Validate
            try
            {
                Actions.WaitForElementVisible(Actions.GetDesktopWindow("QuickBooks Installation"), "Create Account", int.Parse(Sync_Timeout));
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "13_Signup_after_Validate.png", ImageFormat.Png);
                Logger.logMessage("Click on Create Account - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Create Account - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }


            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Signup
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Enter Password
            try
            {
                Actions.SetTextOnElementByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1137", Passwd);
                Logger.logMessage("Enter Password - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Enter Password - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            // Re-enter Password
            try
            {
                Actions.SetTextOnElementByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1140", Passwd);
                Logger.logMessage("Re-enter Password - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Re-enter Password - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            // Enter First Name
            try
            {
                Actions.SetTextOnElementByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1141", firstName);
                Logger.logMessage("Enter FirstName - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Enter FirstName - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            // Enter Last Name
            try
            {
                Actions.SetTextOnElementByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1142", lastName);
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "14_Signup_Details.png", ImageFormat.Png);
                Logger.logMessage("Enter LastName - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Enter LastName - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            // Click on Privacy
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Privacy");
                Logger.logMessage("Click on Privacy - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Privacy - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            // Click on Need sign in help?
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Need sign in help?");
                Logger.logMessage("Click on Need sign in help? - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Need sign in help? - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            // Click on Create Account
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Create Account");
                Logger.logMessage("Click on Create Account - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Create Account - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
        }

        public static void Choose_Install_Location(string installPath)
        {
            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();

            // Click on Explain these options in detail
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Explain these options in detail");
                Logger.logMessage("Click on Explain these options in detail - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Explain these options in detail - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Change Install Location
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////

            if (installPath != "")
            {
                Change_Install_Location(installPath);
            }
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Click on Next
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >");
                Thread.Sleep(1000);
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "08_Ready_to_Install.png", ImageFormat.Png);
                Logger.logMessage("Click on Next - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Next - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
        }

        public static void Change_Install_Location (String installPath)

        {
            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();

            try
            {
                // Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks Installation"));
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Browse");
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "08_Browse_Installation_Location.png", ImageFormat.Png);
                Logger.logMessage("Click on Browse - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Browse - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            try
            {
                Actions.SetTextByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1021", installPath);
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "09_Edit_Installation_Location.png", ImageFormat.Png);
                Logger.logMessage("Click on Browse - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Browse - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "OK");
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "10_Installation_Location.png", ImageFormat.Png);
                Logger.logMessage("Click on OK - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on OK - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            // Click on Next
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >");
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "08_Ready_to_Install.png", ImageFormat.Png);
                Logger.logMessage("Click on Next - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Next - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
        }

        public static void Launch_QB(string SKU)
        {

            Boolean flag =  false;
            int loopCounter = 0;

            switch (SKU)
            {
                case "Enterprise":
               // case "Enterprise Accountant":
                    industryEdition = "Enterprise Solutions General Business";
                    break;

                case "Premier":
                case "Premier Plus":
                case "Premier Accountant":
                    industryEdition = "Premier Edition (General Business)";
                    break;

                case "Enterprise Accountant": // Added
                case "Pro":
                case "Pro Plus":
                industryEdition = "";
                Thread.Sleep(20000);
                break;
            }
            try
            {
                if (industryEdition != "")
                {
                  
                    while (flag == false && loopCounter < 120)
                    {
                        try
                        {
                            flag = Actions.CheckElementExistsByName(Actions.GetDesktopWindow("Select QuickBooks Industry-Specific Edition"), industryEdition);
                            Thread.Sleep(5000);
                            Logger.logMessage(" Industry Edition loop - " + loopCounter);
                        }
                        catch
                        {
                            // MessageBox.Show("CheckElementExistsByName for industry Edition -Failed ");
                        }
                        loopCounter += 1;
                    }

                    if (Actions.CheckDesktopWindowExists("Select QuickBooks Industry-Specific Edition"))
                    {
                        Thread.Sleep(2000);
                        Actions.ClickElementByName(Actions.GetDesktopWindow("Select QuickBooks Industry-Specific Edition"), industryEdition);
                        Thread.Sleep(2000);
                        Actions.ClickElementByName(Actions.GetDesktopWindow("Select QuickBooks Industry-Specific Edition"), "Next >");
                        Thread.Sleep(2000);
                        Actions.ClickElementByName(Actions.GetDesktopWindow("Select QuickBooks Industry-Specific Edition"), "Finish");
                        Thread.Sleep(5000);
                        //MessageBox.Show("QuickBooks Installation Found");
                    }





                    Select_Edition(industryEdition);

                }
                else
                {
                    var MainWindow = Actions.GetDesktopWindow("QuickBooks " + SKU);
                    if (Actions.CheckWindowExists(MainWindow, "Register "))
                    {
                        Actions.ClickElementByName(Actions.GetChildWindow(MainWindow, "Register "), "Remind Me Later");
                        Logger.logMessage("Register QuickBooks window found.");
                        Logger.logMessage("-----------------------------------------------------------------");
                    }
                }
            }

            catch (Exception e)
            {
                Logger.logMessage("Launch QuickBooks - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }


        }

        public static Dictionary<string, string> GetDLLVersions(string readpath, string workSheet, string Range)
        {
            //string readpath = "C:\\Temp\\Parameters.xlsm"; // "C:\\Installation\\Sample.txt";
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(readpath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(workSheet);
            Excel.Range xlRng = (Excel.Range)xlWorkSheet.get_Range(Range, Type.Missing);

            Dictionary<string, string> dic = new Dictionary<string, string>();

            foreach (Excel.Range cell in xlRng)
            {

                string cellIndex = cell.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                string cellValue = Convert.ToString(cell.Value2);
                dic.Add(cellIndex, cellValue);

            }
            return dic;
        }
  
        public static void Select_Edition(string industryEdition)
        {
            
            Boolean flag = false;
            //wait for Product Configuration window to appear
            Thread.Sleep(5000);
            try
            {
                Actions.WaitForWindow("Product Configuration", 30000);

                Window win1 = Actions.GetDesktopWindow("Product Configuration");
                
                while (flag == false)
                {
                    flag = Actions.CheckWindowExists(win1, "QuickBooks Product Configuration");
                    Thread.Sleep(2000);
                }

                Window win2 = Actions.GetChildWindow(win1, "QuickBooks Product Configuration");
                Actions.ClickElementByName(win2, "No");

                Thread.Sleep(45000);

                flag = false;
               
                flag = Actions.CheckDesktopWindowExists("QuickBooks Update Service");
                if (flag == true)
                {
                    //Actions.SetFocusOnWindow(Actions.GetDesktopWindow("QuickBooks Update Service"));
                    SendKeys.SendWait("%l");
                    Logger.logMessage("QuickBooks Update Service Window found.");
                }

                string readpath = "C:\\Temp\\Parameters.xlsm";
                Dictionary<string, string> dic_SKU = new Dictionary<string, string>();

                dic_SKU = Lib.File_Functions.ReadExcelValues(readpath, "Install", "B12");
                string SKU = dic_SKU["B12"];

                qbApp = QuickBooks.GetApp("QuickBooks " + SKU);

                if (qbApp == null)
                {
                    qbApp = QuickBooks.GetApp("QuickBooks: " + SKU);
                }

                // qbApp.WaitWhileBusy();

                // to be updated
                //flag = false;
                //if (SKU == "Premier")
                //{
                //    flag = Actions.CheckDesktopWindowExists("License Agreement");
                //    if (flag == true)
                //    {
                //        win1 = Actions.GetDesktopWindow("License Agreement");
                //       // Actions.ClickElementByName();

                //    }
                //}
 
                CheckWindowsAndClose(SKU);
  
             Logger.logMessage (industryEdition + " Edition selected - Successful");

            }
            catch (Exception e)
            {
                Logger.logMessage("Select Edition " + industryEdition + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
        }

        public static void CheckWindowsAndClose (string SKU)
        {
           
            var MainWindow = Actions.GetDesktopWindow("QuickBooks " + SKU);
            string WindowName = "QuickBooks " + SKU;

            if (MainWindow == null)
            {
                MainWindow = Actions.GetDesktopWindow("QuickBooks: " + SKU);
                WindowName = "QuickBooks: " + SKU;
            }


            if (Actions.CheckWindowExists(Actions.GetDesktopWindow(WindowName), "QuickBooks Setup") == true)
            {
                Window SetupWin = Actions.GetChildWindow(Actions.GetDesktopWindow(WindowName), "QuickBooks Setup");
               // Actions.ClickElementByName(SetupWin, "Close"); // Does not work
                SendKeys.SendWait("%{F4}");
                Logger.logMessage("QuickBooks Setup Window found.");
                Logger.logMessage("-----------------------------------------------------------------");
            }

            if (Actions.CheckWindowExists(Actions.GetDesktopWindow(WindowName), "Set Up an External Accountant User") == true)
            {
                Window ExtAcctWin = Actions.GetChildWindow(Actions.GetDesktopWindow(WindowName), "Set Up an External Accountant User");
                Actions.ClickElementByName(ExtAcctWin, "No");
                Logger.logMessage("Set Up an External Accountant User Window found.");
                Logger.logMessage("-----------------------------------------------------------------");
            }

            if (Actions.CheckWindowExists(Actions.GetDesktopWindow(WindowName), "Automatic Backup") == true)
            {
                SendKeys.SendWait("%N");
                Logger.logMessage("Automatic Backup Window found.");
                Logger.logMessage("-----------------------------------------------------------------");
            }

            if (Actions.CheckWindowExists(Actions.GetDesktopWindow(WindowName), "Accountant Center") == true)
            {

                Window AcctCenWin = Actions.GetChildWindow(Actions.GetDesktopWindow("QuickBooks " + SKU), "Accountant Center");
                Actions.ClickElementByName(AcctCenWin, "Close");
                Logger.logMessage("Accountant Center Window found.");
                Logger.logMessage("-----------------------------------------------------------------");
            }

            if (Actions.CheckWindowExists(MainWindow, "QuickBooks Usage & Analytics Study"))
            {
                Actions.ClickElementByName(Actions.GetChildWindow(MainWindow, "QuickBooks Usage  Analytics Study"), "Continue");
                Thread.Sleep(5000);
                Logger.logMessage("QuickBooks Usage & Analytics Study Window found.");
                Logger.logMessage("-----------------------------------------------------------------");
            }

            if (Actions.CheckWindowExists(MainWindow, "About Automatic Update"))
            {
                Actions.ClickElementByName(Actions.GetChildWindow(MainWindow, "About Automatic Update"), "OK");
                Logger.logMessage("About Automatic Update Window found.");
                Logger.logMessage("-----------------------------------------------------------------");
            }

            if (Actions.CheckWindowExists(MainWindow, "Register "))
            {
                Actions.ClickElementByName(Actions.GetChildWindow(MainWindow, "Register "), "Remind Me Later");
                Logger.logMessage("Register QuickBooks window found.");
                Logger.logMessage("-----------------------------------------------------------------");
            }
           
        }

        public static void Delete_QBDLLs(string installed_path)
        {
            string[] dlls = { "abmapi.DLL", "Accountant.DLL", "AccountRegistersUI.DLL", "ACE.DLL", "ACM.DLL", "ADR.DLL", "acXMLParser.dll", "QBADRHelper.dll" };

            try
            {
                foreach (string dll in dlls)
                {
                    if (File.Exists(installed_path + "\\" + dll))
                        File.Delete(installed_path + "\\" + dll);
                }

                Logger.logMessage("QuickBooks dlls have been deleted");
                Logger.logMessage("-------------------------------------------------------");
            }

            catch (Exception e)
            {
                Logger.logMessage("Error in deleting QuickBooks dlls");
                Logger.logMessage(e.Message);
                Logger.logMessage("-------------------------------------------------------");
            }

        }
       
        public static void CleanUp ()
        {
            string dir = @"C:\ProgramData\Intuit\";
            Logger.logMessage ("Cleanup after Uninstall - Started");
            Logger.logMessage ("-----------------------------------------------------------");

            try
            {
                if (Directory.Exists (dir))
                {
                   DirectoryInfo del_dir = new DirectoryInfo(dir);
                   del_dir.Delete(true);
                   Thread.Sleep(1000);
                   Logger.logMessage(dir + " deletion - Successful");
                 }
            }

            catch (Exception e)
            {
                Logger.logMessage (dir + " deletion - Failed");
                Logger.logMessage (e.Message);
                Logger.logMessage ("-----------------------------------------------------------");
            }


            dir = @"C:\ProgramData\COMMON FILES\Intuit\";

            try
            {
                if (Directory.Exists(dir))
                {
                   DirectoryInfo del_dir = new DirectoryInfo(dir);
                   del_dir.Delete(true);
                   Thread.Sleep(1000);
                   Logger.logMessage(dir + " deletion - Successful");
                }
            }

            catch (Exception e)
            {
                Logger.logMessage (dir + " deletion - Failed");
                Logger.logMessage (e.Message);
                Logger.logMessage ("-----------------------------------------------------------");
            }


            dir = @"C:\Program Files (x86)\Intuit\";

            try 
            {
                if (Directory.Exists(dir))
                {
                    DirectoryInfo del_dir = new DirectoryInfo(dir);
                    del_dir.Delete(true);
                    Thread.Sleep(1000);
                    Logger.logMessage(dir + " deletion - Successful");
                }
            }

            catch (Exception e)
            {
                Logger.logMessage (dir + " deletion - Failed");
                Logger.logMessage (e.Message);
                Logger.logMessage ("-----------------------------------------------------------");
            }


            dir = @"C:\Program Files (x86)\Common Files\Intuit\";
            try
            {
                if (Directory.Exists(dir))
                {
                    DirectoryInfo del_dir = new DirectoryInfo(dir);
                    del_dir.Delete(true);
                    Thread.Sleep(1000);
                    Logger.logMessage(dir + " deletion - Successful");
                }
            }

            catch (Exception e)
            {
                Logger.logMessage (dir + " deletion - Failed");
                Logger.logMessage (e.Message);
                Logger.logMessage ("-----------------------------------------------------------");
            }


            dir = @"C:\Program Files\Intuit\";

            try
            {
                if (Directory.Exists(dir))
                {
                    DirectoryInfo del_dir = new DirectoryInfo(dir);
                    del_dir.Delete(true);
                    Thread.Sleep(1000);
                    Logger.logMessage(dir + " deletion - Successful");
                }
            }
            
            catch (Exception e)
            {
                Logger.logMessage (dir + " deletion - Failed");
                Logger.logMessage (e.Message);
                Logger.logMessage ("-----------------------------------------------------------");
            }

            dir = @"C:\Program Files\Common Files\Intuit\";

            try
            {
                if (Directory.Exists(dir))
                {
                    DirectoryInfo del_dir = new DirectoryInfo(dir);
                    del_dir.Delete(true);
                    Thread.Sleep(1000);
                    Logger.logMessage(dir + " deletion - Successful");
                }
            }

            catch (Exception e)
            {
                Logger.logMessage (dir + " deletion - Failed");
                Logger.logMessage (e.Message);
                Logger.logMessage ("-----------------------------------------------------------");
            }

            // Delete Company Files
            dir = @"C:\Users\Public\Documents\Intuit\";

            try
            {
                if (Directory.Exists(dir))
                {
                    DirectoryInfo del_dir = new DirectoryInfo(dir);
                    foreach (var file in del_dir.GetFiles("*", SearchOption.AllDirectories))
                    {
                        file.Attributes &= ~FileAttributes.ReadOnly;
                        file.Delete();
                        Thread.Sleep(1000);
                    }
                   // del_dir.Delete(true);

                    foreach (System.IO.DirectoryInfo subDirectory in del_dir.GetDirectories())
                    {
                        del_dir.Delete(true);
                        Thread.Sleep(1000);
                    }

                    Logger.logMessage(dir + " deletion - Successful");
                }
             }

            catch (Exception e)
            {
                Logger.logMessage (dir + " deletion - Failed");
                Logger.logMessage (e.Message);
                Logger.logMessage ("-----------------------------------------------------------");
            }

            //try
            //{
            //    if (Environment.Is64BitOperatingSystem)
            //    {
            //        regPath = @"Software\Wow6432Node\";
            //    }
            //    else
            //    {
            //        regPath = @"Software\";
            //    }
            //    Logger.logMessage ("Registry entry " + regPath + " deletion - Successful");
            //}

            //catch (Exception e)
            //{
            //    Logger.logMessage ("Registry entry " + regPath + " deletion - Failed");
            //    Logger.logMessage (e.Message);
            //    Logger.logMessage ("-----------------------------------------------------------");
            //}


            //RegistryKey regKey = Registry.LocalMachine.OpenSubKey(regPath, true);
            //if (regKey != null)
            //{
            //    regKey.DeleteSubKeyTree("Intuit");
            //}
            //regPath = @"Software\";
            //regKey = Registry.LocalMachine.OpenSubKey(regPath, true);

            //if (regKey != null)
            //{
            //    regKey.DeleteSubKeyTree("Intuit");
            //}

            Logger.logMessage ("Cleanup after Uninstall - Completed");
            Logger.logMessage ("-----------------------------------------------------------");

            Logger.logMessage ("***************************************************************************************");
            Logger.logMessage ("Execution of Install Suite - Completed");
            Logger.logMessage ("***************************************************************************************");
        }

        public static void Post_Install ()
        {
            string userName;
            string readpath = "C:\\Temp\\Parameters.xlsm"; 

            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic = Lib.File_Functions.ReadExcelValues(readpath, "Install", "B8:B15");
            SKU = dic["B12"];
            ver = dic["B8"];
            string workflow = dic["B15"];
            reg_ver = Lib.File_Functions.GetRegVer(SKU);
                      

            dic = Lib.File_Functions.ReadExcelValues(readpath, "Install", "E7");
            expected_ver = dic["E7"];

            userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            userName = userName.Remove(0, 5);

            installed_version = File_Functions.GetQBVersion(ver, reg_ver);
            installed_dataPath = File_Functions.GetDataPath(ver, reg_ver);
            dataPath = File_Functions.GetDataPathKey(ver, reg_ver);

            Logger.logMessage ("-----------------------------------------------");
            Logger.logMessage ("-----------------------------------------------");
            Logger.logMessage ("Post Install Checks - Started");
            Logger.logMessage ("-----------------------------------------------");

            // Actions.XunitAssertEuqals(expected_ver, installed_version);

            if (expected_ver == installed_version)
            {
                Logger.logMessage("---------------------------------------------------------");
                Logger.logMessage("Expected version = Installed version = " + expected_ver);
                Logger.logMessage("---------------------------------------------------------");
            }
            else
            {
                Logger.logMessage("---------------------------------------------------------");
                Logger.logMessage("Expected version is NOT EQUAL to Installed version");
                Logger.logMessage("Expected Version: " + expected_ver);
                Logger.logMessage("Installed Version: " + installed_version);
                Logger.logMessage("---------------------------------------------------------");
            }

            if (workflow == "Signin")
            {
                try
                {
                    Assert.True(File.Exists(dataPath + "QBInfo.dat"));
                    Logger.logMessage("File QBInfo.dat is available at " + dataPath + " - Successful");
                    Logger.logMessage("------------------------------------------------------------------------------");
                }
                catch (Exception e)
                {
                    Logger.logMessage("File QBInfo.dat is not available at " + dataPath + " - Failed");
                    Logger.logMessage(e.Message);
                    Logger.logMessage("------------------------------------------------------------------------------");
                }

                try
                {
                    Assert.True(File.Exists(dataPath + "QBInfo.dat"));
                    Logger.logMessage("File QBInfo.dat is available at " + dataPath + " - Successful");
                    Logger.logMessage("------------------------------------------------------------------------------");
                }
                catch (Exception e)
                {
                    Logger.logMessage("File QBInfo.dat is not available at " + dataPath + " - Failed");
                    Logger.logMessage(e.Message);
                    Logger.logMessage("------------------------------------------------------------------------------");
                }

                string fileName = "oauth_" + userName + ".dat";
                try
                {
                    Assert.True(File.Exists(dataPath + @"iamdata\" + fileName));
                    Logger.logMessage("File oauth_<username>.dat is available at " + dataPath + "iamdata\\ - Successful");
                    Logger.logMessage("------------------------------------------------------------------------------");
                }
                catch (Exception e)
                {
                    Logger.logMessage("File oauth_<username>.dat is not available at " + dataPath + "iamdata\\ - Failed");
                    Logger.logMessage(e.Message);
                    Logger.logMessage("------------------------------------------------------------------------------");

                }
            }
            Logger.logMessage ("-----------------------------------------------");
            Logger.logMessage ("Post Install Checks - Completed");
            Logger.logMessage ("-----------------------------------------------");
            Logger.logMessage ("-----------------------------------------------");
        }

        public static void Add_Edition_Automation_Properties(string Edition)
        {
            conf.reload();
            string curr_dir, aut_file;
            curr_dir = Directory.GetCurrentDirectory();
            aut_file = curr_dir + @"\Automation.Properties";
            List<string> prop_value = new List<string>(File.ReadAllLines(aut_file));


            int lineIndex = prop_value.FindIndex(line => line.StartsWith("Edition="));
            if (lineIndex != -1)
            {
                prop_value[lineIndex] = "Edition=" + Edition;
                File.WriteAllLines(aut_file, prop_value);
            }

        }

        public static void Get_QuickBooks_Edition(TestStack.White.Application qbApp, Window qbWindow)
        {
            string Edition = qbWindow.Title;

            if (!Edition.Contains("Pro") && (!Edition.Contains ("Premier")))
            {
                Edition = Edition.Substring(Edition.IndexOf("Intuit ") + "Intuit ".Length);

                if (Edition.Contains("Manufacturing and Wholesale"))
                {
                    Edition = Edition.Replace("Manufacturing and Wholesale", "Mfg and Whsle");
                }

                string left_str = Edition.Substring(0, Edition.LastIndexOf(' '));
                string right_str = Edition.Substring(Edition.LastIndexOf(' ') + 1);

                if (left_str != "QuickBooks Enterprise Solutions")
                {
                    Edition = left_str + " Edition " + right_str;
                }
            }

            else
            {
                Edition = Edition.Substring(Edition.IndexOf("-") + "- ".Length);

                string left_str = Edition.Substring(0, Edition.LastIndexOf(' '));
                string right_str = Edition.Substring(Edition.LastIndexOf(' ') + 1);

                if (left_str.Contains ("Plus"))
                {
                    left_str = left_str.Replace("Plus", "");
                }

                Edition = left_str + " " + right_str;
            }


            Install_Functions.Add_Edition_Automation_Properties(Edition);

        }

    }
}

   
