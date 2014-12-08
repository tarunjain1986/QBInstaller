
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


using Excel = Microsoft.Office.Interop.Excel;


using Installer_Test.Properties.Lib;

using TestStack.BDDfy;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.InputDevices;



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
        public static string resultsPath, LogFilePath;

        public static string testName = "Installer Test Suite";

        public static void Install_US()
        {
            string country, SKU, installType, targetPath, installPath, wkflow, customOpt, License_No, Product_No, UserID, Passwd, firstName, lastName;
            string[] LicenseNo, ProductNo;
 
            Installer_Test.Lib.ScreenCapture sc = new Installer_Test.Lib.ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();
  
            string readpath = "C:\\Temp\\Parameters.xlsm"; 

            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic = Lib.File_Functions.ReadExcelValues(readpath, "Install", "B2:B27");
            country = dic["B5"];
            SKU = dic["B7"];
            installType = dic["B8"];

            targetPath = dic["B12"];
            targetPath = targetPath + @"QBooks\";

            customOpt = dic["B17"];
            wkflow = dic["B18"];
            License_No = dic["B19"];
            Product_No = dic["B20"];
            UserID = dic["B21"];
            Passwd = dic["B22"];
            firstName = dic["B23"];
            lastName = dic["B24"];

            installPath = dic["B27"];

            var regex = new Regex(@".{4}");
            string temp = regex.Replace(License_No, "$&" + "\n");
            LicenseNo = temp.Split('\n');

            regex = new Regex(@".{3}");
            temp = regex.Replace(Product_No, "$&" + "\n");
            ProductNo = temp.Split('\n');
                       
            resultsPath = @"C:\Temp\Results\Install_" + customOpt + "_" + wkflow + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + @"\Screenshots\";
            LogFilePath = @"C:\Temp\Results\Install_" + customOpt + "_" + wkflow + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + @"\Logs\";
            Add_Log_Automation_Properties(LogFilePath);
            //if (!Directory.Exists(LogFilePath))
            //{
            //    Directory.CreateDirectory(LogFilePath);
            //}

            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);

            Logger.logMessage("InstallQB " + targetPath + " - Started..");
            Logger.logMessage("License Number: " + License_No);
            Logger.logMessage("Product Number " + Product_No);

            Logger.logMessage("Function call @ :" + DateTime.Now);

           // Create a folder to save the Screenshots
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
   
               Open_QB(targetPath);

               // Update the Automation.Properties with the new properties
               File_Functions.Update_Automation_Properties();
            }

            catch (Exception e)
            {
                Logger.logMessage("InstallQB " + targetPath + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                Logger.logMessage("------------------------------------------------------------------------------");

            }
     
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

            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic = Lib.File_Functions.ReadExcelValues(readpath, "Install", "B2:B27");
            country = dic["B5"];
            SKU = dic["B7"];
            installType = dic["B8"];

            targetPath = dic["B12"];
            targetPath = targetPath + @"QBooks\";

            customOpt = dic["B17"];
            wkflow = dic["B18"];
            License_No = dic["B19"];
            Product_No = dic["B20"];
            UserID = dic["B21"];
            Passwd = dic["B22"];
            firstName = dic["B23"];
            lastName = dic["B24"];

            installPath = dic["B27"];

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

            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic = Lib.File_Functions.ReadExcelValues(readpath, "Install", "B2:B27");
            country = dic["B5"];
            SKU = dic["B7"];
            installType = dic["B8"];

            targetPath = dic["B12"];
            targetPath = targetPath + @"QBooks\";

            customOpt = dic["B17"];
            wkflow = dic["B18"];
            License_No = dic["B19"];
            Product_No = dic["B20"];
            UserID = dic["B21"];
            Passwd = dic["B22"];
            firstName = dic["B23"];
            lastName = dic["B24"];

            installPath = dic["B27"];

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
                sc.CaptureWindowToFile(pointer, resultsPath + "09_Ready_to_Install_Print.png", ImageFormat.Png);
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
                Actions.ClickElementByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1"); // Click on Install
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "10_Installation_Progress_01.png", ImageFormat.Png);
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

        public static void Open_QB (string targetPath)
        {
            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();

            Boolean flag = false;
            try
            {
                while (flag == false)
                {
                    flag = Actions.CheckElementExistsByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Finish");
                }

                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "15_Finish_QuickBooks.png", ImageFormat.Png);
                Logger.logMessage("Finish button enabled - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            catch (Exception e)
            {
                Logger.logMessage("Finish button not enabled - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            // Click on Finish
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Finish"); // Click on Finish
                Logger.logMessage("Click on Finish - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Finish - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }


            // Click on Open QuickBooks
            try
            {
                flag = false;
                while (flag == false)
                {
                    flag = Actions.CheckElementExistsByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Open QuickBooks");
                }
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Open QuickBooks"); // Launch QuickBooks
                Logger.logMessage("Click on Open QuickBooks - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on Open QuickBooks - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            Logger.logMessage("InstallQB " + targetPath + " - Successful");
            Logger.logMessage("------------------------------------------------------------------------------");
            Logger.logMessage("------------------------------------------------------------------------------");
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
                        // Select the first radio button
                        Actions.ClickElementByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "504");
                        pointer = GetForegroundWindow();
                        sc.CaptureWindowToFile(pointer, resultsPath + "04_InstallationType_Express.png", ImageFormat.Png);
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
                        // Select the second radio button
                        Actions.ClickElementByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1006");
                        pointer = GetForegroundWindow();
                        sc.CaptureWindowToFile(pointer, resultsPath + "04_InstallationType_CustomAndNetworkOptions.png", ImageFormat.Png);
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

        public static void Select_Option (string customOpt, string targetPath, string installPath)
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
                            // Select the first radio button to install QB on "this" machine
                            Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "I'll be using QuickBooks on this computer."); // 1016
                            pointer = GetForegroundWindow();
                            sc.CaptureWindowToFile(pointer, resultsPath + "04_InstallationType_Local.png", ImageFormat.Png);
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
                            // Select the second radio button for a shared installation of QB
                            Actions.ClickElementByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1018"); // Select the second radio button
                            pointer = GetForegroundWindow();
                            sc.CaptureWindowToFile(pointer, resultsPath + "04_InstallationType_Shared.png", ImageFormat.Png);
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
                        // Click on Next
                        Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >");
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

        public static void Enter_License (string [] LicenseNo, string [] ProductNo)
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
                            sc.CaptureWindowToFile(pointer, resultsPath + "05_License_Number.png", ImageFormat.Png);
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
                            sc.CaptureWindowToFile(pointer, resultsPath + "06_Product_Number.png", ImageFormat.Png);
                            Logger.logMessage("Enter Product Numbers - Successful");
                            Logger.logMessage("------------------------------------------------------------------------------");
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
                sc.CaptureWindowToFile(pointer, resultsPath + "11_Installation_Progress_after_Skip.png", ImageFormat.Png);
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
  
        public static void Create_Dir (string resultsPath)
        {
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

        }

        public static void Change_Install_Location (String installPath)

        {
            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();

            try
            {
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
       
        public static void Delete_QBDLLs(string installed_path)
        {
            string[] dlls = { "abmapi.DLL", "Accountant.DLL", "AccountRegistersUI.DLL", "ACE.DLL", "ACM.DLL", "ADR.DLL", "acXMLParser.dll", "QBADRHelper.dll" };

            foreach (string dll in dlls)
            {
                if (File.Exists (installed_path + "\\" + dll))
                File.Delete(installed_path + "\\" + dll);
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

        public static void Accept_License_Agreement (string country)
        {
            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();
            
            // Wait for the QuickBooks Installation dialog to show up
            Actions.WaitForAppWindow("QuickBooks Installation", int.Parse(Sync_Timeout));

            // Wait for the Next button to be enabled
            Actions.WaitForElementEnabled(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >", int.Parse(Sync_Timeout));
            pointer = GetForegroundWindow();
            sc.CaptureWindowToFile(pointer, resultsPath + "01_Install.png", ImageFormat.Png);

            // Verify that the installer dialog box is displayed and click Next 
            try
            {
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >"); // Click on Next
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
                sc.CaptureWindowToFile(pointer, resultsPath + "02_License_Agreement.png", ImageFormat.Png);
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
                Logger.logMessage("Click on License agreement -> Next button - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on License agreement -> Next button - Failed");
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
    }
}

   
