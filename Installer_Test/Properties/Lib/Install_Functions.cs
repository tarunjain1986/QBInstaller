
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

using Microsoft.VisualStudio.TestTools.UnitTesting;

using ScreenShotDemo;

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
        public static string custname, vendorname, itemname;


        public static void Install_QB(string targetPath, string workFlow, string CustomOpt, string[] LicenseNo, string[] ProductNo, string UserID, string Passwd, string firstName, string lastName, string installPath)
        {

            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("InstallQB " + targetPath + " - Started..");
            Logger.logMessage("License Number: " + LicenseNo);
            Logger.logMessage("Product Number " + ProductNo);

            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();
            string resultsPath = @"C:\Temp\Results\Install_" + CustomOpt + "_" + workFlow + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "\\";
         

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

                
                try
                {
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
                        Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "I accept the terms of the license agreement");
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

                    //if (SKU == "CD_SPRO")
                    //{

                    //}

                    // Click on "Explain these options in detail
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

                    if (CustomOpt == "Server")
                    {
                        // If the workflow option is 'Server', there is an entry to be made in
                        AddEntry(targetPath, @"[Languages]");

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
                    if (CustomOpt == "Local" | CustomOpt == "Shared")
                    {
                        
                        switch (CustomOpt)
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

                        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        // Click on Use Your User ID instead
                        if (workFlow == "Signin")
                        {
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

                        }



                   else
                   {

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

                        if (installPath!= "")
                        {
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


                        if (workFlow == "Skip")
                        {
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

                        if (workFlow == "Signup")
                        {
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
                    }
                                 
                   Actions.WaitForElementVisible(Actions.GetDesktopWindow("QuickBooks Installation"), "Open QuickBooks", int.Parse(Sync_Timeout));
                   pointer = GetForegroundWindow();
                   sc.CaptureWindowToFile(pointer, resultsPath + "15_Open_QuickBooks.png", ImageFormat.Png);

                    // Click on Open QuickBooks
                    try
                    {
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

                catch (Exception e)
                {
                    Logger.logMessage("InstallQB " + targetPath + " - Failed");
                    Logger.logMessage(e.Message);
                    Logger.logMessage("------------------------------------------------------------------------------");
                    Logger.logMessage("------------------------------------------------------------------------------");

                }
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

        public static void CheckF2value(TestStack.White.Application qbApp, Window qbWindow)
        {
            //Actions.SelectMenu(qbApp, qbWindow, "File", "New Company...");
            Actions.SendF2ToWindow(qbWindow);
            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();
            string resultsPath = @"C:\Temp\Results\CheckF2_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "\\";
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
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "01_CheckF2.png", ImageFormat.Png);


            }
        }

        public static void SwitchEdition(TestStack.White.Application qbApp, Dictionary<String, String> dic, String exe)
        {
            String edistr;
            try
            {
                foreach (var pair in dic)
                {

                    if (qbApp.HasExited == true)
                    {
                        qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
                    }
                    TestStack.White.UIItems.WindowItems.Window qbWindow = null;
                    qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
                    String title = qbWindow.Title;

                    if (pair.Value.Equals(title))
                    {
                        continue;
                    }

                    else
                    {
                        

                        if (Actions.CheckWindowExists(qbWindow, "QuickBooks Update Service"))
                        {

                            Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks Update Service"), "Install Later");
                        }
                        if (Actions.CheckDesktopWindowExists("QuickBooks Update Service"))
                        {
                            Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Update Service"), "Install Later");
                        }

                        //if (Actions.DesktopInstance_CheckElementExistsByName("QuickBooks Update Service") == true)
                        //{ SendKeys.SendWait("%L"); }
                        // if (Actions.CheckWindowExists(Actions.GetDesktopWindow("Desktop"), "QuickBooks Update Service") == true)
                        // { SendKeys.SendWait("%L"); }
                        Thread.Sleep(1000);
                        if (Actions.CheckWindowExists(Actions.GetDesktopWindow("QuickBooks"), "Register QuickBooks") == true)
                        { SendKeys.SendWait("%L"); }
                        Thread.Sleep(1000);
                        qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
                        Actions.SelectMenu(qbApp, qbWindow, "Help", "Manage My License", "Change to a Different Industry Edition...");
                        Thread.Sleep(3000);
                        
                        Window editionWindow = Actions.GetChildWindow(qbWindow, "Select QuickBooks Industry-Specific Edition");
                        //if (Actions.CheckElementIsEnabled(editionWindow, pair.Key + " - Currently open  "))
                        if (pair.Key == "Enterprise Solutions General Business")
                        {
                            edistr = pair.Key + " - Currently open  ";

                        }
                        else edistr = pair.Key;

                        if (Actions.CheckElementIsEnabled(editionWindow, edistr))
                        {
                            Logger.logMessage(pair.Key + " - Currently open  ");
                            Actions.ClickElementByName(editionWindow, pair.Key);
                        }
                        else
                        {
                           // qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
                            Actions.ClickElementByName(editionWindow,"Cancel");
                            continue;
                        }
                    
                        Thread.Sleep(3000);

                        
                      
                        Actions.ClickElementByName(editionWindow, "Next >");


                        Window editionWindow1 = Actions.GetChildWindow(qbWindow, "Select QuickBooks Industry-Specific Edition");
                        Thread.Sleep(3000);
                        Actions.ClickElementByAutomationID(editionWindow1, "10004");
                        Thread.Sleep(30000);
                        //SendKeys.SendWait("Tab");
                        //SendKeys.SendWait("Enter");
                        //Thread.Sleep(20000);
                        //SendKeys.SendWait("%L");
                        //Thread.Sleep(10000);
                        //SendKeys.SendWait("%L");
                        //Thread.Sleep(30000);
                        Window win1 = Actions.GetDesktopWindow("Product Configuration");
                        Thread.Sleep(1000);
                        Actions.ClickElementByName(Actions.GetChildWindow(win1, "QuickBooks Product Configuration"), "No");
                        Thread.Sleep(30000);

                        if (Actions.CheckDesktopWindowExists("QuickBooks Update Service"))
                        {
                            SendKeys.SendWait("%L");
                            // Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Update Service"), "Install Later");
                        }

                        //  SendKeys.SendWait("%L"); 

                        Thread.Sleep(10000);
                        if (Actions.CheckWindowExists(Actions.GetDesktopWindow("QuickBooks"), "Register QuickBooks") == true)
                        { SendKeys.SendWait("%L"); }

                        Thread.Sleep(30000);



                    }

                }
            }
            catch (Exception e)
            {
                Logger.logMessage("failed" + e.GetBaseException());

            }
        }

        public static void ToggleEdition(TestStack.White.Application qbApp, Dictionary<String, String> dic, String exe)
        {
            try
            {

               

                foreach (var pair in dic)
                {
                    if (qbApp.HasExited == true)
                    {
                        qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
                    }
                    TestStack.White.UIItems.WindowItems.Window qbWindow = null;
                    qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
                    String title = qbWindow.Title;
                    if (pair.Value.Equals(title))
                    {
                        continue;
                    }

                    else
                    {
                        try
                        {
                            if (Actions.CheckWindowExists(qbWindow, "QuickBooks Update Service"))
                            {

                                Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks Update Service"), "Install Later");
                            }
                        }
                        catch(Exception e)
                        {
                            Logger.logMessage(e.ToString());
                        }
                        try
                        {
                            if (Actions.CheckDesktopWindowExists("QuickBooks Update Service"))
                            {
                                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Update Service"), "Install Later");
                            }
                        }
                        catch (Exception e)
                        {
                            Logger.logMessage(e.ToString());
                        }
                        //if (Actions.DesktopInstance_CheckElementExistsByName("QuickBooks Update Service") == true)
                        //{ SendKeys.SendWait("%L"); }
                        // if (Actions.CheckWindowExists(Actions.GetDesktopWindow("Desktop"), "QuickBooks Update Service") == true)
                        // { SendKeys.SendWait("%L"); }
                        Thread.Sleep(1000);
                        if (Actions.CheckWindowExists(Actions.GetDesktopWindow("QuickBooks"), "Register QuickBooks") == true)
                        { SendKeys.SendWait("%L"); }
                        Thread.Sleep(1000);
                        if (Actions.CheckWindowExists(Actions.GetDesktopWindow("QuickBooks"), "Set Up an External Accountant User") == true)
                        {
                            Window ExtAcctWin = Actions.GetChildWindow(Actions.GetDesktopWindow("QuickBooks"), "Set Up an External Accountant User");
                            Actions.ClickElementByName(ExtAcctWin,"No");
                         //   Actions.SendTABToWindow(qbWindow);
                           // Actions.SendENTERoWindow(qbWindow);

                        }
                        if (Actions.CheckWindowExists(Actions.GetDesktopWindow("QuickBooks"), "Accountant Center") == true)
                        {

                            Window AcctCenWin = Actions.GetChildWindow(Actions.GetDesktopWindow("QuickBooks"), "Accountant Center");
                            Actions.ClickElementByName(AcctCenWin,"Close");
                           // Actions.SendESCAPEToWindow(qbWindow);
                        }
                        Thread.Sleep(1000);

                        //Thread.Sleep(1000);
                        qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
                        Actions.SelectMenu(qbApp, qbWindow, "File", "Toggle to Another Edition... ");
                        Thread.Sleep(3000);

                        Window editionWindow = Actions.GetChildWindow(qbWindow, "Select QuickBooks Industry-Specific Edition");
                        Thread.Sleep(3000);

                        Actions.ClickElementByName(editionWindow, pair.Key);
                        Thread.Sleep(1000);
                        Actions.ClickElementByName(editionWindow, "Next >");


                        Window editionWindow1 = Actions.GetChildWindow(qbWindow, "Select QuickBooks Industry-Specific Edition");
                        Thread.Sleep(3000);
                        Actions.ClickElementByName(editionWindow1, "Toggle");
                        Thread.Sleep(30000);
                        //SendKeys.SendWait("Tab");
                        //SendKeys.SendWait("Enter");
                        //Thread.Sleep(20000);
                        //SendKeys.SendWait("%L");
                        //Thread.Sleep(10000);
                        //SendKeys.SendWait("%L");
                        //Thread.Sleep(30000);
                        //Window win1 = Actions.GetDesktopWindow("Product Configuration");
                        //Thread.Sleep(1000);
                        //Actions.ClickElementByName(Actions.GetChildWindow(win1, "QuickBooks Product Configuration"), "No");
                        Thread.Sleep(30000);

                        if (Actions.CheckWindowExists(Actions.GetDesktopWindow("QuickBooks"), "Automatic Backup") == true)
                        {
                            Logger.logMessage("Backup Window Found");
                            //Actions.SendTABToWindow(qbWindow);
                            //Actions.SendENTERoWindow(qbWindow);
                            SendKeys.SendWait("%N");

                        }
                        try
                        { 
                        if (Actions.CheckDesktopWindowExists("QuickBooks Update Service"))
                        {
                            // SendKeys.SendWait("%L");
                            Logger.logMessage("Update window found");
                            Window UpdateWin = Actions.GetDesktopWindow("QuickBooks Update Service");
                            Actions.SendTABToWindow(UpdateWin);
                            Actions.SendENTERoWindow(UpdateWin);
                        
                           // Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Update Service"), "Install Later");
                        }
                        else
                        {
                            SendKeys.SendWait("%L");
                        }
                        }
                        catch(Exception e)
                        {
                            Logger.logMessage(e.ToString());
                        }

                        try
                        {
                            if (Actions.CheckDesktopWindowExists("QuickBooks Product Configuration"))
                            {
                                Logger.logMessage("Update/Product window found");
                                Window ProdConfWin = Actions.GetDesktopWindow("QuickBooks Product Configuration");
                                Actions.SendTABToWindow(ProdConfWin);
                                Actions.SendENTERoWindow(ProdConfWin);
                                //Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Product Configuration"), "Install Later");

                            }

                            else
                            {
                                SendKeys.SendWait("%L");
                            }
                        }
                        catch (Exception e)
                        {
                             Logger.logMessage(e.ToString());

                        }

                        Thread.Sleep(20000);
                        if (Actions.CheckWindowExists(Actions.GetDesktopWindow("QuickBooks"), "Register QuickBooks") == true)
                        {
                            Logger.logMessage("Register window found");
                            Window registerWin = Actions.GetChildWindow(Actions.GetDesktopWindow("QuickBooks"), "Register QuickBooks");
                        
                            Logger.logMessage(registerWin.ToString());
                            Actions.ClickElementByName(registerWin, "Remind Me Later");

                            //SendKeys.SendWait("%L"); }

                            Thread.Sleep(30000);


                        }

                    }
                }

            }





            catch (Exception e)
            {
                Logger.logMessage("failed" + e.GetBaseException());
            }
        }
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
            wppath = wppath + sku + "\\qbwebpatch";
            Logger.logMessage("Function call @ :" + DateTime.Now);

            if(sku=="BEL")
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

// -------------------------Sunder Raj added for creating company file -----------------------------------------------------------------------------------------------------------
        public static void CreateCompanyFile(Dictionary<string, string> refkeyvaluepairdic)
        {
            var qbApp = QuickBooks.GetApp("QuickBooks");
            var qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            string bizName = null, industryList = null, industryType = null, businessType = null, address1 = null, address2 = null, state = null, city = null, country = null, taxid = null, phone = null, zip = null;

            Actions.SelectMenu(qbApp, qbWindow, "File", "New Company...");

            Actions.WaitForChildWindow(qbWindow, "QuickBooks Setup", 999999);

            if ((Actions.CheckElementExistsByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "btnExpressStart")) == true)
            {
                Logger.logMessage("Express Start button found and hence creating company file for Older version of QB");
                Actions.ClickElementByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "btnExpressStart");
            }
            else
            {
                Logger.logMessage("Start button found and hence creating company file for newer version of QB");
                Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "Start Setup");
            }

            //Enter Business Name
            if (refkeyvaluepairdic.ContainsKey("CompanyName"))
            {
                bizName = refkeyvaluepairdic["CompanyName"];
                Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "txtBox_BusinessName", bizName);
            }

            Window QBSetupWindow = Actions.GetChildWindow(qbWindow, "QuickBooks Setup");

            //Enter Industry Type 
            if (refkeyvaluepairdic.ContainsKey("IndustryList"))
            {
                industryList = refkeyvaluepairdic["IndustryList"];
                Actions.SetTextByAutomationID(QBSetupWindow, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Common.Objects.IndustryList_TxtField_AutoID, industryList);
            }
            if (refkeyvaluepairdic.ContainsKey("IndustryType"))
            {
                industryType = refkeyvaluepairdic["IndustryType"];
                Actions.SelectListBoxItemByText(QBSetupWindow, "lstBox_Industry", industryType);
            }
            if (refkeyvaluepairdic.ContainsKey("BusinessType"))
            {
                businessType = refkeyvaluepairdic["BusinessType"];
                Actions.SelectComboBoxItemByText(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "cmbBox_TaxStructure", businessType);

            }
            if (refkeyvaluepairdic.ContainsKey("TaxID"))
            {
                taxid = refkeyvaluepairdic["TaxID"];
                Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "txtBox_TaxID", taxid);
            }
            // Find if the company file consists of single page or 2 pages
            if ((Actions.CheckElementExistsByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "btn_Continue")) == true)
            {
                Logger.logMessage("Continue button found and hence creating company file for Older version of QB");
                Actions.ClickElementByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "btn_Continue");
            }

            //Enter Address 1
            if (refkeyvaluepairdic.ContainsKey("Address1"))
            {
                address1 = refkeyvaluepairdic["Address1"];
                Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "txtBoxAddress1", address1);
            }
            //Enter Address 2
            if (refkeyvaluepairdic.ContainsKey("Address2"))
            {
                address2 = refkeyvaluepairdic["Address2"];
                Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "txtBoxAddress2", address2);
            }

            //Enter City
            if (refkeyvaluepairdic.ContainsKey("City"))
            {
                city = refkeyvaluepairdic["City"];
                Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "txtBoxCity", city);
            }
            //Enter State
            if (refkeyvaluepairdic.ContainsKey("State"))
            {
                state = refkeyvaluepairdic["State"];
                Actions.SelectComboBoxItemByText(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "CmbBox_StateName", state);
            }
            //Enter Country
            if (refkeyvaluepairdic.ContainsKey("Country"))
            {
                country = refkeyvaluepairdic["Country"];
                Actions.SelectComboBoxItemByText(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "CmbBox_CountryName", country);
            }
            //Enter Zip Code
            if (refkeyvaluepairdic.ContainsKey("Zip"))
            {
                zip = refkeyvaluepairdic["Zip"];
                Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "txtBoxZip", zip);
            }
            //Enter Phone Number
            if (refkeyvaluepairdic.ContainsKey("PhoneNo"))
            {
                phone = refkeyvaluepairdic["PhoneNo"];
                Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "txtBoxPhone", phone);
            }

            Actions.ClickElementByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "btnCreateCompany");

            //Wait for the Marketing Page window
            qbWindow = Actions.GetAppWindow(qbApp, "QuickBooks");
            Actions.WaitForChildWindow(qbWindow, "QuickBooks Setup", 1000);
            var qbchild = Actions.GetChildWindow(qbWindow, "QuickBooks Setup");

            // Close the Marketing page window
            if ((Actions.CheckElementExistsByAutomationID(Actions.GetChildWindow(qbchild, "QuickBooks Setup"), "btnCreateCompany")) == true)
            {
                Actions.ClickElementByAutomationID(qbchild, "btnCreateCompany");

            }
            else

            Actions.ClickElementByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "btnCreateCompany");
            Actions.SelectMenu(qbApp, qbWindow, "Windows", "Close");


            String Title = qbWindow.Title;

            if (Title.Contains(bizName) == true)
            {
               Logger.logMessage("Company File Created successfully");
                
            }

            else
                Logger.logMessage("Company file creation failed");

    }
//------------------------End of code for creating company file -----------------------------------------------------------------------------------------------------------

 //------Code by Sunder Raj for Invoke QB-------------------------------------------

        public static void InvokeQB(Dictionary<string, string> refkeyvaluepairdic)

        {
            string qbwin = "Intuit QuickBooks Installer";
            string industryEdition = null;
            //if (Actions.CheckDesktopWindowExists(qbwin))
            //{
            //    Actions.ClickElementByName(Actions.GetDesktopWindow(qbwin), "Open QuickBooks");

                if( Actions.CheckDesktopWindowExists("Select QuickBooks Industry-Specific Edition")==true)
                {
                                Actions.SendTABToWindow(Actions.GetDesktopWindow("Select QuickBooks Industry-Specific Edition"));
                      if (refkeyvaluepairdic.ContainsKey("IndustryEdition"))
                        {
                          industryEdition = refkeyvaluepairdic["IndustryEdition"];
                          
                          Actions.ClickElementByName((Actions.GetDesktopWindow("Select QuickBooks Industry-Specific Edition")), industryEdition);
                          Actions.ClickElementByName(((Actions.GetDesktopWindow("Select QuickBooks Industry-Specific Edition"))), "Exit QuickBooks");

                        }
                }
                
            //}
            //else
            //    Logger.logMessage("Unable to Open QuickBooks");
        }
        
 //------Code by SUnder Raj for Invoke QB-------------------------------------------       
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
            //string readpath = "C:\\Temp\\Parameters.xlsx"; // "C:\\Installation\\Sample.txt";
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

        public static void PerformMIMO(TestStack.White.Application qbApp, Window qbWindow)
        {
            //Setting up the preferences 
            Actions.SelectMenu(qbApp, qbWindow, "Edit", "Preferences...");
            Window PerfWin = Actions.GetChildWindow(qbWindow,"Preferences");
            SendKeys.SendWait("{PGUP}");
            SendKeys.SendWait("{DOWN}");
            SendKeys.SendWait("{DOWN}");
            SendKeys.SendWait("{DOWN}");
            SendKeys.SendWait("{DOWN}");
            SendKeys.SendWait("{DOWN}");
            SendKeys.SendWait("{DOWN}");
            SendKeys.SendWait("{DOWN}");
            SendKeys.SendWait("{DOWN}");
            SendKeys.SendWait("%C");
            SendKeys.SendWait("%I");
            Actions.ClickElementByName(PerfWin,"OK");
          

           
           
            //Setting a new Customer
            Actions.SelectMenu(qbApp, qbWindow, "Customers", "Customer Center");
            Actions.WaitForChildWindow(qbWindow, "Customer Center", 9999);
            Window custcenWin = Actions.GetChildWindow(qbWindow, "Customer Center");
            if (custcenWin.IsCurrentlyActive)
            { 
            Actions.ClickElementByName(custcenWin, "New Customer && Job");
          
            SendKeys.SendWait("{DOWN}");
            SendKeys.SendWait("{ENTER}");
            }
            Window custWin = Actions.GetChildWindow(qbWindow, "New Customer");
            Random _r = new Random();
            custname = "Cust" + _r.Next(1000).ToString();
            Actions.SetTextByAutomationID(custWin,"1001",custname);
            Actions.ClickElementByName(custWin, "OK");
            Actions.CloseAllChildWindows(qbWindow);


            // Item is not created , create an item
            itemname = "item"+ _r.Next(1000).ToString();
            Actions.SelectMenu(qbApp, qbWindow, "Lists", "Item List");         
            Actions.SelectMenu(qbApp, qbWindow, "Edit", "New Item");
            Window itemWin = Actions.GetChildWindow(qbWindow, "New Item");
            Actions.SendTABToWindow(itemWin);
           
            Actions.SetTextOnElementByAutomationID(itemWin, "902",itemname);
            Actions.SendTABToWindow(itemWin);
            if(Actions.CheckElementExistsByName(itemWin,"Enable..."))
            {
                Actions.SendTABToWindow(itemWin);
            }
            Actions.SendTABToWindow(itemWin);
            Actions.SendTABToWindow(itemWin);
            Actions.SendTABToWindow(itemWin);
            Actions.SetTextOnElementByAutomationID(itemWin, "915", "200");
            Actions.SendTABToWindow(itemWin);
            Actions.SetTextOnElementByAutomationID(itemWin, "917", "Rent Expense");
            Actions.ClickElementByName(itemWin, "OK");
            Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");

            //Creating an invoice
            Actions.SelectMenu(qbApp,qbWindow,"Customers","Create Invoices");
            Window invWin = Actions.GetChildWindow(qbWindow, "Create Invoices");
            Actions.SetTextOnElementByAutomationID(invWin,"603",custname);
            Actions.SendTABToWindow(invWin);
            Actions.SendTABToWindow(invWin);
            Actions.SendTABToWindow(invWin);
            Actions.SendTABToWindow(invWin);
            Actions.SendTABToWindow(invWin);
            Actions.SendTABToWindow(invWin);
            Actions.SendTABToWindow(invWin);
            Actions.SendTABToWindow(invWin);
            Actions.SendTABToWindow(invWin);
            Actions.SendTABToWindow(invWin);
            Actions.SendTABToWindow(invWin);
            Actions.SendTABToWindow(invWin);
            Actions.SendTABToWindow(invWin);
            Actions.SetTextOnElementByAutomationID(invWin,"10","1");
            Actions.SendTABToWindow(invWin);
            Actions.SetTextOnElementByAutomationID(invWin,"1",itemname);
            Actions.SendTABToWindow(invWin);
           
            if(Actions.CheckWindowExists(qbWindow,"Warning"))
            {
                Window warWin = Actions.GetChildWindow(qbWindow, "Warning");
                Actions.ClickElementByName(warWin, "OK");
            }
            Actions.SendTABToWindow(invWin);
            Actions.ClickElementByName(invWin,"Save && Close");
            Thread.Sleep(1000);
            
            //Receive Payment

            Actions.SelectMenu(qbApp, qbWindow, "Customers", "Recieve Payments");
            Window rpayWin = Actions.GetChildWindow(qbWindow, "Recieve Payments");
            Actions.SetTextOnElementByAutomationID(rpayWin,"5603",custname);
            Actions.SendTABToWindow(rpayWin);
            Actions.SetTextOnElementByAutomationID(rpayWin,"5604","200");
            Actions.SendTABToWindow(rpayWin);
            Actions.SendTABToWindow(rpayWin);
            Actions.ClickElementByName(rpayWin, "CASH");
            Actions.ClickElementByName(rpayWin, "Auto Apply Payment");
            Actions.ClickButtonByName(qbWindow, "Save && Close");

            //Make a Deposit
            Actions.SelectMenu(qbApp, qbWindow, "Banking", "Make Deposits");

            if (qbWindow.Title.Equals("Need a Bank Account"))
            {
                Actions.ClickButtonByName(Actions.GetChildWindow(qbWindow, "Need a Bank Account"), "Yes");
                Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "Add New Account"), "136", "KTK");
                Actions.ClickButtonByName(Actions.GetChildWindow(qbWindow, "Add New Account"), "Save && Close");
            }
            else
                //Create Purhcase orders
                Actions.SelectMenu(qbApp, qbWindow, "Vendors", " Create Purchase Orders");
            Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "Create Purchase Orders"), "Armani", "603");


            // if Vendor wants to quick add the item
            if (qbWindow.Title.Equals("Venodr Not Found"))
            {
                Actions.ClickButtonByName(Actions.GetChildWindow(qbWindow, "Vendor Not Found"), "Quick Add");
            }
            else
                Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "Create Purchase Orders"), "Skoda", "1");
            Actions.SetTextByAutomationID(qbWindow, "603", "10");
            Actions.ClickButtonByName(qbWindow, "Save && Close");


            Actions.SelectMenu(qbApp, qbWindow, "Vendors", "Receive Items and Enter Bills");
            Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "Enter Bills"), "309", "Armani");

            if (qbWindow.Title.Equals("Open POs Exist"))
            {
                Actions.ClickButtonByName(Actions.GetChildWindow(qbWindow, "Open POs Exist"), "Yes");

            }

            if (qbWindow.Title.Equals("Open Purchase Orders"))
            {
                //Actions.UIA_SelectCheckBoxByName(window, qbWindow, "header Item");
                Actions.ClickButtonByName(Actions.GetChildWindow(qbWindow, "Open Purchase Orders"), "OK");


            }

            else
                Actions.ClickButtonByName(Actions.GetChildWindow(qbWindow, "Enter Bills"), "Save && close");

            //pay bills

            Actions.SelectMenu(qbApp, qbWindow, "vendors", "Pay Bills");

            //Clicking on check box

            Actions.ClickButtonByName(Actions.GetChildWindow(qbWindow, "Pay Bills"), "Pay Selected Bills");
            Actions.ClickButtonByName(Actions.GetChildWindow(qbWindow, "Assign Check Numbers"), " OK");

            if (qbWindow.Title.Equals("Payement Summary"))
            {
                Actions.ClickButtonByName(Actions.GetChildWindow(qbWindow, "Payment Summary"), "Done");

            }




        }
      
    }
}

   
