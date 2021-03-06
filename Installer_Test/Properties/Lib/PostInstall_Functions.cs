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

using Excel = Microsoft.Office.Interop.Excel;

//using Microsoft.VisualStudio.TestTools.UnitTesting;

using ScreenShotDemo;

using TestStack.BDDfy;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.InputDevices;
using Installer_Test.Properties.Lib;


namespace Installer_Test.Lib
{
   
    public class PostInstall_Functions
    {
        [DllImport("User32.dll")]
        public static extern int SetForegroundWindow(IntPtr point);
        [DllImport("User32.dll")]
        private static extern IntPtr GetForegroundWindow();

        public static Property conf = Property.GetPropertyInstance();
        public static string Sync_Timeout = conf.get("SyncTimeOut");
        public string line;
        public static string custname, vendorname, itemname, backuppath;
        public static Random _r = new Random();

        public static void CheckF2value(TestStack.White.Application qbApp, Window qbWindow, string resultsPath, string SKU)
        {
            Logger.logMessage("-------------------------------------------------------------");
            Logger.logMessage("Check F2 - Started");
            
            Actions.SendF2ToWindow(qbWindow);
            ScreenCapture sc = new ScreenCapture();
            System.Drawing.Image img = sc.CaptureScreen();
            IntPtr pointer = GetForegroundWindow();
            pointer = GetForegroundWindow();
            sc.CaptureWindowToFile(pointer, resultsPath + "CheckF2.png", ImageFormat.Png);
            Thread.Sleep(2000);
            Actions.ClickElementByName(Actions.GetChildWindow(Actions.GetDesktopWindow("QuickBooks " + SKU), "Product Information"), "OK");
            Logger.logMessage("Check F2 - Completed");
            Logger.logMessage("-------------------------------------------------------------");
        }
//      /////////////////////////////////////////////////////////////////////////////////////////////////////////
//// The following 2 Functions: SwitchEdition () and ToggleEdition () can be deleted /////////////////////////////////
//        public static void SwitchEdition(TestStack.White.Application qbApp, Dictionary<String, String> dic, String exe, String Bizname, String SearchText)
//        {
//            String edistr;
//            try
//            {
//                foreach (var pair in dic)
//                {

//                    if (qbApp.HasExited == true)
//                    {
//                        qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
//                    }
//                    TestStack.White.UIItems.WindowItems.Window qbWindow = null;
//                    qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
//                    String title = qbWindow.Title;

//                    if ((Bizname + pair.Value).Equals(title))
//                    {
                        
//                        continue;
//                    }

//                    else
//                    {


//                        if (Actions.CheckWindowExists(qbWindow, "QuickBooks Update Service"))
//                        {

//                            Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks Update Service"), "Install Later");
//                        }
//                        if (Actions.CheckDesktopWindowExists("QuickBooks Update Service"))
//                        {
//                            Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Update Service"), "Install Later");
//                        }

                      
//                        Thread.Sleep(1000);
//                        if (Actions.CheckWindowExists(Actions.GetDesktopWindow("QuickBooks"), "Register QuickBooks") == true)
//                        { SendKeys.SendWait("%L"); }
//                        Thread.Sleep(1000);
//                        qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
//                        Actions.SelectMenu(qbApp, qbWindow, "Help", "Manage My License", "Change to a Different Industry Edition...");
//                        Thread.Sleep(3000);

//                        Window editionWindow = Actions.GetChildWindow(qbWindow, "Select QuickBooks Industry-Specific Edition");
//                        //if (Actions.CheckElementIsEnabled(editionWindow, pair.Key + " - Currently open  "))
//                        if (pair.Key == "Enterprise Solutions General Business" || pair.Key == "Premier Edition (General Business)")
//                        {
//                            edistr = pair.Key + " - Currently open  ";

//                        }
//                        else edistr = pair.Key;

//                        if (Actions.CheckElementIsEnabled(editionWindow, edistr))
//                        {

//                            Logger.logMessage(pair.Key + " - Currently open  ");
//                            Actions.ClickElementByName(editionWindow, pair.Key);
//                        }
//                        else
//                        {
//                            // qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
//                            Actions.ClickElementByName(editionWindow, "Cancel");
//                            continue;
//                        }

//                        Thread.Sleep(3000);



//                        Actions.ClickElementByName(editionWindow, "Next >");


//                        Window editionWindow1 = Actions.GetChildWindow(qbWindow, "Select QuickBooks Industry-Specific Edition");
//                        Thread.Sleep(3000);
//                        Actions.ClickElementByAutomationID(editionWindow1, "10004");
//                        Thread.Sleep(500);
//                        try
//                        {
                           
//                            var x = Actions.GetDesktopWindow(Bizname + SearchText);
//                            var t = x.ModalWindows();


//                            if (Actions.CheckWindowExists(x, "Automatic Backup"))
//                            {
//                                Actions.ClickElementByName(Actions.GetChildWindow(x, "Automatic Backup"), "No");
//                               // SendKeys.SendWait("%N");
//                            }
                            
//                        }
//                        catch (Exception e)
//                        {
//                            Logger.logMessage("failed" + e.GetBaseException());

//                        }
//                        Thread.Sleep(30000);


//                        Window win1 = Actions.GetDesktopWindow("Product Configuration");
//                        Thread.Sleep(1000);
//                        Actions.ClickElementByName(Actions.GetChildWindow(win1, "QuickBooks Product Configuration"), "No");
//                        Thread.Sleep(30000);

//                        if (Actions.CheckDesktopWindowExists("QuickBooks Update Service"))
//                        {
//                            SendKeys.SendWait("%L");
//                            // Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Update Service"), "Install Later");
//                        }

//                        //  SendKeys.SendWait("%L"); 

//                        Thread.Sleep(10000);
//                        if (Actions.CheckWindowExists(Actions.GetDesktopWindow("QuickBooks"), "Register QuickBooks") == true)
//                        { SendKeys.SendWait("%L"); }

//                        Thread.Sleep(30000);



//                    }

//                }
//            }
//            catch (Exception e)
//            {
//                Logger.logMessage("failed" + e.GetBaseException());

//            }
//        }

//        public static void ToggleEdition(TestStack.White.Application qbApp, Dictionary<String, String> dic, String exe, String Bizname)
//        {
//            try
//            {

//                foreach (var pair in dic)
//                {
//                    if (qbApp.HasExited == true)
//                    {
//                        qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
//                    }
//                    TestStack.White.UIItems.WindowItems.Window qbWindow = null;
//                    qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
//                    String title = qbWindow.Title;
//                    if ((Bizname + pair.Value).Equals(title))
//                    {
//                        continue;
//                    }

//                    else
//                    {
//                        try
//                        {
//                            if (Actions.CheckWindowExists(qbWindow, "QuickBooks Update Service"))
//                            {

//                                Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks Update Service"), "Install Later");
//                            }
//                        }
//                        catch (Exception e)
//                        {
//                            Logger.logMessage(e.ToString());
//                        }
//                        try
//                        {
//                            if (Actions.CheckDesktopWindowExists("QuickBooks Update Service"))
//                            {
//                                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Update Service"), "Install Later");
//                            }
//                        }
//                        catch (Exception e)
//                        {
//                            Logger.logMessage(e.ToString());
//                        }
//                        try
//                        {
//                            //  Actions.WaitForChildWindow(Actions.GetDesktopWindow("QuickBooks"), "Register QuickBooks", int.Parse(Sync_Timeout));
//                            if (Actions.CheckWindowExists(Actions.GetDesktopWindow("QuickBooks"), "Register QuickBooks") == true)
//                            {
//                                SendKeys.SendWait("%L");
//                            }
//                        }
//                        catch (Exception e)
//                        {
//                            Logger.logMessage(e.ToString());
//                        }

//                        try
//                        {

//                            if (Actions.CheckWindowExists(Actions.GetDesktopWindow("QuickBooks"), "Set Up an External Accountant User") == true)
//                            {
//                                Window ExtAcctWin = Actions.GetChildWindow(Actions.GetDesktopWindow("QuickBooks"), "Set Up an External Accountant User");
//                                Actions.ClickElementByName(ExtAcctWin, "No");
//                            }
//                        }
//                        catch (Exception e)
//                        {
//                            Logger.logMessage(e.ToString());
//                        }
//                        try
//                        {

//                            if (Actions.CheckWindowExists(Actions.GetDesktopWindow("QuickBooks"), "Accountant Center") == true)
//                            {

//                                Window AcctCenWin = Actions.GetChildWindow(Actions.GetDesktopWindow("QuickBooks"), "Accountant Center");
//                                Actions.ClickElementByName(AcctCenWin, "Close");

//                            }
//                        }
//                        catch (Exception e)
//                        {
//                            Logger.logMessage(e.ToString());
//                        }



//                        qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
//                        Actions.SelectMenu(qbApp, qbWindow, "File", "Toggle to Another Edition... ");


//                        Window editionWindow = Actions.GetChildWindow(qbWindow, "Select QuickBooks Industry-Specific Edition");


//                        Actions.ClickElementByName(editionWindow, pair.Key);

//                        Actions.ClickElementByName(editionWindow, "Next >");


//                        Window editionWindow1 = Actions.GetChildWindow(qbWindow, "Select QuickBooks Industry-Specific Edition");

//                        Actions.ClickElementByName(editionWindow1, "Toggle");

//                        Thread.Sleep(30000);

//                        if (Actions.CheckWindowExists(Actions.GetDesktopWindow("QuickBooks"), "Automatic Backup") == true)
//                        {
//                            Logger.logMessage("Backup Window Found");

//                            SendKeys.SendWait("%N");

//                        }
//                        try
//                        {
//                            //Actions.WaitForWindow("QuickBooks Update Service",int.Parse(Sync_Timeout));
//                            if (Actions.CheckDesktopWindowExists("QuickBooks Update Service"))
//                            {

//                                Logger.logMessage("Update window found");
//                                Window UpdateWin = Actions.GetDesktopWindow("QuickBooks Update Service");
//                                SendKeys.SendWait("%L");
//                                //Actions.SendTABToWindow(UpdateWin);
//                                //Actions.SendENTERoWindow(UpdateWin);


//                            }
//                            else
//                            {
//                                SendKeys.SendWait("%L");
//                            }
//                        }
//                        catch (Exception e)
//                        {
//                            Logger.logMessage(e.ToString());
//                        }

//                        try
//                        {
//                            if (Actions.CheckDesktopWindowExists("QuickBooks Product Configuration"))
//                            {
//                                Logger.logMessage("Update/Product window found");
//                                Window ProdConfWin = Actions.GetDesktopWindow("QuickBooks Product Configuration");
//                                SendKeys.SendWait("%L");
//                                //Actions.SendTABToWindow(ProdConfWin);
//                                //Actions.SendENTERoWindow(ProdConfWin);
//                            }

//                            else
//                            {
//                                SendKeys.SendWait("%L");
//                            }
//                        }
//                        catch (Exception e)
//                        {
//                            Logger.logMessage(e.ToString());

//                        }

//                        Thread.Sleep(20000);
//                        if (Actions.CheckWindowExists(Actions.GetDesktopWindow("QuickBooks"), "Register QuickBooks") == true)
//                        {
//                            Logger.logMessage("Register window found");
//                            Window registerWin = Actions.GetChildWindow(Actions.GetDesktopWindow("QuickBooks"), "Register QuickBooks");

//                            Logger.logMessage(registerWin.ToString());
//                            Actions.ClickElementByName(registerWin, "Remind Me Later");






//                        }
//                        Thread.Sleep(30000);
//                    }
//                }

//            }





//            catch (Exception e)
//            {
//                Logger.logMessage("failed" + e.GetBaseException());
//            }
//        }
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // -------------------------Sunder Raj added for creating company file -----------------------------------------------------------------------------------------------------------
        public static void CreateCompanyFile(Dictionary<string, string> refkeyvaluepairdic)
        {
            var qbApp = QuickBooks.GetApp("QuickBooks");
            var qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");

            string bizName = null, industryList = null, industryType = null, businessType = null, address1 = null, address2 = null, state = null, city = null, country = null, taxid = null, phone = null, zip = null;
            string emailid = null, existpassword = null, newpwd = null, confirmpass = null, fname = null, lname = null;

            Logger.logMessage("--------------------------------------------------------------------");
            Logger.logMessage("Create Company File - Started");
            Logger.logMessage("--------------------------------------------------------------------");

            Actions.SelectMenu(qbApp, qbWindow, "File", "New Company...");
            Actions.WaitForChildWindow(qbWindow, "QuickBooks Setup", 999999);

            if ((Actions.CheckElementExistsByName(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "Express Start")) == true)
            {
                Logger.logMessage("Express Start button found and hence creating company file for Older version of QB");
                Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "Express Start");
            }
            else
            {
                Logger.logMessage("Start button found and hence creating company file for newer version of QB");
                Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "Start Setup");
            }

            //Enter Email & password and perform a sign in 

            if ((Actions.CheckElementExistsByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "txt_LoginEmail")) == true)
            {
                if (refkeyvaluepairdic.ContainsKey("EmailID"))
                {
                    emailid = refkeyvaluepairdic["EmailID"];
                    Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "txt_LoginEmail", emailid);
                }
                Actions.ClickElementByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "btn_Next");
            }
            if ((Actions.CheckElementExistsByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "lbl_forgetPwd")) == true)
            {
                if (refkeyvaluepairdic.ContainsKey("ExistPass"))
                {
                    existpassword = refkeyvaluepairdic["ExistPass"];
                    Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "pwd_extUsrPwd", existpassword);
                    if (Actions.CheckElementExistsByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "btn_Continue") == true)
                    {
                        Actions.ClickElementByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "btn_Continue");
                    }

                    if (Actions.CheckElementExistsByName(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "Continue") == true)
                    {
                        Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "Continue");
                        Thread.Sleep(2000);
                    }

                    else
                    {
                        Actions.SendTABToWindow(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"));
                        Actions.SendTABToWindow(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"));
                        Actions.SendENTERoWindow(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"));
                        Thread.Sleep(500);
                    }
                }
            }
           //Enter Email & Sign Up with New email and password
            if ((Actions.CheckElementExistsByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "pwd_NewPwd")) == true)
                {
                    if (refkeyvaluepairdic.ContainsKey("NewPass"))
                    {
                        newpwd = refkeyvaluepairdic["NewPass"];
                        Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "pwd_NewPwd", newpwd);

                    }
                    if (refkeyvaluepairdic.ContainsKey("ConfirmPass"))
                    {
                        confirmpass = refkeyvaluepairdic["ConfirmPass"];
                        Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "pwd_NewConfirm", confirmpass);
                    }
                    if (refkeyvaluepairdic.ContainsKey("FirstName"))
                    {
                        fname = refkeyvaluepairdic["FirstName"];
                        Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "txt_FirstName", fname);
                    }
                    if (refkeyvaluepairdic.ContainsKey("LastName"))
                    {
                        lname = refkeyvaluepairdic["LastName"];
                        Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "txt_LastName", lname);
                        //Actions.WaitForElementEnabled(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "Continue", 1000);
                       // Actions.ClickButtonByName(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "Continue");
                        if (Actions.CheckElementExistsByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "btn_Continue") == true)
                        {
                            Actions.ClickElementByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "btn_Continue");
                        }
                        else
                        {
                            Actions.SendTABToWindow(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"));
                            Actions.SendTABToWindow(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"));
                            Actions.SendENTERoWindow(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"));
                            Thread.Sleep(500);
                        }
                    }

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
            if ((Actions.CheckElementExistsByName(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"),"Continue")) == true)
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
            if (Actions.CheckElementExistsByName(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "Create Company") == true)
            {
                Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "Create Company");
            }
            else
            {
                Actions.ClickElementByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "btnCreateCompany");
            }
            String Title = qbWindow.Title;

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //Actions.ClickElementByAutomationID(Actions.GetChildWindow(qbWindow, "QuickSetup"), "1");

            //Actions.ClickButtonByAutomationID(Actions.GetChildWindow(qbWindow, "Confirm Save As"), "CommandButton_6");

            if (Title.Contains(bizName) == true)
            {
                Logger.logMessage("--------------------------------------------------------------------");
                Logger.logMessage("Company File creation - Successful");
                Logger.logMessage("--------------------------------------------------------------------");
            }

            else
            {
                Logger.logMessage("--------------------------------------------------------------------");
                Logger.logMessage("Company File creation - Failed");
                Logger.logMessage("--------------------------------------------------------------------");
            }
            //Wait for the Marketing Page window
            qbWindow = Actions.GetAppWindow(qbApp, "QuickBooks");
            Actions.WaitForChildWindow(qbWindow, "QuickBooks Setup", 10000);
           var qbchild = Actions.GetChildWindow(qbWindow, "QuickBooks Setup");

            // Close the Marketing page window
            if ((Actions.CheckElementExistsByAutomationID(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "btnCreateCompany")) == true)
            {
                Actions.ClickElementByName(qbchild, "Start Working");

                Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
               // Actions.SendENTERoWindow(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"));
                

            }
            else
            {
                Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "Start Working");
                Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
                // Actions.SendENTERoWindow(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"));
            }
          

        }
          
        //------------------------End of code for creating company file -----------------------------------------------------------------------------------------------------------

        public static void PerformMIMO(TestStack.White.Application qbApp, Window qbWindow)
        {
            Logger.logMessage("----------------------------------------------------");
            Logger.logMessage("Money In Money Out - Started");
            Logger.logMessage("----------------------------------------------------");

            //Setting up the preferences 
            QB_functions.Reset_Preferences(qbApp, qbWindow);

            //Setting a new Customer
            custname = "Cust" + _r.Next(1000).ToString();
            QB_functions.Create_Customer(qbApp, qbWindow, custname);

            // Item is not created , create an item
            itemname = "item" + _r.Next(1000).ToString();
            QB_functions.Create_Item(qbApp, qbWindow, itemname);

            //Creating an invoice
            QB_functions.Create_Invoice(qbApp, qbWindow, custname, itemname);

            //Receive Payment
            QB_functions.Receive_Payments(qbApp, qbWindow, custname);

            //Setting a new Vendor
            vendorname = "Vend" + _r.Next(1000).ToString();
            QB_functions.Create_Vendor(qbApp, qbWindow, vendorname);

            //Create Purhcase orders
            QB_functions.Create_Purchase_Order(qbApp, qbWindow, custname, itemname, vendorname);
            
            //// Creating a bill
            //QB_functions.Create_Bill(qbApp, qbWindow, vendorname, itemname);
           
            ////Pay Bills
            //QB_functions.Pay_Bill(qbApp,qbWindow);

            // Close all windows before resetting preferences
            Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");
            Thread.Sleep(1000);

            //Reseting the preferences 
            QB_functions.Reset_Preferences(qbApp,qbWindow);
            
            Logger.logMessage("Money In Money Out - Successful");
            Logger.logMessage("----------------------------------------------------");
        }

        public static void PerformVerify(TestStack.White.Application qbApp, Window qbWindow)
        {
            Logger.logMessage("----------------------------------------------------");
            Logger.logMessage("Verify Data - Started");
            Logger.logMessage("----------------------------------------------------");

            // Invoking Verify Data from File-> Utility
            Actions.SelectMenu(qbApp, qbWindow, "File", "Utilities", "Verify Data");

            try
            {
                if (Actions.CheckWindowExists(qbWindow, "Verify Data"))
                {
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Verify Data"), "OK");
                    Logger.logMessage("Click on File -> Utilities -> Verify Data - Successful");
                    Logger.logMessage("------------------------------------------------------------------");
                }
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on File -> Utilities -> Verify Data - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("----------------------------------------------------------------------");
            }

            try
            {
                Actions.WaitForChildWindow(qbWindow, "QuickBooks Information", int.Parse(Sync_Timeout));
                if (Actions.CheckWindowExists(qbWindow, "QuickBooks Information"))
                {
                  Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks Information"), "OK");

                  Logger.logMessage("Verify Data - Successful");
                  Logger.logMessage("--------------------------------------------------------------------");
                }
            }
            catch (Exception e)
            {
                Logger.logMessage("Verify Data - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("--------------------------------------------------------------------");
            }
        }

        public static void ApplyWebpatch(String resultsPath)
        {
        
        
        string testName = "WebPatch";
       

        string readpath = "C:\\Temp\\Parameters.xlsm";
        string targetPath, sku;
        var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
        Logger log = new Logger(testName + "_" + timeStamp);
        Dictionary<string, string> dic = new Dictionary<string, string>();
        dic = Lib.File_Functions.ReadExcelValues(readpath, "WebPatch", "B2:B12");
        sku = dic["B7"];
        targetPath = dic["B11"];

        try
        {
            OSOperations.KillProcess("qbw32");
            Logger.logMessage("QuickBooks process killed successfully");
        }
        catch (Exception e)
        {
            Logger.logMessage("Unable to Kill process QBW32" + e.ToString());
        }

        ScreenCapture sc = new ScreenCapture();
        System.Drawing.Image img = sc.CaptureScreen();
        IntPtr pointer = GetForegroundWindow();

        if (sku == "Enterprise" || sku == "Enterprise Accountant")
            OSOperations.InvokeInstaller(targetPath, "en_qbwebpatch.exe");
        else
            OSOperations.InvokeInstaller(targetPath, "qbwebpatch.exe");


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
                Actions.WaitForWindow("QuickBooks Update,Version", 60000);
                Window patchWin1 = Actions.GetDesktopWindow("QuickBooks Update,Version");
                Window updatecomp = Actions.GetChildWindow(patchWin1, "Update complete");
                pointer = GetForegroundWindow();
                sc.CaptureWindowToFile(pointer, resultsPath + "Patch_Applied_Succes.png", ImageFormat.Png);
                Actions.ClickElementByName(updatecomp, "OK");
                Logger.logMessage("Patch Applied Successfully");
            }
        }
        catch (Exception e)
        {
            Logger.logMessage("Patch Failed" + e.ToString());

        }

        }

        public static void PerformRebuild(TestStack.White.Application qbApp, Window qbWindow)
        {
            backuppath = "C:\\Test\\";

            if (!Directory.Exists(backuppath))
            {
                Directory.CreateDirectory(backuppath);
            }

            Actions.SelectMenu(qbApp, qbWindow, "File", "Utilities", "Rebuild Data");

            try
            {
                if (Actions.CheckWindowExists(qbWindow, "QuickBooks Information"))
                {
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks Information"), "OK");
                    Logger.logMessage("Click on QuickBooks Information -> OK - Successful");
                }
            }
            catch (Exception e)
            {
                Logger.logMessage("Click on QuickBooks Information -> OK - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("----------------------------------------------------------------------");
            }
            try
            {
                if (Actions.CheckWindowExists(qbWindow, "Create Backup"))
                {

                    Window bckupWin = Actions.GetChildWindow(qbWindow, "Create Backup");
                    Logger.logMessage("New Backup Window");
                    Actions.ClickElementByName(bckupWin, "Next");
                    Actions.WaitForChildWindow(qbWindow, "Backup Options", int.Parse(Sync_Timeout));
                    if (Actions.CheckWindowExists(qbWindow, "Backup Options"))
                    {

                        Window newbckupWin = Actions.GetChildWindow(qbWindow, "Backup Options");
                        Logger.logMessage("Backup Window to provide the backup path");
                        Actions.SetTextByAutomationID(newbckupWin, "2002", backuppath);
                        Actions.ClickElementByName(newbckupWin, "OK");

                        if (Actions.CheckWindowExists(newbckupWin, "QuickBooks"))
                        {
                            Actions.ClickElementByName(Actions.GetChildWindow(newbckupWin, "QuickBooks"), "Use this Location");
                        }

                        if (Actions.CheckWindowExists(bckupWin, "Save Backup Copy"))
                        {
                            Actions.ClickElementByName(Actions.GetChildWindow(bckupWin, "Save Backup Copy"), "Save");
                            Actions.WaitForChildWindow(qbWindow, "QuickBooks Information", int.Parse(Sync_Timeout));
                            Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks Information"), "OK");
                        }


                    }
                }

                else
                {
                    Window bckupWin1 = Actions.GetChildWindow(qbWindow, "Save Backup Copy");
                    Actions.ClickElementByName(bckupWin1, "Save");
                    Actions.WaitForChildWindow(qbWindow, "QuickBooks Information", int.Parse(Sync_Timeout));
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks Information"), "OK");
                }

                Logger.logMessage("Rebuild Data - Successful");
                Logger.logMessage("--------------------------------------------------------------------");
            }

            catch (Exception e)
            {
                Logger.logMessage("Rebuild Data - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("--------------------------------------------------------------------");
            }

        }
        
        }

       
    }
