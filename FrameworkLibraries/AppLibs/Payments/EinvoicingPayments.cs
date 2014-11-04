using FrameworkLibraries.ActionLibs.WhiteAPI;
using FrameworkLibraries.Utils;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using TestStack.White;
using TestStack.White.UIItems.WindowItems;


namespace FrameworkLibraries.AppLibs.Payments
{
    public class EinvoicingPayments
    {
        public static Property conf = Property.GetPropertyInstance();
        public static string Execution_Speed = conf.get("ExecutionSpeed");
        public static bool bankTransfer, creditCard;

        public static void CreateCustomer(Application quickbooksApp, Window quickbooksWindow, string customer, string companyName, string firstName, string lastName, string jobTitle, string mainPhone, string mainEmail)
        {
            Actions.SelectMenu(quickbooksApp, quickbooksWindow, "Customers", "Customer Center"); //Open Customer Center
            Actions.SendCTRL_KeyToWindow(quickbooksWindow, "n"); //Create New Customer
            Actions.SendKeysToWindow(quickbooksWindow, customer); //Entering Customer Name
            Actions.SendTABToWindow(quickbooksWindow); //Opening
            Actions.SendTABToWindow(quickbooksWindow); //As of
            Actions.SendTABToWindow(quickbooksWindow);
            Actions.SendKeysToWindow(quickbooksWindow, companyName); //Entering Company Name
            Actions.SendTABToWindow(quickbooksWindow);
            Actions.SendKeysToWindow(quickbooksWindow, "Mr."); //Entering Initials
            Actions.SendTABToWindow(quickbooksWindow);
            Actions.SendKeysToWindow(quickbooksWindow, firstName); //Entering First Name
            Actions.SendTABToWindow(quickbooksWindow);
            Actions.SendKeysToWindow(quickbooksWindow, "MI"); //Entering Initials
            Actions.SendTABToWindow(quickbooksWindow);
            Actions.SendKeysToWindow(quickbooksWindow, lastName); //Entering Last Name
            Actions.SendTABToWindow(quickbooksWindow);
            Actions.SendKeysToWindow(quickbooksWindow, jobTitle); //Entering Job Title
            Actions.SendTABToWindow(quickbooksWindow);
            Actions.SendTABToWindow(quickbooksWindow);
            Actions.SendKeysToWindow(quickbooksWindow, mainPhone); //Entering Main Phone
            Actions.SendTABToWindow(quickbooksWindow);
            Actions.SendTABToWindow(quickbooksWindow);
            Actions.SendKeysToWindow(quickbooksWindow, mainEmail); //Entering Main Email
            Actions.SendENTERoWindow(quickbooksWindow); //Closing Customer Window  
            Assert.IsTrue(Actions.CheckWindowExists(quickbooksWindow, "Customer Center: " + customer));
            Logger.logMessage("New Customer has been created");
        }

        public static void CreateInvoiceWithEmailLaterChecked(Application quickbooksApp, Window quickbooksWindow, string customer, string itemName)
        {
            Actions.SelectMenu(quickbooksApp, quickbooksWindow, "Customers", "Create Invoices"); //Open Create Invoice
            Actions.SendTABToWindow(quickbooksWindow);
            bool emailLater = Actions.CheckCheckBoxIsSelected(quickbooksWindow, "EmailLaterChk");
            Assert.IsTrue(emailLater); //Ensuring if Email Later is checked.
            Logger.logMessage("Email Later gets checked");
            for (int i = 0; i < 6; i++)
                Actions.SendTABToWindow(quickbooksWindow);
            Actions.SendKeysToWindow(quickbooksWindow, itemName); //Enter Item Name
            Actions.ClickElementByName(quickbooksWindow, "Save && Close"); //Saving and Closing the Invoice
            if (Actions.CheckWindowExists(quickbooksWindow, "Check Spelling on Form")) //Check for Spell checker for Item Description
                Actions.UIA_ClickItemByName(Actions.UIA_GetAppWindow(quickbooksWindow.Title), Actions.GetChildWindow(quickbooksWindow, "Create Invoices"), "Ignore All"); //Clicking Ignore All button
            if (Actions.CheckWindowExists(quickbooksWindow, "Recording Transaction"))
                Actions.ClickElementByName(quickbooksWindow, "No");
            Thread.Sleep(1000);
            Actions.SendESCAPEToWindow(quickbooksWindow);
            Logger.logMessage("Invoice is created and Email Later was checked");
        }

        public static void OpenSendFormAndSendMail(Application quickbooksApp, Window quickbooksWindow)
        {
            Actions.SelectMenu(quickbooksApp, quickbooksWindow, "File", "Send Forms..."); //Open Send Forms
            Thread.Sleep(int.Parse(Execution_Speed));
            Actions.SendALT_KeyToWindow(quickbooksWindow, "s"); //Clicking "Send Now" Button
            Thread.Sleep(int.Parse(Execution_Speed));
            if (Actions.CheckWindowExists(quickbooksWindow, "Provide Email Information")) //Check if Send Invoice prompts for passowrd
            {
                Actions.UIA_ClickTextByAutomationID(Actions.UIA_GetAppWindow(quickbooksWindow.Title), Actions.GetChildWindow(quickbooksWindow, "Provide Email Information"), "29794");
                Actions.SetTextOnElementByAutomationID(quickbooksWindow, "29794", "Intuit01"); 
                Actions.SendENTERoWindow(quickbooksWindow);
            }
            Actions.WaitForAnyChildWindow(quickbooksWindow, "QuickBooks Information", int.Parse(Execution_Speed));
            if (Actions.CheckWindowExists(quickbooksWindow, "QuickBooks Information")) //Check if Mail has sent successfully
                Actions.SendENTERoWindow(quickbooksWindow);
            Thread.Sleep(1000);
            Logger.logMessage("Mail has sent successfully");
            Actions.CloseAllChildWindows(quickbooksWindow);
        }

        public static void SendInvoiceWebMail(Application quickbooksApp, Window quickbooksWindow, string custName, string itemName)
        {
            Actions.SelectMenu(quickbooksApp, quickbooksWindow, "Customers", "Create Invoices"); //Open Create Invoice
            Actions.SendKeysToWindow(quickbooksWindow, custName); //Entering Customer Name
            for (int i = 0; i < 7; i++)
                Actions.SendTABToWindow(quickbooksWindow);
            Actions.SendKeysToWindow(quickbooksWindow, itemName); //Enter Item Name
            Actions.UIA_ClickItemByAutomationID(Actions.UIA_GetAppWindow(quickbooksWindow.Title), Actions.GetChildWindow(quickbooksWindow, "Create Invoices"), "EmailMenu");
            Actions.UIA_ClickItemByAutomationID(Actions.UIA_GetAppWindow(quickbooksWindow.Title), Actions.GetChildWindow(quickbooksWindow, "Create Invoices"), "EmailBtn"); //Click Email Invoice
            if (Actions.CheckWindowExists(quickbooksWindow, "Check Spelling on Form")) //Check for Spell checker for Item Description
                Actions.UIA_ClickItemByName(Actions.UIA_GetAppWindow(quickbooksWindow.Title), Actions.GetChildWindow(quickbooksWindow, "Create Invoices"), "Ignore All");
            Actions.GetChildWindow(quickbooksWindow, "Send Invoice"); 
            bankTransfer = Actions.CheckElementExistsByName(quickbooksWindow, "Bank Transfer");
            creditCard = Actions.CheckElementExistsByName(quickbooksWindow, "Credit Card");
            Actions.UIA_ClickItemByName(Actions.UIA_GetAppWindow(quickbooksWindow.Title), Actions.GetChildWindow(quickbooksWindow, "Send Invoice"), "Send"); //Clicking Send Button from Send Invoice
            if (Actions.CheckWindowExists(quickbooksWindow, "Provide Email Information")) //Check if Send Invoice prompts for passowrd
            {
                Actions.UIA_ClickTextByAutomationID(Actions.UIA_GetAppWindow(quickbooksWindow.Title), Actions.GetChildWindow(quickbooksWindow, "Provide Email Information"), "29794");
                Actions.SetTextOnElementByAutomationID(quickbooksWindow, "29794", "Intuit01");
                Actions.SendENTERoWindow(quickbooksWindow);
            }
            Actions.GetChildWindow(quickbooksWindow, "QuickBooks Information");
            if (Actions.CheckWindowExists(quickbooksWindow, "QuickBooks Information")) //Check if Mail has sent successfully
                Actions.SendENTERoWindow(quickbooksWindow);
            Logger.logMessage("Mail has sent successfully");
            Actions.CloseAllChildWindows(quickbooksWindow);   
        }

        public static void VoidInvoice(Application quickbooksApp, Window quickbooksWindow)
        {
            Actions.SelectMenu(quickbooksApp, quickbooksWindow, "Customers", "Create Invoices");
            Actions.SendALT_KeyToWindow(quickbooksWindow, "p");
            Actions.SelectMenu(quickbooksApp, quickbooksWindow, "Edit", "Void Invoice");
            Actions.SendALT_KeyToWindow(quickbooksWindow, "a");
            if (Actions.CheckWindowExists(quickbooksWindow, "Recording Transaction"))
                Actions.SendALT_KeyToWindow(quickbooksWindow, "y");
        }

    }
}
