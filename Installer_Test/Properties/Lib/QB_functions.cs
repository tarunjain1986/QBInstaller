using System;
//using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.White.UIItems.WindowItems;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using System.Windows.Forms;
using FrameworkLibraries.Utils;
using System.Threading;

namespace Installer_Test.Properties.Lib
{
   
    public class QB_functions
    {

        public static Property conf = Property.GetPropertyInstance();
        public static string Sync_Timeout = conf.get("SyncTimeOut");

        public static void Create_Customer(TestStack.White.Application qbApp, Window qbWindow, String Customer)
        {
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
            
            Actions.SetTextByAutomationID(custWin, "1001", Customer);
            Actions.ClickElementByName(custWin, "OK");
            Actions.CloseAllChildWindows(qbWindow);

            Logger.logMessage("Customer Created Successfully.");
        }

         public static void Create_Item(TestStack.White.Application qbApp, Window qbWindow, String Item)
        {
            Actions.SelectMenu(qbApp, qbWindow, "Lists", "Item List");
            Actions.SelectMenu(qbApp, qbWindow, "Edit", "New Item");
            Window itemWin = Actions.GetChildWindow(qbWindow, "New Item");
            Actions.SendTABToWindow(itemWin);

            Actions.SetTextOnElementByAutomationID(itemWin, "902", Item);
            Actions.SendTABToWindow(itemWin);
            if (Actions.CheckElementExistsByName(itemWin, "Enable..."))
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

            Logger.logMessage("Item Created Successfully.");
        }

         public static void Create_Invoice(TestStack.White.Application qbApp, Window qbWindow, String Customer,String Item)
         {
             Actions.SelectMenu(qbApp, qbWindow, "Customers", "Create Invoices");
             Window invWin = Actions.GetChildWindow(qbWindow, "Create Invoices");
             Actions.ClickElementByName(invWin, "Maximize");
             Actions.SetTextByAutomationID(invWin, "603", Customer);
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
             Actions.SetTextOnElementByAutomationID(invWin, "10", "1");
             Actions.SendTABToWindow(invWin);
             Actions.SetTextOnElementByAutomationID(invWin, "1", Item);
             Actions.SendTABToWindow(invWin);

             if (Actions.CheckWindowExists(qbWindow, "Warning"))
             {
                 Window warWin = Actions.GetChildWindow(qbWindow, "Warning");
                 Actions.ClickElementByName(warWin, "OK");
             }
             Actions.SendTABToWindow(invWin);
             Actions.ClickElementByName(invWin, "Save && Close");

             if (Actions.CheckWindowExists(qbWindow, "Job Costing - Invoice without Estimate"))
             {
                 Window noestWin = Actions.GetChildWindow(qbWindow, "Job Costing - Invoice without Estimate");
                 Actions.ClickElementByName(noestWin, "No");
             }

             Thread.Sleep(1000);

             Logger.logMessage("Invoice created Successfully.");
         }

         public static void Receive_Payments(TestStack.White.Application qbApp, Window qbWindow, String Customer)
         {
             Actions.SelectMenu(qbApp, qbWindow, "Customers", "Receive Payments");
             Actions.WaitForChildWindow(qbWindow, "Receive Payments", 1000);
             Window rpayWin = Actions.GetChildWindow(qbWindow, "Receive Payments");
             Actions.SetTextOnElementByAutomationID(rpayWin, "5603", Customer);
             Actions.SendTABToWindow(rpayWin);
             Actions.SetTextOnElementByAutomationID(rpayWin, "5604", "200");
             Actions.SendTABToWindow(rpayWin);
             Actions.SendTABToWindow(rpayWin);
             Actions.ClickElementByName(rpayWin, "CASH");
             // Actions.ClickElementByName(rpayWin, "Auto Apply Payment");
             Actions.ClickElementByName(rpayWin, "Save && Close");

             Logger.logMessage("Payment received Successfully.");

             //Make a Deposit
             Actions.SelectMenu(qbApp, qbWindow, "Banking", "Make Deposits");

             if (Actions.CheckWindowExists(qbWindow, "Need a Bank Account"))
             {

                 Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Need a Bank Account"), "Yes");
                 if (Actions.CheckWindowExists(qbWindow, "Add New Account"))
                 {
                     Window bankWin = Actions.GetChildWindow(qbWindow, "Add New Account");
                     Actions.SetTextByAutomationID(bankWin, "136", "HSBC");
                     Actions.ClickElementByName(bankWin, "Save && Close");
                 }

                 Logger.logMessage("Bank created Successfully.");
             }

             if (Actions.CheckWindowExists(qbWindow, "Payments to Deposit"))
             {
                 Window payWin = Actions.GetChildWindow(qbWindow, "Payments to Deposit");
                 Actions.ClickElementByName(payWin, "Select All");
                 Actions.ClickElementByName(payWin, "OK");

             }

             if (Actions.CheckWindowExists(qbWindow, "Make Deposits"))
             {
                 Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Make Deposits"), "Save && Close");
             }

             Logger.logMessage("Deposit to bank created Successfully.");

         }

        public static void Create_Vendor(TestStack.White.Application qbApp, Window qbWindow, String Vendor)
         {
             Actions.SelectMenu(qbApp, qbWindow, "Vendors", "Vendor Center");
             Actions.WaitForChildWindow(qbWindow, "Vendor Center", 9999);
             Window vendcenWin = Actions.GetChildWindow(qbWindow, "Vendor Center");
             if (vendcenWin.IsCurrentlyActive)
             {
                 Actions.ClickElementByName(vendcenWin, "New Vendor...");

                 SendKeys.SendWait("{DOWN}");
                 SendKeys.SendWait("{ENTER}");
             }
             Window venWin = Actions.GetChildWindow(qbWindow, "New Vendor");

             Actions.SetTextByAutomationID(venWin, "1001", Vendor);
             Actions.ClickElementByName(venWin, "OK");
             Actions.CloseAllChildWindows(qbWindow);
             Logger.logMessage("Vendor created Successfully.");

         }

        public static void Create_Purchase_Order(TestStack.White.Application qbApp, Window qbWindow, String Customer,String Item, String Vendor)
        {

            Actions.SelectMenu(qbApp, qbWindow, "Vendors", "Create Purchase Orders");
            if (Actions.CheckWindowExists(qbWindow, "Create Purchase Orders"))
            {
                Window poWin = Actions.GetChildWindow(qbWindow, "Create Purchase Orders");
                Actions.ClickElementByName(poWin, "Maximize");
                Actions.SetTextByAutomationID(poWin, "603", Vendor);
                Actions.SendTABToWindow(poWin);
                Actions.SendTABToWindow(poWin);
                Actions.SendTABToWindow(poWin);
                Actions.SendTABToWindow(poWin);
                Actions.SendTABToWindow(poWin);
                Actions.SendTABToWindow(poWin);
                Actions.SendTABToWindow(poWin);
                Actions.SetTextOnElementByAutomationID(poWin, "1", Item);
                Actions.SendTABToWindow(poWin);
                Actions.SendTABToWindow(poWin);
                Actions.SetTextOnElementByAutomationID(poWin, "10", "1");
                Actions.SendTABToWindow(poWin);
                Actions.SendTABToWindow(poWin);
                Actions.SetTextOnElementByAutomationID(poWin, "25", Customer);
                Actions.SendTABToWindow(poWin);
                Actions.ClickElementByName(poWin, "Save && Close");
            }

            Logger.logMessage("PO created Successfully.");

        }

        public static void Create_Bill(TestStack.White.Application qbApp, Window qbWindow, String Vendor, String Item)
        {
             Actions.SelectMenu(qbApp, qbWindow, "Vendors", "Receive Items and Enter Bill");
            if (Actions.CheckWindowExists(qbWindow, "Enter Bills"))
            {
                Window billWin = Actions.GetChildWindow(qbWindow, "Enter Bills");
                Actions.ClickElementByName(billWin, "Maximize");
                Actions.SetTextByAutomationID(billWin, "309", Vendor);
                Actions.SendTABToWindow(billWin);
                Actions.WaitForChildWindow(qbWindow, "Open POs Exist", int.Parse(Sync_Timeout));

                if (Actions.CheckWindowExists(qbWindow, "Open POs Exist"))
                {
                    Logger.logMessage("POs Window Found");
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Open POs Exist"), "Yes");
                    Thread.Sleep(1000);
                    if (Actions.CheckWindowExists(qbWindow, "Open Purchase Orders"))
                    {
                        Window openpoWin = Actions.GetChildWindow(qbWindow, "Open Purchase Orders");
                        Actions.SendTABToWindow(openpoWin);
                        Actions.SendSPACEToWindow(openpoWin);
                        Actions.ClickElementByName(openpoWin, "OK");
                    }

                    if (Actions.CheckElementExistsByName(billWin, "Save && Close"))
                    {
                        Actions.ClickElementByName(billWin, "Save && Close");
                    }

                    else
                    {
                        Actions.ClickElementByName(billWin, "Close");
                        if (Actions.CheckWindowExists(qbWindow, "Recording Transaction"))
                        {
                            Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Recording Transaction"), "Yes");
                        }

                    }

                }
                else
                {
                    Actions.SendTABToWindow(billWin);
                    Actions.SendTABToWindow(billWin);
                    Actions.SendTABToWindow(billWin);
                    Actions.SendTABToWindow(billWin);
                    Actions.SendTABToWindow(billWin);
                    Actions.SendTABToWindow(billWin);
                    Actions.SendTABToWindow(billWin);
                    Actions.SendKeysToWindow(billWin, Item);
                    Actions.SendTABToWindow(billWin);
                    Actions.SendTABToWindow(billWin);
                    Actions.SetTextByAutomationID(billWin, "2", "1");

                    if (Actions.CheckElementExistsByName(billWin, "Save && Close"))
                    {
                        Actions.ClickElementByName(billWin, "Save && Close");
                    }

                    else
                    {
                        Actions.ClickElementByName(billWin, "Close");
                        if (Actions.CheckWindowExists(qbWindow, "Recording Transaction"))
                        {
                            Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Recording Transaction"), "Yes");
                        }

                    }

                }


            }

            Logger.logMessage("Bill created Successfully.");
        
        }

        public static void Pay_Bill(TestStack.White.Application qbApp, Window qbWindow)
        {

            Actions.SelectMenu(qbApp, qbWindow, "Vendors", "Pay Bills");
            if (Actions.CheckWindowExists(qbWindow, "Pay Bills"))
            {
                Window billpayWin = Actions.GetChildWindow(qbWindow, "Pay Bills");
                Actions.ClickElementByName(billpayWin, "Select All Bills");
                Actions.WaitForElementEnabled(billpayWin, "Pay Selected Bills", int.Parse(Sync_Timeout));
                Actions.ClickElementByName(billpayWin, "Pay Selected Bills");
            }
            if (Actions.CheckWindowExists(qbWindow, "Payment Summary"))
            {
                Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Payment Summary"), "Done");

            }

            Logger.logMessage("Bills payed Successfully.");
        }

        public static void Reset_Preferences(TestStack.White.Application qbApp, Window qbWindow)
        {
            Actions.SelectMenu(qbApp, qbWindow, "Edit", "Preferences...");
            Window PerfWin1 = Actions.GetChildWindow(qbWindow, "Preferences");
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
            Actions.ClickElementByName(PerfWin1, "OK");

            Logger.logMessage("Reseted the preferences Successfully.");

        }

    }

}
