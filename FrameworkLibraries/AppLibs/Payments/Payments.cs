using FrameworkLibraries.ActionLibs.WhiteAPI;
using FrameworkLibraries.Utils;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using TestStack.White;
using TestStack.White.UIItems.WindowItems;

namespace FrameworkLibraries.AppLibs.Payments
{
    public class Payments
    {
        public static Property conf = Property.GetPropertyInstance();
        public static string Execution_Speed = conf.get("ExecutionSpeed");
        public static string Sync_Timeout = conf.get("SyncTimeOut");

        /**********************************************Receive Payments - Processing Credit Card Payment************************************************************************************************************************/
        public static void ProcessCCPayment(Application qbApp, Window qbWindow,String custAmount, String CCNumber, String expDate, String expyear, String custName ,String CCSecurityCode, String billingAddress, String zipCode)
        {
          try
          {
            Actions.SelectMenu(qbApp, qbWindow, "Customers", "Credit Card Processing Activities", "Process Payments");
            Actions.SetTextByAutomationID(qbWindow, "5603", custName);
            Actions.SendTABToWindow(qbWindow);
            Actions.SendKeysToWindow(qbWindow, custAmount);
            Actions.ClickElementByName(qbWindow, "Visa");
            Actions.ClickElementByName(qbWindow, "OK");
            Actions.WaitForChildWindow(qbWindow, "Quickbooks Payments: Process Credit Card", int.Parse(Sync_Timeout));
            SetValuesOnProcessCreditCardPaymentWindow(qbWindow, CCNumber, expDate, expyear, custName, "", "", zipCode);
            Actions.WaitForChildWindow(qbWindow, "Processed Payment Receipt", int.Parse(Sync_Timeout));
            Assert.IsTrue(Actions.CheckWindowExists(qbWindow, "Processed Payment Receipt"));
            Logger.logMessage("Payment is processed successfully");
          }
          catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        /*************************************************Void  Payments************************************************************************************************************************/
        public static void voidCCPayment(Application qbApp, Window qbWindow, String isVoid)
        {
            try
            {
                if (isVoid.Equals("Yes"))
                {
                Actions.ClickElementByName(qbWindow, "Void");
                Actions.WaitForChildWindow(qbWindow, "Processed Void Receipt", int.Parse(Sync_Timeout));
                Assert.IsTrue(Actions.CheckWindowExists(qbWindow, "Processed Void Receipt"));
                Logger.logMessage("Void is processed successfully");
                Actions.SendESCAPEToWindow(qbWindow);
                Actions.SendESCAPEToWindow(qbWindow);
                }
                else
                {
                Actions.SendESCAPEToWindow(qbWindow);
                Actions.SendESCAPEToWindow(qbWindow);
                }
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        /********************************************************************Refund Payments***********************************************************************************/
        public static void RefundPayment(Application qbApp, Window qbWindow, String itemName, String custAmount, String CCNumber, String expDate, String expyear, String custName, String CCSecurityCode, String billingAddress, String zipCode)
        {
            try
            {
                Actions.SelectMenu(qbApp, qbWindow, "Customers", "Create Credit Memos/Refunds");
                Actions.WaitForChildWindow(qbWindow, "Create Credit Memos/Refunds", int.Parse(Sync_Timeout));
                Actions.SetTextByAutomationID(qbWindow, "603", custName);
                for (int i = 1; i < 7; i++)
                {
                    Actions.SendTABToWindow(qbWindow);
                }
                Actions.SendKeysToWindow(qbWindow, itemName);
                Actions.ClickElementByName(qbWindow, "Save && Close");
                Thread.Sleep(2000);
                Actions.WaitForChildWindow(qbWindow, "Available Credit", int.Parse(Sync_Timeout));
                Actions.ClickElementByName(Actions.GetWindow(qbWindow, "Available Credit"), "Give a refund");
                Actions.ClickElementByName(qbWindow, "OK");
                var uiaWin = Actions.UIA_GetAppWindow(qbWindow.Title);
                var payWin = Actions.GetChildWindow(qbWindow, "Issue a Refund");
                Actions.SendTABToWindow(qbWindow);
                Actions.SendTABToWindow(qbWindow);
                Actions.SendTABToWindow(qbWindow);
                Actions.SendKeysToWindow(qbWindow, CCNumber);
                Actions.SendTABToWindow(qbWindow);
                Actions.SendKeysToWindow(qbWindow, expDate);
                Actions.SendTABToWindow(qbWindow);
                Actions.SendKeysToWindow(qbWindow, expyear);
                Actions.ClickElementByName(qbWindow, "OK");
                Thread.Sleep(5000);
                Actions.WaitForChildWindow(qbWindow, "Processed Refund Receipt", int.Parse(Sync_Timeout));
                Assert.IsTrue(Actions.CheckWindowExists(qbWindow, "Processed Refund Receipt"));
                Logger.logMessage("Refund is processed successfully");
                //Actions.ClickElementByName(qbWindow, "Void Refund");
                //Actions.WaitForChildWindow(qbWindow, "Processed Void Receipt", int.Parse(Sync_Timeout));
                //Assert.IsTrue(Actions.CheckWindowExists(qbWindow, "Processed Void Receipt"));
                //Logger.logMessage("Void of Refund is processed successfully");
                //Actions.ClickElementByName(qbWindow, "Close");
                Actions.SendESCAPEToWindow(qbWindow);
                Actions.SendESCAPEToWindow(qbWindow);
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        /**********************************************Process Sales Receipts Transaction************************************************************************************************************************/
        public static void ProcessSalesReceiptPayment(Application qbApp, Window qbWindow, String itemName, String custAmount, String CCNumber, String expDate, String expyear, String custName, String CCSecurityCode, String billingAddress, String zipCode)
        {
            try
            {
                Actions.SelectMenu(qbApp, qbWindow, "Customers", "Enter Sales Receipts");
                Actions.SetTextByAutomationID(qbWindow, "603", custName);
                for (int i = 1; i < 8; i++)
                {
                    Actions.SendTABToWindow(qbWindow);
                }
                Actions.SendKeysToWindow(qbWindow, itemName);
                Actions.ClickElementByName(Actions.GetWindow(qbWindow, "Enter Sales Receipts"), "Visa");
                Actions.WaitForChildWindow(qbWindow, "Quickbooks Payments: Process Credit Card", int.Parse(Sync_Timeout));
                Payments.SetValuesOnProcessCreditCardPaymentWindow(qbWindow, CCNumber, expDate, expyear, custName, "", "", zipCode);
                Actions.WaitForChildWindow(qbWindow, "Processed Payment Receipt", int.Parse(Sync_Timeout));
                Assert.IsTrue(Actions.CheckWindowExists(qbWindow, "Processed Payment Receipt"));
                Logger.logMessage("Payment is processed successfully");
                Actions.ClickElementByName(qbWindow, "Void");
                Actions.WaitForChildWindow(qbWindow, "Processed Void Receipt", int.Parse(Sync_Timeout));
                Assert.IsTrue(Actions.CheckWindowExists(qbWindow, "Processed Void Receipt"));
                Logger.logMessage("Void is processed successfully");
                Actions.SendESCAPEToWindow(qbWindow);
                Actions.SendESCAPEToWindow(qbWindow);
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        /**********************************************Process Authorization Transaction************************************************************************************************************************/
        public static void ProcessAuthorization(Application qbApp, Window qbWindow, String itemName, String custAmount, String CCNumber, String expDate, String expyear, String custName, String CCSecurityCode, String billingAddress, String zipCode)
        {
            try
            {
                Actions.SelectMenu(qbApp, qbWindow, "Customers", "Credit Card Processing Activities", "Authorize/Capture Payments");
                Actions.SetTextByAutomationID(qbWindow, "5603", custName);
                Actions.SendTABToWindow(qbWindow);
                Actions.SendKeysToWindow(qbWindow, custAmount);
                Actions.ClickElementByName(qbWindow, "Visa");
                Actions.WaitForChildWindow(qbWindow, "QuickBooks Merchant Service Message", int.Parse(Sync_Timeout));
                Actions.ClickElementByName(qbWindow, "OK");
                Actions.WaitForChildWindow(qbWindow, "Quickbooks Payments: Authorize Credit Card Funds", int.Parse(Sync_Timeout));
                Payments.ProcessAuthorizationPayment(qbWindow, CCNumber, expDate, expyear, custName, "", "", zipCode);
                Actions.WaitForChildWindow(qbWindow, "Authorization Receipt", int.Parse(Sync_Timeout));
                Assert.IsTrue(Actions.CheckWindowExists(qbWindow, "Authorization Receipt"));
                Logger.logMessage("Authorization is processed successfully");
                Actions.ClickElementByName(qbWindow, "Close");
                Thread.Sleep(2000);
                Actions.SendESCAPEToWindow(qbWindow);
                //Actions.ClickElementByName(qbWindow, "Void");
                //Actions.WaitForChildWindow(qbWindow, "QuickBooks Merchant Service Message", int.Parse(Sync_Timeout));
                //Actions.ClickElementByName(qbWindow, "Yes");
                //Actions.WaitForChildWindow(qbWindow, "Processed Void Receipt", int.Parse(Sync_Timeout));
                //Actions.ClickElementByName(qbWindow, "Close");
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        /**********************************************Process Capture Transaction************************************************************************************************************************/
        public static void ProcessCapture(Application qbApp, Window qbWindow, String custName)
        {
            try
            {
                Actions.SelectMenu(qbApp, qbWindow, "Customers", "Credit Card Processing Activities", "Process Payments");
                Actions.SetTextByAutomationID(Actions.GetWindow(qbWindow, "Receive Payments"), "5603", custName);
                Actions.SendTABToWindow(qbWindow);
                Actions.WaitForChildWindow(qbWindow, "Available Authorizations", int.Parse(Sync_Timeout));
                Actions.ClickElementByName(qbWindow, "OK");
                Actions.ClickElementByName(qbWindow, "Enable Payment");
                Actions.ClickElementByName(qbWindow, "Save && Close");
                Actions.WaitForChildWindow(qbWindow, "Intuit Payment Solutions: Process Credit Card", int.Parse(Sync_Timeout));
                Actions.SendENTERoWindow(qbWindow);
                Actions.WaitForChildWindow(qbWindow, "Capture Payment Receipt", int.Parse(Sync_Timeout));
                Assert.IsTrue(Actions.CheckWindowExists(qbWindow, "Capture Payment Receipt"));
                Logger.logMessage("Capture is processed successfully");
                Actions.ClickElementByName(qbWindow, "Close");
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        /**********************************************Process credit card payment************************************************************************************************************************/
        public static void SetValuesOnProcessCreditCardPaymentWindow(Window paymentWin, string ccNumber, string expMonth, string expYear, string nameOnCard, string secCode, string billingAddr, string zipCode)
        {
            try
            {
                Logger.logMessage("---------------------------------------------------------------------------------");

                var paymentPanel = Actions.GetPaneByName(paymentWin, "Quickbooks Payments: Process Credit Card");

                PropertyCondition editCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
                AutomationElementCollection editElements = paymentPanel.AutomationElement.FindAll(TreeScope.Children, editCondition);
                int count = 0;

                foreach (AutomationElement item in editElements)
                {
                    count = count + 1;
                    TestStack.White.UIItems.TextBox t = new TestStack.White.UIItems.TextBox(item, paymentWin.ActionListener);

                    if (count == 1)
                        t.Text = ccNumber;

                    if (count == 2)
                        t.Text = expMonth;

                    if (count == 3)
                        t.Text = expYear;

                    if (count == 4)
                        t.Text = nameOnCard;

                    if (count == 5)
                        t.Text = secCode;

                    if (count == 6)
                        t.Text = billingAddr;

                    if (count == 7)
                        t.Text = zipCode;
                }
                var uiaWin = Actions.UIA_GetAppWindow(paymentWin.Title);
                var payWin = Actions.GetChildWindow(paymentWin, "Quickbooks Payments: Process Credit Card");
                Actions.UIA_ClickTextByName(uiaWin, payWin, "Process Payment");

                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("---------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        /**********************************************Process Authorization Payment************************************************************************************************************************/
        public static void ProcessAuthorizationPayment(Window paymentWin, string ccNumber, string expMonth, string expYear, string nameOnCard, string secCode, string billingAddr, string zipCode)
        {
            try
            {
                Logger.logMessage("---------------------------------------------------------------------------------");

                var paymentPanel = Actions.GetPaneByName(paymentWin, "Quickbooks Payments: Authorize Credit Card Funds");

                PropertyCondition editCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
                AutomationElementCollection editElements = paymentPanel.AutomationElement.FindAll(TreeScope.Children, editCondition);
                int count = 0;

                foreach (AutomationElement item in editElements)
                {
                    count = count + 1;
                    TestStack.White.UIItems.TextBox t = new TestStack.White.UIItems.TextBox(item, paymentWin.ActionListener);

                    if (count == 1)
                        t.Text = ccNumber;

                    if (count == 2)
                        t.Text = expMonth;

                    if (count == 3)
                        t.Text = expYear;

                    if (count == 4)
                        t.Text = nameOnCard;

                    if (count == 5)
                        t.Text = secCode;

                    if (count == 6)
                        t.Text = billingAddr;

                    if (count == 7)
                        t.Text = zipCode;
                }
                var uiaWin = Actions.UIA_GetAppWindow(paymentWin.Title);
                var payWin = Actions.GetChildWindow(paymentWin, "Quickbooks Payments: Authorize Credit Card Funds");
                Actions.UIA_ClickTextByName(uiaWin, payWin, "Authorize Funds");

                Thread.Sleep(int.Parse(Execution_Speed));
                Logger.logMessage("---------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }

        }

        public static TestStack.White.UIItems.ListViewCells GetFiddlerStackTrace(Window fiddlerWindow)
        {
            try
            {
                Logger.logMessage("---------------------------------------------------------------------------------");
                Logger.logMessage("GetFiddlerStackTrace");

                var allFiddlerWindowElements = fiddlerWindow.Items;
                var sessionPanel = Actions.GetPaneByAutomationID(fiddlerWindow, "pnlSessions");
                var allPanelElements = sessionPanel.Items;
                var list = Actions.GetAllListItems(sessionPanel.Items);
                TestStack.White.UIItems.ListViewRow x = new TestStack.White.UIItems.ListViewRow(list[0].AutomationElement, fiddlerWindow.ActionListener);
                return x.Cells;
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //Utility which tracks the HTTP calls in the background

        public static string trackHttpCalls()
        {

            try
            {

                System.Diagnostics.ProcessStartInfo procStartInfo =

                new System.Diagnostics.ProcessStartInfo("cmd", "/c " + "netstat /f");

                procStartInfo.RedirectStandardOutput = true;

                procStartInfo.UseShellExecute = false;

                procStartInfo.CreateNoWindow = true;

                System.Diagnostics.Process proc = new System.Diagnostics.Process();

                proc.StartInfo = procStartInfo;

                proc.Start();

                return proc.StandardOutput.ReadToEnd();

            }

            catch (Exception e)
            {

                Logger.logMessage(e.Message);

                return null;

            }

        }

        public static void Test_ICN_Calls()
        {

            String checkICNcall = Payments.trackHttpCalls();

            Assert.IsTrue(checkICNcall.Contains("commercerouting-e2e"));

            Logger.logMessage("ICN calls are also going in the back-end" + checkICNcall);

        }

    }
}
