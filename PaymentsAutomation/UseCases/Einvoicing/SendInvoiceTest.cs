using FrameworkLibraries.AppLibs.Payments;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.Utils;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestStack.BDDfy;
using TestStack.White.UIItems.WindowItems;

namespace PaymentsAutomation.UseCases.Einvoicing
{
    public class SendInvoiceTest
    {
        Logger log = new Logger("Send_Invoice_Test.txt");
        Window quickbooksWindow = null;
        public static string exe;
        public static TestStack.White.Application quickbooksApp;
        public static string customerName = "Tarun Test";
        public static string itemName = "Tarun Test";
        public static string newCustomer = "Einvoice Test";
        public static string customer;
        public static Property conf = Property.GetPropertyInstance();
        public static string Execution_Speed = conf.get("ExecutionSpeed");

        [Given(StepTitle = "Verified that QuickBooks Desktop is Up and Running")]
        public void Setup()
        {
            exe = conf.get("QBExePath");
            quickbooksApp = QuickBooks.Initialize(exe);
            quickbooksWindow = QuickBooks.PrepareBaseState(quickbooksApp);
        }

        [When(StepTitle = "Create an Invoice and Send Invoice to customer via WebMail")]
        public void SendWebMail()
        {
            EinvoicingPayments.SendInvoiceWebMail(quickbooksApp, quickbooksWindow, customerName, itemName);
            Assert.IsTrue(EinvoicingPayments.bankTransfer); //Einvoicing Test Case Validation - BankCard and Credit Card should be present
            Assert.IsTrue(EinvoicingPayments.creditCard);
            Logger.logMessage("CC and ACH are present in Send Invoice window");
        }

        [Then(StepTitle = "Ensuring ICN calls are going after the mail has been sent")]
        public void Test_ICN_Calls()
        {
            String checkICNcall = Payments.trackHttpCalls();
            Assert.IsTrue(checkICNcall.Contains("commercerouting-e2e"));
            Logger.logMessage("ICN calls are also going in the back-end" + checkICNcall);
        }
    }
}
