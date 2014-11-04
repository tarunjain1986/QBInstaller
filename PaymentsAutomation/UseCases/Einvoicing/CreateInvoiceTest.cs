using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using FrameworkLibraries.Utils;
using TestStack.White.UIItems.WindowItems;
using TestStack.BDDfy;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.AppLibs.Payments;
using System.Threading;

namespace PaymentsAutomation.UseCases.Einvoicing
{
    public class CreateInvoiceTest
    {
        Logger log = new Logger("Create_Invoice_Test.txt");
        Window quickbooksWindow = null;
        public static string exe;
        public static TestStack.White.Application quickbooksApp;
        public static string customerName = "Tarun Test";
        public static string itemName = "Tarun Test";
        public static string newCustomer = "Einvoice Test";
        public static string customer;
        public static string emailId = "test.qbdt@gmail.com";
        public static Property conf = Property.GetPropertyInstance();
        public static string Execution_Speed = conf.get("ExecutionSpeed");

        [Given(StepTitle = "Verified that QuickBooks Desktop is Up and Running")]
        public void Setup()
        {
            exe = conf.get("QBExePath");
            quickbooksApp = QuickBooks.Initialize(exe);
            quickbooksWindow = QuickBooks.PrepareBaseState(quickbooksApp);
        }

        [When(StepTitle = "Verified that when Invoice gets created with Email Later checked, ICN services should get called")]
        public void createInvoiceWithEmailLaterChecked()
        {
            customer = newCustomer + new Random().Next(int.MinValue, int.MaxValue);
            EinvoicingPayments.CreateCustomer(quickbooksApp, quickbooksWindow, customer, "", "", "", "", "", emailId);
            EinvoicingPayments.CreateInvoiceWithEmailLaterChecked(quickbooksApp, quickbooksWindow, customer, itemName);
            String checkICNcall = Payments.trackHttpCalls(); //Checking if ICN calls are going or not
            Assert.IsTrue(checkICNcall.Contains("commercerouting-e2e"));
            Logger.logMessage("ICN calls are also going in the back-end" + checkICNcall);
        }

        [Then(StepTitle = "Email should appear in the Send Forms and ICN sevices should get called when email has been sent")]
        public void Test_ICN_Call_SendMail_SendForms()
        {
            EinvoicingPayments.OpenSendFormAndSendMail(quickbooksApp, quickbooksWindow);
            String checkICNcall = Payments.trackHttpCalls();
            Assert.IsTrue(checkICNcall.Contains("commercerouting-e2e"));
            Logger.logMessage("ICN calls are also going in the back-end" + checkICNcall);
        }
    }
}
