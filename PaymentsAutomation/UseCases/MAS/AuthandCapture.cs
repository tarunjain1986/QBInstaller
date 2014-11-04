using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using FrameworkLibraries.AppLibs.QBDT;
using System.Threading;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.Payments;
using TestStack.BDDfy;

namespace PaymentsAutomation.UseCases.MAS
{
    public class AuthandCapture
    {
        Logger log = new Logger("CreateAuthCaptureTest.txt");
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static String startupPath = System.IO.Path.GetFullPath("..\\..\\..\\");
        public static Property conf = Property.GetPropertyInstance();
        public string exe = conf.get("QBExePath");
        public static string Execution_Speed = conf.get("ExecutionSpeed");
        public static string Sync_Timeout = conf.get("SyncTimeOut");
        public string testName = "AuthorizePayment";
        public string moduleName = "PaymentAutomation";
        public string exception = "Null";
        public string category = "Null";

        [Given(StepTitle = "QuickBooks Desktop should be up and running")]
        public void Setup()
        {
            exe = conf.get("QBExePath");
            qbApp = QuickBooks.Initialize(exe);
            qbWindow = QuickBooks.PrepareBaseState(qbApp);
        }
        [When(StepTitle = "Authorization should be done successfully")]
        public void AuthorizePayment()
        {
            Payments.ProcessAuthorization(qbApp, qbWindow, "item1", ExcelLibrary.CCPaymentAmount, ExcelLibrary.CCNumber, ExcelLibrary.CCExpDate, ExcelLibrary.CCExpYear, ExcelLibrary.CCCustName, "", "", ExcelLibrary.CCZipCode);
            
        
        }
        [Then(StepTitle = "Capture corressponding to the autorized customer should be done successfully")]
        public void capturePayment()
        {
            Payments.ProcessCapture(qbApp, qbWindow, "cust1");
        }
    }
}
